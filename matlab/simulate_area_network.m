%% simulate_area_network.m  (non‑GUI version)
% ------------------------------------------------------------
% 10 エリア（北海道〜沖縄）の連成スイングを Excel テンプレートから読み込み
% 可変発電機台数に対応、地理マップ上に COI & 各機角を表示
% ------------------------------------------------------------
clear; close all; clc;

excelFile = 'area_parameters_template.xlsx';
if ~isfile(excelFile)
    generate_template_matlab(excelFile);   % 内部生成
end

%% 1) マスター読込
master = readtable(excelFile,'Sheet','Master','PreserveVariableNames',true);
areas  = master.Area;

%% 2) エリア選択ダイアログ
[selIdx, ok] = listdlg('PromptString','可視化対象エリアを選択',...
                       'SelectionMode','multiple',...
                       'ListString', string(areas));
if ok==0, return; end
master = master(selIdx,:);    % フィルタ
areas  = master.Area;
Ns     = height(master);

%% 3) 擾乱設定
distAreaIdx = listdlg('PromptString','擾乱エリアを選択',...
                      'SelectionMode','single',...
                      'ListString',string(areas),...
                      'InitialValue',1);
if isempty(distAreaIdx), return; end
answ = inputdlg({'発電機番号 (1～N_i)','擾乱量 Δδ [rad]'},...
                '擾乱設定',1,{'1','-1.39'});
distGen = str2double(answ{1});
distAmp = str2double(answ{2});

%% 4) パラメータセット
N_each   = master.Generator_Count;
cumN     = [0; cumsum(N_each)];
G        = cumN(end);

p_m_arr   = master.p_m;
b_arr     = master.b;
b_int_arr = master.b_int;
eps_arr   = master.epsilon;

% 緯度経度テーブル (北海道〜沖縄)
allLonLat = [
 141.35 43.06; 140.89 39.70; 139.75 35.68; 137.02 37.15; ...
 136.90 35.18; 135.50 34.70; 133.94 34.39; 134.05 33.56; ...
 130.41 33.59; 127.68 26.21];
baseLonLat = allLonLat(selIdx,:);

%% 5) 接続行列 (静的例：隣接のみ 0.1)
C_DEFAULT = 0.1;
Cmat = zeros(Ns);
adj = { [2],        ... % 北海道
        [1 3],      ... % 東北
        [2 5],      ... % 東京
        [3 6],      ... % 中部
        [4 6],      ... % 北陸
        [4 5 7 8],  ... % 関西
        [6 9],      ... % 中国
        [6],        ... % 四国
        [7],        ... % 九州
        [] };           % 沖縄
for i = 1:Ns
    for j = adj{selIdx(i)}
        jIdx = find(selIdx==j);
        if ~isempty(jIdx)
            Cmat(i,jIdx)=C_DEFAULT; Cmat(jIdx,i)=C_DEFAULT;
        end
    end
end

%% 6) 初期条件
rng('default');                    % 乱数再現性
eps_spread = 0.01;                 % ±0.01 rad の微擾乱
delta0=zeros(G,1); omega0=zeros(G,1);
for i = 1:Ns
    idx = cumN(i)+1 : cumN(i+1);
    delta0(idx) = asin(p_m_arr(i)/b_arr(i)) + eps_spread*randn(size(idx));
end
distGlobalIdx = cumN(distAreaIdx) + min(distGen,N_each(distAreaIdx));
delta0(distGlobalIdx)=distAmp;
init=[delta0;omega0];

%% 7) ODE 解く
t_span=[0 25];
[t,y]=ode45(@(t,y) dyn(t,y,N_each,Ns,cumN,Cmat,...
             p_m_arr,b_arr,b_int_arr,eps_arr), t_span, init);

%% 8) 可視化
visualize_network(t,y,Ns,N_each,cumN,baseLonLat);

%% ---------- 関数 ----------
function dy = dyn(~,y,N_each,Ns,cumN,Cmat,p_m,b,b_int,epsl)
    G=cumN(end); dy=zeros(2*G,1);
    Nh = floor(N_each/2)+1;
    for i=1:Ns
        Ni=N_each(i); base=cumN(i)+1;
        for j=1:Ni
            idx=base+j-1;
            prev=(j==1)*(base+Ni-1)+(j>1)*(idx-1);
            next=(j==Ni)*base      +(j<Ni)*(idx+1);
            g=0;
            nb=find(Cmat(i,:));
            for n=nb
                idxNb = cumN(n)+Nh(n);
                g = g + Cmat(i,n)*sin(y(idx)-y(idxNb));
            end
            delta=y(idx); omega=y(G+idx);
            dy(idx)=omega;
            dy(G+idx)=p_m(i) ...
                     - b(i)*sin(delta) ...
                     - b_int(i)*(sin(delta-y(prev))+sin(delta-y(next))) ...
                     - epsl(i)*b_int(i)*g;
        end
    end
end

function visualize_network(t,y,Ns,N_each,cumN,baseLonLat)
    G=cumN(end);
    SCALE=4; radBase=0.25; radVec=radBase+0.01*N_each;
    load coastlines
    figure; hold on; plot(coastlon,coastlat,'k');
    axis equal; xlim([128 146]); ylim([30 46]);
    title('Area COI vectors & generator angles');
    annotation('textbox',[0.80 0.06 0.18 0.04],'String','Δf [Hz] / ω [rad/s]',...
               'EdgeColor','none','HorizontalAlignment','right','FontSize',8);
    sizeVec=200+8*N_each;
    hArea=scatter(baseLonLat(:,1),baseLonLat(:,2),sizeVec,'k','LineWidth',2);
    th=linspace(0,2*pi,200); hCirc=gobjects(Ns,1); hGen=gobjects(Ns,1);
    for i=1:Ns
        hCirc(i)=plot(baseLonLat(i,1)+radVec(i)*cos(th),...
                      baseLonLat(i,2)+radVec(i)*sin(th),'k:');
        hGen(i)=scatter(nan(1,N_each(i)),nan(1,N_each(i)),36,'filled',...
                        'MarkerFaceColor',[0.2 0.6 1],'MarkerEdgeColor','k');
    end
    timeText=text(0.02,0.95,'','Units','normalized','FontSize',9,'FontWeight','bold');
    for k=1:length(t)
        dMean=zeros(Ns,1); wMean=zeros(Ns,1);
        for i=1:Ns
            idx=cumN(i)+1:cumN(i+1);
            dMean(i)=mean(mod(y(k,idx),2*pi));
            wMean(i)=mean(y(k,G+idx));
        end
        dx=SCALE*wMean.*cos(dMean); dy=SCALE*wMean.*sin(dMean);
        newC=baseLonLat+[dx dy];
        set(hArea,'XData',newC(:,1),'YData',newC(:,2));
        for i=1:Ns
            set(hCirc(i),'XData',newC(i,1)+radVec(i)*cos(th),...
                         'YData',newC(i,2)+radVec(i)*sin(th));
            idx = cumN(i)+1 : cumN(i+1);                % 発電機 idx
            delAbs = mod(y(k, idx), 2*pi);              % 絶対角
            set(hGen(i), 'XData', newC(i,1) + radVec(i)*cos(delAbs), ...
                          'YData', newC(i,2) + radVec(i)*sin(delAbs));
        end
        set(timeText,'String',sprintf('t = %.2f s',t(k)));
        drawnow;
    end
end
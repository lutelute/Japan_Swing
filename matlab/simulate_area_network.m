
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

%% 3) パラメータセット
N_each   = master.Generator_Count;

% --- Excel から取得した発電機台数を確認 -----------------------
disp('▼ Excel から取得した発電機台数');
disp(table(areas, N_each, ...
     'VariableNames', {'Area','Generator_Count'}));

choiceGen = questdlg( ...
    '上記台数設定でシミュレーションを続行しますか?', ...
    'Generator Count Check', ...
    'はい','キャンセル','はい');
if strcmp(choiceGen,'キャンセル')
    disp('キャンセルされたのでスクリプトを終了します。');
    return;
end
cumN     = [0; cumsum(N_each)];
G        = cumN(end);

%% 4) 擾乱設定（複数台対応）
% 擾乱を追加するかどうか確認
addDisturbance = questdlg('擾乱を設定しますか？', '擾乱設定', 'はい', 'いいえ', 'はい');
disturbances = [];

if strcmp(addDisturbance, 'はい')
    while true
        % エリア選択
        distAreaIdx = listdlg('PromptString','擾乱エリアを選択',...
                              'SelectionMode','single',...
                              'ListString',string(areas),...
                              'InitialValue',1);
        if isempty(distAreaIdx), break; end
        
        % 発電機番号と擾乱量の入力
        prompt = {sprintf('発電機番号 (1～%d)', N_each(distAreaIdx)), '擾乱量 Δδ [rad]'};
        answ = inputdlg(prompt, '擾乱設定', 1, {'1', '-1.39'});
        if isempty(answ), break; end
        
        distGen = str2double(answ{1});
        distAmp = str2double(answ{2});
        
        % 入力値チェック
        if distGen < 1 || distGen > N_each(distAreaIdx)
            warndlg(sprintf('発電機番号は1～%dの範囲で入力してください', N_each(distAreaIdx)));
            continue;
        end
        
        % 擾乱情報を保存
        disturbances = [disturbances; distAreaIdx, distGen, distAmp];
        
        % 追加するかどうか確認
        addMore = questdlg('さらに擾乱を追加しますか？', '擾乱追加', 'はい', 'いいえ', 'いいえ');
        if strcmp(addMore, 'いいえ'), break; end
    end
end

% 擾乱設定の表示
if ~isempty(disturbances)
    fprintf('設定された擾乱:\n');
    for i = 1:size(disturbances, 1)
        fprintf('  エリア%d, 発電機%d: %.3f rad\n', ...
                disturbances(i,1), disturbances(i,2), disturbances(i,3));
    end
else
    fprintf('擾乱なしでシミュレーションを実行します\n');
end

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
% 複数擾乱の適用
for i = 1:size(disturbances, 1)
    distAreaIdx = disturbances(i, 1);
    distGen = disturbances(i, 2);
    distAmp = disturbances(i, 3);
    distGlobalIdx = cumN(distAreaIdx) + distGen;
    delta0(distGlobalIdx) = distAmp;
end
init=[delta0;omega0];

%% 7) ODE 解く
t_span=[0 25];
[t,y]=ode45(@(t,y) dyn(t,y,N_each,Ns,cumN,Cmat,...
             p_m_arr,b_arr,b_int_arr,eps_arr), t_span, init);

%% 8) 可視化
visualize_network(t,y,Ns,N_each,cumN,baseLonLat);

%% 9) COI時系列プロット
plot_coi_timeseries(t,y,Ns,N_each,cumN);

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
            % 参考コードの方式：特定位置の発電機同士を結合
            % 各エリアの最初の発電機が前のエリアの中央発電機と結合
            if i > 1 && j == 1
                prevAreaIdx = cumN(i-1) + floor(N_each(i-1)/2) + 1;
                g = g + sin(y(idx) - y(prevAreaIdx));
            end
            % 各エリアの中央発電機が次のエリアの最初の発電機と結合
            if i < Ns && j == floor(Ni/2) + 1
                nextAreaIdx = cumN(i+1) + 1;  % 次エリアの最初の発電機
                g = g + sin(y(idx) - y(nextAreaIdx));
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
    
    % Create figure with subplots (上下配置)
    figure('Position',[100 100 900 1200]);
    
    % Top subplot: Map view (正方形)
    subplot(2,1,1); hold on; plot(coastlon,coastlat,'k');
    axis equal; xlim([128 146]); ylim([30 46]);
    title('Area COI vectors & generator angles');
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
    
    % Bottom subplot: 1D line plot (横長)
    subplot(2,1,2); hold on;
    colors = lines(Ns);
    hLines = gobjects(Ns,1);
    for i=1:Ns
        hLines(i) = plot(1:N_each(i), nan(1,N_each(i)), 'Color', colors(i,:), ...
                        'LineWidth', 2, 'Marker', 'o', 'MarkerSize', 4);
    end
    xlim([0.5 max(N_each)+0.5]);
    ylim([0 2*pi]);
    xlabel('Generator Index');
    ylabel('Generator Angle (rad)');
    title('Generator Angles (1D view)');
    yticks([0 pi/2 pi 3*pi/2 2*pi]);
    yticklabels({'0', '\pi/2', '\pi', '3\pi/2', '2\pi'});
    legend(arrayfun(@(x) sprintf('Area %d', x), 1:Ns, 'UniformOutput', false), ...
           'Location', 'eastoutside');
    timeText2=text(0.02,0.95,'','Units','normalized','FontSize',9,'FontWeight','bold');
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
            % Update map view
            set(hCirc(i),'XData',newC(i,1)+radVec(i)*cos(th),...
                         'YData',newC(i,2)+radVec(i)*sin(th));
            idx = cumN(i)+1 : cumN(i+1);                % 発電機 idx
            delAbs = mod(y(k, idx), 2*pi);              % 絶対角
            set(hGen(i), 'XData', newC(i,1) + radVec(i)*cos(delAbs), ...
                          'YData', newC(i,2) + radVec(i)*sin(delAbs));
            
            % Update 1D line plot
            set(hLines(i), 'YData', delAbs);
        end
        set(timeText,'String',sprintf('t = %.2f s',t(k)));
        set(timeText2,'String',sprintf('t = %.2f s',t(k)));
        drawnow;
    end
end
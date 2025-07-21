function plot_coi_timeseries(t,y,Ns,N_each,cumN)
    G = cumN(end);
    
    % Calculate COI for each area at each time step
    coi_delta = zeros(length(t), Ns);
    coi_omega = zeros(length(t), Ns);
    
    for k = 1:length(t)
        for i = 1:Ns
            idx = cumN(i)+1:cumN(i+1);
            coi_delta(k,i) = mean(mod(y(k,idx), 2*pi));
            coi_omega(k,i) = mean(y(k,G+idx));
        end
    end
    
    % Create separate figure for COI time series
    figure('Position',[1300 100 1000 600]);
    
    % Subplot 1: COI angles vs time
    subplot(2,1,1);
    colors = lines(Ns);
    hold on;
    for i = 1:Ns
        plot(t, coi_delta(:,i), 'Color', colors(i,:), 'LineWidth', 2, ...
             'DisplayName', sprintf('Area %d', i));
    end
    xlabel('Time (s)');
    ylabel('COI Angle (rad)');
    title('Center of Inertia (COI) Angles Over Time');
    yticks([0 pi/2 pi 3*pi/2 2*pi]);
    yticklabels({'0', '\pi/2', '\pi', '3\pi/2', '2\pi'});
    legend('show', 'Location', 'best');
    grid on;
    
    % Subplot 2: COI angular velocities vs time
    subplot(2,1,2);
    hold on;
    for i = 1:Ns
        plot(t, coi_omega(:,i), 'Color', colors(i,:), 'LineWidth', 2, ...
             'DisplayName', sprintf('Area %d', i));
    end
    xlabel('Time (s)');
    ylabel('COI Angular Velocity (rad/s)');
    title('Center of Inertia (COI) Angular Velocities Over Time');
    legend('show', 'Location', 'best');
    grid on;
end
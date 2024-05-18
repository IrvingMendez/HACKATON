% bar_plot_and_sum.m
function suma = GraficaSuma(data)
    % Function to plot the input data as a bar chart and return the sum
    suma = sum(data);
    bar(data);
    title('CONSUMO POR APARTO');
    xlabel('Aparatos');
    ylabel('Voltios');
    
    % Calculate the sum of the elements
end
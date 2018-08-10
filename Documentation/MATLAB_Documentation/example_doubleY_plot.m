%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    % Create Double Y JV and Luminance graph
    templatePath = 'C:\Users\JSB\Documents\OriginLab\2016\User Files\JV-Lum-LogLog-Line.otp';
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Graph Name
    graph_JV_Lum = invoke(originObj, 'CreatePage', 3, 'JV-Luminance'  , templatePath);
    
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    % First the JV layer (Layer 1)
    invoke(originObj, 'Execute', 'layer -s 1;');
    % Find the layer
    graphLayer_JV_Lum = invoke(originObj, 'FindGraphLayer', graph_JV_Lum);
    % Make Layer 1 active
     
    % Get dataplot collection from the graph layer
    dataplots_JV = invoke(graphLayer_JV_Lum, 'DataPlots');
    
    % Add data column by column to the graph
    for i = 1 : (size(luminanceArray,2) - 1)/2
    
        % Create a data range
        dr = invoke(originObj, 'NewDataRange');

        % Add data to data range
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% row   col   row   col
        invoke(dr, 'Add', 'X', wkSheet_lum, 0 ,   0   , -1,   0 ); % x is voltage
        invoke(dr, 'Add', 'Y', wkSheet_lum, 0 , 2*i-1  , -1, 2*i-1 );

        % Add data plot to graph layer
        % 200 -- line
        % 201 -- symbol
        % 202 -- symbol+line
        invoke(dataplots_JV, 'Add', dr, 200);    
    end
    
    % Group data (This allows colors to be automatically incremented and a
        % single legend entry to be created for all the data sets with the 
        % same legend entry)
    invoke(graphLayer_JV_Lum, 'Execute', 'layer -g;');
    
    % For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Layer-cmd
    % Reconstruct the legend (can be done manually in Origin via Ctrl+L)
    invoke(graphLayer_JV_Lum, 'Execute', 'legend -r;');
    % For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Legend-cmd
    
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    % Now luminance layer
    % Make Layer 2 active
    invoke(originObj, 'Execute', 'layer -s 2;'); 
    % Now find the layer
    graphLayer_JV_Lum = invoke(originObj, 'FindGraphLayer', graph_JV_Lum);
    % Get dataplot collection from the graph layer
    dataplots_lum = invoke(graphLayer_JV_Lum, 'DataPlots');
    
    % Add data column by column to the graph
    for i = 1 : (size(luminanceArray,2) - 1)/2
    
        % Create a data range
        dr = invoke(originObj, 'NewDataRange');

        % Add data to data range
        %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% row   col   row   col
        invoke(dr, 'Add', 'X', wkSheet_lum, 0 ,   0   , -1,   0 ); % x is voltage
        invoke(dr, 'Add', 'Y', wkSheet_lum, 0 , 2*i  , -1, 2*i );

        % Add data plot to graph layer
        % 200 -- line
        % 201 -- symbol
        % 202 -- symbol+line
        invoke(dataplots_lum, 'Add', dr, 200);    
    end
    
    % Group data (This allows colors to be automatically incremented and a
        % single legend entry to be created for all the data sets with the 
        % same legend entry)
    invoke(graphLayer_JV_Lum, 'Execute', 'layer -g;');
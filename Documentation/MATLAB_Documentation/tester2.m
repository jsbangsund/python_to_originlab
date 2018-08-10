function createGraph(graphName,templatePath,x_cols,y_cols,LineOrSym,wksObj)
% Create graph page and object
graphObj = invoke(originObj, 'CreatePage', 3, graphName , templatePath);

% Find the graph layer
graphLayer = invoke(originObj, 'FindGraphLayer', graphObj);

% Get dataplot collection from the graph layer
dataPlots = invoke(graphLayer, 'DataPlots');

% Add data column by column to the graph
for ci = 1 : length(x_cols)
    % Create a data range
    dr = invoke(originObj, 'NewDataRange');
    
    % Add data to data range
    %                  worksheet, start row, start col, end row (-1=last), end col
    dr.invoke('Add', 'X', wksObj, 0 , x_cols(ci), -1, x_cols(ci));
    dr.invoke('Add', 'Y', wksObj, 0 , y_cols(ci), -1, y_cols(ci));
    
    % Add data plot to graph layer
    % 200 -- line
    % 201 -- symbol
    % 202 -- symbol+line
    % If specified, plot symbol. By default, plot line
    if any(strcmp(LineOrSym,{'Sym','Symbol','Symbols'}))
        dataPlots.invoke('Add', dr, 201);
    else
        dataPlots.invoke('Add', dr, 200);
    end
end

% Group data (This allows colors to be automatically incremented and a
% single legend entry to be created for all the data sets with the
% same legend entry)
invoke(graphLayer, 'Execute', 'layer -g;');

% For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Layer-cmd
% Reconstruct the legend (can be done manually in Origin via Ctrl+L)
%invoke(graphLayer, 'Execute', 'legend -r;');
% For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Legend-cmd

end

function createGraph_multiwks(graphName,templatePath,x_cols,y_cols,LineOrSym,wksObj)
% wksObj must be a cell array of worksheet Objects
% Each worksheet must be formatted identically
% Create graph page and object
graphObj = invoke(originObj, 'CreatePage', 3, graphName , templatePath);

% Find the graph layer
graphLayer = invoke(originObj, 'FindGraphLayer', graphObj);

% Get dataplot collection from the graph layer
dataPlots = invoke(graphLayer, 'DataPlots');

% Add data column by column to the graph
for wi = 1 : length(wksObj)
    for ci = 1 : length(x_cols)
        % Create a data range
        dr = invoke(originObj, 'NewDataRange');
        
        % Add data to data range
        %                  worksheet, start row, start col, end row (-1=last), end col
        dr.invoke('Add', 'X', wksObj{wi}, 0 , x_cols(ci), -1, x_cols(ci));
        dr.invoke('Add', 'Y', wksObj{wi}, 0 , y_cols(ci), -1, y_cols(ci));
        
        % Add data plot to graph layer
        % 200 -- line
        % 201 -- symbol
        % 202 -- symbol+line
        % If specified, plot symbol. By default, plot line
        if any(strcmp(LineOrSym,{'Sym','Symbol','Symbols'}))
            dataPlots.invoke('Add', dr, 201);
        else
            dataPlots.invoke('Add', dr, 200);
        end
    end
end

% Group data (This allows colors to be automatically incremented and a
% single legend entry to be created for all the data sets with the
% same legend entry)
invoke(graphLayer, 'Execute', 'layer -g;');

% For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Layer-cmd
% Reconstruct the legend (can be done manually in Origin via Ctrl+L)
%invoke(graphLayer, 'Execute', 'legend -r;');
% For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Legend-cmd

end

function []=hello()

graphNames = {'ELPL-Lin','ELPL-Log','V-Lin','dV-Lin','PL-CB-vs-ELRatio'}; % 'tx-vs-L0'
templateNames = {'LT-ELPL-Lin.otp','LT-ELPL-Log.otp','LT-V-Lin.otp','LT-dV-Lin-Sym.otp','PL_CB_vs_ELRatio.otp'};
templateDir = 'C:\Users\JSB\Documents\OriginLab\2016\User Files\';
templateFullPath = [templateDir templateNames{gi}];
% Alternative templates:
% 'LT-dV-Lin-Line.otp'

for gi = 1 : length(graphNames);

end

%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% Create Driving Voltage vs. hours graph
templatePath = 'C:\Users\JSB\Documents\OriginLab\2016\User Files\Lifetime-V-Line.otp';
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Graph Name
graph_V = invoke(originObj, 'CreatePage', 3, 'Driving Voltage'  , templatePath);

% Find the graph layer
graphLayer_V = invoke(originObj, 'FindGraphLayer', graph_V);

% Get dataplot collection from the graph layer
dataplots_Lum = invoke(graphLayer_V, 'DataPlots');

% Add data column by column to the graph
for i = 1 : size(lifetimeArray,2)/3
    hrs_col_idx = 3*i - 3;
    V_col_idx = 3*i - 2;
    
    % Create a data range
    dr = invoke(originObj, 'NewDataRange');
    
    % Add data to data range
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% row   col   row   col
    invoke(dr, 'Add', 'X', wkSheet_lifetime, 0 , hrs_col_idx, -1, hrs_col_idx); % x is hrs
    invoke(dr, 'Add', 'Y', wkSheet_lifetime, 0 , V_col_idx, -1, V_col_idx); % y is V
    
    % Add data plot to graph layer
    % 200 -- line
    % 201 -- symbol
    % 202 -- symbol+line
    invoke(dataplots_Lum, 'Add', dr, 200);
end

% Group data (This allows colors to be automatically incremented and a
% single legend entry to be created for all the data sets with the
% same legend entry)
invoke(graphLayer_V, 'Execute', 'layer -g;');

% For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Layer-cmd
% Reconstruct the legend (can be done manually in Origin via Ctrl+L)
invoke(graphLayer_V, 'Execute', 'legend -r;');
% For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Legend-cmd
%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% Create Double Y L/L0 and V vs hrs graph
templatePath = 'C:\Users\JSB\Documents\OriginLab\2016\User Files\LifetimeYY-Lum-V-Line.otp';
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Graph Name
graph_V_Lum = invoke(originObj, 'CreatePage', 3, 'Luminance-Voltage'  , templatePath);


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% First the L/L0 layer (Layer 1)
invoke(originObj, 'Execute', 'layer -s 1;');
% Find the layer
graphLayer_Lum = invoke(originObj, 'FindGraphLayer', graph_V_Lum);
% Make Layer 1 active

% Get dataplot collection from the graph layer
dataplots_Lum = invoke(graphLayer_Lum, 'DataPlots');

% Add data column by column to the graph
for i = 1 : size(lifetimeArray,2)/3
    hrs_col_idx = 3*i - 3;
    lum_col_idx = 3*i - 1;
    % Create a data range
    dr = invoke(originObj, 'NewDataRange');
    
    % Add data to data range
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% row   col   row   col
    invoke(dr, 'Add', 'X', wkSheet_lifetime, 0 , hrs_col_idx, -1, hrs_col_idx); % x is hrs
    invoke(dr, 'Add', 'Y', wkSheet_lifetime, 0 , lum_col_idx, -1, lum_col_idx); % y is L/L0
    
    % Add data plot to graph layer
    % 200 -- line
    % 201 -- symbol
    % 202 -- symbol+line
    invoke(dataplots_Lum, 'Add', dr, 200);
end

% Group data (This allows colors to be automatically incremented and a
% single legend entry to be created for all the data sets with the
% same legend entry)
invoke(graphLayer_Lum, 'Execute', 'layer -g;');

% For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Layer-cmd
% Reconstruct the legend (can be done manually in Origin via Ctrl+L)
invoke(graphLayer_Lum, 'Execute', 'legend -r;');
% For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Legend-cmd

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Now driving voltage layer
% Make Layer 2 active
invoke(originObj, 'Execute', 'layer -s 2;');
% Now find the layer
graphLayer_V = invoke(originObj, 'FindGraphLayer', graph_V_Lum);
% Get dataplot collection from the graph layer
dataplots_V = invoke(graphLayer_V, 'DataPlots');

% Add data column by column to the graph
for i = 1 : size(lifetimeArray,2)/3
    hrs_col_idx = 3*i - 3;
    V_col_idx = 3*i - 2;
    
    % Create a data range
    dr = invoke(originObj, 'NewDataRange');
    
    % Add data to data range
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% row   col   row   col
    invoke(dr, 'Add', 'X', wkSheet_lifetime, 0 , hrs_col_idx, -1, hrs_col_idx); % x is hrs
    invoke(dr, 'Add', 'Y', wkSheet_lifetime, 0 , V_col_idx, -1, V_col_idx); % y is V
    
    % Add data plot to graph layer
    % 200 -- line
    % 201 -- symbol
    % 202 -- symbol+line
    invoke(dataplots_V, 'Add', dr, 200);
end

% Group data (This allows colors to be automatically incremented and a
% single legend entry to be created for all the data sets with the
% same legend entry)
invoke(graphLayer_V, 'Execute', 'layer -g;');

% For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Layer-cmd
% Reconstruct the legend (can be done manually in Origin via Ctrl+L)
%invoke(graphLayer_JV, 'Execute', 'legend -r;');
% For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Legend-cmd

end
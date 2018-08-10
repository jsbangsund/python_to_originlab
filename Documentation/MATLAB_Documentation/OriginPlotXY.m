function OriginPlotXY(templatePath, savePath, opj_filename, new_or_existing, symbol_or_line, legendEntries, metadataHeader, x_units, y_units, x, y, y_error)
%% Documentation
    % This function allows one to send x, y, and y-error data to an origin
    % project (either new or existing) and apply a user-defined template.
    %
    % y-error is an optional input
    % x and y must be of equal dimension, or x must have only one column
    % metadataHeader should be of equal column number as y
    % legendEntries must be a cell array of strings, corresponding to each
        % column of y data
    % symbol_or_line -> 'symbol', 'line', or 'symbol+line' are acceptable
        % inputs
    % new_or_existing -> 'New' or 'Existing' are acceptable inputs
    % Example usage:
    % templatePath = 'C:\Users\JSB\Documents\OriginLab\2016\User Files\EQE-Line.otp';
    % OriginPlotXY(templatePath, savePath, opj_filename, 'New', 'line', legendEntries, metadataHeader, x_units, y_units, x, y, y_error)
    
numberOfInputs = 12;
%%  Initiate Origin Session and Project
    % Obtain Origin COM Server object
    % Connect to an existing instance of Origin
    % or create a new one if none exists
    originObj = actxserver('Origin.ApplicationSI');
    
    % Open project or make new project
    if strcmp(new_or_existing,'New')
        % Close previous project and make a new one
        invoke(originObj, 'NewProject');
    elseif strcmp(new_or_existing,'Existing')
        openCommand = ['doc -o ' savePath opj_filename '.opj;']; % terminate savePath with \
        invoke(originObj, 'Execute', openCommand);
    else
        error('Acceptable inputs are ''New'' or ''Existing'' for new_or_existing')
    end
    
    % Make the Origin session visible
    invoke(originObj, 'Execute', 'doc -mc 1;');

    % Clear "dirty" flag in Origin to suppress prompt 
    % for saving current project
    invoke(originObj, 'IsModified', 'false');
    
    
    
    %% Send data to workbook
    % Create a workbook
    wkBook = invoke(originObj, 'CreatePage', 2);%, 'EQE'        , 'Origin');
    % Find the worksheet
    wkSheet = invoke(originObj, 'FindWorksheet', wkBook);
    
    % Set number of columns based on inputs and array sizes
    % Make data array of alternating x, y, y-error columns
    if size(x,2) == 1 && nargin < numberOfInputs
        NumOfColumns = size(x,2) + size(y,2);
        dataArray = [x, y];
    elseif size(x,2) == 1 && nargin == numberOfInputs
        NumOfColumns = size(x,2) + size(y,2) + size(y_error,2);
        dataArray = x;
        % Make array of alternating y and y-error columns
        for i = 1 : size(y,2)
            dataArray = [dataArray, y(:,i), y_error(:,i)];
        end
    elseif size(x,2) > 1 && nargin < numberOfInputs
        NumOfColumns = size(x,2) + size(y,2);
        dataArray = [];
        % Make array of alternating y and y-error columns
        for i = 1 : size(y,2)
            dataArray = [dataArray, x(:,i), y(:,i)];
        end
    elseif size(x,2) > 1 && nargin == numberOfInputs
        NumOfColumns = size(x,2) + size(y,2) + size(y_error,2);
        dataArray = [];
        % Make array of alternating y and y-error columns
        for i = 1 : size(y,2)
            dataArray = [dataArray, x(:,i), y(:,i), y_error(:,i)];
        end
    end;
    
    invoke(wkSheet, 'Cols', NumOfColumns);
    
    % Get column collection in the worksheet
    columns = invoke(wkSheet, 'Columns');
    
    % Get user param titles from first column of metadataHeader
    userParams = metadataHeader(:,1);
    % Now make user param rows visible and rename
    for param_idx = 1 : length(userParams)
        % Show the user-defined parameter (so row is visible when worksheet is opened)
        invoke(wkSheet, 'Execute', ['wks.UserParam' num2str(param_idx) '=1;']);
        % Set parameter name
        invoke(wkSheet, 'Execute', ['wks.UserParam' num2str(param_idx) '$="' userParams{param_idx} '";']);     
    end
    
    % Set x column (if only one x column)
    if size(x,2) == 1
        col_x = invoke(columns, 'Item', uint8(0));
        col_x.invoke('Units',x_units{1});
        col_x.invoke('Type',3); % 3 -> x column
    end
    
    for i = 1 : size(y,2);
        if size(x,2) == 1 && nargin < numberOfInputs
            col_y_index = i;
            % get column objects for y and y error
            col_y = invoke(columns, 'Item', uint8(col_y_index));
            % Set column types (e.g. y or error)
            col_y.invoke('Type',0); % 2 -> y error column
            % Set units
            col_y.invoke('Units', y_units{i});
        elseif size(x,2) > 1 && nargin < numberOfInputs % If >1 x column and no error
            col_y_index = 2*i - 1;
            col_x = invoke(columns, 'Item', uint8(col_y_index - 1));
            col_y = invoke(columns, 'Item', uint8(col_y_index));
            % Set column types (e.g. y or error)
            col_x.invoke('Type',3); % 3 -> x column
            col_y.invoke('Type',0); % 2 -> y error column
            % Set units
            col_x.invoke('Units', x_units{i});
            col_y.invoke('Units', y_units{i});
        elseif size(x,2) == 1 && nargin == numberOfInputs % If one x column + y error
            col_y_index = 2*i - 1;
            % get column objects for y and y error
            col_y = invoke(columns, 'Item', uint8(col_y_index));
            col_error = invoke(columns, 'Item', uint8(col_y_index+1));
            % Set column types (e.g. y or error)
            col_y.invoke('Type',0); % 0 -> y column
            col_error.invoke('Type',2); % 2 -> y error column
            % Set units
            col_y.invoke('Units', y_units{i});
            col_error.invoke('Units', y_units{i});
        elseif size(x,2) > 1 && nargin == numberOfInputs % If one x column + y error
            col_y_index = 3*i - 2;
            % get column objects for y and y error
            col_x     = invoke(columns, 'Item', uint8(col_y_index - 1));
            col_y     = invoke(columns, 'Item', uint8(col_y_index));
            col_error = invoke(columns, 'Item', uint8(col_y_index + 1));
            % Set column types (e.g. y or error)
            col_x.invoke('Type',3); % 3 -> x column
            col_y.invoke('Type',0); % 0 -> y column
            col_error.invoke('Type',2); % 2 -> y error column
            % Set units
            col_x.invoke('Units', x_units{i});
            col_y.invoke('Units', y_units{i});
            col_error.invoke('Units', y_units{i});
        end
      
        % Enter metadata into user-defined parameter rows:
        for param_idx = 1 : length(userParams)
            % syntax: 'col(ColumnIndex)[UserParameterRowName]$ = "UserParameterMetadata";'
            % e.g.: 'col(1)[Date]$ = "2016-01-05";'
            % Note: who knows what origin developers were thinking, but in
            % this case, column indexing starts from 1, not 0. Consistency
            % nightmare. The +1 accounts for this
            % Metadata is placed in same column as y data
            invoke(wkSheet,'Execute',['col(' num2str(col_y_index+1) ')[' ...
                userParams{param_idx} ']$ = "' metadataHeader{param_idx,i+1} '";']);
        end        
        
        % Add comments text, which is accessed by the legend and determines
        % color incrementing/grouping of data
        col_y.invoke('Comments', legendEntries{i}); 
        
    end
    %% Send data to graph
    % Set data to the columns
    % I use num2cell because origin can interpret NaN values in a cell
    % array but not in a double array for some reason.
    invoke(wkSheet, 'SetData', num2cell(dataArray), 0, 0);
    
    % Create a graph
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Graph Name
    graph = invoke(originObj, 'CreatePage', 3, 'Graph1', templatePath);
    
    % Find the graph layer
    graphLayer = invoke(originObj, 'FindGraphLayer', graph);
    
    % Get dataplot collection from the graph layer
    dataplots = invoke(graphLayer, 'DataPlots');
    
    % Add data column by column to the graph
    for i = 1 : size(y,2)
        % Create a data range
        dr = invoke(originObj, 'NewDataRange');
        
        if size(x,2) == 1 && nargin < numberOfInputs
            col_y_index = i;
            % Add data to data range
            %%%%%%%%%%%%%%%%%%%%%%%%%%%%%% row  col   row   col
            invoke(dr, 'Add', 'X', wkSheet, 0 , 0 , -1, 0 );
            invoke(dr, 'Add', 'Y', wkSheet, 0 , col_y_index   , -1, col_y_index );
        elseif size(x,2) > 1 && nargin < numberOfInputs % If >1 x column and no error
            col_y_index = 2*i - 1;
            col_x_index = col_y_index - 1;
            % Add data to data range
            %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% row   col   row   col
            invoke(dr, 'Add', 'X', wkSheet, 0 , col_x_index , -1, col_x_index );
            invoke(dr, 'Add', 'Y', wkSheet, 0 , col_y_index  , -1, col_y_index );
        elseif size(x,2) == 1 && nargin == numberOfInputs % If one x column + y error
            col_y_index = 2*i - 1;
            col_error_index = col_y_index + 1;
            % Add data to data range
            %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% row   col   row   col
            invoke(dr, 'Add', 'X', wkSheet, 0 , 0 , -1, 0 );
            invoke(dr, 'Add', 'Y', wkSheet, 0 , col_y_index  , -1, col_y_index );
            invoke(dr, 'Add', 'ED', wkSheet, 0 , col_error_index  , -1, col_error_index );
        elseif size(x,2) > 1 && nargin == numberOfInputs % If one x column + y error
            col_y_index = 3*i - 2;
            col_x_index = col_y_index - 1;
            col_error_index = col_y_index + 1;
            % Add data to data range
            %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% row   col   row   col
            invoke(dr, 'Add', 'X', wkSheet, 0 , col_x_index , -1, col_x_index );
            invoke(dr, 'Add', 'Y', wkSheet, 0 , col_y_index  , -1, col_y_index );
            invoke(dr, 'Add', 'ED', wkSheet, 0 , col_error_index  , -1, col_error_index );
        end
        
        % Add data plot to graph layer
        % 200 -- line
        % 201 -- symbol
        % 202 -- symbol+line
        if strcmp(symbol_or_line,'symbol')
            invoke(dataplots, 'Add', dr, 201);    
        elseif strcmp(symbol_or_line,'line')
            invoke(dataplots, 'Add', dr, 200);  
        elseif strcmp(symbol_or_line,'symbol+line')
            invoke(dataplots, 'Add', dr, 202);
        end
    end
    
    % Group data (This allows colors to be automatically incremented and a
    % single legend entry to be created for all the data sets with the same
    % legend entry)
    invoke(graphLayer, 'Execute', 'layer -g;');
    
    % Reconstruct the legend (can be done manually in Origin via Ctrl+L)
    invoke(graphLayer, 'Execute', 'legend -r;');

    %% Commands to rescale the graph and set axes labels (unnecessary with template)    
%     % Auto-Rescale the graph layer
%     invoke(gl, 'Execute', 'rescale;');
%     
%     % Change the bottom x' title
%     invoke(gl, 'Execute', 'xb.text$ = "J (mA/cm^2)";');
%     % Change the left y's title
%     invoke(gl, 'Execute', 'yl.text$ = "EQE (%)";');
%     
%     %show the top and right axes
%     invoke(gl, 'Execute', 'range ll = !;');
%     invoke(gl, 'Execute', 'll.x2.showAxes=3;');
%     invoke(gl, 'Execute', 'll.y2.showAxes=3;');
%     
%     %set the x axis scale
%     invoke(gl, 'Execute', 'll.x.from=1E-4;');
%     invoke(gl, 'Execute', 'll.x.to=1E3;');
%
%     %set the y axis scale
%     invoke(gl, 'Execute', 'll.y.from=-2.5;');
%     invoke(gl, 'Execute', 'll.y.to=22.5;');
%     
%     %set the x axis Major tick increment. 
%     invoke(gl, 'Execute', 'll.x.inc=10;');

    %% Save and exit
    %Save the current project using the specified path and filename
    if strcmp(new_or_existing,'New')
        saveCommand = ['save ' savePath opj_filename '.opj;'];
        invoke(originObj, 'Execute', saveCommand);  
    elseif strcmp(new_or_existing,'Existing')
        saveCommand = ['save ' savePath opj_filename '.opj;']; % May need to append $ at end of filename??
        invoke(originObj, 'Execute', saveCommand);
    end
    % saveCommand should have form 'save C:\myPath\myOriginFile.opj'
    
    % Close the project (this will allow you to delete or modify the
    % project. Otherwise the project will stay open.)
    invoke(originObj, 'Execute', 'doc -d;');
    % For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Document-cmd
    
    % Release
    release(originObj);
    
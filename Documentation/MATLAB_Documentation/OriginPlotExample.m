%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% This m-file shows some base operations between a MATLAB client and an 
% Origin Server application.
%
% This example does the following:
%   -> Connect to an existing Origin server application or create a new 
%      one if none exists.
%   -> Create workbook and find workseet, and then change worksheet name.
%   -> Get columns from worksheet, and set column's type, and set data 
%      to column.
%   -> Create graph and add x-y-error data as scatter line plot to graph.
%   -> Customize plot, such as axes' label, legend, range, etc.
%   -> Save project.
%
% Usage:
%   data = [0.1:0.1:3; sin(0.1:0.1:3)]';
%   OriginPlotExample(data, {'x data','sin(x)'}, 'testPlot.opj', 'C:\Users\...\');
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% Documentation resources:
% Automation server forum: http://www.originlab.com/forum/forum.asp?FORUM_ID=18
% Lab talk documentation: http://www.originlab.com/doc/LabTalk/
% In origin: Help -> Programming -> Automation Server


function OriginPlotExample(data, legendEntries, opjFilename, path)
    
    %% Initiate session and project
    % longNames must be a cell array of strings, corresponding to each
        % column of EQE data
    
    % Obtain Origin COM Server object
        % Connect to an existing instance of Origin
        % or create a new one if none exists
    originObj = actxserver('Origin.ApplicationSI');
    
    % Close previous project and make a new one
    invoke(originObj, 'NewProject');
    
    % Make the Origin session visible (I don't know why this is necessary)
    invoke(originObj, 'Execute', 'doc -mc 1;');

    % Clear "dirty" flag in Origin to suppress prompt 
        % for saving current project
    invoke(originObj, 'IsModified', 'false');
    
    % Create a workbook
    % 2 -> workbook, 3 -> graph, more documentation at: http://www.originlab.com/doc/COM/Classes/Application/CreatePage
    %                                        Type of Page,  Workbook name     , Template
    workbook = invoke(originObj, 'CreatePage',      2     , 'My Workbook name', 'Origin');
    
    % Find the worksheet
    worksheet = invoke(originObj, 'FindWorksheet', workbook);
    
    % Rename the worksheet
    invoke(worksheet, 'Name', 'My Worksheet Name');

    %% Insert data into worksheet
    % Set number of columns based on data
    NumberOfColumns = size(data,2);
    invoke(worksheet, 'Cols', NumberOfColumns);
    
    % Get column collection in the worksheet
    cols = invoke(worksheet, 'Columns');
    
    % NOTE: column indexing in origin starts from 0

    %                           Column # (starting from 0)
    colx = invoke(cols, 'Item', uint8(0)); % I don't know why they use the uint8() command. Should work either way.
    coly = invoke(cols, 'Item', uint8(1));
    
    % Write long name, units, and comments:
    coly.invoke('LongName', 'Date, Sample, etc.'); % Here I mix syntax. Both invoke(coly,...) and coly.invoke(...) work.
    coly.invoke('Units', 'mA/cm^2');
    coly.invoke('Comments', legendEntries{i});

    % Set column type
    % 3 -> x, 0 -> y, 2 -> error
    invoke(colx, 'Type', 3);  % x column
    invoke(coly, 'Type', 0);  % y column

    
    % Insert data into the columns
    %                                            starting col, starting row
    invoke(worksheet, 'SetData', num2cell(data), 0           , 0);
    % Note: Origin interprets cell arrays better than doubles for some
    % reason. For example, they will truncate NaN's, whereas if passing a
    % double array with NaN's, Origin will freak out and add seemingly
    % random points. The num2cell() conversion is not required, however.

    %% Create a graph
    
    % Set template file path
    templatePath = 'C:\Users\username\Documents\OriginLab\2016\User Files\myTemplate.otp';
    % Make graph object
    graphObject = invoke(originObj, 'CreatePage', 3, 'My Graph Name', templatePath);
    % Note: replace template path with 'Origin' if you do not wish to apply
    % a template
    
    % Find the graph layer
    graphLayer = invoke(originObj, 'FindGraphLayer', graphObject);
    
    % Get dataplot collection from the graph layer
    dataPlots = invoke(graphLayer, 'DataPlots');

    % Create a data range
    dataRange = invoke(originObj, 'NewDataRange');
    
    % Add data to data range
    %                                        Start Row?  Start Col Index   End Row??   End Col Index
    invoke(dataRange, 'Add', 'X', worksheet, 0         , 0               , -1       ,   0        );
    invoke(dataRange, 'Add', 'Y', worksheet, 0         , 1               , -1       ,   1        );
    % I don't really understand what the third index does, but -1 has
    % worked for me.
    % You may be able to add multiple Y columns at once by changing the End
    % Col index, but I haven't tested thi.
    
    % Add data plot to graph layer
    % 200 -- line
    % 201 -- symbol
    % 202 -- symbol+line
    invoke(dataPlots, 'Add', dataRange, 200);  % 200 for line plot   
    
    % If you want to add multiple columns of data in a loop, you need to
    % call each of these commands in each loop:
%     dataRange = invoke(originObj, 'NewDataRange');
%     invoke(dr, 'Add', 'X', worksheet, 0 , x_column_index , -1, x_column_index );
%     invoke(dr, 'Add', 'Y', worksheet, 0 , y_column_index , -1, y_column_index );
%     invoke(dataPlots, 'Add', dataRange, 202);  % 202 for symbol+line plot
    


    
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

    %% Group data and handle the legend
    % Group data (This allows colors to be automatically incremented and a
        % single legend entry to be created for all the data sets with the 
        % same legend entry)
    invoke(graphLayer, 'Execute', 'layer -g;');
    % For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Layer-cmd
    
    % Reconstruct the legend (can be done manually in Origin via Ctrl+L)
    invoke(graphLayer, 'Execute', 'legend -r;');
    % For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Legend-cmd

    %% Save the .opj file and close out origin
    %Save the current project to a .opj using the specified path and filename
    saveCommand = ['save ' path opjFilename '.opj'];
    invoke(originObj, 'Execute', saveCommand);  
    % saveCommand should have form 'save C:\myPath\myOriginFile.opj'
    
    % Close the project (this will allow you to delete or modify the
        % project. Otherwise the project will stay open.)
    invoke(originObj, 'Execute', 'doc -d;')
    % For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Document-cmd
    
   
    % Release
    release(originObj);
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
%   x = [0.1:0.1:3; 10 * sin(0.1:0.1:3); 20 * cos(0.1:0.1:3)]';
%   MATLABCallOrigin(x);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function MATLABCallOrigin(x)

    % Obtain Origin COM Server object
    % Connect to an existing instance of Origin
    % or create a new one if none exists
    
    originObj = actxserver('Origin.ApplicationSI');
    
    % Make the Origin session visible
    invoke(originObj, 'Execute', 'doc -mc 1;');

    % Clear "dirty" flag in Origin to suppress prompt 
    % for saving current project
    invoke(originObj, 'IsModified', 'false');
    
    % Create a workbook
    strBook = invoke(originObj, 'CreatePage', 2, '', 'Origin');
    
    % Find the worksheet
    wks = invoke(originObj, 'FindWorksheet', strBook);
    
    % Rename the worksheet to "MySheet"
    invoke(wks, 'Name', 'MySheet');
    
    % Set 3 columns
    invoke(wks, 'Cols', 6);
    
    % Get column collection in the worksheet
    cols = invoke(wks, 'Columns');
    
    % Get the columns
    colx = invoke(cols, 'Item', uint8(0));
    coly = invoke(cols, 'Item', uint8(1));
    colerr = invoke(cols, 'Item', uint8(3));
    
    % Set column type
    invoke(colx, 'Type', 3);  % x column
    invoke(coly, 'Type', 0);  % y column
    invoke(colerr, 'Type', 2);  % y error
   
    % Set data to the columns
    invoke(wks, 'SetData', x, 0, 0);
    
    % Create a graph
    strGraph = invoke(originObj, 'CreatePage', 3, '', 'Origin');
    
    % Find the graph layer
    gl = invoke(originObj, 'FindGraphLayer', strGraph);
    
    % Get dataplot collection from the graph layer
    dps = invoke(gl, 'DataPlots');
    
    % Create a data range
    dr = invoke(originObj, 'NewDataRange');
    
    % Add data to data range
    invoke(dr, 'Add', 'X', wks, 0, 0, -1, 0);
    invoke(dr, 'Add', 'Y', wks, 0, 1, -1, 1);
    invoke(dr, 'Add', 'ED', wks, 0, 2, -1, 2);
    
    % Add data plot to graph layer
    invoke(dps, 'Add', dr, 202);  % 202 for symbol+line plot
    
    % Rescale the graph layer
    invoke(gl, 'Execute', 'rescale;');
    
    % Change the bottom x' title
    invoke(gl, 'Execute', 'xb.text$ = "Channel";');
    % Change the left y's title
    invoke(gl, 'Execute', 'yl.text$ = "Amplitude";');
    
    %show the top and right axes
    invoke(gl, 'Execute', 'range ll = !;');
    invoke(gl, 'Execute', 'll.x2.showAxes=3;');
    invoke(gl, 'Execute', 'll.y2.showAxes=3;');
    
    %set the x axis scale
    invoke(gl, 'Execute', 'll.x.from=0;');
    invoke(gl, 'Execute', 'll.x.to=3;');
    
    %set the x axis Major tick increment. 
    invoke(gl, 'Execute', 'll.x.inc=10;');

    %delete the legend
    invoke(gl, 'Execute', 'label -r legend;');
    
    %Save the current project using the specified path and filename
    invoke(originObj, 'Execute', 'save C:\Users\JSB\Google Drive\Research\Scripts\MATLABCallOrigin\MatlabCallOrigin.opj');  
    
    % Release
    release(originObj);
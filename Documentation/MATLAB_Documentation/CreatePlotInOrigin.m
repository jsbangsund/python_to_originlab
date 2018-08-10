%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%
% This m-file shows a basic example of how to communicate between a MATLAB client and an Origin Server application.
% This example does the following:
%       -> Connects to Origin server application if an instance exists, or launches a new instance of Origin server
%       -> Loads a previously customized Origin Project file which has a worksheet and a custom graph
%       -> Sends some data over to the Origin worksheet, which then updates the graph
%       -> Rescales the graph and copies the graph image to the system clipboard
%
% For documentation on all methods and properties supported by the Origin Automation Server, please refer to the
% Programming Help File, and the topic "Calling Origin from Other Applications (Automation Server Support)".
%
% Note: This m-file was tested only in MATLAB 6.1
% You may need to download a new DLL from the MathWorks website:
%        ftp://ftp.mathworks.com/pub/tech-support/solutions/s29502/actxcli.dll
% and replace the older DLL found in the following MATLAB subfolder
%       $MATLAB\toolbox\matlab\winfun\@activex\private\actxcli.dll
% to fix a bug in MATLAB ActiveX 
%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%
function CreatePlotInOrigin()   
    %
    % Obtain Origin COM Server object
    % This will connect to an existing instance of Origin, or create a new one if none exist
    originObj=actxserver('Origin.ApplicationSI');

    % Make the Origin session visible
    invoke(originObj, 'Execute', 'doc -mc 1;');
       
    % Clear "dirty" flag in Origin to suppress prompt for saving current project
    % You may want to replace this with code to handling of saving current project
    invoke(originObj, 'IsModified', 'false');
    
    % Load the custom project CreateOriginPlot.OPJ found in Samples area of Origin installation
    invoke(originObj, 'Execute', 'syspath$=system.path.program$;');
    strPath='';
    strPath = invoke(originObj, 'LTStr', 'syspath$');
    invoke(originObj, 'Load', strcat(strPath, 'Samples\COM Server and Client\Matlab\CreatePlotInOrigin.OPJ'));

    % Create some data to send over to Origin - create three vectors
    % This can be replaced with real data such as result of computation in MATLAB
    mdata = [0.1:0.1:3; 10 * sin(0.1:0.1:3); 20 * cos(0.1:0.1:3)];
    % Transpose the data so that it can be placed in Origin worksheet columns
    mdata = mdata';
    % Send this data over to the Data1 worksheet
    invoke(originObj, 'PutWorksheet', 'Data1', mdata);
    
    % Rescale the two layers in the graph and copy graph to clipboard
    invoke(originObj, 'Execute', 'page.active = 1; layer - a; page.active = 2; layer - a;');
    % '4' makes it an OLE object (clicking opens origin)
    invoke(originObj, 'CopyPage', 'Graph1', '4'); 
    
    % You can now get the graph image from clipboard and paste in PowerPoint etc. 
    %
% end
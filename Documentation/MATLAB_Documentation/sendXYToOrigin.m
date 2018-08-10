%% Example of sending x-y data to origin
% Note: sometimes Origin likes to freeze and MATLAB will just continue
% running. Ctrl+C doesn't work in this case, so go to task manager and
% manually close Origin. It is also advisable to have Origin closed when
% accessing Origin from MATLAB.


% Get path to save origin file
savePath = mfilename('fullpath');
subtract = length(mfilename());
savePath = savePath(1:end-subtract); % subtract off length of .m filename

% load data
load PlotXY_example_data.mat

% Select a template
% Origin typically saves templates in your user files directory
% e.g. 'C:\Users\JSB\Documents\OriginLab\2016\User Files\Y-Error.otp'
templatePath = [savePath 'Y-Error.otp'];
% Alternatively you can use built-in templates:
% templatePath = 'Scatter';

% Name the .opj file
opj_filename = 'SummaryPlots';

% Selecting 'New' will overwrite a .opj file with the same name
% Selecting 'Existing' will add new graphs/workbooks to an existing .opj
new_or_existing = 'New'; % 'Existing' or 'New'
symbol_or_line = 'symbol+line'; % 'symbol', 'line', or 'symbol+line'

% Make sure metadataHeader is a cell array of strings
metadataHeader = cellfun(@(x) num2str(x),metadataHeader,'UniformOutput',false);

% Send everything over to the origin plot function
% Note: y_error is optional and you must either have 1 x column (same for
% all y data) or the same number of x columns as y data
OriginPlotXY(templatePath, savePath, opj_filename, new_or_existing, symbol_or_line, ...
    legendEntries, metadataHeader, x_units, y_units, x, y, y_error)
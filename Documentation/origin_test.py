# For documentation see: http://www.originlab.com/doc/COM/Classes
# Wiki: http://wiki.originlab.com/~originla/wiki/index.php?title=Category:COM_Server
# Some LabTalk commands don't seem to work on python, but I might be missing something
# Labtalk documentation at: http://www.originlab.com/doc/LabTalk/
# Some crappy documentation on PyOrigin: http://www.originlab.com/doc/python
import win32com.client
import os
# ------------------------------------ User inputs
# Directory where templates are stored
templateDir = os.path.abspath('C:\\Users\\JSB\\Documents\\OriginLab\\2016\\User Files\\')
# ------------------------------------
# Connect to Origin
origin = win32com.client.Dispatch("Origin.ApplicationSI")
# To open a new origin session (not overlap on current open session use:
# origin = win32com.client.Dispatch("Origin.Application")
# Make session visible
origin.Visible=1
#origin.Exit()
# Close previous project and make a new one
origin.NewProject
# Wait for origin to compile
origin.Execute("sec -poc 3.5") 
# Preferred way to create a workbook is to use CreatePage and reference the 
# workbook by name. This just returns unicode of the name, and not an instance
workbookName = origin.CreatePage(2, 'EL-PL' , 'Origin') # 2 for workbook
# get workbook instance from name
workbook = origin.WorksheetPages(workbookName)
# Alternative command to make a workbook
#workbook = origin.WorksheetPages.Add # For some reason this method results in odd formatting
workbook.Name = 'TestName'
# Get worksheet instance, index starts at 0
worksheet = workbook.Layers(0)
# Set number of columns
worksheet.Cols = 5
# Get column instance, index starts at 0
column = worksheet.Columns(0)
# Change column Units, Long Name, or Comments]
column.LongName = 'This is a long name'
column.Units = 'megameters'
column.Comments = 'blah blah'
# Set type to X
column.Type = 3
# Set type to Y
column.Type = 0
# Increase Width of Column:
width = 10;
col_idx = 0;
origin.Execute('wks.col' + str(col_idx+1) + '.width=' + str(width) + ';')
# Set user parameter rows, make visible, and rename
# These header rows can be used to store reference information
userParams = ['Voltage','Efficiency','t50']
for param_idx in range(0,len(userParams)):
    # Show the user-defined parameter (so row is visible when worksheet is opened)
    # Index starts at 1
    worksheet.Execute('wks.UserParam' + str(param_idx+1) + '=1;')
    # Set parameter name
    worksheet.Execute('wks.UserParam' + str(param_idx+1) + '$="' + userParams[param_idx] + '";')
    # set value in user param row
    value = 'my metadata'
    worksheet.Execute('col(' + str(col_idx+1) + ')[' + userParams[param_idx] + ']$ = "' + str(value) + '";')
# Add another worksheet to the active page
worksheet2 = origin.ActivePage.Layers.Add
worksheet2.Name = 'Test2'
# Make worksheet active (indexing starts at 0?)

##### This stuff might not work
# origin.Execute('page.active$ = ' + str(0))
# Now that the worksheet is active, find it and get an instance
worksheet = origin.FindWorksheet(workbookName)
##### 

# Save origin project
savepath = 'C:\Users\JSB\Google Drive\Research\Scripts\Python\\'
filename = 'test'
# This isn't working for some reason
#origin.Save(str(os.path.join(savepath,filename)))
# Labtalk command can be used instead
saveCommand = 'save ' + savepath + filename + '.opj'
#origin.Execute(saveCommand)
# Close origin, if desired
#origin.Execute('doc -d;')
#origin.Save(savepath + filename)
# Put the X and Y data into the worksheet
origin.PutWorksheet('['+workbook.Name+']'+worksheet.Name, xData, 0, 0) # row 0, col 0
origin.PutWorksheet('['+workbook.Name+']'+worksheet.Name, yData, 0, 1) # row 0, col 1

# Make the Origin session visible
#origin.Execute('doc -mc 1;')

# Clear "dirty" flag in Origin to suppress prompt 
# for saving current project
#origin.IsModified('false')

def createGraph(origin,graphName,templatePath,worksheets,x_cols,y_cols,LineOrSym):
    # origin is an origin app instance
    # graphName is a string of the name of the graph
    # templatePath is a string of the path and template name to be used    
    # worksheets must be worksheet instance or a list of worksheet instances
    # Each worksheet must be formatted identically
    # For differently formatted worksheets, would need to make x_cols a list of lists
    # x_col, y_col, and LineOrSym should be same length
    # x_col and y_col are lists of indices for x and y columns that are paired element-wise
    # LineOrSym is a list of strings specifying 'Symbol', 'Line+Symbol', or 'Line'

    # Create graph page and instance
    # Standard template is 'Origin'
    graph = origin.CreatePage(3, graphName , templatePath);
    graph = origin.GraphPages(graph)
    # Find the graph layer
    graphLayer = origin.FindGraphLayer(graph)
    
    # Get dataplot collection from the graph layer
    dataPlots = graphLayer.DataPlots
    
    # Add data column by column to the graph
    # loop over worksheets within column loops so that data from same column
    # can be grouped. E.g. all PL data is in same column and will be grouped.
    for ci in range(0,len(x_cols)):
        for worksheet in worksheets:
            # Create a data range
            dr = origin.NewDataRange
            
            # Add data to data range
            #                  worksheet, start row, start col, end row (-1=last), end col
            dr.Add('X', worksheet, 0 , x_cols[ci], -1, x_cols[ci])
            dr.Add('Y', worksheet, 0 , y_cols[ci], -1, y_cols[ci])
            
            # Add data plot to graph layer
            # 200 -- line
            # 201 -- symbol
            # 202 -- symbol+line
            # If specified, plot symbol. By default, plot line
            if LineOrSym[ci] == 'Symbol':
                dataPlots.Add(dr, 201)
            elif LineOrSym[ci] == 'Line+Symbol':
                dataPlots.Add(dr, 202)
            else:
                dataPlots.Add(dr, 200)
        
        
        # Group each column (This allows colors to be automatically incremented
        # and a single legend entry to be created for all the data sets with
        # the same legend entry)
        BeginIndex = (ci-1)*len(worksheets) + 1
        EndIndex = BeginIndex + len(worksheets)-1
        graphLayer.Execute('layer -g ' + str(BeginIndex) + ' '  + str(EndIndex) + ';')
    
    
    # Rescales axes
    #Rescale type: 1 = manual, 2 = normal, 3 = auto, 4 = fixed from, and 5 = fixed to.
    #graphLayer.Execute(layer.axis.rescale=3');
    graphLayer.Execute('Rescale')
    
    # For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Layer-cmd
    # Reconstruct the legend (can be done manually in Origin via Ctrl+L)
    #graphLayer.Execute('legend -r;');
    # For further documentation see: http://www.originlab.com/doc/LabTalk/ref/Legend-cmd
    
    

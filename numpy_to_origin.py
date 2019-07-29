import numpy as np
import win32com.client
import os
import datetime
import time

def numpy_to_origin(
    data_array,column_axis=0,types=None,
    long_names=None,comments=None,units=None,
    user_defined=None,
    origin=None,project_filename='project.opj',
    origin_version=2018,
    worksheet_name='Sheet',workbook_name='Book',
    graph_name='Graph',template_name='LINE.otp',
    template_path='OriginTemplates'):
    '''
    Sends 2d numpy array to originlab worksheet
    Inputs:
    data_array = numpy array object
    column_axis = integer (0 or 1) for axis to interpret as worksheet columns
    long_names,comments,units = lists for header rows, length = # of columns
    user_defined = list of (key,value) tuples for metadata for a sheet
    origin = origin session, which is returned from previous calls to this program
             if passed, a new session will not be created, and graph will be added to 
             current session
    origin_version = 2016 other year, right now >2016 handles DataRange differently
    '''
    # If no origin session has been passed, start a new one
    if origin is None:
        # Connect to Origin client
        origin = win32com.client.Dispatch("Origin.ApplicationSI")
        # To open a new origin session (not overlap on current open session use:
        # origin = win32com.client.Dispatch("Origin.Application")
        # Make session visible
        origin.Visible=1
        # Session can be later closed using origin.Exit()
        # Close previous project and make a new one
        origin.NewProject
        # Wait for origin to compile
        origin.Execute("sec -poc 3.5")
        time.sleep(2)
        
    # Check if workbook exists. If not create a new workbook page with this name
    layer_idx=None
    if origin.WorksheetPages(workbook_name) is None:
        print(workbook_name)
        print(origin.WorksheetPages(workbook_name))
        workbook_name = origin.CreatePage(2, workbook_name , 'Origin') # 2 for workbook
        # Use Sheet1 if workbook is newly made
        layer_idx=0
    # get workbook instance from name
    wb = origin.WorksheetPages(workbook_name)
    if layer_idx is None:
        wb.Layers.Add() # Add a worksheet
        #then find the last worksheet to modify (to avoid overwriting other data)
        i = 0
        while not (wb.Layers(i) is None):
            i+=1
        layer_idx = i-1
    ws=wb.Layers(layer_idx) # Get worksheet instance, index starts at 0.
    ws.Name=worksheet_name # Set worksheet name
    # For now, assume only x and y data for each line (ignore error data)
    ws.Cols=data_array.shape[column_axis] # Set number of columns in worksheet
    # Change column Units, Long Name, or Comments]
    for col_idx in range(0,data_array.shape[column_axis]):
        col=ws.Columns(col_idx) # Get column instance, index starts at 0
        if (not long_names is None) and (len(long_names)>col_idx):
            col.LongName=long_names[col_idx]
        if (not units is None) and (len(units)>col_idx):
            col.Units=units[col_idx]
        if (not comments is None) and (len(comments)>col_idx):
            col.Comments=comments[col_idx]
        if not (types is None) and (len(types)>col_idx):
            # Set column data type to ( 0=Y, 3=X , ?=X error, ?=Y error)
            col.Type=types[col_idx]
        origin.PutWorksheet('['+wb.Name+']'+ws.Name, np.float64(data_array[col_idx,:]).tolist(), 0, col_idx) # start row, start col
    #origin.PutWorksheet('['+wb.Name+']'+ws.Name, np.float64(data_array).T.tolist(), 0, col_idx) # start row, start col
    if not user_defined is None:
        # User Param Rows
        for idx,param in enumerate(user_defined):
            ws.Execute('wks.UserParam' + str(idx+1) + '=1; wks.UserParam' + str(idx+1) + '$="' + param[0] + '";')
            ws.Execute('col(1)[' + param[0] + ']$="' + param[1] + '";')
        origin.Execute('wks.col1.width=10;')
    return origin,wb,ws
    
def  createGraph_multiwks(origin,graphName,template,templatePath,worksheets,x_cols,y_cols,LineOrSym,
                       origin_version=2018):
    # worksheets must be a list of worksheets
    # template is the full path and template filename
    # Each worksheet must be formatted identically
    # x_col, y_col, and LineOrSym should be same length
    # Create graph page and object
    templateFullPath=os.path.join(templatePath,template)
    # Create graph if doesn't already exist
    graphLayer = origin.FindGraphLayer(graphName)
    if graphLayer is None:
        graphName = origin.CreatePage(3, graphName, templateFullPath)
        # Find the graph layer
        graphLayer = origin.FindGraphLayer(graphName)

    # Get dataplot collection from the graph layer
    dataPlots = graphLayer.DataPlots

    # Add data column by column to the graph
    # loop over worksheets within column loops so that data from same column
    # can be grouped. E.g. all PL data is in same column and will be grouped.
    for ci,x_col in enumerate(x_cols):
        for wi,worksheet in enumerate(worksheets):
            # Create a data range
            # Tested only on origin 2016 and 2018
            if origin_version<=2016:
                dr = origin.NewDataRange # Make a new datarange
            elif origin_version>2016:
                dr = origin.NewDataRange()
            
            # Add data to data range
            #                  worksheet, start row, start col, end row (-1=last), end col
            dr.Add('X', worksheet, 0 , x_col,       -1, x_col)
            dr.Add('Y', worksheet, 0 , y_cols[ci], -1, y_cols[ci])
            # Add data plot to graph layer
            # 200 -- line
            # 201 -- symbol
            # 202 -- symbol+line
            # If specified, plot symbol. By default, plot line
            if LineOrSym[ci] in ['Sym','Symbol','Symbols']:
                dataPlots.Add(dr, 201)
            elif LineOrSym[ci] == 'Line+Sym':
                dataPlots.Add(dr,202)
            else:
                dataPlots.Add(dr, 200)

        # Group each column (This allows colors to be automatically incremented
        # and a single legend entry to be created for all the data sets with
        # the same legend entry)
        BeginIndex = ci*len(worksheets) + 1;
        EndIndex = BeginIndex + len(worksheets) - 1;
        graphLayer.Execute('layer -g ' + str(BeginIndex) + ' ' + str(EndIndex) + ';')
    # Rescales axes
    #Rescale type: 1 = manual, 2 = normal, 3 = auto, 4 = fixed from, and 5 = fixed to.
    #graphLayer.invoke('Execute','layer.axis.rescale=3');
    graphLayer.Execute('Rescale')
    graphLayer.Execute('legend -r')
    return graphName
    # To exit, call origin.Exit()
import datetime
import time
import numpy as np
import win32com.client
import os
import matplotlib.pyplot as plt
import matplotlib.colors as colors

try:
    import OriginExt
    app = OriginExt.Application()
    origin_version = app.GetLTVar("@V")
    # Origin 2015 9.2xn
    # Origin 2016 9.3xn
    # Origin 2017 9.4xn
    # Origin 2018 >= 9.50n and < 9.55n
    # Origin 2018b >= 9.55n
    # Origin 2019 >= 9.60n and < 9.65n (Fall 2019)
    # Origin 2019b >= 9.65n (Spring 2020)
except:
    print('OriginExt not installed, t')

# Ideas for improvements:
# - Compile line data (labels, format, color) into df
#   then use df to sort and group lines in origin
# - support for subplots / multiple layers
# - support for double y or double x axes
# - support for errorbars
def matplotlib_to_origin(
            fig,ax,
            origin=None,project_filename='project.opj',
            worksheet_name='Sheet',workbook_name='Book',
            graph_name='Graph',template_name='LINE.otp',
            template_path='OriginTemplates'):
    '''
    Inputs:
    fig = matplotlib figure object
    ax = matplotlib axis object
    template = origin template name for desired plot, if exists
    templatePath = path on local computer to template folder
    origin = origin session, which is returned from previous calls to this program
             if passed, a new session will not be created, and graph will be added to
             current session
    origin_version = 2016 other year, right now >2016 handles DataRange differently
    '''
    # If no origin session has been passed, start a new one
    if origin==None:
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
    # Create a workbook page
    workbook= origin.CreatePage(2, workbook_name , 'Origin') # 2 for workbook
    # get workbook instance from name
    wb = origin.WorksheetPages(workbook)
    # Get worksheet instance, index starts at 0. Can add more worksheets with wb.Layers.Add
    # wb.Layers.Add() for origin_version>2016
    ws=wb.Layers(0)
    ws.Name=worksheet_name # Set worksheet name
    # For now, assume only x and y data for each line (ignore error data)
    ws.Cols=len(ax.lines)*2 # Set number of columns in worksheet

    # Make graph page
    template=os.path.join(template_path,template_name) # Pick template
    graph = origin.CreatePage(3, graph_name, template) # Make a graph with the template
    graph_page = origin.GraphPages(graph) # Get graph page
    graph_layer = origin.FindGraphLayer(graph) # Get graph layer
    data_plots = graph_layer.DataPlots # Get dataplots
    # Grouping indices (not yet implemented)
    group_start_idx = 0
    group_end_idx = 0
    for line_idx,line in enumerate(ax.lines):
        # Indices for x and y columns
        x_col_idx = line_idx * 2
        y_col_idx = x_col_idx + 1
        col=ws.Columns(x_col_idx) # Get column instance, index starts at 0
        # Change column Units, Long Name, or Comments]
        col.LongName='X'
        col.Units='Unit'
        col.Comments=''
        col.Type=3 # Set column data type to ( 0=Y, 3=X , ?=X error, ?=Y error)
        col=ws.Columns(y_col_idx)
        col.Type=0
        col.LongName='Y'
        col.Units='Unit'
        # By default, lines have the label _line#
        # If the first character isn't "_", put this label
        # In the comments row
        if not line.get_label()[0] == '_':
            col.Comments = line.get_label()

        origin.PutWorksheet('['+wb.Name+']'+ws.Name, np.float64(line.get_xdata()).tolist(), 0, x_col_idx) # start row, start col
        origin.PutWorksheet('['+wb.Name+']'+ws.Name, np.float64(line.get_ydata()).tolist(), 0, y_col_idx) # start row, start col

        # Tested only on origin 2016 and 2018
        if origin_version<=2016:
            dr = origin.NewDataRange # Make a new datarange
        elif origin_version>2016:
            dr = origin.NewDataRange()
        # Add data to data range
        # Column type, worksheet, start row, start col, end row (-1=last), end col
        dr.Add('X', ws, 0 , x_col_idx, -1, x_col_idx)
        dr.Add('Y', ws, 0 , y_col_idx, -1, y_col_idx)
        # Add data plot to graph layer
        # 200 -- line
        # 201 -- symbol
        # 202 -- symbol+line
        # Symbols
        # https://www.originlab.com/doc/LabTalk/ref/List-of-Symbol-Shapes
        # https://www.originlab.com/doc/LabTalk/ref/Options_for_Symbols
        #0 = no symbol, 1 = square, 2 = circle, 3 = up triangle, 4 = down triangle,
        #5 = diamond, 6 = cross (+), 7 = cross (x), 8 = star (*), 9 = bar (-), 10 = bar (|),
        # 11 = number, 12 = LETTER, 13 = letter, 14 = right arrow, 15 = left triangle,
        #16 = right triangle, 17 = hexagon, 18 = star, 19 = pentagon, 20 = sphere
        # Symbol interior
        #0 = no symbol, 1 = solid, 2 = open, 3 = dot center, 4 = hollow, 5 = + center,
        # 6 = x center, 7 = - center, 8 = | center, 9 = half up, 10 = half right,
        # 11 = half down, 12 = half left
        # https://matplotlib.org/api/markers_api.html
        mpl_sym_conv = {'s':'1','o':'2','^':'3','v':'4','D':'5','+':'6','x':'7',
                                    '*':'8','_':'9','|':'10','h':'17','p':'19'}
        #Line
        if plt.getp(line,'marker')=='None':
            data_plots.Add(dr,200)
            lc = colors.to_hex(plt.getp(line,'color'))
            # Set line color and line width
            graph_layer.Execute(
                'range rr = !' + str(line_idx+1) + '; ' +
                'set rr -cl color('+lc+');' + # line color
                'set rr -w 500*'+str(plt.getp(line,'linewidth'))+';') # line width

        #Symbol
        elif plt.getp(line,'linestyle')=='None':
            data_plots.Add(dr,201)
            # Set symbol size, edge color, face color
            mec = colors.to_hex(plt.getp(line,'mec'))
            mfc = colors.to_hex(plt.getp(line,'mfc'))
            graph_layer.Execute(
                'range rr = !' + str(line_idx+1) + '; ' +
                'set rr -k '+mpl_sym_conv[plt.getp(line,'marker')]+';' + # symbol type
                'set rr -kf 2;' + # symbol interior
                'set rr -z '+str(plt.getp(line,'ms'))+';' + # symbol size
                'set rr -c color('+mec+');'+ # edge color
                'set rr -cf color('+mfc+');'+ # face color
                'set rr -kh 10*'+str(plt.getp(line,'mew'))+';')# edge width
        #Line+Symbol
        else:
            data_plots.Add(dr,202)
            # Set symbol size, edge color, face color
            lc = colors.to_hex(plt.getp(line,'color'))
            mec = colors.to_hex(plt.getp(line,'mec'))
            mfc = colors.to_hex(plt.getp(line,'mfc'))
            graph_layer.Execute(
                'range rr = !' + str(line_idx+1) + '; ' +
                'set rr -k '+mpl_sym_conv[plt.getp(line,'marker')]+';' + # symbol type
                'set rr -kf 2;' + # symbol interior
                'set rr -z '+str(plt.getp(line,'ms'))+');' + # symbol size
                'set rr -c color('+mec+');'+ # edge color
                'set rr -cf color('+mfc+');'+ # face color
                'set rr -kh 10*'+str(plt.getp(line,'mew'))+';' + # edge width
                'set rr -cl color('+lc+');' + # line color
                'set rr -w 500*'+str(plt.getp(line,'linewidth'))+';') # line width



    # For labtalk documentation of graph formatting, see:
    # https://www.originlab.com/doc/LabTalk/guide/Formatting-Graphs
    # https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-Label-obj
    # For matplotlib documentation, see:
    # https://matplotlib.org/api/axes_api.html
    # Get figure dimensions
    # Set figure dimensions
    # Get axes ranges
    x_axis_range = ax.get_xlim()
    y_axis_range = ax.get_ylim()
    # Get axes scale types
    x_axis_scale = ax.get_xscale()
    y_axis_scale = ax.get_yscale()
    # Get axes labels
    x_axis_label = ax.get_xlabel()
    y_axis_label = ax.get_ylabel()
    title = ax.get_title()
    # Set axes titles (xb for bottom axis, yl for left y-axis, etc.)
    graph_layer.Execute('label -xb ' + x_axis_label + ';')
    graph_layer.Execute('label -yl ' + y_axis_label + ';')
    # Set fontsizes
    #graph_layer.Execute('layer.x.label.pt = 12;')
    #graph_layer.Execute('layer.y.label.pt = 12;')
    #graph_layer.Execute('xb.fsize = 16;')
    #graph_layer.Execute('yl.fsize = 16;')

    # Axis label number format:
    # 1 = decimal without commas, 2 = scientific,
    # 3 = engineering, and 4 = decimal with commas (for date).
    # https://www.originlab.com/doc/LabTalk/ref/Layer-Axis-Label-obj
    # Set x-axis properties
    if x_axis_scale == 'linear':
        graph_layer.Execute('layer.x.type = 0;')
        # Change number format to decimal
        graph_layer.Execute('layer.x.label.numFormat=1')
    elif x_axis_scale == 'log':
        graph_layer.Execute('layer.x.type = 2;')
        # Change tick label number type to scientific
        graph_layer.Execute('layer.x.label.numFormat=2')
    graph_layer.Execute('layer.x.from = ' + str(x_axis_range[0]) + '; ' +
                           'layer.x.to = ' + str(x_axis_range[1]) + ';')
    # Set y-axis properties
    if y_axis_scale == 'linear':
        graph_layer.Execute('layer.y.type = 0;')
        # Change tick label number type to decimal
        graph_layer.Execute('layer.y.label.numFormat=1')
    elif y_axis_scale == 'log':
        graph_layer.Execute('layer.y.type = 2;')
        # Change tick label number type to scientific
        graph_layer.Execute('layer.y.label.numFormat=2')
    graph_layer.Execute('layer.y.from = '+str(y_axis_range[0])+'; '+
                           'layer.y.to = '+str(y_axis_range[1])+';')

    # Set page dimensions based on figure size
    # figure_size_inches = fig.get_size_inches()
    # graph_page.Execute('page.width= page.resx*'+str(figure_size_inches[0])+'; '+
                         # 'page.height= page.resy*'+str(figure_size_inches[1])+';')
    # Units 1 = % page, 2 = inches, 3 = cm, 4 = mm, 5 = pixel, 6 = points, and 7 = % of linked layer.
    # graph_layer.Execute('layer.unit=2; ' +
                           # 'layer.width='+str(figure_size_inches[0])+'; '+
                           # 'layer.height='+str(figure_size_inches[1])+';')
    # Group each column (This allows colors to be automatically incremented
            # and a single legend entry to be created for all the data sets with
            # the same legend entry)
    #graph_layer.Execute('layer -g ' + str(group_start_idx) + ' '  + str(group_end_idx) + ';')
    #graph_layer.Execute('Rescale')
    graph_layer.Execute('legend -r') # re-construct legend
    return origin

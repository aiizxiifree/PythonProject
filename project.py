'''Graph creator'''
import xlrd
import pygal
def command(status):
    '''This function will recieve command.
    Have 4 commands.
        1.) Graph = Bring the data to be graph.
        2.) Compare = First, bring the datas calculated. After that bring the data to be graph.
        3.) Two object compare = First, bring the datas calculated.  After that merge the data in the same graph.
        4.) Two object hraph = Merge the data in the same graph.'''
    if status == 'Compare':
        compare(status)
    elif status == 'Graph':
        excel_reader(status)
    elif status == 'Two object graph' or status == "Two object compare":
        two_object(status)
    else:
        command(input('"Compare"\n"Graph"\n"Two object compare"\n"Two object graph"\nEnter command : '))

def excel_reader(status):
    '''This function will recieve excel file and open.
    Sent to data manage fuction.'''
    print('-' * 100+'\nExcel reader')
    file_name = input('Enter file name : ')
    file_location = 'C:/Users/Home/Desktop/Project/whatever-master/%s.xlsx' % file_name # Change to your own file location.
    workbook = xlrd.open_workbook(file_location)
    sheet_count = workbook.nsheets # int values.
    sheet_name = workbook.sheet_names() # list of sheet's name.
    print('The total number of sheets %d' % sheet_count)
    for name in range(sheet_count):
        print(str(name+1)+'.)', sheet_name[name], end=' ')
    print()
    select_sheet = int(input('Select Once by enter number : ')) - 1
    sheet = workbook.sheet_by_index(select_sheet) #<-- This is sheet you want to open.
    sheet_rows = sheet.nrows # row of sheet
    sheet_columns = sheet.ncols # column of sheet
    if file_name == 'reg': # <-- Change to your own file's name (register).
        if select_sheet == 0:
            return register_info_manage(sheet, sheet_rows, sheet_columns, {}, status, select_sheet, file_name, sheet_name[select_sheet])
        else:
            terminations_of_collegians_info_manage(sheet, sheet_rows, sheet_columns, {}, sheet_name[select_sheet])
    elif file_name == 'complete': # <-- Change to your own file's name (graduate and job).
        return graduate_and_job_info(sheet, sheet_rows, sheet_columns, {}, status, select_sheet, file_name, sheet_name[select_sheet])

def register_info_manage(sheet, sheet_rows, sheet_columns, data, status, select_sheet, file_name, graph_name):
    '''Recieve raw data of excel.
    Turn raw data to information that you want.'''
    print('-' * 100+'\nData manage')
    for row in range(sheet_rows):
        data_in_line = []
        if row == 1 or row == 6 or row == 10:
            for column in range(2, sheet_columns):
                if sheet.cell_value(row, column) == '-':
                    data_in_line.append(None)
                elif 'float' in str(type(sheet.cell_value(row, column))):
                    data_in_line.append(int(sheet.cell_value(row, column)))
            if len(data_in_line) > 1:
                data[str(row)] = data_in_line
    for key in sorted(data.keys()):
        if key == '6':
            all_register = data[key] # <------ ข้อมูลรวมของนักศึกษาที่สมัครสอบ (list)
        elif key == '10':
            amount_of_eligible_students = data[key] # <------ ข้อมูลรวมของนักศึกษาที่มีสิทธิ์เข้าศึกษา (list)
        else:
            years = data[key] # <------ ตั้งแต่ปี 2545-2559 (list)
    if status == 'Compare' or status == 'Two object compare':
        return years, amount_of_eligible_students, select_sheet, file_name
    elif status == 'Graph':
        objects_plot_graph(years, all_register, amount_of_eligible_students, 'The '+graph_name, 'Registers', 'Holders')

def terminations_of_collegians_info_manage(sheet, sheet_rows, sheet_columns, data, graph_name):
    '''Recieve raw data of excel.
    Turn raw data to information that you want.'''
    print('-' * 100+'\nData manage')
    for row in range(sheet_rows):
        for column in range(sheet_columns):
            if column == 0 and sheet.cell_value(row, column).isdigit():
                if str(sheet.cell_value(row, column)) in data:
                    data[str(sheet.cell_value(row, column))] += int(sheet.cell_value(row, column+2))
                else:
                    data[str(sheet.cell_value(row, column))] = int(sheet.cell_value(row, column+2))
    years = [int(year) for year in sorted(data)]
    terminations_of_collegians = [int(data[year]) for year in sorted(data)]
    object_plot_graph(years, data, 'The '+graph_name)

def graduate_and_job_info(sheet, sheet_rows, sheet_columns, data, status, select_sheet, file_name, graph_name):
    '''Recieve raw data of excel.
    Turn raw data to information that you want.'''
    print('-' * 100+'\nData manage')
    for row in range(sheet_rows):
        data_in_line = []
        if row == 3 or row == 5:
            for column in range(1, sheet_columns-2):
                if sheet.cell_value(row, column) == '':
                    data_in_line.append(None)
                elif 'float' in str(type(sheet.cell_value(row, column))):
                    data_in_line.append(int(sheet.cell_value(row, column)))
            if len(data_in_line) > 1:
                data[str(row)] = data_in_line
    for key in data.keys():
        if key == '3':
            years = data[key]
        else:
            balchelor = data[key]
    years = [year - 3 for year in years]
    if status == 'Compare' or status == 'Two object compare':
        return years, balchelor, select_sheet, file_name
    elif status == 'Graph':
        object_plot_graph(years, balchelor, 'The '+graph_name)
    elif status == 'Two object graph':
        return years, balchelor, 'The '+graph_name

def compare(status):
    '''Bring the informations and match them together.
    Sent pair of infortions to calcutate function.'''
    print('-' * 100+'\nDatas compare')
    years, first_data, selector_first_data, file_name_first_data = excel_reader(status)
    years, second_data, selector_second_data, file_name_second_data = excel_reader(status)
    if file_name_second_data == file_name_first_data:
        return differnce_between_graduate_and_job(years, first_data, second_data, selector_first_data, selector_second_data, status, 'The percents of balchelors that got a job')
    else:
        return differnce_between_graduate_and_students(years, file_name_second_data, file_name_first_data, first_data, second_data, status, 'The percents of collegians that graduated')

def differnce_between_graduate_and_job(years, first_data, second_data, selector_first_data, selector_second_data, status, graph_name):
    '''This function will recieve informations and calculate information as percenet format.'''
    print('-' * 100+'\nDatas calculate')
    list_result = []
    if selector_first_data == 2:
        first_data, second_data = second_data, first_data
    for position in range(len(first_data)):
        if second_data[position] == None:
            result = None
        else:
            result = first_data[position] * 100 / second_data[position]
            result = '%.2f' % result
            result = float(result)
        list_result.append(result)
    if status == 'Two object compare':
        return years, list_result, graph_name
    elif status == 'Compare':
        object_plot_graph(years, list_result, graph_name)

def differnce_between_graduate_and_students(years, file_name_second_data, file_name_first_data, first_data, second_data, status, graph_name):
    '''This function will recieve informations and calculate information as percenet format.'''
    print('-' * 100+'\nDatas calculate')
    list_result = []
    if file_name_first_data == 'reg':
        first_data, second_data = second_data, first_data
    for position in range(len(first_data)):
        if first_data[position] == None:
            result = None
        else:
            result = first_data[position] * 100 / second_data[position-3]
            result = '%.2f' % result
            result = float(result)
        list_result.append(result)
    if status == 'Two object compare':
        return years, list_result, graph_name
    elif status == 'Compare':
        object_plot_graph(years, list_result, graph_name)

def objects_plot_graph(years, first_data, second_data, graph_name, name_of_fitst_data, name_of_second_data):
    '''Recieve informaions and plot it in graph.'''
    print('-' * 100+'\nPlot graph')
    chart_types = {'Line' : pygal.Line(), 'Bar' : pygal.Bar()}
    chart_selector = input('"Line"\n"Bar"\nEnter : ')
    chart = chart_types[chart_selector]
    chart.title = graph_name
    chart.x_labels = map(str, years)
    chart.add(name_of_fitst_data, first_data)
    chart.add(name_of_second_data, second_data)
    chart.render_to_file('line.svg')
    command(input('"Compare"\n"Graph"\n"Two object compare"\n"Two object graph"\nEnter command : '))

def object_plot_graph(years, data, graph_name):
    '''Recieve informaions and plot it in graph.'''
    print('-' * 100+'\nPlot graph')
    chart_types = {'Line' : pygal.Line(), 'Bar' : pygal.Bar() \
    , 'Gauge' : pygal.SolidGauge(half_pie=True, inner_radius=0.70, style=pygal.style.styles['default'](value_font_size=10))}
    chart_selector = input('"Line"\n"Bar"\n"Gauge"\nEnter : ')
    chart = chart_types[chart_selector]
    if chart_selector == 'Bar' or chart_selector == 'Line':
        chart.title = graph_name
        chart.x_labels = map(str, years)
        chart.add('Balchelor', data)
        chart.render_to_file('line.svg')
        command(input('"Compare"\n"Graph"\n"Two object compare"\n"Two object graph"\nEnter command : '))
    else:
        percent_formatter = lambda x: '{:.10g}%'.format(x)
        chart.value_formatter = percent_formatter
        chart.title = graph_name
        for i in range(len(years)):
            chart.add(str(years[i]), data[i])
        chart.render_to_file('line.svg')
        command(input('"Compare"\n"Graph"\n"Two object compare"\n"Two object graph"\nEnter command : '))

def two_object(status):
    '''Recieve command and do it follow the command.'''
    print('-' * 100+'\nTwo datas compare')
    if status == 'Two object compare':
        years, first_data, first_data_name = compare(status)
        years, second_data, second_data_name = compare(status)
    else:
        years, first_data, first_data_name = excel_reader(status)
        years, second_data, second_data_name = excel_reader(status)
    objects_plot_graph(years, first_data, second_data, first_data_name+' and '+second_data_name, first_data_name, second_data_name)

command(input('"Compare"\n"Graph"\n"Two object compare"\n"Two object graph"\nEnter command : '))

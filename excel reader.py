''''''
import xlrd
def excel_reader_reg(): #data manage of amount_of_all_register and amount_of_eligible_students for 2545 to 2559
    '''This function will read register_stat.
    Turn it to data for using as list.'''
    data = {}
    file_reg = 'reg' # Change string to your name's file.
    file_location = 'C:/Users/administrator_/Desktop/whatever-master/%s.xlsx' % file_reg # Change to your own file location.
    workbook = xlrd.open_workbook(file_location)
    sheet_count_reg = workbook.nsheets
    sheet = workbook.sheet_by_index(0)
    rows = sheet.nrows
    columns = sheet.ncols
    #print(columns) #<----- มีกี่คอลัมในชีท(แนวนอน)
    #print(rows) #<----- มีกี่แถวในชีท(แนวตั้ง)
    for row in range(rows):
        data_in_line = []
        if row == 6 or row == 10 or row == 1:
            for column in range(columns):
                if sheet.cell_value(row,column) == '' or sheet.cell_value(row,column) == '-':
                    data_in_line.append(0)
                elif 'float' in str(type(sheet.cell_value(row,column))):
                    data_in_line.append(int(sheet.cell_value(row,column)))
                else:
                    data_in_line.append(sheet.cell_value(row,column))
            if len(set(data_in_line)) > 1:
                data[row] = data_in_line
    for key in sorted(data.keys()):
        if key == 6:
            all_register = data[key][2:] # <------ ข้อมูลรวมของนักศึกษาที่สมัครสอบ (list)
        elif key == 10:
            amount_of_eligible_students = data[key][2:] # <------ ข้อมูลรวมของนักศึกษาที่มีสิทธิ์เข้าศึกษา (list)
        else:
            years_reg = [str(i) for i in data[key][2:]] # <------ ตั้งแต่ปี 2545-2559 (list)
    #print('จำนวนคนสมัครทั้งหมด :', all_register)
    #print('มีสิทธิ์เรียนที่นี่ทั้งหมด :', amount_of_eligible_students)
    #print('ปี :', years_reg)
def excel_reader_job(): 
    ''''''
    data = {}
    file_job = 'graduated_and_getwork' # Change string to your name's file.
    file_location = 'C:/Users/administrator_/Desktop/whatever-master/%s.xlsx' % file_job # Change to your own file location.
    workbook = xlrd.open_workbook(file_location)
    sheet_count_job = workbook.nsheets
    sheet = workbook.sheet_by_index(0)
    rows = sheet.nrows
    columns = sheet.ncols
    for i in range(sheet_count_job):
        sheet = workbook.sheet_by_index(i)
        for row in range(rows):
            data_in_line = []
            if row == 5 or row == 3:
                for column in range(columns-2):
                    if sheet.cell_value(row,column) == '' or sheet.cell_value(row,column) == '-':
                        data_in_line.append(0)
                    elif 'float' in str(type(sheet.cell_value(row,column))):
                        data_in_line.append(int(sheet.cell_value(row,column)))
                    else:
                        data_in_line.append(sheet.cell_value(row,column))
                if len(set(data_in_line)) > 1:
                    data[row+i] = data_in_line
    print(data)
excel_reader_job()

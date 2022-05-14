import PySimpleGUI as sg
import xlsxwriter



def generate_excel_worksheet(table_array, headings):
    # Add headings to begin
    table_array = [headings] + table_array
    print(table_array)
    workbook = xlsxwriter.Workbook('Contact.xlsx')
    worksheet = workbook.add_worksheet()

    row_index = 0
    column_index = 0

    for row in table_array:
        for column in row:
            worksheet.write(row_index, column_index, column)
            column_index +=1
        column_index = 0
        row_index +=1
    
    workbook.close()
    
    return row_index

def create(contact_information_array, headings):

    contact_information_window_layout = [
        [sg.Table(values=contact_information_array, headings=headings, max_col_width=35,
                    auto_size_columns=True,
                    display_row_numbers=True,
                    justification='right',
                    num_rows=10,
                    key='-TABLE-',
                    row_height=35,
                    tooltip='Reservations Table')],
        [sg.Button('Export to Excel Spreadsheet')]
    ]

    contact_information_window = sg.Window("Contact Information Window", 
    contact_information_window_layout, modal=True)


    while True:
        event, values = contact_information_window.read()
        if event == "Exit" or event == sg.WIN_CLOSED:
            break
        elif event == 'Export to Excel Spreadsheet':
            generate_excel_worksheet(contact_information_array, headings)
            
        
    contact_information_window.close()
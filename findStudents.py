from openpyxl import load_workbook

wb = load_workbook('cs106b.xlsx')
main_sheet = wb.get_sheet_by_name('Main')


# Sift students with section 'my_section' and returns their mails
def find_students(my_section):
    
    # Dictionary (student_name -> student_mail)
    my_students = {} 

    # Collect my students
    for i in range(1, 150):
        cur_section = main_sheet['G' + str(i)].value
        if cur_section != my_section:
            continue
        cur_name = main_sheet['D' + str(i)].value
        cur_mail = main_sheet['E' + str(i)].value
        my_students[cur_name] = cur_mail

    return my_students
    

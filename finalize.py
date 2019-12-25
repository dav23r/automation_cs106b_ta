import sys # for reading assing number from cli
import subprocess # for using shell
from findStudents import find_students # for retrieving my folks
from openpyxl import load_workbook # for manipulatin xlsx files

# Script reads data from grades file of corresponding assignment,
# fills excel file and send emails to students as well as to seminarists.


def parseMarking(path, my_students):
    
    def stripped(line):
        return line.strip()

    marks = ['Plus Plus', 'Plus', 'Check Plus',
             'Check', 'Check Minus',
             'Minus', 'Minus Minus', '0']

    marking = {}
    cur_comment = []
    final_mark = ''
    cur_name = ''

    with open(path, 'r') as input:
        for line in input:
            st = stripped(line)
            if st in my_students:
                if cur_name != '':
                    marking[cur_name] = (''.join(cur_comment), final_mark)
                cur_comment = []
                cur_name = st
            else:
                if 'წინასწარი შეფასება:' in st:
                    final_mark = stripped(st[st.index(':')+1:])
                    assert (final_mark in marks)
                cur_comment.append(line)

        marking[cur_name] = (''.join(cur_comment), final_mark)

    return marking
                

def sendMail(subject, addr, *content, attach = None):
    disclaimer = "Disclaimer: მეილი დაგენერირებულია ავტომატურად, შეიძლება ხარვეზები იყოს"
    mail_command = ["mail", "-s", subject, addr]
    if attach != None:
        mail_command[1:1] = ['-a', attach]
    p = subprocess.Popen(mail_command, stdin = subprocess.PIPE)
    content = '\n'.join(content).rstrip()
    content += '\n\n' + disclaimer
    print (addr + '\n' + content + 'END\n')
    p.communicate(bytearray(content.encode()))


def notifyStudents(grades, my_students, assignNum):
    subject = "პროგრამირების აბსტრაქციები"
    welcome = "სალამი, გიგზავნით მე-%d დავალების შენიშვნებს/კომენტარებს/რჩევებს." \
              "თუ რამე გაუგებარს ნახავთ ან სხვა ტიპის შეკითხვა გაგიჩნდებათ, feel free to ask." % (assignNum)
    for g in grades:
        sendMail(subject, my_students[g], welcome, grades[g][0])
     


def sendExcel(seminarists, excel_file, assignNum, my_section):
    subject = "აბსტრაქციების მე-%d გასწორებული დავალება, სექცია - %d" % (assignNum, my_section)
    welcome = ""
    for sem in seminarists:
        sendMail(subject, sem, welcome, attach = excel_file)


def updateExcel(grades, my_students, excel_file, assignNum):

    wb = load_workbook(excel_file)

    grades_sheet = wb.get_sheet_by_name('Results')
    main_sheet = wb.get_sheet_by_name('Main')

    for student in my_students:
        for i in range(1, 150):
            cur_val = grades_sheet['A' + str(i)].value
            if cur_val != None and main_sheet[cur_val[6:]].value == student:
                grades_sheet[chr(ord('B') + 2 * assignNum) 
                            + str(i)].value = grades[student][0]
                grades_sheet[chr(ord('B') + 1 + 2 * assignNum)
                            + str(i)].value = grades[student][1]

    wb.save('cs106b.xlsx')


def createBackup(excel_file):
    subprocess.call(['cp', excel_file, excel_file + '_backup.xlsx'])

    

if __name__ == '__main__':

    if len(sys.argv) != 2:
        print ("Provide assignment number as argument to script")
        exit()

    seminarists = ['g.bochorishvili@freeuni.edu.ge', 
                   'n.tsimakuridze@freeuni.edu.ge', 
                   'nbarb09@freeuni.edu.ge', 
                   'ndzam10@freeuni.edu.ge']

    my_section = 6
    assignNum = int(sys.argv[1])
    
    my_students = find_students(my_section) 

    # acquire dictionary (student name -> cooment string)
    grades = parseMarking('/home/dav23r/Desktop/cs106b/assign%d' % (assignNum) + '/grades',
                           my_students)

    excel_file = '/home/dav23r/Desktop/cs106b/cs106b.xlsx'
    createBackup(excel_file)    
    updateExcel(grades, my_students, excel_file, assignNum)
    notifyStudents(grades, my_students, assignNum)
    sendExcel(seminarists, excel_file, assignNum, my_section)
    

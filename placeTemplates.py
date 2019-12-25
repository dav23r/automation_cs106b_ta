from findStudents import find_students
import subprocess # for using shell

# Initialize template file and place in all assignment directories
my_students = find_students()
template = open('grades', 'w')
template.write('\n\n'.join(map(lambda x: x + '\nწინასწარი შეფასება:\n', my_students.keys())))

template.close()

subprocess.call(["find", ".", "-type", "d", "-name", "assign*", # find every directory whose name contains 'assign'
                 "!", "-exec", "test", "-e", "grades", ";",     # and doesn't already have 'grades' file in it
                "-exec", "cp", "grades", "{}", ";"])            # and copy grades file to each of them.



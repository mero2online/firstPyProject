from tkinter import filedialog
from tkinter import messagebox
from tkinter import *
from tkcalendar import Calendar
from datetime import datetime
import os

destination_path = ''
source_path = ''

currentDay = datetime.now().day
currentMonth = datetime.now().month
currentYear = datetime.now().year

cwd = os.getcwd()
# runBatFile = os.system(f'{cwd}\\backup.bat')


def createTask():
    createTaskCommand = f'schtasks /create /tn Elevated-WellBackup /tr {cwd}\\backup.bat /sc daily /sd {startDateValue} /st {hour_text.get()}:{minute_text.get()} /ed {endDateValue} > out.txt'
    print(createTaskCommand)
    # os.system(createTaskCommand)
    # f = open('out.txt', 'r')
    # out = f.read()
    # f.close()
    # return f'{out} Created'


def getAllBackupTasks():
    commandToFindTask = 'schtasks /query /fo LIST /v | findstr "Elevated" > out.txt'
    os.system(f'{commandToFindTask}')
    f = open('out.txt', 'r')
    tasks = f.read()
    f.close()
    tasksNames = tasks.replace(
        "TaskName:                             \\", '').split("\n")
    del tasksNames[-1]
    # print(tasksNames)
    # afterDelete = deleteTask(tasksNames[1])
    # print(afterDelete)
    return tasksNames


def deleteTask(taskName):
    os.system(f'schtasks /delete /tn {taskName} /f > out.txt')
    f = open('out.txt', 'r')
    out = f.read()
    f.close()
    return f'{out} Deleted'


def runCommand():
    command = f'Xcopy "{source_path}" "{destination_path}/{source_path.split("/")[-1]}" /E /H /C /I > out.txt'
    saveCommand = open("backup.bat", "w")
    saveCommand.write(command)
    saveCommand.close()

    # open and read the file after the appending:
    # f = open('out.txt', 'r')
    # print(f.read())
    # f.close()
    # os.remove('out.txt')


def browse_source_button():
    global source_path
    source_path = filedialog.askdirectory()

    if(source_path == destination_path):
        source_path
        messagebox.showerror(
            'Error', 'Source path can`t be same as destination path')
        return source_label.config(text='Choose another path')
    else:
        return source_label.config(text=source_path)


def browse_destination_button():
    global destination_path
    destination_path = filedialog.askdirectory()
    if(destination_path == source_path):
        destination_path
        messagebox.showerror(
            'Error', 'Destination path can`t be same as source path')
        return destination_label.config(text='Choose another path')
    else:
        return destination_label.config(text=destination_path)


# Create window object
app = Tk()

# Buttons

browse_source_btn = Button(
    app, text='Browse Source Folder', background='#A3E4DB', command=browse_source_button)
browse_source_btn.grid(row=0, column=0, pady=20, padx=20)

browse_destination_btn = Button(
    app, text='Browse Destination Folder', background='#A3E4DB', command=browse_destination_button)
browse_destination_btn.grid(row=0, column=1, pady=20, padx=20)

source_label = Label(app, text='Source Path')
source_label.grid(row=1, column=0)

destination_label = Label(app, text='Destination Path')
destination_label.grid(row=1, column=1)

# Add Calendar Start
calStart = Calendar(app, selectmode='day',
                    year=currentYear, month=currentMonth,
                    day=currentDay, date_pattern='mm/dd/yyyy')
calStart.grid(pady=5, row=2, column=0)


def grad_start_date():
    global startDateValue
    startDateValue = calStart.get_date()
    startDate.config(text="Start Date is: " + calStart.get_date())


# Add Calendar End
calEnd = Calendar(app, selectmode='day',
                  year=currentYear, month=currentMonth,
                  day=currentDay+1, date_pattern='mm/dd/yyyy')
calEnd.grid(pady=5, row=2, column=1)


def grad_end_date():
    global endDateValue
    endDateValue = calEnd.get_date()
    endDate.config(text="End Date is: " + calEnd.get_date())

# Hour Minute Limit


def limitSizeHour(*args):
    value = hour_text.get()
    if len(value) > 2:
        hour_text.set(value[:2])


def limitSizeMinute(*args):
    value = minute_text.get()
    if len(value) > 2:
        minute_text.set(value[:2])


# Hour
hour_text = StringVar(app, value='11')
hour_text.trace('w', limitSizeHour)
hour_label = Label(app, text='hour', font=('bold', 14))
hour_label.grid(row=3, column=0, sticky=E)
hour_entry = Entry(app, textvariable=hour_text, width=2)
hour_entry.grid(row=4, column=0, pady=5, sticky=E)
# Minute
minute_text = StringVar(app, value='00')
minute_text.trace('w', limitSizeMinute)
minute_label = Label(app, text='minute', font=('bold', 14))
minute_label.grid(row=3, column=1, sticky=W)
minute_entry = Entry(app, textvariable=minute_text, width=2)
minute_entry.grid(row=4, column=1, pady=5, sticky=W)

# Add Button and Label
Button(app, text="Set Start Date",
       command=grad_start_date).grid(row=6, column=0, pady=5)

startDate = Label(app, text="Start Date")
startDate.grid(row=5, column=0, pady=5)

# Add Button and Label
Button(app, text="Set End Date",
       command=grad_end_date).grid(row=6, column=1, pady=5)

endDate = Label(app, text="End Date")
endDate.grid(row=5, column=1, pady=5)

run_command_btn = Button(app, text='Run command',
                         background='#A3E4DB', command=createTask)
run_command_btn.grid(pady=20, sticky=E)

app.title('Command Runner')
app.geometry('700x600')
app.configure(bg='#000')

# To center all app columns
app.grid_columnconfigure((0, 1), weight=1)

p1 = PhotoImage(file='backup_icon.png')

# Setting icon of master window
app.iconphoto(False, p1)

# Start program
app.mainloop()

from tkinter import filedialog
from tkinter import messagebox
from tkinter import *
from tkcalendar import Calendar
from datetime import datetime
import os

destination_path = ''
source_path = ''
selected_item = ''
startDateValue = ''
endDateValue = ''

currentDay = datetime.now().day
currentMonth = datetime.now().month
currentYear = datetime.now().year

cwd = os.getcwd()
# runBatFile = os.system(f'{cwd}\\backup.bat')


def createTask():
    if(startDateValue == ''):
        messagebox.showerror('Error', 'Start Date can`t be empty')
        return False
    elif(endDateValue == ''):
        messagebox.showerror('Error', 'End Date can`t be empty')
        return False
    time = f'{hour_text.get()}:{minute_text.get()}'
    taskName = f'{source_path.split("/")[-1]}-Backup-Elevated'
    batchFileName = f'{source_path.split("/")[-1]}-Backup.bat'
    # createTaskCommand = f'schtasks /create /tn {source_path.split("/")[-1]}-Backup-Elevated /tr {cwd}\\{source_path.split("/")[-1]}-Backup.bat /sc daily /sd {startDateValue} /st {hour_text.get()}:{minute_text.get()} /ed {endDateValue} > out.txt'
    createTaskCommand = f'powershell $Action= New-ScheduledTaskAction -Execute "cmd.exe" -Argument \'/c start /min "" "{cwd}\\{batchFileName}"\'; $Stt = New-ScheduledTaskTrigger -Daily -At {time}; Register-ScheduledTask -TaskName "{taskName}" -Action $Action -Trigger $Stt; $TargetTask = Get-ScheduledTask -TaskName "{taskName}"; $endTime=([DateTime]\'{endDateValue} {time}\').ToString(\'yyyy-MM-dd"T"HH:mm:ss\'); $TargetTask.Triggers[0].EndBoundary=$endTime; Set-ScheduledTask -InputObject $TargetTask;'

    print(createTaskCommand)
    os.system(createTaskCommand)
    f = open('out.txt', 'r')
    out = f.read()
    f.close()
    return f'{out} Created'


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


def deleteTask():
    os.system(f'schtasks /delete /tn {selected_item} /f > out.txt')
    f = open('out.txt', 'r')
    out = f.read()
    f.close()
    populate_list()
    os.remove(f'{"-".join(selected_item.split("-")[:-1])}.bat')
    messagebox.showinfo('Info', f'Task {selected_item} deleted successfully')
    return f'{out} Deleted'


def saveCommand():
    if(source_path == ''):
        messagebox.showerror('Error', 'Source path can`t be empty')
        return False
    elif(destination_path == ''):
        messagebox.showerror('Error', 'Destination path can`t be empty')
        return False
    source = source_path.replace("/", "\\")
    dest = destination_path.replace("/", "\\")
    destFolder = source_path.split("/")[-1]
    command = f'md "{dest}\Folder"\nXcopy "{source}" "{dest}\Folder\{destFolder}" /E /H /C /I > {cwd}\out.txt\nREN "{dest}\Folder" "%date:~-4,4%"-"%date:~-10,2%"-"%date:~7,2% %Time::=.%"\nexit'
    saveCommand = open(
        f'{source_path.split("/")[-1]}-Backup.bat', "w")
    saveCommand.write(command)
    saveCommand.close()

    # open and read the file after the appending:
    # f = open('out.txt', 'r')
    # print(f.read())
    # f.close()
    # os.remove('out.txt')


def addTask():
    if(saveCommand() == False):
        return
    if(createTask() == False):
        return
    populate_list()
    messagebox.showinfo(
        'Info', f'Task {source_path.split("/")[-1]}-Backup-Elevated created successfully')


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


def populate_list():
    os.remove('out.txt')
    tasks_list.delete(0, END)
    tasks = getAllBackupTasks()
    for task in tasks:
        tasks_list.insert(END, task)


def select_item(event):
    try:
        global selected_item
        index = tasks_list.curselection()
        selected_item = tasks_list.get(index)
        print(selected_item)
    except IndexError:
        pass


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

tasks_list = Listbox(app, height=8, width=50, border=0)
tasks_list.grid(row=7, column=0, columnspan=3, rowspan=6, pady=20, padx=20)

Button(app, text="Delete Task",
       command=deleteTask).grid(row=13, column=0, pady=5)

run_command_btn = Button(app, text='Add Task',
                         background='#A3E4DB', command=addTask)
run_command_btn.grid(pady=20, sticky=E)

populate_list()
tasks_list.bind('<<ListboxSelect>>', select_item)

app.title('Command Runner')
app.geometry('700x700')
app.configure(bg='#000')

# To center all app columns
app.grid_columnconfigure((0, 1), weight=1)

p1 = PhotoImage(file='backup_icon.png')

# Setting icon of master window
app.iconphoto(False, p1)

# Start program
app.mainloop()

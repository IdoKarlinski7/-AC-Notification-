import datetime
import win32com.client

TASK_ACTION_EXEC = 0
TASK_TRIGGER_TIME = 1
TRIGGER_DAILY = 2

# Create scheduler instance
scheduler = win32com.client.Dispatch('Schedule.Service')
scheduler.Connect()
root_folder = scheduler.GetFolder('\\')
print(type(scheduler))
task_def = scheduler.NewTask(0)
print(type(task_def))

# Create trigger for daily task
start_time = start_time = datetime.datetime.now() + datetime.timedelta(minutes=1)
trigger = task_def.Triggers.Create(TRIGGER_DAILY)
trigger.Repetition.Duration = "P30D"
trigger.Repetition.Interval = "PT1D"
trigger.StartBoundary = start_time.isoformat()
trigger.Id = "DailyTrigger"
trigger.Enabled = True

# Create daily action
get_action = task_def.Actions.Create(TASK_ACTION_EXEC)
get_action.ID = 'GET action (Daily)'
get_action.Path = 'C:\pythonProject\Notification Sys\daily_exec\daily_exec.exe'
get_action.Arguments = ''

# Set parameters
task_def.RegistrationInfo.Description = 'GET action (Daily)'
task_def.Settings.Enabled = True
task_def.Settings.StopIfGoingOnBatteries = False

# Register task. If task already exists, it will be updated
TASK_CREATE_OR_UPDATE = 6
TASK_LOGON_NONE = 0
root_folder.RegisterTaskDefinition('GET action (Daily)', task_def, TASK_CREATE_OR_UPDATE, '', '', TASK_LOGON_NONE)

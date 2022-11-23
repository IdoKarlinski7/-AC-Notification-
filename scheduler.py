import datetime
import win32com.client



TASK_ACTION_EXEC = 0
TASK_TRIGGER_TIME = 1

scheduler = win32com.client.Dispatch('Schedule.Service')
scheduler.Connect()
root_folder = scheduler.GetFolder('\\')
task_def = scheduler.NewTask(0)

# Create trigger
start_time = datetime.datetime.now() + datetime.timedelta(seconds=30)
trigger = task_def.Triggers.Create(TASK_TRIGGER_TIME)
trigger.StartBoundary = start_time.isoformat()

# Create triggers for notifications
first_time = datetime.datetime.now() + datetime.timedelta(days=152)  # five months from calibration
trigger_first = task_def.Triggers.Create(TASK_TRIGGER_TIME)
trigger_first.Id = "MonthAheadTrigger"
trigger_first.StartBoundary = first_time.isoformat()

second_time = first_time + datetime.timedelta(days=14)  # five months and two weeks from calibration
trigger_second = task_def.Triggers.Create(TASK_TRIGGER_TIME)
trigger_second.Id = "TwoWeeksAheadTrigger"
trigger_second.StartBoundary = second_time.isoformat()

# Create action
action = task_def.Actions.Create(TASK_ACTION_EXEC)
action.ID = 'Test Event'
action.Path = 'C:/pythonProject/designer.exe'
action.Arguments = ''

# Set parameters
task_def.RegistrationInfo.Description = 'Test Task'
task_def.Settings.Enabled = True
task_def.Settings.StopIfGoingOnBatteries = False

# Register task
# If task already exists, it will be updated
TASK_CREATE_OR_UPDATE = 6
TASK_LOGON_NONE = 0
root_folder.RegisterTaskDefinition('Test Event', task_def, TASK_CREATE_OR_UPDATE, '', '', TASK_LOGON_NONE)

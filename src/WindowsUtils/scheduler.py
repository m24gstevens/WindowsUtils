"""
scheduler.py
====================================
Methods scheduling a task for use by Windows Task Scheduler
"""

from comtypes.client import CreateObject
from comtypes.automation import VARIANT
from WindowsUtils.helpers.schedtypes import *
import datetime

SUNDAY = 0x1
MONDAY = 0x2
TUESDAY = 0x4
WEDNESDAY = 0x8
THURSDAY = 0x10
FRIDAY = 0x20
SATURDAY = 0x40

class TaskFactory:
    """
    Class for use in creating tasks

    A task should be created in the following steps:
    First initialize a TaskFactory, then call these methods in order:
    - configure_task_settings
    - configure_***_trigger
    where *** is one of logon, daily, weekly, or time
    - configure_program
    - register_task

    Once a task has been registered, you should delete the class instance.
    Failure of any of the methods means you should start afresh with
    a new instance
    """
    def __init__(self, task_name, author_name = "WindowsUtils"):
        """Initializes name and folder information

        Parameters
        ----------
        task_name : str
            name of the task
        author_name : str, optional
            name of the task author. Defaults to 'WindowsUtils'
        """
        CLSID = GUID("{0F87369F-A4E5-4CFC-BD3E-73E6154572DD}")
        servicer = CreateObject(CLSID, interface = ITaskService)
        if servicer.Connect(VARIANT(), VARIANT(),
                            VARIANT(), VARIANT()):
            raise Exception("Couldn't connect to the scheduler")
        try:
            self.name = task_name
            self.root_folder = servicer.GetFolder("")
        except Exception:
            raise Exception("Couldn't get the tasks root folder")
        else:
            try:
                self.task = servicer.NewTask(0)
            except Exception:
                raise Exception("Couldn't create a new task instance")
            else:
                try:
                    reg_info = self.task.get_RegistrationInfo()
                except Exception:
                    raise Exception("Can't put identification info")
                else:
                    try:
                        reg_info = self.task.get_RegistrationInfo()
                        reg_info.put_Author(author_name)
                    except Exception:
                        raise Exception("Can't put identification info")
                    else:
                        try:
                            principal = self.task.get_Principal()
                            principal.put_LogonType(LOGON_INTERACTIVE_TOKEN)
                        except Exception:
                            raise Exception("Can't put principal info")

    def configure_task_settings(self,
                                start_when_available = None, if_network_available = None,
                                stop_on_batteries = None, wake_to_run = None):
        """Configures settings which may start or stop task from running

        Parameters
        ----------
        start_when_available : bool, optional
            True indicates that the task should start as soon as it can,
            up to an unpredictable time delay.
        if_network_available : bool, optional
            True indicates that the task will only run if network connection is
            available
        stop_on_batteries : bool, optional
            True indicates that the task will stop if on battery power
        wake_to_run : bool, optional
            True indicates that the task will wake up the computer when it runs
        """
        try:
            task_settings = self.task.get_Settings()
            if not (start_when_available is None):
                val = VARIANT_TRUE if start_when_available else VARIANT_FALSE
                task_settings.put_StartWhenAvailable(val)
            if not (if_network_available is None):
                val = VARIANT_TRUE if network_available else VARIANT_FALSE
                task_settings.put_RunOnlyIfNetworkAvailable(val)
            if not (stop_on_batteries is None):
                val = VARIANT_TRUE if stop_on_batteries else VARIANT_FALSE
                task_settings.put_DisallowStartIfOnBatteries(val)
            if not (wake_to_run is None):
                val = VARIANT_TRUE if wake_to_run else VARIANT_FALSE
                task_settings.put_WakeToRun(val)
        except Exception:
            raise Exception("Couldn't configure the settings")

    def _start_config_trigger(self, trigger_code, interface,
                              start_boundary = datetime.datetime.now(),
                              end_boundary = None):
        formatting_string = "%Y-%m-%dT%H:%M:%S"
        try:
            trig_collection = self.task.get_Triggers()
            trigger = trig_collection.Create(trigger_code)
            custom_trigger = trigger.QueryInterface(interface)
            custom_trigger.put_Id("Trigger1")
            if start_boundary:
                start_date_string = start_boundary.strftime(formatting_string)
                custom_trigger.put_StartBoundary(start_date_string)
            if end_boundary:
                end_date_string = end_boundary.strftime(formatting_string)
                custom_trigger.put_EndBoundary(end_date_string)
        except Exception:
            raise Exception("Couldn't set trigger start and end")
        else:
            return custom_trigger

    def configure_weekly_trigger(self, start_boundary = datetime.datetime.now(),
                                 end_boundary = None,
                                 weekday_mask = 0x7F, interval = 1):
        """
        Configures a weekly trigger on a task.

        Parameters:
        -----------
        start_boundary : datetime, optional
            a datetime object indicating the first time a task can run.
            defaults to the current time
        end_boundary : datetime, optional
            a datetime object indicating the last time a task can run
        weekday_mask : int, optional
            bitmask of the days of the week the task will run.
            bitpatterns contained in SUNDAY, MONDAY, ..., SATURDAY variables
            defaults to every possible day
        interval : int, optional
            number of weeks in between each weekly run.
            defaults to 1
        """
        weekly_trigger = self._start_config_trigger(TRIGGER_WEEKLY, IWeeklyTrigger,
                                                    start_boundary,
                                                    end_boundary)
        try:
            weekly_trigger.put_WeeksInterval(interval)
            weekly_trigger.put_DaysOfWeek(weekday_mask)
        except Exception:
            raise Exception("Couldn't set a weekly trigger")

    def configure_daily_trigger(self, start_boundary = datetime.datetime.now(),
                                end_boundary = None, interval = 1):
        """
        Configures a daily trigger on a task.

        Parameters:
        -----------
        start_boundary : datetime, optional
            a datetime object indicating the first time a task can run.
            defaults to the current time
        end_boundary : datetime, optional
            a datetime object indicating the last time a task can run
        interval : int, optional
            number of days in between each run.
            defaults to 1
        """
        daily_trigger = self._start_config_trigger(TRIGGER_DAILY, IDailyTrigger,
                                                    start_boundary,
                                                    end_boundary)
        try:
            daily_trigger.put_DaysInterval(interval)
        except Exception:
            raise Exception("Couldn't set a daily trigger")

    def configure_logon_trigger(self, start_boundary = datetime.datetime.now(),
                                end_boundary = None):
        """
        Configures a logon trigger on a task.

        Currently, the task will run on logon for any user
        
        Parameters:
        -----------
        start_boundary : datetime, optional
            a datetime object indicating the first time a task can run.
            defaults to the current time
        end_boundary : datetime, optional
            a datetime object indicating the last time a task can run
        """
        logon_trigger = self._start_config_trigger(TRIGGER_LOGON, ILogonTrigger,
                                                    start_boundary,
                                                    end_boundary)

    def configure_time_trigger(self, start_boundary = datetime.datetime.now(),
                               end_boundary = None):
        """
        Configures a time trigger on a task.

        Parameters:
        -----------
        start_boundary : datetime, optional
            a datetime object indicating the first time a task can run.
            defaults to the current time
        end_boundary : datetime, optional
            a datetime object indicating the last time a task can run
        """
        time_trigger = self._start_config_trigger(TRIGGER_TIME, ITimeTrigger,
                                                  start_boundary, end_boundary)


    def configure_program(self, executable, args = None):
        """Configures an action for our task in form of running an executable

        Parameters
        ----------
        executable : str
            absolute path to the executable the task will run
        args : str, optional
            command line arguments to pass to the process
        """
        try:
            action_collection = self.task.get_Actions()
            action = action_collection.Create(ACTION_EXEC)
            action.put_Id("Action1")
            exec_action = action.QueryInterface(IExecAction)
        except Exception:
            raise Exception("Couldn't configure an action")
        else:
            try:
                exec_action.put_Path(executable)
                if args:
                    exec_action.put_Arguments(args)
            except Exception:
                raise Exception("Couldn't associate process with the task")

    def register_task(self):
        """Registers the task in the root task registration folder, and
        closes the task creation session
        """
        try:
            registered_task = self.root_folder.RegisterTaskDefinition("Test",
                                                                      self.task,
                                                                      6,
                                                                      VARIANT(),
                                                                      VARIANT(),
                                                                      0,
                                                                      VARIANT(""))
        except Exception:
            raise Exception("Error saving the task")

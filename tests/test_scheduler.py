import unittest
import os
from WindowsUtils import scheduler
from WindowsUtils.helpers.schedtypes import (ITaskDefinition, ITaskFolder,
                                             ITaskSettings, ITrigger,
                                             IAction, IExecAction,
                                             IWeeklyTrigger)
from comtypes import POINTER
from comtypes.automation import VARIANT
import datetime

notepad_path = os.environ['WINDIR'] + "\\system32\\notepad.exe"
bat_path = os.path.normpath(os.path.join(__file__, "..\\bat_test.bat"))

class Test(unittest.TestCase):
    def setUp(self):
        self.tf = scheduler.TaskFactory("WindowsUtils test")

    def test_logon_trigger(self):
        self.assertIsInstance(self.tf.task, POINTER(ITaskDefinition))
        self.assertIsInstance(self.tf.root_folder, POINTER(ITaskFolder))
        self.tf.configure_task_settings(start_when_available = True,
                                        stop_on_batteries = False)
        settings = self.tf.task.get_Settings()
        self.assertIsInstance(settings, POINTER(ITaskSettings))
        self.assertTrue(bool(settings.get_StartWhenAvailable()))
        self.assertFalse(bool(settings.get_DisallowStartIfOnBatteries()))

        start_dt = datetime.datetime.now() + datetime.timedelta(minutes = 3)
        end_dt = start_dt + datetime.timedelta(days = 1)
        self.tf.configure_logon_trigger(start_dt, end_dt)
        trig_collection = self.tf.task.get_Triggers()
        ct = trig_collection.get_Count()
        self.assertGreater(ct, 0)
        while (ct > 0):
            trig = trig_collection.get_Item(ct)
            self.assertIsInstance(trig, POINTER(ITrigger))
            trig_id = trig.get_Id()
            if trig_id == "Trigger1":
                self.assertEqual(trig.get_Type(), 9)    #Logon type
                break
            ct -= 1

        self.tf.configure_program(notepad_path)
        act_collection = self.tf.task.get_Actions()
        ct = act_collection.get_Count()
        self.assertGreater(ct, 0)
        while (ct > 0):
            act = act_collection.get_Item(ct)
            self.assertIsInstance(act, POINTER(IAction))
            act_id = act.get_Id()
            if act_id == "Action1":
                self.assertEqual(act.get_Type(), 0)   #Exec type
                exec_action = act.QueryInterface(IExecAction)
                self.assertEqual(exec_action.get_Path(), notepad_path)
                break
            ct -= 1

        self.assertIsNone(self.tf.register_task())
        
    def test_weekly_trigger(self):
        self.assertIsInstance(self.tf.task, POINTER(ITaskDefinition))
        self.assertIsInstance(self.tf.root_folder, POINTER(ITaskFolder))
        self.tf.configure_task_settings(start_when_available = True,
                                        stop_on_batteries = False)
        settings = self.tf.task.get_Settings()
        self.assertIsInstance(settings, POINTER(ITaskSettings))
        self.assertTrue(bool(settings.get_StartWhenAvailable()))
        self.assertFalse(bool(settings.get_DisallowStartIfOnBatteries()))

        start_dt = datetime.datetime.now() + datetime.timedelta(minutes = 3)
        end_dt = start_dt + datetime.timedelta(days = 1)
        self.tf.configure_weekly_trigger(start_dt, end_dt)
        trig_collection = self.tf.task.get_Triggers()
        ct = trig_collection.get_Count()
        self.assertGreater(ct, 0)
        while (ct > 0):
            trig = trig_collection.get_Item(ct)
            self.assertIsInstance(trig, POINTER(ITrigger))
            trig_id = trig.get_Id()
            if trig_id == "Trigger1":
                self.assertEqual(trig.get_Type(), 3)    #Weekly type
                weekly_trig = trig.QueryInterface(IWeeklyTrigger)
                self.assertIsInstance(weekly_trig, POINTER(IWeeklyTrigger))
                self.assertEqual(weekly_trig.get_DaysOfWeek(), 0x7F)
                self.assertEqual(weekly_trig.get_WeeksInterval(), 1)
                break
            ct -= 1

        self.tf.configure_program(notepad_path, bat_path)
        act_collection = self.tf.task.get_Actions()
        ct = act_collection.get_Count()
        self.assertGreater(ct, 0)
        while (ct > 0):
            act = act_collection.get_Item(ct)
            self.assertIsInstance(act, POINTER(IAction))
            act_id = act.get_Id()
            if act_id == "Action1":
                self.assertEqual(act.get_Type(), 0)   #Exec type
                exec_action = act.QueryInterface(IExecAction)
                self.assertEqual(exec_action.get_Path(), notepad_path)
                self.assertEqual(exec_action.get_Arguments(), bat_path)
                break
            ct -= 1

        self.assertIsNone(self.tf.register_task())


if __name__ == "__main__":
    unittest.main()
    

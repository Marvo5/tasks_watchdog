"""This script is writen to automate task scheduler verification after reboot"""

from __future__ import annotations
from datetime import datetime
import zoneinfo
import win32com.client
import colorama
from variables import LAST_BOOT, TRIGGER_TYPE_NAME, error_dict
from colour_print import colour_print, GREEN, RED, BLACK, BOLD

class Task:
    """Class to represent a task.
    """
    def __init__(self, task: object) -> None:
        self._name = task.Name
        self._path = task.Path
        self._enabled = task.Enabled
        self._last_run = task.LastRunTime.replace(tzinfo=zoneinfo.ZoneInfo("Europe/Warsaw")) # replace necessary due to wrong tzinfo from win32api
        self._next_run = task.NextRunTime.replace(tzinfo=zoneinfo.ZoneInfo("Europe/Warsaw")) # replace necessary due to wrong tzinfo from win32api
        self._last_result = task.LastTaskResult
        self.triggers = []
        self.actions = []
            
    def add_trigger(self, trigger: Trigger):
        """Add a new trigger to the list"""
        self.triggers.append(trigger)

    def add_action(self, action: Action):
        """Add a new action to the list"""
        self.actions.append(action)

    def get_result(self):
        return f"{self._last_result:#010x}"
    
    def health_check(self):
        if (self._last_result == 0x0 or self._last_result ==  0x41301) and self._last_run > LAST_BOOT:
            return True
        return False
        # TODO: Add more logic for task run succesfully after reboot

    def print_func(self):
        print(f"Name        : {self._name}")
        print(f"Path        : {self._path}")
        print(f"Last Run    : {self._last_run}")
        print(f"Next Run    : {self._next_run}")
        try: error_msg = error_dict[self._last_result]
        except KeyError: error_msg = "UNKNOWN ERROR CODE"
        print(f"Last Result : {self._last_result:#07x} {error_msg}")
        for action in self.actions:
            action.print_func()
        for trigger in self.triggers:
            trigger.print_func()
        if self._enabled:
            if self.health_check():
                colour_print("OK", GREEN)
            else:
                colour_print("FAIL", RED)
        else:
            print("DISABLED")
        print()

    last_result = property(get_result)


class Trigger:
    """Class to represent a trigger.
    """
    def __init__(self, trigger: object) -> None:
        self._type = trigger.Type
        self._enabled = trigger.Enabled
        self._execution_time_limit = trigger.ExecutionTimeLimit
        self._id = trigger.Id
        self._repetition = (trigger.Repetition.Duration,
                            trigger.Repetition.Interval,
                            trigger.Repetition.StopAtDurationEnd)
        if trigger.StartBoundary:
            self._start_boundary = datetime.fromisoformat(trigger.StartBoundary).astimezone()
        if trigger.EndBoundary:
            self._end_boundary = datetime.fromisoformat(trigger.EndBoundary).astimezone()
    def get_type_name(self):
        return TRIGGER_TYPE_NAME[self._type]

    _type_name = property(get_type_name)

    def print_func(self):
        if self._enabled:
            colour_print(f"Trigger     : {self._type_name}")
        else:
            colour_print(f"Trigger     : {self._type_name} [DISABLED]", BLACK)

class Action:
    """Class to represent an action.
    """
    def __init__(self, action: object) -> None:
        self._type = action.Type
        self._id = action.Id
        self._data = self.type_data(action)

    def type_data(self, action):
        if self._type == win32com.client.constants.TASK_ACTION_COM_HANDLER:  # =5
            coma = win32com.client.CastTo(action, "IComHandlerAction")
            return("COM Handler Action", (coma.ClassId, coma.Data))
        if self._type == win32com.client.constants.TASK_ACTION_EXEC:  # =0
            execa = win32com.client.CastTo(action, "IExecAction")
            return("Exec Action", (execa.Path, execa.Arguments))
        if self._type == win32com.client.constants.TASK_ACTION_SEND_EMAIL:  # =6
            maila = win32com.client.CastTo(action, "IEmailAction")
            return("Send Email Action", (maila.Subject, maila.To))
        if self._type == win32com.client.constants.TASK_ACTION_SHOW_MESSAGE:  # =7
            showa = win32com.client.CastTo(action, "IShowMessageAction")
            return("Show Message Action", (showa.Title, showa.MessageBody))
        
    def print_func(self):
        print(f"Action      : {self._data[0]} | {self._data[1]}")


def load_tasks() -> list:
    task_list = []
    scheduler = win32com.client.gencache.EnsureDispatch("Schedule.Service")
    scheduler.Connect()
    folders =[]
    with open("config.cfg", encoding="utf-8") as config_file:
        for line in config_file:
            if line[0] == "#":
                continue
            line = line.rstrip()
            folders += [scheduler.GetFolder(line)]
    if len(folders) == 0:
        folders = [scheduler.GetFolder("\\")]
    while folders:
        folder = folders.pop(0)
        # add new folders recursive
        folders += list(folder.GetFolders(0))
        for task in folder.GetTasks(0):
            new_task = Task(task)
            task_list.append(new_task)
            # Add all created triggers to Task object.
            triggers = task.Definition.Triggers
            for trigger in triggers:
                new_trigger = Trigger(trigger)
                new_task.add_trigger(new_trigger)
            # Add all created actions to Task object
            actions = task.Definition.Actions
            for action in actions:
                new_action = Action(action)
                new_task.add_action(new_action)
    return task_list


if __name__ == "__main__":
    tasks_list = load_tasks()
    counter = [0 , 0]
    colorama.init()
    for task in tasks_list:
        if task._enabled:
            if task.health_check():
                counter[0] += 1
            else:
                counter[1] += 1
        task.print_func()

    print("-"*80, "\n")
    if counter == [0, 0]:
        colour_print("All task(s) DISABLED", BOLD)
    elif counter[1] == 0:
        colour_print(f"{counter[0]} task(s) started succesfully after reboot.", GREEN, BOLD)
    else:
        colour_print(f"{counter[1]} task(s) failed to start after reboot. {counter[0]} task(s) run successfully.", RED, BOLD)
    colorama.deinit()

    input()
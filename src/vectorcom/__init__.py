import logging
from time import sleep
from typing import Callable, ClassVar, Optional

import pythoncom
import win32com.client
from win32com.client.dynamic import CDispatch

from vectorcom.configuration import PLPath

from .configuration import Configuration
from .version import Version

LOG = logging.getLogger("VectorCOM")


class Canoe:
    class _Events:
        OnOpenCbk: ClassVar[Callable[str, None]] = lambda fullname: LOG.debug(
            "Configuration opened: %s", fullname
        )
        OnQuitCbk: ClassVar[Callable[..., None]] = lambda: LOG.debug("Closing CANoe ...")

        OnOpenPending: bool = False
        OnQuitPending: bool = False

        def OnOpen(self, fullname: str):
            self.__class__.OnOpenCbk(fullname)
            self.OnOpenPending = False

        def OnQuit(self):
            self.__class__.OnQuitCbk()
            self.OnQuitPending = False

        def waitOnOpen(self):
            while self.OnOpenPending:
                pythoncom.PumpWaitingMessages()
                sleep(0.1)

        def waitOnQuit(self):
            while self.OnQuitPending:
                pythoncom.PumpWaitingMessages()
                sleep(0.1)

    _COM: ClassVar[CDispatch]

    @property
    def FullName(self) -> str:
        return self._COM.FullName

    @property
    def Name(self) -> str:
        return self._COM.Name

    @property
    def Path(self) -> PLPath:
        return PLPath(self._COM.Path)

    @property
    def Visible(self) -> bool:
        return self._COM.Visible

    @Visible.setter
    def Visible(self, value: bool) -> None:
        self._COM.Visible = value

    @property
    def Version(self) -> Version:
        return Version(self._COM.Version)

    @property
    def Configuration(self) -> Configuration:
        return Configuration(self._COM.Configuration)

    def Open(
        self,
        path: PLPath,
        autoSave: Optional[bool] = None,
        promptUser: Optional[bool] = None,
    ) -> None:
        self._Events.OnOpenPending = True
        if autoSave and promptUser:
            self._COM.Open(path, autoSave, promptUser)
        elif autoSave:
            self._COM.Open(path, autoSave)
        else:
            self._COM.Open(path)
        self._COM.waitOnOpen()

    def Quit(self) -> None:
        self._Events.OnQuitPending = True
        self._COM.Quit()
        self._COM.waitOnQuit()

    def __init__(self) -> None:
        LOG.info("Opening CANoe application...")
        self.__class__._COM = win32com.client.DispatchWithEvents(
            "CANoe.Application", self._Events
        )

    def __rich_repr__(self):
        yield "FullName", self.FullName
        yield "Name", self.Name
        yield "Path", self.Path
        yield "Visible", self.Visible
        yield self.Version
        yield self.Configuration

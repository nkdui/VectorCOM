import logging
from types import NotImplementedType
from typing import Callable, ClassVar, Optional

import rich.repr
import win32com.client
from win32com.client.dynamic import CDispatch

from vectorcom.configuration import PLPath

from .common import RefBool, waitEventFinished
from .configuration import Configuration
from .version import Version

LOG = logging.getLogger("VectorCOM")


@rich.repr.auto
class Canoe:
    class _Events:
        OnOpenCbk: ClassVar[Callable[[str], None]] = lambda fullname: LOG.debug(
            "Opened CANoe configuration file: '%s'", fullname
        )
        OnQuitCbk: ClassVar[Callable[..., None]] = lambda: LOG.debug(
            "Quitting CANoe ..."
        )
        OnOpenFinished: ClassVar[RefBool] = RefBool(True)
        OnQuitFinished: ClassVar[RefBool] = RefBool(True)

        @classmethod
        def OnOpen(cls, fullname: str):
            cls.OnOpenCbk(fullname)
            cls.OnOpenFinished.true

        @classmethod
        def OnQuit(cls):
            cls.OnQuitCbk()
            cls.OnQuitFinished.true

    _COM: ClassVar[CDispatch]

    @property
    def Bus(self) -> NotImplementedType:
        return NotImplemented

    @property
    def CAPL(self) -> NotImplementedType:
        return NotImplemented

    @property
    def ChannelMappingName(self) -> str:
        return self._COM.ChannelMappingName

    @ChannelMappingName.setter
    def ChannelMappingName(self, value: str) -> None:
        self._COM.ChannelMappingName = value

    @property
    def Configuration(self) -> Configuration:
        return Configuration(self._COM.Configuration)

    @property
    def Environment(self) -> NotImplementedType:
        return NotImplemented

    @property
    def FullName(self) -> str:
        return self._COM.FullName

    @property
    def Measurement(self) -> NotImplementedType:
        return NotImplemented

    @property
    def Name(self) -> str:
        return self._COM.Name

    @property
    def Networks(self) -> NotImplementedType:
        return NotImplemented

    @property
    def Path(self) -> PLPath:
        return PLPath(self._COM.Path)

    @property
    def Performance(self) -> NotImplementedType:
        return NotImplemented

    @property
    def Simulation(self) -> NotImplementedType:
        return NotImplemented

    @property
    def System(self) -> NotImplementedType:
        return NotImplemented

    @property
    def UI(self) -> NotImplementedType:
        return NotImplemented

    @property
    def Visible(self) -> bool:
        return self._COM.Visible

    @Visible.setter
    def Visible(self, value: bool) -> None:
        self._COM.Visible = value

    @property
    def Version(self) -> Version:
        return Version(self._COM.Version)

    @classmethod
    def Open(
        cls,
        path: PLPath,
        autoSave: Optional[bool] = None,
        promptUser: Optional[bool] = None,
    ) -> None:
        cls._Events.OnOpenFinished.false
        if autoSave and promptUser:
            cls._COM.Open(path, autoSave, promptUser)
        elif autoSave:
            cls._COM.Open(path, autoSave)
        else:
            cls._COM.Open(path)
        waitEventFinished(cls._Events.OnOpenFinished)

    @classmethod
    def Quit(cls) -> None:
        cls._Events.OnQuitFinished.false
        cls._COM.Quit()
        waitEventFinished(cls._Events.OnQuitFinished)

    @property
    def OnOpen(self) -> Callable[[str], None]:
        return self._Events.OnOpenCbk

    @OnOpen.setter
    def OnOpen(self, callback: Callable[[str], None]) -> None:
        self._Events.OnOpenCbk = callback

    @property
    def OnQuit(self) -> Callable[[], None]:
        return self._Events.OnQuitCbk

    @OnQuit.setter
    def OnQuit(self, callback: Callable[[], None]) -> None:
        self._Events.OnQuitCbk = callback

    def __init__(self) -> None:
        self.__class__._COM = win32com.client.DispatchWithEvents(
            "CANoe.Application", self._Events
        )

    def __rich_repr__(self):
        yield "Bus", self.Bus
        yield "CAPL", self.CAPL
        yield "ChannelMappingName", self.ChannelMappingName
        yield self.Configuration
        yield "Environment", self.Environment
        yield "FullName", self.FullName
        yield "Measurement", self.Measurement
        yield "Name", self.Name
        yield "Networks", self.Networks
        yield "Path", self.Path
        yield "Performance", self.Performance
        yield "Simulation", self.Simulation
        yield "System", self.System
        yield "UI", self.UI
        yield "Visible", self.Visible
        yield self.Version

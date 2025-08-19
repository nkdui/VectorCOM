from __future__ import annotations
from enum import IntEnum
from pathlib import Path as PLPath
from typing import Callable, ClassVar, Optional

from win32com.client.dynamic import CDispatch
from win32com.client import WithEvents


class CfgMode(IntEnum):
    Online = 0
    Offline = 1


class CfgExeVariant(IntEnum):
    Win32 = 0
    Win64 = 1


class Configuration:
    _COM: ClassVar[CDispatch]
    ON_CLOSE_CBK: ClassVar[Optional[Callable[[], None]]] = None
    ON_SYSVARCHANGED_CBK: ClassVar[Optional[Callable[[], None]]] = None

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
    def Mode(self) -> CfgMode:
        return CfgMode(self._COM.Mode)

    @Mode.setter
    def Mode(self, new_mode: CfgMode) -> None:
        self._COM.Mode = new_mode.value

    @property
    def Modified(self) -> bool:
        return self._COM.Modified

    @property
    def ExecutionEnvironment(self) -> CfgExeVariant:
        return CfgExeVariant(self._COM.ExecutionEnvironment)

    @ExecutionEnvironment.setter
    def ExecutionEnvironment(self, new_env: CfgExeVariant) -> None:
        self._COM.ExecutionEnvironment = new_env.value

    @property
    def ReadOnly(self) -> bool:
        return self._COM.ReadOnly

    @property
    def Saved(self) -> bool:
        return self._COM.Saved

    @property
    def NETTargetFramework(self) -> int:
        return self._COM.NETTargetFramework

    @property
    def Comment(self) -> str:
        return self._COM.Comment

    def __init__(self, configuration: CDispatch) -> None:
        if not hasattr(self.__class__, "_COM"):
            self.__class__._COM = configuration

    def __rich_repr__(self):
        yield "FullName", self.FullName
        yield "Name", self.Name
        yield "Path", self.Path
        # yield "Mode", self.Mode
        yield "Modified", self.Modified
        yield "ExecutionEnvironment", self.ExecutionEnvironment
        yield "ReadOnly", self.ReadOnly
        yield "Saved", self.Saved
        yield "NETTargetFramework", self.NETTargetFramework
        yield "Comment", self.Comment

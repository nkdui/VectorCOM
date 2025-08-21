import logging
from enum import IntEnum
from pathlib import Path as PLPath
from types import NotImplementedType
from typing import Callable, ClassVar

import rich.repr
from win32com.client import WithEvents
from win32com.client.dynamic import CDispatch

from .common import RefBool
from .testconfiguration import TestConfigurations

LOG = logging.getLogger("VectorCOM")


class CfgMode(IntEnum):
    Online = 0
    Offline = 1


class CfgExeVariant(IntEnum):
    Win32 = 0
    Win64 = 1


class CfgFDXTL(IntEnum):
    FDXTL_UDP_IPv4 = 1
    FDXTL_UDP_IPv6 = 2
    FDXTL_TCP_IPv4 = 3
    FDXTL_TCP_IPv6 = 4


@rich.repr.auto
class Configuration:
    class _Events:
        OnCloseCbk: ClassVar[Callable[..., None]] = lambda: LOG.debug(
            "Closed CANoe configuration"
        )
        OnSysVarDefChangedCbk: ClassVar[Callable[..., None]] = lambda: LOG.debug(
            "System variable definition changed"
        )
        OnCloseFinished: ClassVar[RefBool] = RefBool(True)
        OnSysVarDefChangedFinished: ClassVar[RefBool] = RefBool(True)

        @classmethod
        def OnClose(cls):
            cls.OnCloseCbk()
            cls.OnCloseFinished.true

        @classmethod
        def OnSystemVariablesDefinitionChanged(cls):
            cls.OnSysVarDefChangedCbk()
            cls.OnSysVarDefChangedFinished.true

    _com: ClassVar[CDispatch]

    @property
    def AsynchronousCheckEvaluationEnabled(self) -> bool:
        return self._com.AsynchronousCheckEvaluationEnabled

    @AsynchronousCheckEvaluationEnabled.setter
    def AsynchronousCheckEvaluationEnabled(self, value: bool) -> None:
        self._com.AsynchronousCheckEvaluationEnabled = value

    @property
    def CANoe4Server(self) -> NotImplementedType:
        return NotImplemented

    @property
    def CLibraries(self) -> NotImplementedType:
        return NotImplemented

    @property
    def Comment(self) -> str:
        return self._com.Comment

    @Comment.setter
    def Comment(self, value: str) -> None:
        self._com.Comment = value

    @property
    def CommunicationSetup(self) -> NotImplementedType:
        return NotImplemented

    @property
    def DistributedMode(self) -> NotImplementedType:
        return NotImplemented

    @property
    def EthernetBusSystem(self) -> NotImplementedType:
        return NotImplemented

    @property
    def ExecutionEnvironment(self) -> CfgExeVariant:
        return CfgExeVariant(self._com.ExecutionEnvironment)

    @ExecutionEnvironment.setter
    def ExecutionEnvironment(self, new_env: CfgExeVariant) -> None:
        self._com.ExecutionEnvironment = new_env.value

    @property
    def FDXEnabled(self) -> bool:
        return self._com.FDXEnabled

    @FDXEnabled.setter
    def FDXEnabled(self, value: bool) -> None:
        self._com.FDXEnabled = value

    @property
    def FDXFiles(self) -> NotImplementedType:
        return NotImplemented

    @property
    def FDXPort(self) -> int:
        return self._com.FDXPort

    @FDXPort.setter
    def FDXPort(self, value: int) -> None:
        self._com.FDXPort = value

    @property
    def FDXTransportLayer(self) -> CfgFDXTL:
        return CfgFDXTL(self._com.FDXTransportLayer)

    @FDXTransportLayer.setter
    def FDXTransportLayer(self, new_value: CfgFDXTL) -> None:
        self._com.FDXTransportLayer = new_value.value

    @property
    def FullName(self) -> str:
        return self._com.FullName

    @property
    def GeneralSetup(self) -> NotImplementedType:
        return NotImplemented

    @property
    def GlobalTcpIpStackSetting(self) -> NotImplementedType:
        return NotImplemented

    @property
    def HardwareConfigurations(self) -> NotImplementedType:
        return NotImplemented

    @property
    def HwConfigurationSelection(self) -> int:
        return self._com.HwConfigurationSelection

    @HwConfigurationSelection.setter
    def HwConfigurationSelection(self, value: int) -> None:
        self._com.HwConfigurationSelection = value

    @property
    def IOHardware(self) -> NotImplementedType:
        return NotImplemented

    @property
    def Mode(self) -> CfgMode:
        return CfgMode(self._com.mode)

    @Mode.setter
    def Mode(self, new_mode: CfgMode) -> None:
        self._com.Mode = new_mode.value

    @property
    def Modified(self) -> bool:
        return self._com.Modified

    @property
    def MultiCANoe(self) -> NotImplementedType:
        return NotImplemented

    @property
    def Name(self) -> str:
        return self._com.Name

    @property
    def NETTargetFramework(self) -> int:
        return self._com.NETTargetFramework

    @property
    def OfflineSetup(self) -> NotImplementedType:
        return NotImplemented

    @property
    def OnlineSetup(self) -> NotImplementedType:
        return NotImplemented

    @property
    def OpenConfigurationResult(self) -> NotImplementedType:
        return NotImplemented

    @property
    def Path(self) -> PLPath:
        return PLPath(self._com.Path)

    @property
    def ReadOnly(self) -> bool:
        return self._com.ReadOnly

    @property
    def Saved(self) -> bool:
        return self._com.Saved

    @property
    def Sensor(self) -> NotImplementedType:
        return NotImplemented

    @property
    def ServiceGeneratorActive(self) -> bool:
        return self._com.ServiceGeneratorActive

    @ServiceGeneratorActive.setter
    def ServiceGeneratorActive(self, value: bool) -> None:
        self._com.ServiceGeneratorActive = value

    @property
    def SimulationSetup(self) -> NotImplementedType:
        return NotImplemented

    @property
    def SplitOverlappingFlexRayNMFrames(self) -> bool:
        return self._com.SplitOverlappingFlexRayNMFrames

    @SplitOverlappingFlexRayNMFrames.setter
    def SplitOverlappingFlexRayNMFrames(self, value: bool) -> None:
        self._com.SplitOverlappingFlexRayNMFrames = value

    @property
    def StandaloneMode(self) -> NotImplementedType:
        return NotImplemented

    @property
    def StartValueList(self) -> NotImplementedType:
        return NotImplemented

    @property
    def SymbolMappings(self) -> NotImplementedType:
        return NotImplemented

    @property
    def TestConfigurations(self) -> TestConfigurations:
        return TestConfigurations(self._com.TestConfigurations)

    @property
    def TestSetup(self) -> NotImplementedType:
        return NotImplemented

    @property
    def UserFiles(self) -> NotImplementedType:
        return NotImplemented

    @property
    def UseShortLabel(self) -> bool:
        return self._com.UseShortLabel

    @UseShortLabel.setter
    def UseShortLabel(self, value: bool) -> None:
        self._com.UseShortLabel = value

    @property
    def VTSystem(self) -> NotImplementedType:
        return NotImplemented

    @property
    def XILAPIEnabled(self) -> bool:
        return self._com.XILAPIEnabled

    @property
    def XILAPIPort(self) -> int:
        return self._com.XILAPIPort

    @XILAPIPort.setter
    def XILAPIPort(self, value: int) -> None:
        self._com.XILAPIPort = value

    @classmethod
    def CompileAndVerify(cls):
        cls._com.CompileAndVerify()

    @property
    def OnClose(self) -> Callable[[], None]:
        return self._Events.OnCloseCbk

    @OnClose.setter
    def OnClose(self, callback: Callable[[], None]) -> None:
        self._Events.OnCloseCbk = callback

    @property
    def OnSystemVariablesDefinitionChanged(self) -> Callable[[], None]:
        return self._Events.OnSysVarDefChangedCbk

    @OnSystemVariablesDefinitionChanged.setter
    def OnSystemVariablesDefinitionChanged(self, callback: Callable[[], None]) -> None:
        self._Events.OnSysVarDefChangedCbk = callback

    def __init__(self, configuration: CDispatch) -> None:
        self.__class__._com = configuration
        self._events = WithEvents(configuration, self._Events)

    def __rich_repr__(self):
        yield (
            "AsynchronousCheckEvaluationEnabled",
            self.AsynchronousCheckEvaluationEnabled,
        )
        yield "CANoe4Server", self.CANoe4Server
        yield "CLibraries", self.CLibraries
        yield "Comment", self.Comment
        yield "CommunicationSetup", self.CommunicationSetup
        yield "DistributedMode", self.DistributedMode
        yield "EthernetBusSystem", self.EthernetBusSystem
        yield "ExecutionEnvironment", self.ExecutionEnvironment
        yield "FDXEnabled", self.FDXEnabled
        yield "FDXFiles", self.FDXFiles
        yield "FDXPort", self.FDXPort
        yield "FDXTransportLayer", self.FDXTransportLayer
        yield "FullName", self.FullName
        yield "GeneralSetup", self.GeneralSetup
        yield "GlobalTcpIpStackSetting", self.GlobalTcpIpStackSetting
        yield "HardwareConfigurations", self.HardwareConfigurations
        yield "HwConfigurationSelection", self.HwConfigurationSelection
        yield "IOHardware", self.IOHardware
        yield "Mode", self.Mode
        yield "Modified", self.Modified
        yield "MultiCANoe", self.MultiCANoe
        yield "Name", self.Name
        yield "NETTargetFramework", self.NETTargetFramework
        yield "OfflineSetup", self.OfflineSetup
        yield "OnlineSetup", self.OnlineSetup
        yield "OpenConfigurationResult", self.OpenConfigurationResult
        yield "Path", self.Path
        yield "ReadOnly", self.ReadOnly
        yield "Saved", self.Saved
        yield "Sensor", self.Sensor
        yield "ServiceGeneratorActive", self.ServiceGeneratorActive
        yield "SimulationSetup", self.SimulationSetup
        yield "SplitOverlappingFlexRayNMFrames", self.SplitOverlappingFlexRayNMFrames
        yield "StandaloneMode", self.StandaloneMode
        yield "SymbolMappings", self.SymbolMappings
        yield "TestSetup", self.TestSetup
        yield "UserFiles", self.UserFiles
        yield "UseShortLabel", self.UseShortLabel
        yield "VTSystem", self.VTSystem
        yield "XILAPIEnabled", self.XILAPIEnabled
        yield "XILAPIPort", self.XILAPIPort
        yield self.TestConfigurations

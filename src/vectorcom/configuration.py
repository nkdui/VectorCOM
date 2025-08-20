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

    _COM: ClassVar[CDispatch]

    @property
    def AsynchronousCheckEvaluationEnabled(self) -> bool:
        return self._COM.AsynchronousCheckEvaluationEnabled

    @AsynchronousCheckEvaluationEnabled.setter
    def AsynchronousCheckEvaluationEnabled(self, value: bool) -> None:
        self._COM.AsynchronousCheckEvaluationEnabled = value

    @property
    def CANoe4Server(self) -> NotImplementedType:
        return NotImplemented

    @property
    def CLibraries(self) -> NotImplementedType:
        return NotImplemented

    @property
    def Comment(self) -> str:
        return self._COM.Comment

    @Comment.setter
    def Comment(self, value: str) -> None:
        self._COM.Comment = value

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
        return CfgExeVariant(self._COM.ExecutionEnvironment)

    @ExecutionEnvironment.setter
    def ExecutionEnvironment(self, new_env: CfgExeVariant) -> None:
        self._COM.ExecutionEnvironment = new_env.value

    @property
    def FDXEnabled(self) -> bool:
        return self._COM.FDXEnabled

    @FDXEnabled.setter
    def FDXEnabled(self, value: bool) -> None:
        self._COM.FDXEnabled = value

    @property
    def FDXFiles(self) -> NotImplementedType:
        return NotImplemented

    @property
    def FDXPort(self) -> int:
        return self._COM.FDXPort

    @FDXPort.setter
    def FDXPort(self, value: int) -> None:
        self._COM.FDXPort = value

    @property
    def FDXTransportLayer(self) -> CfgFDXTL:
        return CfgFDXTL(self._COM.FDXTransportLayer)

    @FDXTransportLayer.setter
    def FDXTransportLayer(self, new_value: CfgFDXTL) -> None:
        self._COM.FDXTransportLayer = new_value.value

    @property
    def FullName(self) -> str:
        return self._COM.FullName

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
        return self._COM.HwConfigurationSelection

    @HwConfigurationSelection.setter
    def HwConfigurationSelection(self, value: int) -> None:
        self._COM.HwConfigurationSelection = value

    @property
    def IOHardware(self) -> NotImplementedType:
        return NotImplemented

    @property
    def Mode(self) -> CfgMode:
        return CfgMode(self._COM.mode)

    @Mode.setter
    def Mode(self, new_mode: CfgMode) -> None:
        self._COM.Mode = new_mode.value

    @property
    def Modified(self) -> bool:
        return self._COM.Modified

    @property
    def MultiCANoe(self) -> NotImplementedType:
        return NotImplemented

    @property
    def Name(self) -> str:
        return self._COM.Name

    @property
    def NETTargetFramework(self) -> int:
        return self._COM.NETTargetFramework

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
        return PLPath(self._COM.Path)

    @property
    def ReadOnly(self) -> bool:
        return self._COM.ReadOnly

    @property
    def Saved(self) -> bool:
        return self._COM.Saved

    @property
    def Sensor(self) -> NotImplementedType:
        return NotImplemented

    @property
    def ServiceGeneratorActive(self) -> bool:
        return self._COM.ServiceGeneratorActive

    @ServiceGeneratorActive.setter
    def ServiceGeneratorActive(self, value: bool) -> None:
        self._COM.ServiceGeneratorActive = value

    @property
    def SimulationSetup(self) -> NotImplementedType:
        return NotImplemented

    @property
    def SplitOverlappingFlexRayNMFrames(self) -> bool:
        return self._COM.SplitOverlappingFlexRayNMFrames

    @SplitOverlappingFlexRayNMFrames.setter
    def SplitOverlappingFlexRayNMFrames(self, value: bool) -> None:
        self._COM.SplitOverlappingFlexRayNMFrames = value

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
        return TestConfigurations(self._COM.TestConfigurations)

    @property
    def TestSetup(self) -> NotImplementedType:
        return NotImplemented

    @property
    def UserFiles(self) -> NotImplementedType:
        return NotImplemented

    @property
    def UseShortLabel(self) -> bool:
        return self._COM.UseShortLabel

    @UseShortLabel.setter
    def UseShortLabel(self, value: bool) -> None:
        self._COM.UseShortLabel = value

    @property
    def VTSystem(self) -> NotImplementedType:
        return NotImplemented

    @property
    def XILAPIEnabled(self) -> bool:
        return self._COM.XILAPIEnabled

    @property
    def XILAPIPort(self) -> int:
        return self._COM.XILAPIPort

    @XILAPIPort.setter
    def XILAPIPort(self, value: int) -> None:
        self._COM.XILAPIPort = value

    @classmethod
    def CompileAndVerify(cls):
        cls._COM.CompileAndVerify()

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
        self.__class__._COM = configuration
        WithEvents(configuration, self._Events)

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

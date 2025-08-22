import logging
from types import NotImplementedType
from typing import Callable, ClassVar, Optional

import rich.repr
from win32com.client import CDispatch, WithEvents

from .common import RefBool, StopReason, TestElementType, Verdict, waitEventFinished
from .testtree import TestTreeElements
from .testunit import TestUnits

LOG = logging.getLogger("VectorCOM")


@rich.repr.auto
class TestConfiguration:
    class _Events:
        OnStartCbk: Callable[..., None]
        OnStopCbk: Callable[[StopReason], None]
        OnVerdictChangedCbk: Callable[[Verdict], None]
        OnVerdictFailCbk: Callable[..., None]
        OnStartFinished: RefBool = RefBool(True)
        OnStopFinished: RefBool = RefBool(True)
        OnVerdictChangedFinished: RefBool = RefBool(True)
        OnVerdictFailFinished: RefBool = RefBool(True)

        def OnStart(self):
            self.OnStartCbk()
            self.OnStartFinished.true

        def OnStop(self, reason: StopReason):
            self.OnStopCbk(reason)
            self.OnStopFinished.true

        def OnVerdictChanged(self, verdict: Verdict):
            self.OnVerdictChangedCbk(verdict)
            self.OnVerdictChangedFinished.true

        def OnVerdictFail(self):
            self.OnVerdictFailCbk()
            self.OnVerdictFailFinished.true

    _com: CDispatch

    @property
    def Caption(self) -> Optional[str]:
        try:
            return self._com.Caption
        except AttributeError as attr_e:
            if attr_e.name == "Caption":
                return None
            raise

    @property
    def Elements(self) -> Optional[TestTreeElements]:
        try:
            return TestTreeElements(self._com.Elements)
        except AttributeError as attr_e:
            if attr_e.name == "Elements":
                return None
            raise

    @property
    def Enabled(self) -> str:
        return self._com.Enabled

    @property
    def Id(self) -> Optional[str]:
        try:
            return self._com.Id
        except AttributeError as attr_e:
            if attr_e.name == "Id":
                return None
            raise

    @property
    def Name(self) -> str:
        return self._com.Name

    @property
    def PortCreation(self) -> Optional[int]:
        try:
            return self._com.PortCreation
        except AttributeError as attr_e:
            if attr_e.name == "PortCreation":
                return None
            raise

    @property
    def Report(self) -> NotImplementedType:
        return NotImplemented

    @property
    def Running(self) -> bool:
        try:
            return self._com.Running
        except AttributeError as attr_e:
            if attr_e.name == "Running":
                return None
            raise

    @property
    def Settings(self) -> NotImplementedType:
        return NotImplemented

    @property
    def TcpIpStackSetting(self) -> NotImplementedType:
        return NotImplemented

    @property
    def TestUnits(self) -> TestUnits:
        return TestUnits(self._com.TestUnits)

    @property
    def Type(self) -> Optional[TestElementType]:
        try:
            return TestElementType(self._com.Type)
        except AttributeError as attr_e:
            if attr_e.name == "Type":
                return None
            raise

    @property
    def Verdict(self) -> Verdict:
        return Verdict(self._com.Verdict)

    def Start(self):
        if self.Running:
            return
        self.events.OnStartFinished.false
        self._com.Start()
        waitEventFinished(self.events.OnStartFinished)

    def Stop(self):
        if not self.Running:
            return
        self.events.OnStopFinished.false
        self._com.Stop()
        waitEventFinished(self.events.OnStopFinished)

    def __init__(self, testcfg: CDispatch) -> None:
        self._com = testcfg
        self.events = WithEvents(testcfg, self._Events)
        self.events.OnStartCbk = lambda: LOG.debug(
            "Test Configuration %s started", self.Name
        )
        self.events.OnStopCbk = lambda reason: LOG.debug(
            "Test Configuration %s stopped with reason %s", self.Name, reason
        )
        self.events.OnVerdictChangedCbk = lambda verdict: LOG.debug(
            "Test Configuration %s verdict changed to %s", self.Name, verdict
        )
        self.events.OnVerdictFailCbk = lambda: LOG.debug(
            "Test Configuration %s failed", self.Name
        )

    def __rich_repr__(self):
        yield "Caption", self.Caption
        yield "Elements", self.Elements
        yield "Enabled", self.Enabled
        yield "Id", self.Id
        yield "Name", self.Name
        yield "PortCreation", self.PortCreation
        yield "Report", self.Report
        yield "Running", self.Running
        yield "Settings", self.Settings
        yield "TcpIpStackSetting", self.TcpIpStackSetting
        yield "Type", self.Type
        yield "Verdict", self.Verdict
        yield self.TestUnits


@rich.repr.auto
class TestConfigurations:
    _COM: ClassVar[CDispatch]

    @property
    def Count(self) -> int:
        return self._COM.Count

    def Item(self, index: int) -> TestConfiguration:
        return TestConfiguration(self._COM.Item(index))

    def __init__(self, testcfgs: CDispatch) -> None:
        self.__class__._COM = testcfgs

    def __iter__(self):
        for i in range(1, self.Count + 1):
            yield self.Item(i)

    def __getitem__(self, index: int) -> TestConfiguration:
        return self.Item(index)

    def __rich_repr__(self):
        yield "Count", self.Count
        yield list(self)

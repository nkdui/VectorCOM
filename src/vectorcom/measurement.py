import logging
from typing import Callable, ClassVar, cast

import rich.repr
from win32com.client import CDispatch, WithEvents

from .common import RefBool, waitEventFinished

LOG = logging.getLogger("VectorCOM")


@rich.repr.auto
class Measurement:
    class _Events:
        OnExitCbk: ClassVar[Callable[..., None]] = lambda: LOG.debug(
            "Exiting measurement ..."
        )
        OnInitCbk: ClassVar[Callable[..., None]] = lambda: LOG.debug(
            "Initializing measurement ..."
        )
        OnStartCbk: ClassVar[Callable[..., None]] = lambda: LOG.debug(
            "Starting measurement ..."
        )
        OnStopCbk: ClassVar[Callable[..., None]] = lambda: LOG.debug(
            "Stopping measurement ..."
        )

        OnExitFinished: ClassVar[RefBool] = RefBool(True)
        OnInitFinished: ClassVar[RefBool] = RefBool(True)
        OnStartFinished: ClassVar[RefBool] = RefBool(True)
        OnStopFinished: ClassVar[RefBool] = RefBool(True)

        @classmethod
        def OnExit(cls):
            cls.OnExitCbk()
            cls.OnExitFinished.true

        @classmethod
        def OnInit(cls):
            cls.OnInitCbk()
            cls.OnInitFinished.true

        @classmethod
        def OnStart(cls):
            cls.OnStartCbk()
            cls.OnStartFinished.true

        @classmethod
        def OnStop(cls):
            cls.OnStopCbk()
            cls.OnStopFinished.true

    _com: ClassVar[CDispatch]
    events: ClassVar[_Events]

    @property
    def AnimationDelay(self) -> int:
        return self._com.AnimationDelay

    @AnimationDelay.setter
    def AnimationDelay(self, new_value: int):
        self._com.AnimationDelay = new_value

    @property
    def MeasurementIndex(self) -> int:
        return self._com.MeasurementIndex

    @MeasurementIndex.setter
    def MeasurementIndex(self, new_value: int):
        self._com.MeasurementIndex = new_value

    @property
    def Running(self) -> bool:
        return self._com.Running

    @Running.setter
    def Running(self, new_value: bool):
        self._com.Running = new_value

    @classmethod
    def Animate(cls):
        cls._com.Animate()

    @classmethod
    def Break(cls):
        cls._com.Break()

    @classmethod
    def Reset(cls):
        cls._com.Reset()

    @classmethod
    def Start(cls):
        cls.events.OnStartFinished.false
        cls._com.Start()
        waitEventFinished(cls.events.OnStartFinished)

    @classmethod
    def Step(cls):
        cls._com.Step()

    @classmethod
    def StopEx(cls):
        cls.events.OnStopFinished.false
        cls._com.StopEx()
        waitEventFinished(cls.events.OnStopFinished)

    def __init__(self, measurement: CDispatch) -> None:
        Measurement._com = measurement
        Measurement.events = cast(
            Measurement._Events, WithEvents(measurement, self._Events)
        )

    def __rich_repr__(self):
        yield "AnimationDelay", self.AnimationDelay
        yield "MeasurementIndex", self.MeasurementIndex
        yield "Running", self.Running

from enum import IntEnum
from time import sleep, time

import pythoncom


class RefBool:
    @property
    def true(self) -> None:
        self._value = True

    @property
    def false(self) -> None:
        self._value = False

    def __init__(self, value: bool):
        self._value = value

    def __bool__(self):
        return self._value


class Verdict(IntEnum):
    VerdictNotAvailable = 0
    VerdictPassed = 1
    VerdictFailed = 2
    VerdictNone = 3
    VerdictInconclusive = 4
    VerdictErrorInTestSystem = 5


class TestElementType(IntEnum):
    TestTypeReserved = 0
    TestConfiguration = 1
    TestUnit = 2
    TestGroup = 3
    TestSequence = 4
    TestCase = 5
    TestFixture = 6
    TestCaseList = 7
    TestSequenceList = 8


class StopReason(IntEnum):
    StopReasonEnd = 0
    StopReasonUserAbort = 1
    StopReasonGeneralError = 2
    StopReasonVerdictImpact = 3


def waitEventFinished(event: RefBool, timeout: float = 0, waitstep: float = 0.1):
    start_time = time()
    while (timeout == 0) or (time() - start_time < timeout):
        pythoncom.PumpWaitingMessages()
        if event:
            return
        sleep(waitstep)
    raise TimeoutError("Event did not finish in time")

from types import NotImplementedType

import rich.repr
from win32com.client import CDispatch

from .common import TestElementType, Verdict
from .testtree import TestTreeElements


@rich.repr.auto
class TestUnit:
    _com: CDispatch

    @property
    def Caption(self) -> str:
        try:
            return self._com.Caption
        except AttributeError as attr_e:
            if attr_e.name == "Caption":
                return NotImplemented
            raise

    @property
    def Elements(self) -> TestTreeElements:
        try:
            return TestTreeElements(self._com.Elements)
        except AttributeError as attr_e:
            if attr_e.name == "Elements":
                return NotImplemented
            raise

    @property
    def Enabled(self) -> bool:
        return self._com.Enabled

    @Enabled.setter
    def Enabled(self, value: bool) -> None:
        self._com.Enabled = value

    @property
    def Id(self) -> str:
        try:
            return self._com.Id
        except AttributeError as attr_e:
            if attr_e.name == "Id":
                return NotImplemented
            raise

    @property
    def Name(self) -> str:
        return self._com.Name

    @property
    def Report(self) -> NotImplementedType:
        return NotImplemented

    @property
    def Type(self) -> TestElementType:
        try:
            return TestElementType(self._com.Type)
        except AttributeError as attr_e:
            if attr_e.name == "Type":
                return NotImplemented
            raise

    @property
    def Verdict(self) -> Verdict:
        return Verdict(self._com.Verdict)

    def __init__(self, testunit: CDispatch) -> None:
        self._com = testunit

    def __rich_repr__(self):
        yield "Caption", self.Caption
        yield "Enabled", self.Enabled
        yield "Id", self.Id
        yield "Name", self.Name
        yield "Report", self.Report
        yield "Type", self.Type
        yield "Verdict", self.Verdict
        yield self.Elements


@rich.repr.auto
class TestUnits:
    _com: CDispatch

    @property
    def Count(self) -> int:
        return self._com.Count

    def Item(self, index: int) -> TestUnit:
        return TestUnit(self._com.Item(index))

    def __init__(self, testunits: CDispatch) -> None:
        self._com = testunits

    def __iter__(self):
        for i in range(1, self.Count + 1):
            yield self.Item(i)

    def __rich_repr__(self):
        yield "Count", self.Count
        yield list(self)

from __future__ import annotations

import rich.repr
from win32com.client import CDispatch

from .common import TestElementType, Verdict


@rich.repr.auto
class TestTreeElement:
    _com: CDispatch

    @property
    def Caption(self) -> str:
        return self._com.Caption

    @property
    def Elements(self) -> TestTreeElements:
        return TestTreeElements(self._com.Elements)

    @property
    def Enabled(self) -> bool:
        return self._com.Enabled

    @Enabled.setter
    def Enabled(self, value: bool) -> None:
        self._com.Enabled = value

    @property
    def Id(self) -> str:
        return self._com.Id

    @property
    def Title(self) -> str:
        return self._com.Title

    @property
    def Type(self) -> TestElementType:
        return TestElementType(self._com.Type)

    @property
    def Verdict(self) -> Verdict:
        return Verdict(self._com.Verdict)

    def __init__(self, testtree: CDispatch) -> None:
        self._com = testtree

    def __rich_repr__(self):
        yield "Caption", self.Caption
        yield "Enabled", self.Enabled
        yield "Id", self.Id
        yield "Title", self.Title
        yield "Type", self.Type
        yield "Verdict", self.Verdict
        yield self.Elements


@rich.repr.auto
class TestTreeElements:
    _com: CDispatch

    @property
    def Count(self) -> int:
        return self._com.Count

    def Item(self, index: int) -> TestTreeElement:
        return TestTreeElement(self._com.Item(index))

    def __init__(self, testtrees: CDispatch) -> None:
        self._com = testtrees

    def __iter__(self):
        for i in range(1, self.Count + 1):
            yield self.Item(i)

    def __rich_repr__(self):
        yield "Count", self.Count
        yield list(self)

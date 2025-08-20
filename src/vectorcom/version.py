import rich
import rich.repr
from win32com.client.dynamic import CDispatch


@rich.repr.auto
class Version:
    @property
    def FullName(self) -> str:
        return self._com.FullName

    @property
    def Name(self) -> str:
        return self._com.Name

    @property
    def major(self) -> int:
        return self._com.major

    @property
    def minor(self) -> int:
        return self._com.minor

    @property
    def Build(self) -> int:
        return self._com.Build

    @property
    def Patch(self) -> int:
        return self._com.Patch

    def __init__(self, version: CDispatch) -> None:
        self._com = version

    def __rich_repr__(self):
        yield "FullName", self.FullName
        yield "Name", self.Name
        yield "major", self.major
        yield "minor", self.minor
        yield "Build", self.Build
        yield "Patch", self.Patch

from typing import Callable, Generator, Union, NewType

PathLike = NewType('PathLike', str)
FolderPath = Union[str, PathLike]
FilePath = Union[str, PathLike]

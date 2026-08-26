from __future__ import annotations

from contextlib import contextmanager
from importlib import resources
from typing import BinaryIO, Iterator

from hr_toolkit import templates as _template_package


TEMPLATE_PACKAGE = _template_package.__name__


@contextmanager
def open_template_resource(resource_name: str) -> Iterator[BinaryIO]:
    files = getattr(resources, "files", None)
    if files is not None:
        with files(_template_package).joinpath(resource_name).open("rb") as handle:
            yield handle
        return

    # Python 3.8 的 open_binary() 要求目标是有实际 loader 的普通包。
    # 显式导入模板包也确保 PyInstaller 不会把只有数据文件的目录降级成
    # 无 location 的 namespace package。
    with resources.open_binary(_template_package, resource_name) as handle:
        yield handle

from __future__ import annotations

from contextlib import contextmanager
from importlib import resources
from typing import BinaryIO, Iterator


TEMPLATE_PACKAGE = "hr_toolkit.templates"


@contextmanager
def open_template_resource(resource_name: str) -> Iterator[BinaryIO]:
    files = getattr(resources, "files", None)
    if files is not None:
        with files(TEMPLATE_PACKAGE).joinpath(resource_name).open("rb") as handle:
            yield handle
        return

    with resources.open_binary(TEMPLATE_PACKAGE, resource_name) as handle:
        yield handle

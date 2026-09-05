"""资料打包的阶段计数：完成一个工作项后才前进，未知总量不显示百分比。"""

from __future__ import annotations

import time
from typing import Any, Callable, Iterable, Iterator


class MaterialProgress:
    def __init__(self, callback: Callable[[int, int, str], None] | None) -> None:
        self.callback = callback
        self.context = ""
        self.phase = "准备资料"
        self.current = 0
        self.total = 0
        self.unit = "项"
        self._last_emit = 0.0

    def emit(self, detail: str = "", *, force: bool = False) -> None:
        if self.callback is None:
            return
        now = time.monotonic()
        if not force and now - self._last_emit < 0.1:
            return
        self._last_emit = now
        count = f"已完成 {self.current}/{self.total} {self.unit}" if self.total else "正在处理，总量尚未确定"
        message = f"【{self.phase}】{self.context}{count}"
        if detail:
            message += "；" + detail
        self.callback(self.current, self.total, message)

    def begin(self, phase: str, total: int = 0, unit: str = "项", detail: str = "") -> None:
        self.phase, self.current, self.total, self.unit = phase, 0, total, unit
        self.emit(detail, force=True)

    def advance(self, detail: str = "") -> None:
        self.current += 1
        self.emit(detail, force=self.total > 0 and self.current == self.total)

    def detail(self, _current: int, _total: int, message: str) -> None:
        # PDF 页计数仅描述当前文件，不能覆盖外层文件计数。
        self.emit(message)

    def items(self, items: Iterable[Any], describe: Callable[[Any], str]) -> Iterator[Any]:
        for item in items:
            self.emit("正在处理：" + describe(item), force=self.current == 0)
            yield item
            # continue 也会到达这里；取消或异常退出不虚增完成数量。
            self.current += 1
            if self.total > 0 and self.current == self.total:
                self.emit("本阶段工作项已处理完毕", force=True)

    def skipped_tail(self, count: int) -> None:
        # 提前满足材料要求时单独列明跳过数，绝不把这些文件记成已 OCR。
        self.emit(f"本阶段结束：已核对 {self.current} 个；材料已满足，另有 {count} 个无需识别", force=True)

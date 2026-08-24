from __future__ import annotations

from datetime import date, datetime
import hashlib
import os
import shutil
import subprocess
import sys
from pathlib import Path


SUPPORTED_EXCEL_SUFFIXES = {".xlsx", ".xls"}


def is_supported_excel_file(path: Path) -> bool:
    return path.suffix.lower() in SUPPORTED_EXCEL_SUFFIXES and not path.name.startswith(("~$", ".~"))


def ensure_xlsx_workbook(path: Path, temp_dir: Path) -> Path:
    path = Path(path).expanduser().resolve()
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_EXCEL_SUFFIXES:
        raise ValueError(f"Excel 文件仅支持 .xlsx 或 .xls：{path}")

    file_kind = _detect_excel_file_kind(path)
    if suffix == ".xlsx" and file_kind == "xlsx":
        return path
    output_dir = _conversion_dir(path, temp_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{path.stem}.xlsx"
    if output_path.exists():
        return output_path

    if file_kind == "xlsx":
        shutil.copyfile(path, output_path)
    else:
        _convert_xls_to_xlsx(path, output_path, temp_dir=temp_dir)
    if not output_path.exists():
        raise RuntimeError(f".xls 转换失败，未生成文件：{output_path}")
    return output_path


def _detect_excel_file_kind(path: Path) -> str:
    with path.open("rb") as file:
        header = file.read(8)
    if header.startswith(b"PK\x03\x04"):
        return "xlsx"
    if header.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return "xls"
    return path.suffix.lower().lstrip(".")


def _conversion_dir(path: Path, temp_dir: Path) -> Path:
    digest = hashlib.sha1(str(path).encode("utf-8")).hexdigest()[:12]
    return temp_dir / "xls_converted" / digest


def _convert_with_xlrd(source: Path, output_path: Path) -> None:
    """使用纯 Python xlrd + openpyxl 在内存中将 .xls 转为 .xlsx。

    特点：
    1. 纯只读文件流解析，绝不修改或触碰源文件（100% 解决 Windows 下批次文件校验哈希变化的问题）；
    2. 无需启动外部 Excel/WPS/COM 进程，速度快（0.01 秒级），跨平台稳定。
    """
    import xlrd
    import openpyxl

    rb = xlrd.open_workbook(str(source), formatting_info=False)
    wb = openpyxl.Workbook()
    try:
        # 移除默认新建的 Sheet
        wb.remove(wb.active)

        for sheet_name in rb.sheet_names():
            rs = rb.sheet_by_name(sheet_name)
            ws = wb.create_sheet(title=sheet_name)
            for row_idx in range(rs.nrows):
                row_vals: list[object] = []
                for col_idx in range(rs.ncols):
                    cell = rs.cell(row_idx, col_idx)
                    val = cell.value
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        try:
                            dt_tuple = xlrd.xldate_as_tuple(val, rb.datemode)
                            if dt_tuple[3:] == (0, 0, 0):
                                val = date(dt_tuple[0], dt_tuple[1], dt_tuple[2])
                            else:
                                val = datetime(*dt_tuple)
                        except Exception:
                            pass
                    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                        val = bool(val)
                    elif cell.ctype == xlrd.XL_CELL_ERROR:
                        val = None
                    elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                        val = None
                    row_vals.append(val)
                ws.append(row_vals)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(str(output_path))
    finally:
        wb.close()
        rb.release_resources()


def _convert_xls_to_xlsx(source: Path, output_path: Path, temp_dir: Path | None = None) -> None:
    errors: list[str] = []

    # 1. 优先使用纯 Python xlrd 内存转换（极速、跨平台、绝不触碰源文件）
    try:
        _convert_with_xlrd(source, output_path)
        if output_path.exists() and output_path.stat().st_size > 0:
            return
    except Exception as exc:
        errors.append(f"内置 xlrd 转换失败：{exc}")

    # 2. 如果源文件为特殊格式导致 xlrd 失败，在临时沙箱副本中通过 COM / LibreOffice 转换
    sandbox_dir = (temp_dir or output_path.parent) / "xls_sandbox"
    sandbox_dir.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha1(str(source).encode("utf-8")).hexdigest()[:8]
    sandbox_source = sandbox_dir / f"{digest}_{source.name}"
    try:
        shutil.copyfile(source, sandbox_source)
    except Exception as exc:
        errors.append(f"创建只读转换副本失败：{exc}")
        raise RuntimeError(
            "无法创建 .xls 临时转换副本；为保护源文件，已停止外部转换。"
            + (" 详细信息：" + "；".join(errors) if errors else "")
        ) from exc

    try:
        if sys.platform.startswith("win"):
            try:
                _convert_with_windows_com(sandbox_source, output_path)
                return
            except Exception as exc:
                errors.append(f"Excel/WPS 转换失败：{exc}")
        try:
            _convert_with_libreoffice(sandbox_source, output_path)
            return
        except Exception as exc:
            errors.append(f"LibreOffice 转换失败：{exc}")
    finally:
        if sandbox_source != source and sandbox_source.exists():
            try:
                sandbox_source.unlink()
            except Exception:
                pass

    raise RuntimeError(
        "无法将 .xls 转换为 .xlsx。请确认文件格式完整，或本机已安装 Microsoft Excel、WPS 表格或 LibreOffice。"
        + (" 详细信息：" + "；".join(errors) if errors else "")
    )


def _convert_with_windows_com(source: Path, output_path: Path) -> None:
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "缺少 Windows .xls 转换依赖 pywin32。请在 Windows 打包环境执行 "
            "`python -m pip install -r requirements.txt` 后重新打包。"
        ) from exc

    pythoncom.CoInitialize()
    app = None
    workbook = None
    try:
        last_error: Exception | None = None
        for prog_id in ("Excel.Application", "Ket.Application", "KET.Application", "ET.Application", "et.Application"):
            try:
                app = win32com.client.DispatchEx(prog_id)
                break
            except Exception as exc:
                last_error = exc
        if app is None:
            raise RuntimeError(f"未找到 Excel 或 WPS COM 组件：{last_error}")
        app.Visible = False
        app.DisplayAlerts = False
        workbook = app.Workbooks.Open(
            str(source),
            UpdateLinks=0,
            ReadOnly=True,
            AddToMru=False,
        )
        workbook.SaveAs(str(output_path), FileFormat=51)
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _convert_with_libreoffice(source: Path, output_path: Path) -> None:
    executable = _find_libreoffice()
    if executable is None:
        raise RuntimeError(_diagnose_missing_libreoffice())
    output_path.parent.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [
            executable,
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(output_path.parent),
            str(source),
        ],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        env=_build_conversion_env(),
        timeout=300,
    )
    if result.returncode != 0:
        raise RuntimeError((result.stderr or result.stdout or "").strip() or f"退出码 {result.returncode}")
    converted = output_path.parent / f"{source.stem}.xlsx"
    if converted.exists() and converted != output_path:
        shutil.move(str(converted), str(output_path))


def _find_libreoffice() -> str | None:
    """Locate the LibreOffice / soffice executable on the host.

    macOS .app bundles launched from Finder inherit a minimal ``PATH`` that
    only contains ``/usr/bin:/bin:/usr/sbin:/sbin``. That means Homebrew
    installs at ``/opt/homebrew/bin`` (Apple Silicon) and
    ``/usr/local/bin`` (Intel) — the two most common install locations —
    are *not* visible to ``shutil.which`` even when ``soffice`` is installed
    on the machine. Without these fallbacks the user sees a confusing
    "未找到 libreoffice/soffice 命令" error even though the binary exists.
    """

    search_paths: list[str] = []
    path_env = os.environ.get("PATH", "")
    if path_env:
        search_paths.extend(path_env.split(os.pathsep))

    homebrew_candidates = (
        "/opt/homebrew/bin",
        "/usr/local/bin",
        "/opt/local/bin",
    )
    for candidate in homebrew_candidates:
        if candidate not in search_paths:
            search_paths.append(candidate)

    mac_app_paths = (
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "~/Applications/LibreOffice.app/Contents/MacOS/soffice",
    )
    for candidate in mac_app_paths:
        expanded = os.path.expanduser(candidate)
        if Path(expanded).exists():
            return expanded

    for command in ("soffice", "libreoffice"):
        executable = shutil.which(command, path=os.pathsep.join(search_paths))
        if executable:
            return executable

    return None


def _build_conversion_env() -> dict[str, str]:
    """Provide a stable PATH so subprocess can locate LibreOffice on macOS.

    Rebuilds ``PATH`` from the current environment plus the well-known
    Homebrew prefixes so the spawned ``soffice`` process can find its own
    resources (system fonts, locale data, etc.) — not just the binary.
    """

    env = os.environ.copy()
    candidates = ("/opt/homebrew/bin", "/usr/local/bin", "/opt/local/bin")
    current_path = env.get("PATH", "")
    parts = [p for p in current_path.split(os.pathsep) if p]
    for candidate in candidates:
        if Path(candidate).exists() and candidate not in parts:
            parts.append(candidate)
    env["PATH"] = os.pathsep.join(parts) if parts else current_path
    return env


def _diagnose_missing_libreoffice() -> str:
    """Return a helpful explanation when LibreOffice cannot be located."""

    path_env = os.environ.get("PATH", "")
    checked: list[str] = []
    for candidate in (
        "/opt/homebrew/bin/soffice",
        "/usr/local/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        checked.append(f"{candidate}={'存在' if Path(candidate).exists() else '缺失'}")
    return (
        "未找到 libreoffice/soffice 命令。当前 PATH="
        f"{path_env or '<空>'}；已检查："
        + "，".join(checked)
        + "。请安装 LibreOffice（brew install --cask libreoffice），或将其可执行文件添加到 PATH。"
    )

from __future__ import annotations

import argparse
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from build_windows import (
    PE_MACHINE_AMD64,
    WIN7_REQUIRED_7ZIP_FILES,
    WIN7_REQUIRED_UCRT_FILES,
    WIN7_REQUIRED_VC_RUNTIME_FILES,
    read_pe_machine,
)


SEVEN_ZIP_VERSION = "26.02"
SEVEN_ZIP_URL = "https://github.com/ip7z/7zip/releases/download/26.02/7z2602-x64.exe"
SEVEN_ZIP_SHA256 = "6745fa76dc2ea031596d8678f6f6b99c3c1b435b4164a63485adbbc7b8d82ef0"
UCRT_VERSION = "10.0.14393.795"
UCRT_URL = (
    "https://download.microsoft.com/download/C/D/8/"
    "CD8533F8-5324-4D30-824C-B834C5AD51F9/standalonesdk/Installers/"
    "948a611cd2aca64b1e5113ffb7b95d5f.cab"
)
UCRT_SHA256 = "13719323a05589c3ee7f600794c5d385d692ba1a30d87ce86987b23d511de976"
VC_REDIST_VERSION = "14.29.30157"
VC_REDIST_URL = (
    "https://download.visualstudio.microsoft.com/download/pr/"
    "35564904-a9a4-4911-813b-6acd16f6f0d5/"
    "6AFAE68A783F11292149175844AED0E2CE3F247BC0250F6CB18C931295B3F399/"
    "VC_redist.x64.exe"
)
VC_REDIST_SHA256 = "6afae68a783f11292149175844aed0e2ce3f247bc0250f6cb18c931295b3f399"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="准备固定版本的 Win7 app-local UCRT 与官方 7-Zip x64 运行时。"
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=REPO_ROOT / "build" / "win7-runtime",
    )
    parser.add_argument("--ucrt-source", type=Path, help="可选的 Windows SDK UCRT x64 目录")
    args = parser.parse_args(argv)

    if os.name != "nt":
        raise RuntimeError("Win7 运行时只能在 Windows runner 上准备。")
    output_dir = args.output_dir.resolve()
    seven_zip_dir = output_dir / "7zip"
    ucrt_dir = output_dir / "ucrt"
    vc_runtime_dir = output_dir / "vc-runtime"
    _prepare_seven_zip(seven_zip_dir)
    _prepare_ucrt(ucrt_dir, explicit=args.ucrt_source)
    _prepare_vc_runtime(vc_runtime_dir)
    print(f"7-Zip 运行时：{seven_zip_dir}")
    print(f"app-local UCRT：{ucrt_dir}")
    print(f"Visual C++ app-local runtime：{vc_runtime_dir}")
    return 0


def _prepare_seven_zip(target_dir: Path) -> None:
    target_dir.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="hr_toolkit_7zip_") as temporary:
        installer = Path(temporary) / "7zip-x64.exe"
        _download_verified(SEVEN_ZIP_URL, installer, SEVEN_ZIP_SHA256)
        install_dir = Path(temporary) / "installed"
        subprocess.run(
            [str(installer), "/S", f"/D={install_dir}"],
            check=True,
            timeout=120,
        )
        _replace_with_selected_files(
            source_dir=install_dir,
            target_dir=target_dir,
            names=WIN7_REQUIRED_7ZIP_FILES,
            label=f"7-Zip {SEVEN_ZIP_VERSION}",
        )


def _prepare_ucrt(target_dir: Path, *, explicit: Path | None = None) -> None:
    if explicit is not None:
        _replace_with_selected_files(
            source_dir=explicit.resolve(),
            target_dir=target_dir,
            names=WIN7_REQUIRED_UCRT_FILES,
            label=f"Windows SDK {UCRT_VERSION} app-local UCRT",
        )
        return

    target_dir.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="hr_toolkit_ucrt_") as temporary:
        temporary_dir = Path(temporary)
        archive = temporary_dir / "ucrt-x64.cab"
        _download_verified(UCRT_URL, archive, UCRT_SHA256)
        extracted_dir = temporary_dir / "extracted"
        extracted_dir.mkdir()
        subprocess.run(
            ["expand.exe", "-F:*", str(archive), str(extracted_dir)],
            check=True,
            timeout=120,
        )
        selected = _discover_ucrt_files(extracted_dir)
        _replace_with_discovered_mapping(
            selected=selected,
            target_dir=target_dir,
            names=WIN7_REQUIRED_UCRT_FILES,
            label=f"Windows SDK {UCRT_VERSION} app-local UCRT",
        )


def _prepare_vc_runtime(target_dir: Path) -> None:
    target_dir.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="hr_toolkit_vcredist_") as temporary:
        temporary_dir = Path(temporary)
        installer = temporary_dir / "VC_redist.x64.exe"
        _download_verified(VC_REDIST_URL, installer, VC_REDIST_SHA256)
        layout_dir = temporary_dir / "layout"
        layout_dir.mkdir()
        # VC_redist /layout writes its payload into the current directory.
        subprocess.run(
            [str(installer), "/layout", "/quiet", "/norestart"],
            cwd=layout_dir,
            check=True,
            timeout=180,
        )
        minimum_msi = _find_casefold_name(layout_dir, "vc_runtimeMinimum_x64.msi")
        extracted_dir = temporary_dir / "extracted"
        subprocess.run(
            [
                "msiexec.exe",
                "/a",
                str(minimum_msi),
                "/qn",
                f"TARGETDIR={extracted_dir}",
            ],
            check=True,
            timeout=180,
        )
        _replace_with_discovered_files(
            source_dir=extracted_dir,
            target_dir=target_dir,
            names=WIN7_REQUIRED_VC_RUNTIME_FILES,
            label=f"Visual C++ 2015-2019 {VC_REDIST_VERSION}",
        )


def _discover_ucrt_files(extracted_dir: Path) -> dict[str, Path]:
    selected: dict[str, Path] = {}
    required = {name.casefold() for name in WIN7_REQUIRED_UCRT_FILES}
    for path in sorted(extracted_dir.rglob("*")):
        if not path.is_file() or not path.name.casefold().startswith("fil"):
            continue
        if read_pe_machine(path) != PE_MACHINE_AMD64:
            raise RuntimeError(f"Windows SDK {UCRT_VERSION} UCRT 包含非 x64 PE：{path.name}")
        name = _embedded_ucrt_name(path, required=required)
        if name in selected:
            raise RuntimeError(f"Windows SDK {UCRT_VERSION} UCRT 文件重复：{name}")
        selected[name] = path
    return selected


def _embedded_ucrt_name(path: Path, *, required: set[str]) -> str:
    data = path.read_bytes()
    decoded_values = (
        data.decode("utf-16le", errors="ignore"),
        data.decode("latin-1", errors="ignore"),
    )
    if any("ucrtbase.dll" in decoded.casefold() for decoded in decoded_values):
        return "ucrtbase.dll"
    matches = {
        match.casefold()
        for decoded in decoded_values
        for match in re.findall(r"api-ms-win-[a-z0-9-]+\.dll", decoded, flags=re.IGNORECASE)
        if match.casefold() in required
    }
    if len(matches) != 1:
        raise RuntimeError(
            f"Windows SDK {UCRT_VERSION} UCRT 文件名无法唯一识别："
            f"{path.name} -> {sorted(matches)}"
        )
    return matches.pop()


def _download_verified(url: str, destination: Path, expected_sha256: str) -> None:
    request = urllib.request.Request(url, headers={"User-Agent": "HRToolkit-Build/1.0"})
    digest = hashlib.sha256()
    with urllib.request.urlopen(request, timeout=60) as response, destination.open("wb") as output:
        while True:
            chunk = response.read(1024 * 1024)
            if not chunk:
                break
            output.write(chunk)
            digest.update(chunk)
    actual = digest.hexdigest()
    if actual != expected_sha256:
        destination.unlink(missing_ok=True)
        raise RuntimeError(
            f"运行时下载校验失败：期望 {expected_sha256}，实际 {actual}"
        )


def _find_casefold_name(root: Path, name: str) -> Path:
    matches = sorted(
        path
        for path in root.rglob("*")
        if path.is_file() and path.name.casefold() == name.casefold()
    )
    if len(matches) != 1:
        raise RuntimeError(f"运行时布局中 {name} 数量不正确：{len(matches)}")
    return matches[0]


def _replace_with_discovered_files(
    *,
    source_dir: Path,
    target_dir: Path,
    names: tuple[str, ...],
    label: str,
) -> None:
    selected: dict[str, Path] = {}
    for name in names:
        matches = sorted(
            path
            for path in source_dir.rglob("*")
            if path.is_file() and path.name.casefold() == name.casefold()
        )
        if not matches:
            raise RuntimeError(f"{label} 缺少文件：{name}")
        digests = {_sha256_file(path) for path in matches}
        if len(digests) != 1:
            raise RuntimeError(f"{label} 包含内容不同的同名文件：{name}")
        selected[name] = matches[0]
    if target_dir.exists():
        shutil.rmtree(target_dir)
    target_dir.mkdir(parents=True)
    for name, source in selected.items():
        shutil.copy2(source, target_dir / name)


def _replace_with_discovered_mapping(
    *,
    selected: dict[str, Path],
    target_dir: Path,
    names: tuple[str, ...],
    label: str,
) -> None:
    missing = [name for name in names if name.casefold() not in selected]
    unexpected = sorted(set(selected) - {name.casefold() for name in names})
    if missing or unexpected:
        raise RuntimeError(
            f"{label} 内容不完整：缺少 {missing}，多出 {unexpected}"
        )
    if target_dir.exists():
        shutil.rmtree(target_dir)
    target_dir.mkdir(parents=True)
    for name in names:
        shutil.copy2(selected[name.casefold()], target_dir / name)


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _replace_with_selected_files(
    *,
    source_dir: Path,
    target_dir: Path,
    names: tuple[str, ...],
    label: str,
) -> None:
    missing = [name for name in names if not (source_dir / name).is_file()]
    if missing:
        raise RuntimeError(f"{label} 缺少文件：{missing}")
    if target_dir.exists():
        shutil.rmtree(target_dir)
    target_dir.mkdir(parents=True)
    for name in names:
        shutil.copy2(source_dir / name, target_dir / name)


if __name__ == "__main__":
    raise SystemExit(main())

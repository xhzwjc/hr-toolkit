"""Build and verify the compact NumPy runtime used by Intel macOS releases.

The upstream NumPy x86_64 wheel bundles OpenBLAS and makes the complete DMG
larger than the release host's attachment limit.  HR Toolkit does not perform
large linear-algebra workloads; RapidOCR only needs NumPy's ordinary array and
small-vector operations.  Building the same pinned NumPy source with its
documented internal fallback preserves those results without shipping the
unused OpenBLAS payload.

This script is intentionally restricted to native CPython 3.12/x86_64 macOS.
All other release targets keep the ordinary pinned upstream wheel.
"""

from __future__ import annotations

import argparse
import hashlib
import os
import platform
import subprocess
import sys
from pathlib import Path
from typing import Mapping, Optional, Sequence


REPO_ROOT = Path(__file__).resolve().parents[1]
NUMPY_VERSION = "1.26.4"
MACOS_DEPLOYMENT_TARGET = "11.0"
EXPECTED_WHEEL_NAME = (
    f"numpy-{NUMPY_VERSION}-cp312-cp312-macosx_11_0_x86_64.whl"
)
MAX_COMPACT_WHEEL_BYTES = 10 * 1024 * 1024


class MacX64RuntimeError(RuntimeError):
    """Raised when the Intel macOS release runtime is not reproducibly compact."""


def _run(
    command: Sequence[str],
    *,
    env: Optional[Mapping[str, str]] = None,
) -> None:
    print("执行：" + " ".join(str(part) for part in command), flush=True)
    result = subprocess.run(list(command), check=False, env=dict(env) if env else None)
    if result.returncode != 0:
        raise MacX64RuntimeError(
            f"命令失败（{result.returncode}）：{' '.join(str(part) for part in command)}"
        )


def require_native_intel_macos() -> None:
    if sys.platform != "darwin" or platform.machine() != "x86_64":
        raise MacX64RuntimeError("紧凑运行时只能在原生 Intel macOS 构建机上生成")
    if sys.version_info[:2] != (3, 12):
        raise MacX64RuntimeError("紧凑运行时必须使用 CPython 3.12 构建")


def _safe_prepare_wheel_dir(wheel_dir: Path) -> Path:
    resolved = wheel_dir.expanduser().resolve()
    allowed_root = (REPO_ROOT / "build").resolve()
    if resolved == allowed_root or allowed_root not in resolved.parents:
        raise MacX64RuntimeError(f"wheel 目录必须位于仓库 build/ 子目录：{resolved}")
    resolved.mkdir(parents=True, exist_ok=True)
    for path in resolved.glob("numpy-*.whl"):
        if path.is_symlink() or not path.is_file():
            raise MacX64RuntimeError(f"拒绝清理异常 wheel 路径：{path}")
        path.unlink()
    return resolved


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def validate_compact_wheel(wheel_path: Path) -> None:
    if wheel_path.name != EXPECTED_WHEEL_NAME:
        raise MacX64RuntimeError(
            "NumPy wheel 的版本、Python ABI、最低 macOS 版本或架构不符合发布要求："
            f"{wheel_path.name}"
        )
    size = wheel_path.stat().st_size
    if size <= 0 or size > MAX_COMPACT_WHEEL_BYTES:
        raise MacX64RuntimeError(
            f"NumPy 紧凑 wheel 体积异常：{size} 字节（上限 {MAX_COMPACT_WHEEL_BYTES}）"
        )


def configuration_uses_internal_fallback(configuration: Mapping[str, object]) -> bool:
    dependencies = configuration.get("Build Dependencies")
    if not isinstance(dependencies, Mapping):
        return False
    blas = dependencies.get("blas")
    lapack = dependencies.get("lapack")
    if not isinstance(blas, Mapping) or not isinstance(lapack, Mapping):
        return False
    return (
        str(blas.get("name", "")).casefold() == "none"
        and str(lapack.get("detection method", "")).casefold() == "internal"
    )


def verify_installed_runtime() -> None:
    import numpy

    if numpy.__version__ != NUMPY_VERSION:
        raise MacX64RuntimeError(
            f"安装后的 NumPy 版本错误：{numpy.__version__} != {NUMPY_VERSION}"
        )
    configuration = numpy.show_config(mode="dicts")
    if not isinstance(configuration, Mapping) or not configuration_uses_internal_fallback(
        configuration
    ):
        raise MacX64RuntimeError("安装后的 NumPy 仍依赖外部 BLAS/LAPACK")
    if float(numpy.linalg.norm(numpy.asarray([3.0, 4.0]))) != 5.0:
        raise MacX64RuntimeError("NumPy 内部线性代数回退自检失败")

    package_root = Path(numpy.__file__).resolve().parent
    if (package_root / "numpy.libs").exists():
        raise MacX64RuntimeError("安装后的 NumPy 仍包含外部运行库目录 numpy.libs")
    for binary in sorted(package_root.rglob("*.so")):
        result = subprocess.run(
            ["otool", "-L", str(binary)],
            check=False,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        if result.returncode != 0:
            raise MacX64RuntimeError(f"无法审计 NumPy Mach-O 依赖：{binary}")
        if "openblas" in result.stdout.casefold():
            raise MacX64RuntimeError(f"NumPy 扩展仍链接 OpenBLAS：{binary}")


def prepare_runtime(wheel_dir: Path) -> Path:
    require_native_intel_macos()
    wheel_dir = _safe_prepare_wheel_dir(wheel_dir)
    environment = os.environ.copy()
    environment["MACOSX_DEPLOYMENT_TARGET"] = MACOS_DEPLOYMENT_TARGET
    environment["ARCHFLAGS"] = "-arch x86_64"
    _run(
        [
            sys.executable,
            "-m",
            "pip",
            "wheel",
            "--disable-pip-version-check",
            "--no-deps",
            "--no-binary=numpy",
            "--no-cache-dir",
            "--wheel-dir",
            str(wheel_dir),
            "-Csetup-args=-Dblas=none",
            "-Csetup-args=-Dlapack=none",
            "-Csetup-args=-Dallow-noblas=true",
            f"numpy=={NUMPY_VERSION}",
        ],
        env=environment,
    )
    wheels = sorted(wheel_dir.glob("numpy-*.whl"))
    if len(wheels) != 1:
        raise MacX64RuntimeError(f"预期生成一个 NumPy wheel，实际为：{wheels}")
    wheel_path = wheels[0]
    validate_compact_wheel(wheel_path)
    _run(
        [
            sys.executable,
            "-m",
            "pip",
            "install",
            "--disable-pip-version-check",
            "--force-reinstall",
            "--no-deps",
            str(wheel_path),
        ]
    )
    verify_installed_runtime()
    _run([sys.executable, "-m", "pip", "check"])
    print(
        "Intel macOS 紧凑 NumPy 运行时验证通过："
        f"{wheel_path.name}，{wheel_path.stat().st_size} 字节，SHA256={_sha256(wheel_path)}",
        flush=True,
    )
    return wheel_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="准备并验证 Intel macOS 紧凑发布运行时")
    parser.add_argument(
        "--wheel-dir",
        type=Path,
        default=REPO_ROOT / "build" / "macos-x64-runtime-wheels",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    prepare_runtime(args.wheel_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

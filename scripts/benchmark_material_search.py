from __future__ import annotations

import argparse
import hashlib
import json
import statistics
import sys
import tempfile
import time
from pathlib import Path

from openpyxl import Workbook


REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from hr_toolkit.tools import material_collector as mc


def _employee_name(index: int) -> str:
    # Three CJK characters keep the fixture inside the production name parser.
    return "测" + chr(0x4E00 + index // 80) + chr(0x4E00 + index % 80)


def _id_card(index: int) -> str:
    return f"11010119900101{index:04d}"


def _build_roster(path: Path, employee_count: int) -> None:
    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet("人员名单")
    worksheet.append(("姓名", "身份证号码"))
    for index in range(employee_count):
        worksheet.append((_employee_name(index), _id_card(index)))
    workbook.save(path)


def _build_cached_library(
    library: Path,
    cache_path: Path,
    *,
    employee_count: int,
    file_count: int,
) -> None:
    cache = mc._new_ocr_cache()
    source_dir = library / "平铺资料"
    source_dir.mkdir(parents=True)
    for index in range(file_count):
        employee_index = index % employee_count
        relative = f"平铺资料/{index:06d}.jpg"
        source = library / relative
        payload = index.to_bytes(8, "little")
        source.write_bytes(payload)
        stat = source.stat()
        digest = hashlib.sha256(payload).hexdigest()
        cache_key = f"{digest}_{len(payload)}"
        cache["entries"][cache_key] = {
            "content_hash": digest,
            "source_size": len(payload),
            "source_mtime": stat.st_mtime,
            "material_type": "身份证",
            "match_method": "ocr_id_front",
            "subtype": "正面",
            "extracted_name": _employee_name(employee_index),
            "extracted_names": [_employee_name(employee_index)],
            "extracted_id_hash": mc._hash_id_card(_id_card(employee_index)),
            "extracted_phone_hash": "",
            "ocr_text": "",
            "verified_at": mc._beijing_now_str(),
            "sample_filename": source.name,
            "source_relpath": relative,
            "analysis_state": "complete",
            "index_scope": mc.LIBRARY_MODE_FLAT_OCR,
        }
        cache["paths"][relative] = {
            "cache_key": cache_key,
            "source_size": stat.st_size,
            "source_mtime_ns": stat.st_mtime_ns,
            "source_change_token": mc._file_change_token(source, stat),
        }
    if not mc._save_ocr_cache(cache_path, cache):
        raise RuntimeError("无法建立压力测试 OCR 缓存。")


def _peak_rss_bytes() -> int:
    try:
        import resource

        peak = int(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
        return peak if sys.platform == "darwin" else peak * 1024
    except ImportError:
        try:
            import psutil

            return int(psutil.Process().memory_info().rss)
        except Exception:
            return 0


def run_benchmark(*, employee_count: int, file_count: int, repeat: int) -> dict[str, object]:
    if not 1 <= employee_count <= 10_000:
        raise ValueError("employee_count 必须在 1 到 10000 之间。")
    if not employee_count <= file_count <= 50_000:
        raise ValueError("file_count 必须不少于员工数且不超过 50000。")
    if repeat <= 0:
        raise ValueError("repeat 必须大于 0。")

    with tempfile.TemporaryDirectory(prefix="hr_material_benchmark_") as temporary:
        root = Path(temporary)
        library = root / "资料库"
        library.mkdir()
        roster = root / "人员名单.xlsx"
        cache_path = library / mc._OCR_CACHE_FILE_NAME

        prepare_started = time.perf_counter()
        _build_roster(roster, employee_count)
        _build_cached_library(
            library,
            cache_path,
            employee_count=employee_count,
            file_count=file_count,
        )
        prepare_seconds = time.perf_counter() - prepare_started

        durations: list[float] = []
        match_counts: list[int] = []
        for run_index in range(repeat):
            output = root / f"output-{run_index}"
            started = time.perf_counter()
            result = mc.collect_employee_materials(
                library,
                output,
                roster_source=roster,
                material_types=["身份证"],
                library_mode=mc.LIBRARY_MODE_FLAT_OCR,
                collect_all=True,
                generate_report=False,
                use_ocr_cache=True,
                ocr_cache_path=cache_path,
            )
            durations.append(time.perf_counter() - started)
            match_counts.append(len(result.matches))
            if len(result.target_employees) != employee_count:
                raise RuntimeError("名单解析数量异常。")
            if len(result.matches) != file_count:
                raise RuntimeError("缓存 OCR 匹配数量异常。")

        return {
            "employees": employee_count,
            "image_files": file_count,
            "repeat": repeat,
            "fixture_prepare_seconds": round(prepare_seconds, 4),
            "durations_seconds": [round(value, 4) for value in durations],
            "median_seconds": round(statistics.median(durations), 4),
            "matches": match_counts,
            "peak_rss_bytes": _peak_rss_bytes(),
            "ocr_mode": "prebuilt-cache-hit",
        }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="压力测试数万图片、数千人名单的 OCR 缓存扫描与匹配链路。"
    )
    parser.add_argument("--employees", type=int, default=5000)
    parser.add_argument("--files", type=int, default=10000)
    parser.add_argument("--repeat", type=int, default=2)
    args = parser.parse_args()
    print(
        json.dumps(
            run_benchmark(
                employee_count=args.employees,
                file_count=args.files,
                repeat=args.repeat,
            ),
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()

"""Excel comparison and safe export, independent of the desktop interface."""

import os
import tempfile
from pathlib import Path

import pandas as pd


def _unique(columns):
    return list(dict.fromkeys(columns))


def compare_tables(source, target, source_columns, target_columns, keys, operation):
    """Compare by common keys; preserve unmatched source rows in a difference.

    Intersection keeps every matching pair, including many-to-many matches.
    Difference returns selected source columns only. Null keys follow pandas'
    merge semantics: null matches null.
    """
    if operation not in ("intersection", "difference"):
        raise ValueError("未知的比较操作。")
    keys = _unique(keys)
    source_columns = _unique(source_columns)
    target_columns = _unique(target_columns)
    if not keys:
        raise ValueError("请至少选择一列作为共同筛选列。")
    if not source.columns.is_unique or not target.columns.is_unique:
        raise ValueError("表头包含重复列名，请先为每一列设置唯一名称。")
    if not source_columns and (operation == "difference" or not target_columns):
        raise ValueError("请至少选择一列输出列；差集需要选择源表输出列。")
    for frame, requested, name in (
        (source, _unique(source_columns + keys), "源表"),
        (target, _unique((target_columns if operation == "intersection" else []) + keys), "目标表"),
    ):
        missing = [column for column in requested if column not in frame.columns]
        if missing:
            raise ValueError(f"{name}缺少列：{', '.join(map(str, missing))}。请重新选择文件和列。")

    if operation == "difference":
        # A unique marker avoids collisions with user-provided column labels.
        marker = "_comparison_status"
        while marker in keys:
            marker += "_"
        target_keys = target.loc[:, keys].drop_duplicates().copy()
        target_keys[marker] = True
        membership = source.loc[:, keys].merge(
            target_keys, on=keys, how="left", sort=False, validate="many_to_one",
        )
        unmatched = membership[marker].isna().to_numpy()
        return source.loc[unmatched, source_columns].reset_index(drop=True)

    overlap = (set(source_columns) & set(target_columns)) - set(keys)
    output_columns = [f"{c}_源表" if c in overlap else c for c in source_columns]
    output_columns.extend(
        f"{c}_目标表" if c in overlap else c
        for c in target_columns if c not in keys or c not in source_columns
    )
    if len(set(output_columns)) != len(output_columns):
        raise ValueError("输出列名发生冲突，请调整列选择或修改原始表头。")
    merged = source.loc[:, _unique(source_columns + keys)].merge(
        target.loc[:, _unique(target_columns + keys)],
        on=keys, how="inner", suffixes=("_源表", "_目标表"), sort=False,
    )
    return merged.loc[:, output_columns].reset_index(drop=True)


def validate_destination(destination, input_paths):
    """Reject input aliases, including symlinks and existing hard links."""
    destination = Path(destination).expanduser().resolve()
    if destination.suffix.lower() != ".xlsx":
        raise ValueError("结果文件必须使用 .xlsx 扩展名。")
    for input_path in input_paths:
        input_path = Path(input_path).expanduser().resolve()
        same_path = os.path.normcase(str(destination)) == os.path.normcase(str(input_path))
        same_file = False
        if destination.exists() and input_path.exists():
            same_file = os.path.samefile(destination, input_path)
        if same_path or same_file:
            raise ValueError("保存位置不能与任一输入文件相同，请选择其他文件名。")
    if not destination.parent.is_dir():
        raise ValueError("保存目录不存在，请选择有效目录。")
    if destination.exists() and not destination.is_file():
        raise ValueError("保存位置是目录，请选择文件名。")
    return destination


def save_result(result, destination, input_paths, *, overwrite=False):
    """Finish writing a temporary workbook before replacing the destination."""
    destination = validate_destination(destination, input_paths)
    if destination.exists() and not overwrite:
        raise FileExistsError("结果文件已存在，请确认覆盖或选择其他文件名。")
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(
            suffix=".xlsx", prefix=".excel-result-", dir=destination.parent, delete=False,
        ) as handle:
            temporary = Path(handle.name)
        result.to_excel(temporary, index=False, engine="openpyxl")
        if overwrite:
            os.replace(temporary, destination)
        elif os.name == "nt":
            # Windows rename refuses to replace an existing destination.
            os.rename(temporary, destination)
        else:
            # POSIX rename replaces files; linking instead provides no-clobber.
            os.link(temporary, destination)
        return destination
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)

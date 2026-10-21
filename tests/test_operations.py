import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from pandas.testing import assert_frame_equal

from excel_operations import compare_tables, save_result


class ComparisonTests(unittest.TestCase):
    def setUp(self):
        self.source = pd.DataFrame({"id": [1, 1, 2], "value": ["a", "b", "c"], "secret": ["x", "y", "z"]})
        self.target = pd.DataFrame({"id": [2], "other": ["d"]})

    def compare(self, operation, source_columns=("id", "value"), target_columns=("id", "other"), keys=("id",)):
        return compare_tables(self.source, self.target, source_columns, target_columns, keys, operation)

    def test_intersection_excludes_unmatched_rows(self):
        result = self.compare("intersection")
        assert_frame_equal(result, pd.DataFrame({"id": [2], "value": ["c"], "other": ["d"]}))

    def test_difference_preserves_unmatched_duplicate_keys_and_order(self):
        result = self.compare("difference")
        assert_frame_equal(result, self.source.loc[:1, ["id", "value"]])

    def test_difference_excludes_unselected_columns(self):
        result = self.compare("difference", source_columns=["value"])
        self.assertEqual(result.columns.tolist(), ["value"])
        self.assertEqual(result["value"].tolist(), ["a", "b"])

    def test_keys_need_not_be_exported(self):
        result = self.compare("intersection", ["value"], ["other"])
        assert_frame_equal(result, pd.DataFrame({"value": ["c"], "other": ["d"]}))

    def test_target_only_export_is_supported(self):
        result = self.compare("intersection", [], ["other"])
        assert_frame_equal(result, pd.DataFrame({"other": ["d"]}))

    def test_duplicate_target_keys_do_not_multiply_difference(self):
        self.target = pd.DataFrame({"id": [2, 2], "other": ["d", "e"]})
        self.assertEqual(self.compare("difference")["id"].tolist(), [1, 1])

    def test_many_to_many_intersection_preserves_pairs(self):
        self.target = pd.DataFrame({"id": [1, 1, 1], "other": ["d", "e", "f"]})
        result = self.compare("intersection")
        self.assertEqual(len(result), 6)
        self.assertEqual(set(zip(result["value"], result["other"])), {
            (a, b) for a in ("a", "b") for b in ("d", "e", "f")
        })

    def test_composite_keys(self):
        self.source["group"] = ["x", "y", "z"]
        self.target = pd.DataFrame({"id": [1], "group": ["x"]})
        result = self.compare("difference", ["value"], [], ["id", "group"])
        self.assertEqual(result["value"].tolist(), ["b", "c"])

    def test_numeric_column_labels(self):
        source = pd.DataFrame({2025: [1, 2], "name": ["a", "b"]})
        target = pd.DataFrame({2025: [2], "extra": [3]})
        result = compare_tables(source, target, ["name"], ["extra"], [2025], "intersection")
        assert_frame_equal(result, pd.DataFrame({"name": ["b"], "extra": [3]}), check_column_type=False)

    def test_colliding_payload_names_have_explicit_suffixes(self):
        self.target = pd.DataFrame({"id": [2], "value": ["target"]})
        result = self.compare("intersection", ["value"], ["value"])
        assert_frame_equal(result, pd.DataFrame({"value_源表": ["c"], "value_目标表": ["target"]}))

    def test_suffix_collision_is_reported(self):
        self.source["value_源表"] = "existing"
        self.target = pd.DataFrame({"id": [2], "value": ["target"]})
        with self.assertRaisesRegex(ValueError, "冲突"):
            self.compare("intersection", ["value", "value_源表"], ["value"])

    def test_null_keys_match_null_and_duplicates_remain(self):
        self.source = pd.DataFrame({"id": [None, 1, 1], "value": ["null", "a", "b"]})
        self.target = pd.DataFrame({"id": [float("nan")], "other": ["null-target"]})
        self.assertEqual(self.compare("difference")["value"].tolist(), ["a", "b"])

    def test_non_default_duplicate_index_does_not_change_difference(self):
        self.source.index = [5, 5, 1]
        self.assertEqual(self.compare("difference")["value"].tolist(), ["a", "b"])

    def test_empty_target_keeps_every_source_row(self):
        self.target = self.target.iloc[:0]
        self.assertEqual(self.compare("difference")["value"].tolist(), ["a", "b", "c"])
        self.assertTrue(self.compare("intersection").empty)

    def test_empty_source_exports_expected_columns(self):
        self.source = self.source.iloc[:0]
        result = self.compare("difference")
        self.assertTrue(result.empty)
        self.assertEqual(result.columns.tolist(), ["id", "value"])

    def test_indicator_named_key_is_supported(self):
        source = pd.DataFrame({"_comparison_status": [1, 2]})
        target = pd.DataFrame({"_comparison_status": [2]})
        result = compare_tables(source, target, source.columns, [], source.columns, "difference")
        self.assertEqual(result.iloc[:, 0].tolist(), [1])

    def test_pandas_internal_indicator_labels_can_be_keys(self):
        for key in ("_left_indicator", "_right_indicator"):
            source = pd.DataFrame({key: [1, 1, 2]})
            target = pd.DataFrame({key: [2]})
            result = compare_tables(source, target, [key], [], [key], "difference")
            self.assertEqual(result[key].tolist(), [1, 1])

    def test_missing_keys_and_missing_columns_are_reported(self):
        with self.assertRaisesRegex(ValueError, "筛选"):
            self.compare("intersection", keys=[])
        with self.assertRaisesRegex(ValueError, "缺少列"):
            self.compare("intersection", keys=["missing"])
        with self.assertRaisesRegex(ValueError, "输出列"):
            self.compare("difference", source_columns=[])

    def test_repeated_selection_is_deduplicated(self):
        result = self.compare("intersection", ["id", "id"], ["id", "id"], ["id", "id"])
        assert_frame_equal(result, pd.DataFrame({"id": [2]}))


class ExportTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)
        self.source = self.directory / "source.xlsx"
        self.target = self.directory / "target.xlsx"
        self.original = pd.DataFrame({"id": [1, 2], "secret": ["a", "b"]})
        self.original.to_excel(self.source, index=False)
        self.original.to_excel(self.target, index=False)
        self.result = pd.DataFrame({"id": [1]})
        self.inputs = [self.source, self.target]

    def test_round_trip_export(self):
        output = self.directory / "result.xlsx"
        save_result(self.result, output, self.inputs)
        assert_frame_equal(pd.read_excel(output), self.result)
        self.assertFalse(list(self.directory.glob(".excel-result-*")))

    def test_each_input_is_protected_even_with_overwrite(self):
        for path in self.inputs:
            before = path.read_bytes()
            with self.assertRaisesRegex(ValueError, "输入文件"):
                save_result(self.result, path, self.inputs, overwrite=True)
            self.assertEqual(path.read_bytes(), before)

    def test_existing_results_require_explicit_overwrite(self):
        output = self.directory / "result.xlsx"
        output.write_bytes(b"old result")
        with self.assertRaises(FileExistsError):
            save_result(self.result, output, self.inputs)
        self.assertEqual(output.read_bytes(), b"old result")
        save_result(self.result, output, self.inputs, overwrite=True)
        assert_frame_equal(pd.read_excel(output), self.result)

    def test_failed_write_preserves_existing_file_and_cleans_temporary(self):
        output = self.directory / "result.xlsx"
        output.write_bytes(b"old result")
        with patch.object(pd.DataFrame, "to_excel", side_effect=OSError("disk full")):
            with self.assertRaises(OSError):
                save_result(self.result, output, self.inputs, overwrite=True)
        self.assertEqual(output.read_bytes(), b"old result")
        self.assertFalse(list(self.directory.glob(".excel-result-*")))

    def test_a_file_created_during_export_is_not_overwritten(self):
        output = self.directory / "result.xlsx"
        real_writer = pd.DataFrame.to_excel

        def write_and_create(frame, path, **kwargs):
            real_writer(frame, path, **kwargs)
            output.write_bytes(b"concurrent file")

        with patch.object(pd.DataFrame, "to_excel", write_and_create):
            with self.assertRaises(FileExistsError):
                save_result(self.result, output, self.inputs)
        self.assertEqual(output.read_bytes(), b"concurrent file")
        self.assertFalse(list(self.directory.glob(".excel-result-*")))

    def test_hard_link_to_input_is_protected(self):
        alias = self.directory / "alias.xlsx"
        try:
            os.link(self.source, alias)
        except OSError:
            self.skipTest("Hard links are unavailable on this filesystem")
        with self.assertRaisesRegex(ValueError, "输入文件"):
            save_result(self.result, alias, self.inputs, overwrite=True)

    def test_invalid_extensions_and_missing_directories_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "xlsx"):
            save_result(self.result, self.directory / "result.xls", self.inputs)
        with self.assertRaisesRegex(ValueError, "目录"):
            save_result(self.result, self.directory / "missing" / "result.xlsx", self.inputs)


if __name__ == "__main__":
    unittest.main()


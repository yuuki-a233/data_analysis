import importlib.util
import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import pandas as pd

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
QT_AVAILABLE = importlib.util.find_spec("PyQt5") is not None
if QT_AVAILABLE:
    from PyQt5.QtCore import Qt
    from PyQt5.QtWidgets import QApplication, QMessageBox
    import test as app_module


@unittest.skipUnless(QT_AVAILABLE, "Install PyQt5 to run GUI regression tests")
class WindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)
        self.source = self.directory / "source.xlsx"
        self.target = self.directory / "target.xlsx"
        pd.DataFrame({"id": [1, 1, 2], "name": ["a", "b", "c"]}).to_excel(self.source, index=False)
        pd.DataFrame({"id": [2], "extra": ["d"]}).to_excel(self.target, index=False)
        self.paths = [str(self.source), str(self.target)]
        self.window = app_module.MainWindow()
        self.addCleanup(self.window.close)
        self.critical = patch.object(app_module.QMessageBox, "critical").start()
        self.warning = patch.object(app_module.QMessageBox, "warning").start()
        self.information = patch.object(app_module.QMessageBox, "information").start()
        self.addCleanup(patch.stopall)

    def choose(self, paths):
        with patch.object(app_module.QFileDialog, "getOpenFileNames", return_value=(paths, "")):
            self.window.Select_multiple_files()

    def comparison(self):
        self.choose(self.paths)
        self.window.open_newwindow()
        view = self.window.new_window
        view.listwidget_3.item(0).setSelected(True)
        return view

    def test_cancel_preserves_headers_and_open_window(self):
        view = self.comparison()
        before = list(self.window.source_columns)
        with patch.object(app_module.pd, "read_excel") as reader:
            self.choose([])
        reader.assert_not_called()
        self.assertEqual(self.window.source_columns, before)
        self.assertIs(self.window.new_window, view)

    def test_pair_replaces_one_file_instead_of_appending_three(self):
        self.choose(self.paths[:1])
        self.choose(list(reversed(self.paths)))
        self.assertEqual(self.window.files, list(reversed(self.paths)))
        self.assertEqual(self.window.source_columns, ["id", "extra"])
        self.assertEqual(self.window.qlist.count(), 2)

    def test_two_single_selections_form_a_pair(self):
        self.choose(self.paths[:1])
        self.choose(self.paths[1:])
        self.assertEqual(self.window.files, self.paths)
        self.assertEqual(self.window.target_columns, ["id", "extra"])

    def test_single_file_replaces_completed_pair(self):
        self.choose(self.paths)
        self.choose(self.paths[1:])
        self.assertEqual(self.window.files, self.paths[1:])
        self.assertEqual(self.window.target_columns, [])

    def test_too_many_files_preserves_previous_selection(self):
        self.choose(self.paths)
        self.choose(self.paths + self.paths[:1])
        self.assertEqual(self.window.files, self.paths)
        self.warning.assert_called_once()

    def test_failed_load_rolls_back_files_and_headers(self):
        view = self.comparison()
        with patch.object(app_module.pd, "read_excel", side_effect=ValueError("invalid workbook")):
            self.choose(list(reversed(self.paths)))
        self.assertEqual(self.window.files, self.paths)
        self.assertEqual(self.window.source_columns, ["id", "name"])
        self.assertIs(self.window.new_window, view)
        self.critical.assert_called_once()

    def test_clear_closes_comparison_and_resets_state(self):
        view = self.comparison()
        self.window.clear_list()
        self.assertEqual(self.window.files, [])
        self.assertEqual(self.window.source_columns, [])
        self.assertEqual(self.window.target_columns, [])
        self.assertIsNone(self.window.new_window)
        self.assertFalse(view.isVisible())
        self.assertEqual(self.window.qlist.count(), 0)

    def test_reopened_window_has_no_stale_key_selection(self):
        previous = self.comparison()
        previous.listwidget.item(0).setSelected(False)
        self.window.open_newwindow()
        current = self.window.new_window
        self.assertEqual(current._selected(current.listwidget_3), [])
        self.assertEqual(current._selected(current.listwidget), ["id", "name"])

    def test_main_windows_do_not_share_file_state(self):
        self.choose(self.paths)
        other = app_module.MainWindow()
        self.addCleanup(other.close)
        self.assertEqual(other.files, [])

    def test_non_text_labels_and_identical_display_text_preserve_types(self):
        date = datetime(2025, 1, 1)
        view = app_module.NewWindow(self.paths, [2025, "2025", date], [2025, "2025", date])
        self.addCleanup(view.close)
        selected = view._selected(view.listwidget)
        self.assertEqual(selected, [2025, "2025", date])
        self.assertIsInstance(selected[0], int)
        self.assertIsInstance(selected[1], str)
        self.assertEqual(view.listwidget.item(0).text(), "2025")
        self.assertEqual(view.listwidget.item(1).text(), "2025")

    def test_intersection_export_uses_current_output_selection(self):
        view = self.comparison()
        view.listwidget.item(0).setSelected(False)
        view.listwidget_2.item(0).setSelected(False)
        output = self.directory / "intersection"
        with patch.object(app_module.QFileDialog, "getSaveFileName", return_value=(str(output), "")):
            view.overlap()
        result = pd.read_excel(output.with_suffix(".xlsx"))
        self.assertEqual(result.columns.tolist(), ["name", "extra"])
        self.assertEqual(result.to_dict("list"), {"name": ["c"], "extra": ["d"]})
        self.critical.assert_not_called()

    def test_difference_export_preserves_duplicates(self):
        view = self.comparison()
        view.listwidget.item(0).setSelected(False)
        output = self.directory / "difference.xlsx"
        with patch.object(app_module.QFileDialog, "getSaveFileName", return_value=(str(output), "")):
            view.filt()
        result = pd.read_excel(output)
        self.assertEqual(result.to_dict("list"), {"name": ["a", "b"]})

    def test_export_rejects_input_path_before_overwrite_prompt(self):
        view = self.comparison()
        original = self.source.read_bytes()
        with patch.object(app_module.QFileDialog, "getSaveFileName", return_value=(str(self.source), "")), patch.object(app_module.QMessageBox, "question") as question:
            view.overlap()
        question.assert_not_called()
        self.critical.assert_called_once()
        self.assertEqual(self.source.read_bytes(), original)

    def test_declined_overwrite_preserves_existing_result(self):
        view = self.comparison()
        output = self.directory / "result.xlsx"
        output.write_bytes(b"old")
        with patch.object(app_module.QFileDialog, "getSaveFileName", return_value=(str(output), "")), patch.object(app_module.QMessageBox, "question", return_value=QMessageBox.No):
            view.overlap()
        self.assertEqual(output.read_bytes(), b"old")
        self.information.assert_not_called()

    def test_confirmed_overwrite_exports_result(self):
        view = self.comparison()
        output = self.directory / "result.xlsx"
        output.write_bytes(b"old")
        with patch.object(app_module.QFileDialog, "getSaveFileName", return_value=(str(output), "")), patch.object(app_module.QMessageBox, "question", return_value=QMessageBox.Yes):
            view.overlap()
        self.assertEqual(pd.read_excel(output)["id"].tolist(), [2])
        self.information.assert_called_once()

    def test_save_dialog_cancel_does_not_write(self):
        view = self.comparison()
        with patch.object(app_module.QFileDialog, "getSaveFileName", return_value=("", "")), patch.object(app_module, "save_result") as save:
            view.overlap()
        save.assert_not_called()
        self.information.assert_not_called()

    def test_missing_key_and_incomplete_files_are_reported(self):
        self.window.open_newwindow()
        self.warning.assert_called_once()
        self.choose(self.paths)
        self.window.open_newwindow()
        with patch.object(app_module.QFileDialog, "getSaveFileName") as dialog:
            self.window.new_window.overlap()
        dialog.assert_not_called()
        self.critical.assert_called_once()


if __name__ == "__main__":
    unittest.main()


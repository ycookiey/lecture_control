import sys
import os
import json
import shutil
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QFileDialog,
    QMessageBox,
    QInputDialog,
)
from PyQt6.QtCore import Qt, QTimer
import win32com.client


class TimeTableApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("授業フォルダ、ショートカット生成")
        self.timetables = {}
        self.current_timetable = "default"
        self.auto_save_timer = QTimer(self)
        self.auto_save_timer.timeout.connect(self.auto_save)
        self.auto_save_timer.start(60000)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)

        folder_layout = QHBoxLayout()
        self.summary_folder = QLineEdit()
        self.shortcut_folder = QLineEdit()
        summary_button = QPushButton("選択...")
        shortcut_button = QPushButton("選択...")
        summary_button.clicked.connect(
            lambda: self.select_folder(
                self.summary_folder, "授業フォルダの配置場所を選択"
            )
        )
        shortcut_button.clicked.connect(
            lambda: self.select_folder(
                self.shortcut_folder, "ショートカットの配置場所を選択"
            )
        )

        folder_layout.addWidget(QLabel("授業フォルダディレクトリ:"))
        folder_layout.addWidget(self.summary_folder)
        folder_layout.addWidget(summary_button)
        folder_layout.addWidget(QLabel("ショートカットディレクトリ:"))
        folder_layout.addWidget(self.shortcut_folder)
        folder_layout.addWidget(shortcut_button)
        main_layout.addLayout(folder_layout)

        self.timetable = QTableWidget(5, 5)
        self.timetable.setHorizontalHeaderLabels(["月", "火", "水", "木", "金"])
        self.timetable.setVerticalHeaderLabels(["1", "2", "3", "4", "5"])
        self.timetable.setMinimumHeight(300)
        self.timetable.itemChanged.connect(self.on_timetable_changed)
        main_layout.addWidget(self.timetable)

        button_layout = QHBoxLayout()
        generate_button = QPushButton("フォルダとショートカットを生成")
        generate_button.clicked.connect(self.generate_folders_and_shortcuts)
        save_button = QPushButton("名前をつけて保存")
        save_button.clicked.connect(self.save_timetable)
        load_button = QPushButton("読み込み")
        load_button.clicked.connect(self.load_timetable)
        clear_shortcut_button = QPushButton("ショートカットフォルダをクリア")
        clear_shortcut_button.clicked.connect(self.clear_shortcut_folder)
        button_layout.addWidget(generate_button)
        button_layout.addWidget(save_button)
        button_layout.addWidget(load_button)
        button_layout.addWidget(clear_shortcut_button)
        main_layout.addLayout(button_layout)

        self.load_timetables()
        self.auto_load_last_timetable()

    def select_folder(self, line_edit, caption):
        folder = QFileDialog.getExistingDirectory(self, caption)
        if folder:
            line_edit.setText(folder)
            self.auto_save()

    def generate_folders_and_shortcuts(self):
        summary_folder = self.summary_folder.text()
        shortcut_folder = self.shortcut_folder.text()

        if not summary_folder or not shortcut_folder:
            QMessageBox.warning(
                self,
                "エラー",
                "フォルダを指定してください。",
            )
            return

        shell = win32com.client.Dispatch("WScript.Shell")

        for row in range(5):
            for col in range(5):
                item = self.timetable.item(row, col)
                if item and item.text():
                    class_name = item.text()
                    period = str(row + 1)
                    day = ["月", "火", "水", "木", "金"][col]

                    class_folder = os.path.join(summary_folder, class_name)
                    os.makedirs(class_folder, exist_ok=True)

                    shortcut_name = f"{period}.{class_name}.lnk"
                    shortcut_path = os.path.join(shortcut_folder, shortcut_name)

                    shortcut = shell.CreateShortCut(shortcut_path)
                    shortcut.TargetPath = class_folder
                    shortcut.WorkingDirectory = class_folder
                    shortcut.Description = f"{col+1}{day}"
                    shortcut.IconLocation = (
                        f"{os.environ['SystemRoot']}\\System32\\shell32.dll,3"
                    )
                    shortcut.save()

        QMessageBox.information(
            self, "完了", "フォルダとショートカットの生成が完了しました。"
        )

    def save_timetable(self):
        name, ok = QInputDialog.getText(
            self, "時間割の保存", "時間割の名前を入力してください:"
        )
        if ok and name:
            self.current_timetable = name
            self.auto_save()
            QMessageBox.information(
                self, "保存完了", f"時間割 '{name}' を保存しました。"
            )

    def load_timetable(self):
        names = list(self.timetables.keys())
        if not names:
            QMessageBox.warning(self, "エラー", "保存された時間割がありません。")
            return
        name, ok = QInputDialog.getItem(
            self, "時間割の読み込み", "時間割を選択してください:", names, 0, False
        )
        if ok and name:
            self.load_timetable_by_name(name)

    def load_timetable_by_name(self, name):
        if name in self.timetables:
            timetable_data = self.timetables[name]
            self.summary_folder.setText(timetable_data["summary_folder"])
            self.shortcut_folder.setText(timetable_data["shortcut_folder"])
            self.timetable.blockSignals(True)
            self.timetable.clearContents()
            for class_data in timetable_data["classes"]:
                self.timetable.setItem(
                    class_data["row"],
                    class_data["col"],
                    QTableWidgetItem(class_data["name"]),
                )
            self.timetable.blockSignals(False)
            self.current_timetable = name
        else:
            QMessageBox.warning(self, "エラー", f"時間割 '{name}' が見つかりません。")

    def save_timetables(self):
        data_to_save = {
            "timetables": self.timetables,
            "last_used": self.current_timetable,
        }
        with open("timetables.json", "w", encoding="utf-8") as f:
            json.dump(data_to_save, f, ensure_ascii=False, indent=2)

    def load_timetables(self):
        try:
            with open("timetables.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                self.timetables = data.get("timetables", {})
                self.current_timetable = data.get("last_used", "default")
        except FileNotFoundError:
            self.timetables = {}
            self.current_timetable = "default"

    def on_timetable_changed(self, item):
        self.auto_save()

    def auto_save(self):
        timetable_data = {
            "summary_folder": self.summary_folder.text(),
            "shortcut_folder": self.shortcut_folder.text(),
            "classes": [],
        }
        for row in range(5):
            for col in range(5):
                item = self.timetable.item(row, col)
                if item and item.text():
                    timetable_data["classes"].append(
                        {"name": item.text(), "row": row, "col": col}
                    )
        self.timetables[self.current_timetable] = timetable_data
        self.save_timetables()

    def auto_load_last_timetable(self):
        if self.current_timetable in self.timetables:
            self.load_timetable_by_name(self.current_timetable)
        elif self.timetables:
            last_timetable = list(self.timetables.keys())[-1]
            self.load_timetable_by_name(last_timetable)

    def clear_shortcut_folder(self):
        shortcut_folder = self.shortcut_folder.text()
        if not shortcut_folder:
            QMessageBox.warning(
                self, "エラー", "ショートカットフォルダが指定されていません。"
            )
            return

        reply = QMessageBox.question(
            self,
            "確認",
            "ショートカットフォルダ内のすべてのファイルとフォルダを削除しますか？\n"
            "この操作は元に戻せません。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            try:
                for filename in os.listdir(shortcut_folder):
                    file_path = os.path.join(shortcut_folder, filename)
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                QMessageBox.information(
                    self, "完了", "ショートカットフォルダをクリアしました。"
                )
            except Exception as e:
                QMessageBox.warning(
                    self,
                    "エラー",
                    f"ショートカットフォルダのクリア中にエラーが発生しました：{str(e)}",
                )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TimeTableApp()
    window.show()
    sys.exit(app.exec())

import sys
import os
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
)
from PyQt6.QtCore import Qt
import win32com.client


class TimeTableApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("授業フォルダ、ショートカット生成")

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
        main_layout.addWidget(self.timetable)

        generate_button = QPushButton("フォルダとショートカットを生成")
        generate_button.clicked.connect(self.generate_folders_and_shortcuts)
        main_layout.addWidget(generate_button)

    def select_folder(self, line_edit, caption):
        folder = QFileDialog.getExistingDirectory(self, caption)
        if folder:
            line_edit.setText(folder)

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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TimeTableApp()
    window.show()
    sys.exit(app.exec())

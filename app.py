# Copyright (c) 2026 Akshith Vuppala
# Licensed under the MIT License.
# See LICENSE file in the project root for full license information.

import os
import sys

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

# Import your logic from the other file
# from logic import run_calculation
from logic import audit_excel

os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Auditing files")
        self.setMinimumWidth(500)

        # 1. Main Vertical Layout (Top to Bottom)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 2. Instructions Section
        instructions = QLabel(
            "### Usage Instructions\n"
            "1. **File Selection:** Ensure both input files are selected and accessible.\n"
            "2. **Data Consistency:** All sheets within the files must share the **same header row** and index.\n"
            "3. **Header Formatting:** Headers for **Reference column** must have exact spaces and dots (e.g., 'Serial. No' vs 'Serial No'). Case sensitivity is ignored.\n"
            "4. **File Access:** Close the files in Excel/LibreOffice before clicking execute to avoid 'Permission Denied' errors.\n"
            "5. **Value Constraints:** Ensure integer values are within the allowed range to prevent calculation overflows.\n"
            "6. **Output:** Choose an output path where you have write permissions.\n"
            "7. **Amount** : Make sure that **Credit** and **Debit** Columns have no special characters (e.g., '.' or '*')\n\n"
            "**Click 'Execute Logic' to begin the process.**"
        )
        instructions.setTextFormat(Qt.MarkdownText)
        instructions.setWordWrap(True)
        main_layout.addWidget(instructions)

        # 3. Input Grid (2 columns: Label | Input)
        grid = QGridLayout()
        main_layout.addLayout(grid)

        # -- Filepath 1 --
        grid.addWidget(QLabel("Buyer Excel Sheet:"), 0, 0)
        self.path1_edit = QLineEdit()
        self.path1_edit.setText("./sheets/CARE.xlsx")
        grid.addWidget(self.path1_edit, 0, 1)
        btn_browse1 = QPushButton("Browse")
        btn_browse1.clicked.connect(lambda: self.get_file(self.path1_edit))
        grid.addWidget(btn_browse1, 0, 2)

        grid.addWidget(QLabel("Buyer Header Row index"), 1, 0)
        self.int1 = QLineEdit()
        self.int1.setText("12")
        grid.addWidget(self.int1, 1, 1)

        grid.addWidget(QLabel("Buyer Document Number Header"), 2, 0)
        self.str1 = QLineEdit()
        self.str1.setText("Vch No.")
        grid.addWidget(self.str1, 2, 1)

        # -- Filepath 2 --
        grid.addWidget(QLabel("Seller Excel Sheet"), 3, 0)
        self.path2_edit = QLineEdit()
        self.path2_edit.setText("./sheets/jiva.xlsx")
        grid.addWidget(self.path2_edit, 3, 1)
        btn_browse2 = QPushButton("Browse")
        btn_browse2.clicked.connect(lambda: self.get_file(self.path2_edit))
        grid.addWidget(btn_browse2, 3, 2)

        grid.addWidget(QLabel("Seller Header Row Index"), 4, 0)
        self.int2 = QLineEdit()
        self.int2.setText("24")
        grid.addWidget(self.int2, 4, 1)

        grid.addWidget(QLabel("Seller Document Number Header"), 5, 0)
        self.str2 = QLineEdit()
        self.str2.setText("Document Number")
        grid.addWidget(self.str2, 5, 1)

        # -- Output Filepath --
        grid.addWidget(QLabel("Output Path:"), 6, 0)
        self.out_path = QLineEdit()
        self.out_path.setText("output/output1.xlsx")
        grid.addWidget(self.out_path, 6, 1)
        btn_out = QPushButton("Save As")
        btn_out.clicked.connect(self.save_file)
        grid.addWidget(btn_out, 6, 2)

        # 4. Action Button
        self.run_btn = QPushButton("Execute Logic")
        self.run_btn.setStyleSheet(
            "background-color: #2b5797; color: #282828; height: 40px;"
        )
        self.run_btn.clicked.connect(self.run_logic)
        main_layout.addWidget(self.run_btn)

    def get_file(self, target_line_edit):
        fname, _ = QFileDialog.getOpenFileName(self, "Select File")
        if fname:
            target_line_edit.setText(fname)

    def save_file(self):
        fname, _ = QFileDialog.getSaveFileName(self, "Save Output As")
        if fname:
            self.out_path.setText(fname)

    def run_logic(self):
        # 1. Gather data from your UI
        p1 = self.path1_edit.text()
        p2 = self.path2_edit.text()

        # 2. Pre-validation (Instruction check)
        if not p1 or not p2:
            # Simple alert if files are missing
            QMessageBox.warning(
                self, "Input Error", "Make sure both files are provided!"
            )
            return

        try:
            # 3. Call your backend logic (from logic.py)
            # We pass the data using the dictionary unpacking we discussed
            data = {
                "filepath1": p1,
                "filepath2": p2,
                "head_row1": int(self.int1.text()),
                "head_row2": int(self.int2.text()),
                "key_1": self.str1.text(),
                "key_2": self.str2.text(),
                "out_put": self.out_path.text(),
            }

            audit_excel(**data)

            # If successful:
            QMessageBox.information(self, "Success", "Processing complete!")

        except Exception as e:
            # --- OPTION 3: THE DETAILED LOG BOX ---
            # This triggers if the backend crashes or has a data error
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("Data Format Error")
            msg.setText("The backend could not process the files.")
            msg.setInformativeText(
                "Check if your headers have correct spaces and dots."
            )

            # This puts the actual Python error/console log into a hidden "Show Details" section
            msg.setDetailedText(f"Console Log / Traceback:\n{str(e)}")

            msg.exec()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

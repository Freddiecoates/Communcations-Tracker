import sys
import csv
from PyQt5.QtWidgets import (
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
    QMessageBox,
    QFileDialog,
    QCheckBox,
    QDialog,
)

from PyQt5.QtCore import QDate
from PyQt5.QtGui import QPalette, QColor
import sqlite3
import os
import sys
import shutil


def get_database_path():
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
        database_source = os.path.join(base_path, "comm_tracker.db")
        database_dest = os.path.join(os.path.dirname(sys.executable), "comm_tracker.db")

        if not os.path.exists(database_dest):
            shutil.copy(database_source, database_dest)
        return database_dest
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "comm_tracker.db")


DATABASE = get_database_path()


def create_database():
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS comms (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            comm_name TEXT NOT NULL,
            comm_type TEXT NOT NULL,
            template_id TEXT NOT NULL,
            created_by TEXT DEFAULT '',
            checked_by TEXT DEFAULT '',
            links_qr_correct TEXT DEFAULT 'No',
            format_correct TEXT DEFAULT 'No',
            cta_correct TEXT DEFAULT 'No',
            peer_reviewed TEXT DEFAULT 'No'
        )
        """
    )

    cursor.execute("PRAGMA table_info(comms)")
    existing_columns = [column[1] for column in cursor.fetchall()]
    new_columns = [
        ('links_qr_correct', 'TEXT DEFAULT "No"'),
        ('format_correct', 'TEXT DEFAULT "No"'),
        ('cta_correct', 'TEXT DEFAULT "No"'),
        ('peer_reviewed', 'TEXT DEFAULT "No"'),
    ]

    for col_name, col_def in new_columns:
        if col_name not in existing_columns:
            cursor.execute(f"ALTER TABLE comms ADD COLUMN {col_name} {col_def}")

    conn.commit()
    conn.close()


create_database()


class CommTrackerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Communication Tracker")
        self.setGeometry(100, 100, 800, 600)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2E3440;
            }
            QLabel {
                color:rgb(98, 100, 104);
                font-size: 14px;
            }
            QPushButton {
                background-color: #4C566A;
                color: #D8DEE9;
                border: 1px solid #5E81AC;
                padding: 8px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #5E81AC;
            }
            QTableWidget {
                background-color: #3B4252;
                color: #D8DEE9;
                gridline-color: #4C566A;
                font-size: 14px;
            }
            QHeaderView::section {
                background-color: #4C566A;
                color: #D8DEE9;
                padding: 6px;
                font-size: 14px;
            }
            QLineEdit {
                background-color: #4C566A;
                color: #D8DEE9;
                border: 1px solid #5E81AC;
                padding: 6px;
                font-size: 14px;
                border-radius: 4px;
            }
            QCheckBox {
                color: #D8DEE9;
                font-size: 14px;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
            }
        """)

        # main widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # search bar and button
        self.search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search by Comm Name or Template ID")
        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.search_records)
        self.search_layout.addWidget(self.search_input)
        self.search_layout.addWidget(self.search_button)
        self.layout.addLayout(self.search_layout)

        # table to display records
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(
            ["ID", "Date", "Comm Name", "Comm Type", "Template ID", "Created By", "Checked By"]
        )
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.layout.addWidget(self.table)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.doubleClicked.connect(self.show_details_dialog)

        # buttons for actions
        self.add_button = QPushButton("Add Record")
        self.add_button.clicked.connect(self.show_add_dialog)
        self.edit_button = QPushButton("Edit Record")
        self.edit_button.clicked.connect(self.show_edit_dialog)
        self.delete_button = QPushButton("Delete Record")
        self.delete_button.clicked.connect(self.delete_record)
        self.import_button = QPushButton("Import CSV")
        self.import_button.clicked.connect(self.import_csv)
        self.export_button = QPushButton("Export CSV")
        self.export_button.clicked.connect(self.export_csv)
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.delete_button)
        button_layout.addWidget(self.import_button)
        button_layout.addWidget(self.export_button)
        self.layout.addLayout(button_layout)

        # load data into the table
        self.load_data()

    def load_data(self):
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM comms")
        rows = cursor.fetchall()
        conn.close()
        self.table.setRowCount(len(rows))
        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(value) if value is not None else ""))

    def search_records(self):
        search_term = self.earch_input.text().strip()
        if not search_term:
            self.load_data()
            return
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT * FROM comms
            WHERE comm_name LIKE ? OR template_id LIKE ?
            """,
            (f"%{search_term}%", f"%{search_term}%"),
        )
        rows = cursor.fetchall()
        conn.close()
        self.table.setRowCount(len(rows))
        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem(str(value) if value is not None else ""))

    def show_add_dialog(self):
        self.add_dialog = QWidget()
        self.add_dialog.setWindowTitle("Add Record")
        self.add_dialog.setGeometry(200, 200, 400, 300)
        layout = QVBoxLayout(self.add_dialog)

        # input fields
        self.date_input = QLineEdit(QDate.currentDate().toString("yyyy-MM-dd"))
        self.date_input.setEnabled(False)
        self.comm_name_input = QLineEdit()
        self.comm_type_input = QLineEdit()
        self.template_id_input = QLineEdit()
        layout.addWidget(QLabel("Date:"))
        layout.addWidget(self.date_input)
        layout.addWidget(QLabel("Comm Name:"))
        layout.addWidget(self.comm_name_input)
        layout.addWidget(QLabel("Comm Type:"))
        layout.addWidget(self.comm_type_input)
        layout.addWidget(QLabel("Template ID:"))
        layout.addWidget(self.template_id_input)

        # save button
        save_button = QPushButton("Save")
        save_button.clicked.connect(self.save_record)
        layout.addWidget(save_button)
        self.add_dialog.setLayout(layout)
        self.add_dialog.show()

    def save_record(self):
        date = self.date_input.text()
        comm_name = self.comm_name_input.text()
        comm_type = self.comm_type_input.text()
        template_id = self.template_id_input.text()
        if not comm_name or not comm_type or not template_id:
            QMessageBox.warning(self, "Error", "All fields are required!")
            return

        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO comms (date, comm_name, comm_type, template_id)
            VALUES (?, ?, ?, ?)
            """,
            (date, comm_name, comm_type, template_id),
        )
        conn.commit()
        conn.close()
        self.load_data()
        self.add_dialog.close()

    def show_edit_dialog(self):
        selected_row = self.table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "Error", "Please select a record to edit!")
            return
        self.edit_dialog = QWidget()
        self.edit_dialog.setWindowTitle("Edit Record")
        self.edit_dialog.setGeometry(200, 200, 400, 300)
        layout = QVBoxLayout(self.edit_dialog)

        # get selected record data
        record_id = self.table.item(selected_row, 0).text()
        comm_name = self.table.item(selected_row, 2).text()
        created_by = self.table.item(selected_row, 5).text()
        checked_by = self.table.item(selected_row, 6).text()

        # input fields
        self.edit_comm_name_input = QLineEdit(comm_name)
        self.edit_created_by_input = QLineEdit(created_by if created_by else "")
        self.edit_checked_by_input = QLineEdit(checked_by if checked_by else "")

        # disable created_by and checked_by if already filled
        if created_by:
            self.edit_created_by_input.setEnabled(False)
        if checked_by:
            self.edit_checked_by_input.setEnabled(False)

        layout.addWidget(QLabel("Comm Name:"))
        layout.addWidget(self.edit_comm_name_input)
        layout.addWidget(QLabel("Created By:"))
        layout.addWidget(self.edit_created_by_input)
        layout.addWidget(QLabel("Checked By:"))
        layout.addWidget(self.edit_checked_by_input)

        # save button
        save_button = QPushButton("Save")
        save_button.clicked.connect(lambda: self.update_record(record_id))
        layout.addWidget(save_button)
        self.edit_dialog.setLayout(layout)
        self.edit_dialog.show()

    def update_record(self, record_id):
        comm_name = self.edit_comm_name_input.text()
        created_by = self.edit_created_by_input.text()
        checked_by = self.edit_checked_by_input.text()

        if not comm_name:
            QMessageBox.warning(self, "Error", "Comm Name is required!")
            return

        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        cursor.execute(
            """
            UPDATE comms
            SET comm_name = ?, created_by = ?, checked_by = ?
            WHERE id = ?
            """,
            (comm_name, created_by or '', checked_by or '', record_id),
        )

        conn.commit()
        conn.close()
        self.load_data()
        self.edit_dialog.close()

    def delete_record(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Error", "Please select records to delete!")
            return

        # get unique row indices
        selected_rows = set(item.row() for item in selected_items)
        ids = [self.table.item(row, 0).text() for row in selected_rows]
        confirm = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Delete {len(ids)} selected records?",
            QMessageBox.Yes | QMessageBox.No,
        )

        if confirm == QMessageBox.Yes:
            conn = sqlite3.connect(DATABASE)
            cursor = conn.cursor()
            placeholders = ",".join("?" for _ in ids)
            cursor.execute(f"DELETE FROM comms WHERE id IN ({placeholders})", ids)
            conn.commit()
            conn.close()
            self.load_data()

    def import_csv(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open CSV File", "", "CSV Files (*.csv)"
        )

        if not file_path:
            return
        try:
            conn = sqlite3.connect(DATABASE)
            cursor = conn.cursor()
            with open(file_path, "r") as file:
                reader = csv.reader(file)
                next(reader)  # skip header row
                for row in reader:
                    if len(row) < 4:
                        continue  # skip invalid rows

                    # extract base fields
                    date = row[0] if row[0] else QDate.currentDate().toString("yyyy-MM-dd")
                    comm_name = row[1]
                    comm_type = row[2]
                    template_id = row[3]
                    created_by = row[4] if len(row) > 4 else ""
                    checked_by = row[5] if len(row) > 5 else ""
                    links_ok = row[6] if len(row) > 6 else 'No'
                    format_ok = row[7] if len(row) > 7 else 'No'
                    cta_ok = row[8] if len(row) > 8 else 'No'
                    peer_ok = row[9] if len(row) > 9 else 'No'

                    cursor.execute(
                        """
                        INSERT INTO comms (
                            date, comm_name, comm_type, template_id,
                            created_by, checked_by,
                            links_qr_correct, format_correct, cta_correct, peer_reviewed
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,

                        (date, comm_name, comm_type, template_id,
                         created_by, checked_by,
                         links_ok, format_ok, cta_ok, peer_ok)
                    )

            conn.commit()
            conn.close()
            self.load_data()
            QMessageBox.information(self, "Success", "CSV imported successfully!")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"CSV Import Failed: {str(e)}")

    def export_csv(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save CSV File", "communications_export.csv", "CSV Files (*.csv)"
        )

        if not file_path:
            return
        try:
            conn = sqlite3.connect(DATABASE)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM comms")
            rows = cursor.fetchall()
            conn.close()
            with open(file_path, 'w', newline='') as file:
                writer = csv.writer(file)

                # write header
                writer.writerow([
                    'ID', 'Date', 'Comm Name', 'Comm Type', 'Template ID',
                    'Created By', 'Checked By', 'Links/QR Correct',
                    'Format Correct', 'CTA Correct', 'Peer Reviewed'
                ])

                # write data
                writer.writerows(rows)
            QMessageBox.information(self, "Success", "CSV exported successfully!")



        except Exception as e:
            QMessageBox.critical(self, "Error", f"CSV Export Failed: {str(e)}")

    def show_details_dialog(self):
        selected_row = self.table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "Error", "Please select a record to view!")
            return
        record_id = self.table.item(selected_row, 0).text()
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM comms WHERE id = ?", (record_id,))
        record = cursor.fetchone()
        conn.close()

        if not record:
            QMessageBox.warning(self, "Error", "Record not found!")
            return

        # create details
        dialog = QDialog(self)
        dialog.setWindowTitle("Record Details")
        dialog.setStyleSheet(self.styleSheet())
        layout = QVBoxLayout()

        # date
        layout.addWidget(QLabel(f"Date: {record[1]}"))

        # comm name
        self.details_comm_name = QLineEdit(record[2])
        layout.addWidget(QLabel("Comm Name:"))
        layout.addWidget(self.details_comm_name)

        # comm type and ID
        layout.addWidget(QLabel(f"Comm Type: {record[3]}"))
        layout.addWidget(QLabel(f"Template ID: {record[4]}"))

        # created and checked by
        self.details_created_by = QLineEdit(record[5] or "")
        self.details_checked_by = QLineEdit(record[6] or "")
        layout.addWidget(QLabel("Created By:"))
        layout.addWidget(self.details_created_by)
        layout.addWidget(QLabel("Checked By:"))
        layout.addWidget(self.details_checked_by)

        # checkbox's
        self.links_check = QCheckBox("Links/QR Correct?")
        self.links_check.setChecked(record[7] == 'Yes')
        layout.addWidget(self.links_check)
        self.format_check = QCheckBox("Format Correct?")
        self.format_check.setChecked(record[8] == 'Yes')
        layout.addWidget(self.format_check)
        self.cta_check = QCheckBox("CTAs Correct?")
        self.cta_check.setChecked(record[9] == 'Yes')
        layout.addWidget(self.cta_check)
        self.peer_review_check = QCheckBox("Peer Reviewed?")
        self.peer_review_check.setChecked(record[10] == 'Yes')
        layout.addWidget(self.peer_review_check)

        # save button
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(lambda: self.save_details(record_id, dialog))
        layout.addWidget(save_btn)
        dialog.setLayout(layout)
        dialog.exec_()

    def save_details(self, record_id, dialog):
        # get new values
        comm_name = self.details_comm_name.text()
        created_by = self.details_created_by.text()
        checked_by = self.details_checked_by.text()
        links_ok = 'Yes' if self.links_check.isChecked() else 'No'
        format_ok = 'Yes' if self.format_check.isChecked() else 'No'
        cta_ok = 'Yes' if self.cta_check.isChecked() else 'No'
        peer_ok = 'Yes' if self.peer_review_check.isChecked() else 'No'

        # update database
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        cursor.execute(
            """
            UPDATE comms SET
                comm_name = ?, created_by = ?, checked_by = ?,
                links_qr_correct = ?, format_correct = ?,
                cta_correct = ?, peer_reviewed = ?
            WHERE id = ?
            """,
            (comm_name, created_by, checked_by, links_ok, format_ok, cta_ok, peer_ok, record_id)
        )

        conn.commit()
        conn.close()
        self.load_data()  # refresh table
        dialog.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CommTrackerApp()
    window.show()
    sys.exit(app.exec_())


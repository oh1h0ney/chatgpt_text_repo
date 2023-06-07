import sys
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QTableWidgetItem, QVBoxLayout,\
    QDialog, QLabel, QComboBox, QPushButton, QTableWidget, QHeaderView, QWidget, QLineEdit
import pandas as pd
import json


class FieldMappingDialog(QDialog):
    def __init__(self, fields, files):
        super().__init__()

        self.setWindowTitle("Field Mapping")
        self.resize(400, 200)

        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Field", "Source File", "Mapped Field", "Mapped Field Edit"])

        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.buttons_layout = QVBoxLayout()
        self.add_mapping_button = QPushButton("Add Mapping")
        self.reset_button = QPushButton("Reset")
        self.apply_save_button = QPushButton("Apply and Save")
        self.load_button = QPushButton("Load Mapping")

        self.add_mapping_button.clicked.connect(self.add_mapping)
        self.reset_button.clicked.connect(self.reset_mappings)
        self.apply_save_button.clicked.connect(self.apply_and_save_mappings)
        self.load_button.clicked.connect(self.load_mappings)

        self.buttons_layout.addWidget(self.add_mapping_button)
        self.buttons_layout.addWidget(self.reset_button)
        self.buttons_layout.addWidget(self.apply_save_button)
        self.buttons_layout.addWidget(self.load_button)

        self.layout.addWidget(self.table)
        self.layout.addLayout(self.buttons_layout)

        self.field_file_map = {}
        self.original_mappings = {}

        for field in fields:
            for file in files:
                try:
                    data = pd.read_excel(file) if file.endswith('.xlsx') else pd.read_csv(file)
                    if field in data.columns:
                        self.field_file_map[field] = file
                        self.add_row(field, file)
                        break
                except Exception as e:
                    QMessageBox.warning(self, 'Error', f'Failed to read file: {file}\n{str(e)}')

        self.populate_mapped_field_combos()

    def add_row(self, field, file):
        row = self.table.rowCount()
        self.table.insertRow(row)

        field_item = QTableWidgetItem(field)
        field_item.setFlags(field_item.flags() & ~Qt.ItemIsEditable)

        file_item = QTableWidgetItem(file)
        file_item.setFlags(file_item.flags() & ~Qt.ItemIsEditable)

        self.table.setItem(row, 0, field_item)
        self.table.setItem(row, 1, file_item)

        mapped_field_combo = QComboBox()
        self.table.setCellWidget(row, 2, mapped_field_combo)

        mapped_field_edit = QLineEdit()
        self.table.setCellWidget(row, 3, mapped_field_edit)

    def populate_mapped_field_combos(self):
        for row in range(self.table.rowCount()):
            mapped_field_combo = self.table.cellWidget(row, 2)
            if mapped_field_combo is not None:
                combo_field = self.table.item(row, 0).text()
                combo_file = self.table.item(row, 1).text()
                mapped_field_combo.addItem(combo_field)  # Default option (use field as mapping)
                for r in range(self.table.rowCount()):
                    # if self.table.item(r, 1).text() == combo_file and r != row:
                    mapped_field_combo.addItem(self.table.item(r, 0).text())

    def add_mapping(self):
        row = self.table.rowCount()
        if row > 0:
            field_item = self.table.item(row - 1, 0)
            if field_item is not None:
                field = field_item.text()
                self.add_row(field, "")

    def reset_mappings(self):
        self.table.clearContents()
        self.table.setRowCount(0)
        self.populate_mapped_field_combos()

    def apply_and_save_mappings(self):
        mappings = self.get_mappings()
        if mappings:
            save_dialog = QFileDialog()
            save_dialog.setAcceptMode(QFileDialog.AcceptSave)
            save_dialog.setDefaultSuffix('.json')
            save_dialog.setWindowTitle("Save Mappings")
            if save_dialog.exec() == QDialog.Accepted:
                file_path = save_dialog.selectedFiles()[0]
                with open(file_path, 'w') as file:
                    json.dump(mappings, file, indent=4)
                QMessageBox.information(self, 'Success', 'Mappings saved successfully.')

    def load_mappings(self):
        load_dialog = QFileDialog()
        load_dialog.setFileMode(QFileDialog.ExistingFile)
        load_dialog.setWindowTitle("Load Mappings")
        if load_dialog.exec() == QDialog.Accepted:
            file_path = load_dialog.selectedFiles()[0]
            with open(file_path, 'r') as file:
                try:
                    mappings = json.load(file)
                    self.original_mappings = mappings
                    self.populate_table()
                    QMessageBox.information(self, 'Success', 'Mappings loaded successfully.')
                except Exception as e:
                    QMessageBox.warning(self, 'Error', f'Failed to load mappings.\n{str(e)}')

    def get_mappings(self):
        mappings = {}
        for row in range(self.table.rowCount()):
            field_item = self.table.item(row, 0)
            mapped_field_combo = self.table.cellWidget(row, 2)
            if field_item is not None and mapped_field_combo is not None:
                field = field_item.text()
                mapped_field = mapped_field_combo.currentText()
                if mapped_field != "" and mapped_field != field:
                    mappings[field] = mapped_field
        return mappings

    def populate_table(self):
        self.table.clearContents()
        self.table.setRowCount(0)
        for field, mapped_field in self.original_mappings.items():
            file = self.field_file_map.get(field, "")
            self.add_row(field, file)
            row = self.table.rowCount() - 1
            mapped_field_combo = self.table.cellWidget(row, 2)
            if mapped_field_combo is not None:
                index = mapped_field_combo.findText(mapped_field)
                if index >= 0:
                    mapped_field_combo.setCurrentIndex(index)


class MergeSplitTool(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel CSV Tool")
        self.resize(400, 200)

        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        self.merge_button = QPushButton("Merge Files")
        self.split_button = QPushButton("Split File")
        self.mapping_button = QPushButton("Field Mapping")

        self.layout.addWidget(self.merge_button)
        self.layout.addWidget(self.split_button)
        self.layout.addWidget(self.mapping_button)

        self.merge_button.clicked.connect(self.merge_files)
        self.split_button.clicked.connect(self.split_file)
        self.mapping_button.clicked.connect(self.field_mapping)

        self.mappings = {}

    def merge_files(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        files = file_dialog.getOpenFileNames(self, 'Select Files to Merge', filter='Excel Files (*.xlsx);;CSV Files (*.csv)')
        if files:
            save_dialog = QFileDialog()
            save_dialog.setDefaultSuffix('.xlsx')
            save_file, _ = save_dialog.getSaveFileName(self, 'Save Merged File', 'output', filter='Excel Files (*.xlsx)')
            if save_file:
                field_mappings = self.get_field_mappings(files[0])
                if not field_mappings:
                    return

                merged_data = pd.DataFrame()
                for file in files[0]:
                    try:
                        data = pd.read_excel(file) if file.endswith('.xlsx') else pd.read_csv(file)
                        data = self.apply_field_mappings(data, field_mappings)
                        merged_data = pd.concat([merged_data, data], ignore_index=True)
                    except Exception as e:
                        QMessageBox.warning(self, 'Error', f'Failed to read file: {file}\n{str(e)}')
                        return

                merged_data.to_excel(save_file, index=False)
                QMessageBox.information(self, 'Success', 'Files merged successfully.')

    def split_file(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file, _ = file_dialog.getOpenFileName(self, 'Select File to Split', filter='Excel Files (*.xlsx);;CSV Files (*.csv)')
        if file:
            try:
                data = pd.read_excel(file) if file.endswith('.xlsx') else pd.read_csv(file)
            except Exception as e:
                QMessageBox.warning(self, 'Error', f'Failed to read file: {file}\n{str(e)}')
                return

            field_names = data.columns.tolist()
            selected_fields, ok = self.select_fields_dialog('Select Fields to Split', 'Select fields:', field_names)
            if ok:
                split_dialog = QFileDialog()
                split_dialog.setFileMode(QFileDialog.Directory)
                save_dir = split_dialog.getExistingDirectory(self, 'Select Save Directory')
                if save_dir:
                    try:
                        for field in selected_fields:
                            field_data = data.groupby(field)
                            for group_name, group_data in field_data:
                                file_name = f'{field}_{group_name}'
                                save_path = f'{save_dir}/{file_name}.xlsx'
                                group_data.to_excel(save_path, index=False)

                        QMessageBox.information(self, 'Success', 'File split successfully.')
                    except Exception as e:
                        QMessageBox.warning(self, 'Error', f'Failed to split file.\n{str(e)}')

    def get_field_mappings(self, files):
        field_mappings = {}
        field_names = set()
        file_fields_map = {}
        for file in files:
            try:
                data = pd.read_excel(file) if file.endswith('.xlsx') else pd.read_csv(file)
                file_fields_map[file] = data.columns.tolist()
                field_names.update(data.columns.tolist())
            except Exception as e:
                QMessageBox.warning(self, 'Error', f'Failed to read file: {file}\n{str(e)}')
                return None

        dialog = FieldMappingDialog(field_names, files)
        if dialog.exec() == QDialog.Accepted:
            mappings = dialog.get_mappings()
            for field, mapped_field in mappings.items():
                for file, fields in file_fields_map.items():
                    if field in fields:
                        field_mappings[(file, field)] = (file, mapped_field)

        return field_mappings

    def select_fields_dialog(self, title, label, field_names):
        dialog = QDialog(self)
        dialog.setWindowTitle(title)

        layout = QVBoxLayout()
        field_label = QLabel(label)
        layout.addWidget(field_label)

        field_combo = QComboBox()
        field_combo.addItems(field_names)
        layout.addWidget(field_combo)

        ok_button = QPushButton('OK')
        ok_button.clicked.connect(dialog.accept)
        layout.addWidget(ok_button)

        dialog.setLayout(layout)

        if dialog.exec() == QDialog.Accepted:
            selected_fields = [field_combo.itemText(i) for i in range(field_combo.count())]
            return selected_fields, True

        return None, False

    def apply_field_mappings(self, data, field_mappings):
        for (file, field), (mapped_file, mapped_field) in field_mappings.items():
            if file == mapped_file:
                data.rename(columns={field: mapped_field}, inplace=True)
        return data

    def field_mapping(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        files = file_dialog.getOpenFileNames(self, 'Select Files for Field Mapping', filter='Excel Files (*.xlsx);;CSV Files (*.csv)')
        if files:
            self.mappings = self.get_field_mappings(files[0])
            if self.mappings:
                QMessageBox.information(self, 'Success', 'Field mapping completed.')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    tool = MergeSplitTool()
    tool.show()
    sys.exit(app.exec())


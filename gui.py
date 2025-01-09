import os
import sys
from functools import partial

from PyQt6.QtWidgets import QApplication, QWidget, QGridLayout, QLineEdit, QComboBox, QPushButton, QFileDialog
from PyQt6.QtCore import QStandardPaths

from dataframing import csv_load, filter_csv
from error_msg import error_msg
from pdf_creation import pdf


class FileExplorer:
    def __init__(self, csv_file_browser, save_file_browser, campaign_dropdown):
        self.csv_file_browser = csv_file_browser
        self.save_file_browser = save_file_browser
        self.campaign_dropdown = campaign_dropdown

    def open_file_dialog(self, button_id):
        dialog = QFileDialog()
        downloads_folder = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DownloadLocation)

        if button_id == 'csv_browse':
            # Set the directory for Windows (Downloads folder)
            dialog.setDirectory(downloads_folder)

            # Set the directory for Linux
            #dialog.setDirectory("/home/ben/Downloads/")

            dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
            dialog.setNameFilter("Text (*.csv)")
            dialog.setViewMode(QFileDialog.ViewMode.List)

            if dialog.exec():
                filename = dialog.selectedFiles()
                if filename:
                    # Load CSV and update the UI
                    self.csv_file_browser.setText(filename[0])

        elif button_id == 'save_browse':
            dialog.setDirectory(downloads_folder)
            dialog.setFileMode(QFileDialog.FileMode.Directory)
            dialog.setViewMode(QFileDialog.ViewMode.List)

            if dialog.exec():
                filename = dialog.selectedFiles()
                if filename:
                    self.save_file_browser.setText(filename[0])


class CsvToPdfApp(QWidget):
    def __init__(self):
        super().__init__()  # super init is used to pull from parent class

        self.setWindowTitle('CSV to PDF v0.1')
        self.setGeometry(100, 100, 280, 80)

        # Create layout
        self.layout = QGridLayout()
        self.setLayout(self.layout)

        # Create widgets
        self.create_widgets()

        # File explorer integration
        self.file_explorer = FileExplorer(self.csv_file_browser, self.save_file_browser, self.campaign_dropdown)

        # Connect buttons to the file explorer methods
        self.csv_browse_button.clicked.connect(partial(self.file_explorer.open_file_dialog, 'csv_browse'))
        self.save_browse_button.clicked.connect(partial(self.file_explorer.open_file_dialog, 'save_browse'))

    def create_widgets(self):
        # Create labels and input fields
        self.csv_file_browser = QLineEdit()
        self.csv_file_browser.setPlaceholderText('Select CSV')
        self.layout.addWidget(self.csv_file_browser, 0, 0, 1, 2)

        self.save_file_browser = QLineEdit()
        self.save_file_browser.setPlaceholderText('Select Save Location')
        self.layout.addWidget(self.save_file_browser, 1, 0, 1, 2)

        self.file_name = QLineEdit()
        self.file_name.setPlaceholderText('Save As (report name)')
        self.layout.addWidget(self.file_name, 4, 0, 1, 3)

        # Add dropdown menus
        # todo re-add the report type functionality as i am going to remove it for this release
        self.dropdown = QComboBox()
        self.dropdown.setPlaceholderText('Select Report Type')
        #self.dropdown.addItems(['Summary Stats', 'Complete Stats'])
        self.dropdown.addItems(['Complete Stats'])
        self.layout.addWidget(self.dropdown, 2, 0, 1, 2)

        self.campaign_dropdown = QComboBox()
        self.campaign_dropdown.setPlaceholderText('Select Campaign')
        self.layout.addWidget(self.campaign_dropdown, 3, 0, 1, 3)
        self.campaign_dropdown.setEnabled(False)

        # dropdowns to allow user to choose the step and version, but ive set them to do it automatically currently
        #self.step_dropdown = QComboBox()
        #self.step_dropdown.setPlaceholderText('Select Step')
        #self.step_dropdown.addItems(['1', '2', '3', '4'])
        #self.layout.addWidget(self.step_dropdown, 3, 1, 1, 1)

        #self.version_dropdown = QComboBox()
        #self.version_dropdown.setPlaceholderText('Select Version')
        #self.version_dropdown.addItems(['A', 'B', 'C'])
        #self.layout.addWidget(self.version_dropdown, 3, 2, 1, 1)

        # Create buttons
        self.generate_button = QPushButton('Generate')
        self.generate_button.clicked.connect(partial(self.button_click, 'generate'))
        self.layout.addWidget(self.generate_button, 5, 0, 1, 3)

        self.csv_browse_button = QPushButton('Browse')
        self.csv_browse_button.clicked.connect(partial(self.button_click, 'csv_browse'))
        self.layout.addWidget(self.csv_browse_button, 0, 2)

        self.save_browse_button = QPushButton('Browse')
        self.save_browse_button.clicked.connect(partial(self.button_click, 'save_browse'))
        self.layout.addWidget(self.save_browse_button, 1, 2)

        self.apply_button = QPushButton('Apply')
        self.apply_button.clicked.connect(partial(self.button_click, 'apply'))
        self.layout.addWidget(self.apply_button, 2, 2)

    def report_type_check(self):

        if self.dropdown.currentText() == 'Complete Stats':
            correct_input = False
            self.campaign_dropdown.setEnabled(True)
            button_options = csv_load(self.csv_file_browser.text())
            self.campaign_dropdown.addItems(
                button_options['campaign'].drop_duplicates().tolist()
            )

        elif self.dropdown.currentText() == 'Summary Stats':
            correct_input = True
        else:
            error_msg('Something has gone horribly wrong')

        return correct_input

    def button_click(self, action):
        if action == 'generate':
            # makes sure all inputs have been filled before generating save location or pdf
            if (self.csv_file_browser.text() != '' and self.save_file_browser.text() != ''
                    and self.file_name.text() != '' and self.dropdown.currentIndex() != -1):
                save_location = os.path.join(self.save_file_browser.text(), self.file_name.text())
                filtered_csv = filter_csv(csv_load(self.csv_file_browser.text()), self.campaign_dropdown.currentText())

                pdf(filtered_csv, save_location)
            else:
                error_msg('Missing Inputs')

        elif action == 'apply':
            if self.dropdown.currentIndex() != -1 and self.csv_file_browser.text() != '':
                if self.dropdown.currentText() == 'Summary Stats':
                    report_type = ('Report type set to ' + self.dropdown.currentText())
                    error_msg(report_type)
                else:
                    self.campaign_dropdown.clear()  # Clear existing items
                    correct_input = False
                    while not correct_input:
                        try:
                            correct_input = self.report_type_check()
                            report_type = ('Report type set to ' + self.dropdown.currentText())
                            error_msg(report_type)
                            break
                        except KeyError as e:
                            error_msg('No campaigns found in CSV, ensure this file is not a summary statistics')



            else:
                error_msg('Select CSV and Report Type (if applicable) before applying')


if __name__ == '__main__':
    app = QApplication([])
    app.setStyle('Fusion')
    window = CsvToPdfApp()
    window.show()
    sys.exit(app.exec())

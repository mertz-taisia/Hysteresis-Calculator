import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QSpinBox, QFrame
from PyQt5.QtGui import QPainter, QColor, QFont
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QLabel, QStackedWidget
import sys
import analysis
import os
import openpyxl
from PyQt5.QtWidgets import QFileDialog
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QPalette
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QSpacerItem, QSizePolicy



class CustomWidget(QWidget):
    def __init__(self, color, parent=None):
        super().__init__(parent)
        self.color = color
        self.ramp_count = 1

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), self.color)

class MainPage(QMainWindow):
    def __init__(self, stacked_widget):
        super().__init__()
        self.stacked_widget = stacked_widget
        self.ramp_count = 1

        self.setWindowTitle("Ramp Analyzer")
        self.setGeometry(100, 100, 675, 370)
        self.setWindowFlags(Qt.MSWindowsFixedSizeDialogHint)

        self.setAutoFillBackground(True)
        p = self.palette()
        p.setColor(self.backgroundRole(), QColor(217, 217, 217))
        self.setPalette(p)

        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Top Row (Custom RGB: 100, 149, 237)
        top_row = CustomWidget(QColor(0, 128, 199))
        top_row.setFixedHeight(self.height() * 27 // 100)
        layout.addWidget(top_row)

        # Middle Row (Custom RGB: 255, 0, 0)
        middle_row = CustomWidget(QColor(255, 255, 255))
        middle_row.setFixedHeight(self.height()  * 55 // 100)

        middle_row_layout = QVBoxLayout(middle_row)
        middle_row_layout.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        data_container = QWidget()
        data_input_layout = QHBoxLayout(data_container)

        label1 = QLabel("Data(.xlsx)")
        label1.setAlignment(Qt.AlignLeft)
        label1.setStyleSheet("font-size: 18px; color: #6E6E6E;")
        font = QFont()
        font.setPointSize(20)
        label1.setFont(font)
        data_input_layout.addWidget(label1)

        file_input_layout = QHBoxLayout()

        file_frame = QFrame() 
        file_frame.setFrameShape(QFrame.StyledPanel) 
        file_frame.setContentsMargins(0, 0, 0, 0) 
        file_frame.setStyleSheet("background: #FFFFFF; border: 1px solid #989898;")

        file_input_layout.addWidget(file_frame)

        frame_layout = QHBoxLayout(file_frame)
        
        self.file_path_input = QLineEdit()
        self.file_path_input.setReadOnly(True)
        self.file_path_input.setStyleSheet("border: transparent")
        frame_layout.addWidget(self.file_path_input)

        select_file_button = QPushButton("Browse")

        select_file_button.setStyleSheet("""
            QPushButton {
                background: #0080C7; 
                border: 1px solid #989898; 
                font-size: 14px; 
                padding-left: 15px; 
                padding-right: 15px; 
                padding-top: 3px; 
                padding-bottom: 3px; 
                color: #EDF3F3
            }

            QPushButton:hover {
                background-color: #EDF3F3;
                border: 1px solid #E5E8E9;
                color: #464646;
            }

            QPushButton:pressed {
                background-color: #E0E8E8;
                border: 1px solid #B9B9B9;
            }
        """)
        select_file_button.clicked.connect(self.browse_file)
        frame_layout.addWidget(select_file_button)


        data_input_layout.addSpacing(100)
        data_input_layout.addLayout(file_input_layout)

        middle_row_layout.addWidget(data_container)

        middle_row_layout.addSpacing(5)

        ramps_container = QWidget()
        ramps_input_layout = QHBoxLayout(ramps_container)

        label2 = QLabel("Number of Ramps")
        label2.setAlignment(Qt.AlignLeft)
        label2.setStyleSheet("font-size: 18px; color: #6E6E6E;")
        label2.setFont(font)
        ramps_input_layout.addWidget(label2)
        
        ramps_input_layout.addSpacing(43)

        ramp_frame = QFrame()
        ramp_frame.setFrameShape(QFrame.Box)

        ramp_num_layout = QHBoxLayout(ramp_frame)

        subtract = QPushButton("-")
        subtract.setFixedSize(25, 25) 
        subtract.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                color: #6B6B6B;
                background-color: transparent;
                border: 1px solid #E5E8E9;
                padding: 5px;
            }

            QPushButton:hover {
                background-color: #EDF3F3;
                border: 1px solid #E5E8E9;
                color: #464646;
            }

            QPushButton:pressed {
                background-color: #E0E8E8;
                border: 1px solid #B9B9B9;
            }
        """)

        add = QPushButton("+")
        add.setFixedSize(25, 25) 
        add.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                color: #6B6B6B;
                background-color: transparent;
                border: 1px solid #E5E8E9;
                padding: 5px;
            }

            QPushButton:hover {
                background-color: #EDF3F3;
                border: 1px solid #E5E8E9;
                color: #464646;
            }

            QPushButton:pressed {
                background-color: #E0E8E8;
                border: 1px solid #B9B9B9;
            }
        """)
        subtract.clicked.connect(self.decrement_value)  
        add.clicked.connect(self.increment_value)

        self.count_label = QLabel('1 Ramp(s)')
        self.count_label.setStyleSheet("border: transparent;")


        ramp_num_layout.addWidget(self.count_label)
        ramp_num_layout.addStretch(1)  
        ramp_num_layout.addWidget(subtract)
        ramp_num_layout.addWidget(add)
        
        container = QWidget()
        container.setLayout(ramp_num_layout)
        container.setStyleSheet("background: transparent; border: 1px solid #989898;")
 
        main_layout = QVBoxLayout()
        main_layout.addWidget(container)

        ramps_input_layout.addWidget(container)
        
        middle_row_layout.addWidget(ramps_container)

        layout.addWidget(middle_row)

        # Bottom Row (Custom RGB: 237,243,243)
        bottom_row = CustomWidget(QColor(237,243,243))
        bottom_row.setFixedHeight(self.height() * 23 // 100)

        button_layout = QHBoxLayout()
        button_layout.setAlignment(Qt.AlignRight)

        self.button1 = QPushButton("Run Analysis")
        self.button1.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                color: #6B6B6B;
                background-color: #FFFFFF;
                border: 2px solid #DEE2E3;
                padding: 5px;
                
            }

            QPushButton:hover {
                background-color: #BDC7C7;
                border: 1px solid #ADB4B4;
                color: #464646;
            }

            QPushButton:pressed {
                background: #0080C7; 
                border: 1px solid #989898; 
                color: #EDF3F3
            }
        """)
        
        self.button1.setFixedSize(140, 35) 
        
        button2 = QPushButton("Cancel")
        button2.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                color: #6B6B6B;
                background-color: #EDF3F3;
                border: 2px solid #DEE2E3;
                padding: 5px;
            }

            QPushButton:hover {
                background-color: #BDC7C7;
                border: 1px solid #ADB4B4;
                color: #464646;
            }

            QPushButton:pressed {
                background-color: #D0D0D0;
                border: 1px solid #B9B9B9;
            }
        """)
                
        button2.setFixedSize(140, 35) 
        button2.clicked.connect(QApplication.instance().quit)
        
        self.button1.clicked.connect(self.goToPage2)

        button_layout.addWidget(button2)
        button_layout.addWidget(self.button1)

        button_layout.addSpacing(8)

        bottom_row.setLayout(button_layout)

        layout.addWidget(bottom_row)

        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

        
    def browse_file(self):
        default_dir = os.path.expanduser("~")  
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_dialog.setDirectory(default_dir)
        file_dialog.setNameFilter("XLSX files (*.xlsx)")

        if file_dialog.exec_():
            selected_files = file_dialog.selectedFiles()
            if len(selected_files) > 0:
                file_path = selected_files[0]
                self.file_path_input.setText(file_path)

    def increment_value(self):
        self.ramp_count += 1
        self.update_count_label()

    def decrement_value(self):
        if self.ramp_count != 1:
            self.ramp_count -= 1
        self.update_count_label()

    def update_count_label(self):
        self.count_label.setText('{} Ramp(s)'.format(self.ramp_count))
    
    def validate_data_format(self, file_path, ramp_count):
        try:
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active
            row_count = sheet.max_row

            if row_count % ramp_count != 0:
                return False

            return True

        except Exception as e:
            return False
    

    def goToPage2(self):
        self.button1.setText("Running...")
        QtWidgets.QApplication.processEvents()

        file_path = self.file_path_input.text()
        if not file_path:
            error_dialog = QtWidgets.QMessageBox()
            error_dialog.setIcon(QtWidgets.QMessageBox.Warning)
            error_dialog.setWindowTitle("Error")
            error_dialog.setText("Please add a file.")
            error_dialog.exec_()
            self.button1.setText("Run Analysis")
        elif not self.validate_data_format(file_path, self.ramp_count):
            error_dialog = QtWidgets.QMessageBox()
            error_dialog.setIcon(QtWidgets.QMessageBox.Warning)
            error_dialog.setWindowTitle("Error")
            error_dialog.setText("Invalid file format.")
            error_dialog.exec_()
            self.button1.setText("Run Analysis")
        else:
            self.stacked_widget.setCurrentIndex(1)
            self.stacked_widget.widget(1).update_values(self.ramp_count, file_path)
            self.button1.setText("Run Analysis")


class Page2(QWidget):
    def __init__(self):
        super().__init__()
        self.ramp_count = None
        self.file_path_input = None
        self.meanResult = None
        self.combinedResult = None

        self.setGeometry(100, 100, 675, 370)
        self.setWindowFlags(Qt.MSWindowsFixedSizeDialogHint)  

        self.setAutoFillBackground(True)
        p = self.palette()
        p.setColor(self.backgroundRole(), QColor(217, 217, 217))
        self.setPalette(p)

        layout = QVBoxLayout()        
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Top Row (Custom RGB: 100, 149, 237)
        top_row = CustomWidget(QColor(0, 128, 199))
        top_row.setFixedHeight(self.height() * 20 // 100)
        layout.addWidget(top_row)

        # Middle Row (Custom RGB: 255, 0, 0)
        middle_row = CustomWidget(QColor(255, 255, 255))
        middle_row.setFixedHeight(self.height() * 60 // 100)
        middle_row_layout = QVBoxLayout(middle_row) 
        layout.addWidget(middle_row)
        middle_row_layout.setSpacing(5)

        spacer1 = QSpacerItem(5, 5, QSizePolicy.Minimum, QSizePolicy.Expanding)
        middle_row_layout.addItem(spacer1)

        button_layout = QHBoxLayout()

        button_layout.addSpacing(10)

        self.saveSingleRampButton = QPushButton("Save Single Ramp Data")
        self.saveSingleRampButton.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                color: #6B6B6B;
                background-color: #FFFFFF;
                border: 2px solid #DEE2E3;
                padding: 5px;
                
            }

            QPushButton:hover {
                background-color: #BDC7C7;
                border: 1px solid #ADB4B4;
                color: #464646;
            }

            QPushButton:pressed {
                background: #0080C7; 
                border: 1px solid #989898; 
                color: #EDF3F3
            }
        """)

        self.saveSingleRampButton.clicked.connect(self.save_single_file)
        button_layout.addWidget(self.saveSingleRampButton)

        button_layout.addSpacing(10)

        self.saveOverallRampButton = QPushButton("Save Overall Ramp Data")
        self.saveOverallRampButton.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                color: #6B6B6B;
                background-color: #FFFFFF;
                border: 2px solid #DEE2E3;
                padding: 5px;
                
            }

            QPushButton:hover {
                background-color: #BDC7C7;
                border: 1px solid #ADB4B4;
                color: #464646;
            }

            QPushButton:pressed {
                background: #0080C7; 
                border: 1px solid #989898; 
                color: #EDF3F3
            }
        """)
        self.saveOverallRampButton.clicked.connect(self.save_overall_file)
        button_layout.addWidget(self.saveOverallRampButton)

        button_layout.addSpacing(10)

        middle_row_layout.addLayout(button_layout)
        
        spacer2 = QSpacerItem(5, 5, QSizePolicy.Minimum, QSizePolicy.Expanding)
        middle_row_layout.addItem(spacer2)

        # Bottom Row (Custom RGB: 237,243,243)
        bottom_row = CustomWidget(QColor(237,243,243))
        bottom_row.setFixedHeight(self.height() * 20 // 100)

        button_layout = QHBoxLayout()
        button_layout.setAlignment(Qt.AlignRight)

        self.button1 = QPushButton("Return to Main Page")
        self.button1.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                color: #6B6B6B;
                background-color: #FFFFFF;
                border: 2px solid #DEE2E3;
                padding: 5px;
                
            }

            QPushButton:hover {
                background-color: #BDC7C7;
                border: 1px solid #ADB4B4;
                color: #464646;
            }

            QPushButton:pressed {
                background: #0080C7; 
                border: 1px solid #989898; 
                color: #EDF3F3
            }
        """)
        
        self.button1.setFixedSize(140, 35) 
        
        button2 = QPushButton("Cancel")
        button2.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                color: #6B6B6B;
                background-color: #EDF3F3;
                border: 2px solid #DEE2E3;
                padding: 5px;
            }

            QPushButton:hover {
                background-color: #BDC7C7;
                border: 1px solid #ADB4B4;
                color: #464646;
            }

            QPushButton:pressed {
                background-color: #D0D0D0;
                border: 1px solid #B9B9B9;
            }
        """)
                
        button2.setFixedSize(140, 35) 
        button2.clicked.connect(QApplication.instance().quit)
        self.button1.clicked.connect(self.goToMainPage)

        button_layout.addWidget(button2)
        button_layout.addWidget(self.button1)

        button_layout.addSpacing(8)

        bottom_row.setLayout(button_layout)
        layout.addWidget(bottom_row)

        self.setLayout(layout)
    
    def update_values(self, ramp_count, file_path_input):
        self.ramp_count = ramp_count
        self.file_path_input = file_path_input
        analysis.dataAnalysis(self.ramp_count, self.file_path_input)
        self.meanResult = analysis.getMean()
        self.combinedResult = analysis.getCombinedData()

    
    def goToMainPage(self):
        self.saveSingleRampButton.setText("Save Single Ramp Data")
        self.saveOverallRampButton.setText("Save Overall Ramp Data")
        self.parent().setCurrentIndex(0)
        self.saveSingleRampButton.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                color: #6B6B6B;
                background-color: #FFFFFF;
                border: 2px solid #DEE2E3;
                padding: 5px;
                
            }

            QPushButton:hover {
                background-color: #BDC7C7;
                border: 1px solid #ADB4B4;
                color: #464646;
            }

            QPushButton:pressed {
                background: #0080C7; 
                border: 1px solid #989898; 
                color: #EDF3F3
            }
        """)
        self.saveOverallRampButton.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                color: #6B6B6B;
                background-color: #FFFFFF;
                border: 2px solid #DEE2E3;
                padding: 5px;
                
            }

            QPushButton:hover {
                background-color: #BDC7C7;
                border: 1px solid #ADB4B4;
                color: #464646;
            }

            QPushButton:pressed {
                background: #0080C7; 
                border: 1px solid #989898; 
                color: #EDF3F3
            }
        """)

            
    def save_single_file(self):
        self.saveSingleRampButton.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                padding: 5px;
                background-color: #BDC7C7;
                border: 1px solid #ADB4B4;
                color: #464646;
            }
        """)
        self.saveSingleRampButton.setText("Saving...")
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        default_file_path = self.file_path_input  
        default_dir, file_name = os.path.split(default_file_path)
        default_file_name = "results " + file_name  
        file_filter = "Excel Files (*.xlsx)"
        save_file_name, _ = QFileDialog.getSaveFileName(self, "Save File", os.path.join(default_dir, default_file_name), file_filter, options=options)

        if save_file_name:  
            try:
                self.meanResult.to_excel(save_file_name + ".xlsx", index=False)
            except Exception as e:
                print(f"An error occurred while saving the file: {e}")
        self.saveSingleRampButton.setText("Saved")
        self.saveSingleRampButton.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                padding: 5px;
                background: #0080C7; 
                border: 1px solid #989898; 
                color: #EDF3F3
            }
        """)
    
    def save_overall_file(self):
        self.saveOverallRampButton.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                padding: 5px;
                background-color: #BDC7C7;
                border: 1px solid #ADB4B4;
                color: #464646;
            }
        """)
        self.saveOverallRampButton.setText("Saving...")
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        default_file_path = self.file_path_input 
        default_dir, file_name = os.path.split(default_file_path)
        default_file_name = "combined " + file_name  
        file_filter = "Excel Files (*.xlsx)"
        save_file_name, _ = QFileDialog.getSaveFileName(self, "Save File", os.path.join(default_dir, default_file_name), file_filter, options=options)

        if save_file_name:  
            try:
                self.combinedResult.to_excel(save_file_name + ".xlsx", index=False)
            except Exception as e:
                print(f"An error occurred while saving the file: {e}")
        self.saveOverallRampButton.setText("Saved")
        self.saveOverallRampButton.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                padding: 5px;
                background: #0080C7; 
                border: 1px solid #989898; 
                color: #EDF3F3
            }
        """)


if __name__ == '__main__':
    app = QApplication(sys.argv)

    main_window = QMainWindow()
    stacked_widget = QStackedWidget()

    main_page = MainPage(stacked_widget)
    page2 = Page2()

    stacked_widget.addWidget(main_page)
    stacked_widget.addWidget(page2)

    stacked_widget.setFixedSize(675, 370)

    main_window.setCentralWidget(stacked_widget)
    main_window.show()

    sys.exit(app.exec_())

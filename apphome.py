from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

# Application Home Page/Frame Widget
class AppHome(QWidget):
    """
    Frame/Page Widget "Application Home/Menu":

    This Frame Module is a simple "Menu/Home page" for our Application,
    here we have a button for every page.
    """

    def __init__(self, parent, controller):
        """
        Frame/Page Widget "Application Home/Menu":

        Class AppHome(parent, controller):
        parent: This widget parent, usually the container widget.
        controller: QT Main window, to control changing pages.

        Initialization of the Application Home Class page.
        """
        super().__init__(parent=parent)
        # first create our widgets
        self.CreateWidgets(controller)
        # then set our grid layout and config
        self.GridConfigs()

    # Creation of widgets on screen
    def CreateWidgets(self, controller):
        """
        Function CreateWidgets(controller)
        controller: Our main window / controller widget.

        Used to create new widgets (labels, buttons, etc.),
        controller is used in buttons to show another page/frame widget.
        """

        # Label Title
        self.LabelTitle = QLabel(self, text="Excel - Automatação")
        # Excel Columns Automation
        self.buttonExcelXlsAuto = QPushButton(self, text="Exportação de Colunas - Excel")
        self.buttonExcelXlsAuto.setToolTip(
            "<b>Exportação de Colunas:</b>\nExporta e ordena as colunas da planilha selecionada para um arquivo."
        )
        self.buttonExcelXlsAuto.clicked.connect(lambda: controller.show_Page("ExcelXlsAuto"))
        # Excel Concat Automation
        self.buttonExcelConcat = QPushButton(self, text="Concatenação de Planilhas/Arquivos - Excel")
        self.buttonExcelConcat.setToolTip(
            "<b>Concatenação de Planilhas/Arquivos:</b>\nConcatena arquivos e suas planilhas."
        )
        self.buttonExcelConcat.clicked.connect(lambda: controller.show_Page("ExcelConcat"))

    # Grid Configuration
    def GridConfigs(self):
        """
        Function GridConfigs()

        Used to configure this frame grid (columns and rows) for our widgets.
        """

        # Grid Creation
        myGridLayout = QGridLayout(self)
        # space between widgets
        myGridLayout.setSpacing(10)
        # stretch last row
        myGridLayout.setRowStretch(3, 1)
        # Label - Title
        myGridLayout.addWidget(self.LabelTitle, 0, 0, Qt.AlignmentFlag.AlignCenter)
        # Button - Excel Columns Automation
        myGridLayout.addWidget(self.buttonExcelXlsAuto, 1, 0, Qt.AlignmentFlag.AlignCenter)
        # Button - Excel Concatenate Automation
        myGridLayout.addWidget(self.buttonExcelConcat, 2, 0, Qt.AlignmentFlag.AlignCenter)
        # set this widget layout to the grid layout
        self.setLayout(myGridLayout)

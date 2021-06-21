import os
import pandas as pd
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

# Excel Automation - Concatenate Page/Frame Widget
class ExcelConcat(QWidget):
    """
    Frame/Page Widget "Excel Concatenate Automation":

    This Frame Module is responsible to load excel files,
    get all .xls/.xlsx/.xlsm, concatenate/append each and export them
    into a single Excel file.
    """

    def __init__(self, parent, controller):
        """
        Frame/Page Widget "Excel Concatenation":

        Class ExcelConcat(parent, controller):
        parent: This widget parent, usually the container widget.
        controller: QT Main window, to control changing pages.

        Initialization of the Excel Concatenate Automation Class page.
        """

        super().__init__(parent=parent)
        # first create our widgets
        self.CreateWidgets(controller)
        # then set our grid layout and config
        self.GridConfigs()
        # our DataFrame
        self.dataFrame = pd.DataFrame()
        # init our excelFile as None
        self.excelFile = None
        # Date by sheet name bool
        self.DatebySheetName = False

    # Creation of widgets on screen
    def CreateWidgets(self, controller):
        """
        Function CreateWidgets(controller)
        controller: Our main window / controller widget.

        Used to create new widgets (labels, buttons, etc.),
        controller is used in buttons to show another page/frame widget.
        """

        # Label - Title
        self.LabelTitle = QLabel(self, text="Concatenação de Planilhas/Arquivos - Excel")
        # Label - File Status
        self.LabelFileStatus = QLabel(self, text="Esperando Arquivo...")
        # Label - File Status
        self.LabelDevNote = QLabel(
            self, text="As colunas devem ser iguais, se não será adicionado como uma nova coluna."
        )
        # CheckBox - Date by Sheet Name
        self.CheckDateBySheetName = QCheckBox(self, text="Data baseada no nome das planilhas nas pastas de trabalho?")
        self.CheckDateBySheetName.stateChanged.connect(self.IsDateBySheetName)
        # Load Excel Button
        self.LoadExcelFileButton = QPushButton(self, text="Abrir Arquivos do Excel")
        self.LoadExcelFileButton.setToolTip(
            "<b>Abrir Arquivos do Excel:</b> Escolha um ou vários arquivos para carregar."
        )
        self.LoadExcelFileButton.clicked.connect(lambda: self.LoadExcelFiles())
        # Export to Excel Button
        self.ExportExcelFileButton = QPushButton(self, text="Exportar Arquivo do Excel")
        self.ExportExcelFileButton.setToolTip("<b>Exporte Arquivos do Excel:</b> Exporte seu arquivo concatenado.")
        self.ExportExcelFileButton.clicked.connect(lambda: self.ExportExcelFile())
        # Home Page Button
        self.buttonAppHome = QPushButton(self, text="Menu Inicial")
        self.buttonAppHome.clicked.connect(lambda: controller.show_Page("AppHome"))

    # Grid Configuration
    def GridConfigs(self):
        """
        Function GridConfigs()

        Used to configure this frame grid (columns and rows) for our widgets.
        """

        # Grid Creation
        myGridLayout = QGridLayout(self)
        # space between widgets
        myGridLayout.setSpacing(25)
        # stretch last row
        myGridLayout.setRowStretch(6, 1)
        # Label - Title
        myGridLayout.addWidget(self.LabelTitle, 0, 0, 1, 2, Qt.AlignmentFlag.AlignCenter)
        # Label - File Status
        myGridLayout.addWidget(self.LabelFileStatus, 1, 0, 1, 2, Qt.AlignmentFlag.AlignCenter)
        # CheckBox - Date by Sheet Name
        myGridLayout.addWidget(self.CheckDateBySheetName, 2, 0, 1, 2, Qt.AlignmentFlag.AlignCenter)
        # Button Load Excel
        myGridLayout.addWidget(self.LoadExcelFileButton, 3, 0, Qt.AlignmentFlag.AlignCenter)
        # Button Export Excel
        myGridLayout.addWidget(self.ExportExcelFileButton, 3, 1, Qt.AlignmentFlag.AlignCenter)
        # Label - DevNote
        myGridLayout.addWidget(self.LabelDevNote, 4, 0, 1, 2, Qt.AlignmentFlag.AlignCenter)
        # Button - Application Home
        myGridLayout.addWidget(self.buttonAppHome, 5, 0, 1, 2, Qt.AlignmentFlag.AlignCenter)
        # set this widget layout to the grid layout
        self.setLayout(myGridLayout)

    def IsDateBySheetName(self, state):
        """
        Function IsDateBySheetName(state)
        state: The current state of our CheckBox.

        Used to set our date by sheet name, when we change our CheckBox.
        """

        if state == Qt.Checked:
            self.DatebySheetName = True
            print("Checked")
        else:
            self.DatebySheetName = False
            print("Unchecked")

    # Load Excel files from a FileDialog
    def LoadExcelFiles(self):
        """
        Function LoadMultExcelFiles()

        Load multiple excel files for our Dataframe from a FileDialog,
        it also append/concatenate everything into our DataFrame, so later we can
        export it to a Excel File.
        """

        # Make our DataFrame empty each time we load our files
        # To not append data wrong to our DataFrame
        self.dataFrame = pd.DataFrame()
        # Make our Excelfile equals None each time we load our files
        self.excelFile = None
        # File Dialog Filter
        FilesFilter = "Excel 2010 (*.xlsx);; Excel 2003 (*.xls);; Todos Arquivos (*.*)"
        # A File dialog for our App
        # To load multiple files
        Files = QFileDialog.getOpenFileNames(
            parent=self,
            caption="Abrir Arquivos do Excel",
            directory="",
            filter=FilesFilter,
            initialFilter="Excel 2010 (*.xlsx)",
        )
        if Files:
            print(Files)
            try:
                # for each file picked by our file dialog
                for File in Files[0]:
                    self.excelFile = pd.ExcelFile(File)
                    # append all our sheets into one
                    for sheets in self.excelFile.sheet_names:
                        # first parse each sheet
                        appendSheets = pd.DataFrame(self.excelFile.parse(sheets))
                        if self.DatebySheetName == True:
                            # Add a column for the Dates
                            dates = appendSheets.insert(0, "Data", sheets)
                            # format to date our sheet name
                            appendSheets["Data"] = pd.to_datetime(appendSheets["Data"], format="%d%m%Y")
                            # format to DD/MM/YYYY
                            appendSheets["Data"] = appendSheets["Data"].dt.strftime("%d/%m/%Y")
                        # append / concatenate to our data frame
                        self.dataFrame = self.dataFrame.append(appendSheets)
                        # Just for debug purposes
                        print(appendSheets)
                        print("Concatenating sheets and putting to our DataFrame!")
                    print("File loaded and Appended: " + os.path.basename(File))
                # show that the file are ready to export
                self.LabelFileStatus.setText("Arquivo prontos para a exportação!!!")
                # if our files paths are bigger than 0
                if len(Files[0]) > 0:
                    # create a new message box
                    msg = QMessageBox()
                    msg.setWindowTitle("Concatenação completa!")
                    msg.setIcon(QMessageBox.Information)
                    msg.setText("Os seguintes arquivos e suas planilhas, foram carregados e concatenados:")
                    # File loaded String for our MessageBox
                    filesLoadedString = ""
                    # for each path in our files
                    for pathString in Files[0]:
                        # add new lines for each path to separate each file
                        filesLoadedString += pathString + "\n"
                    # set our informative text, showing our loaded files.
                    msg.setInformativeText(filesLoadedString)
                    msg.setStandardButtons(QMessageBox.Ok)
                    # execute the message
                    msg.exec_()
            # Error handling
            except ValueError:
                self.LabelFileStatus.setText("Não é possível abrir esse arquivo: " + os.path.basename(File))
                print("Unable to open this file: " + os.path.basename(File))
            # Obsolete with FileDialog
            # Just here in case, we switch to a automatic directory file loading
            except FileNotFoundError:
                self.LabelFileStatus.setText("O seguinte Arquivo não foi encontrado: " + os.path.basename(File))
                print("File not found, please try another file: " + os.path.basename(File))

    # Export / Save our data frame to a excel file
    def ExportExcelFile(self):
        """
        Function ExportExcelFile()

        Responsible to export / save our appended DataFrame to a
        single Excel File.
        """

        # show a warning in case there's nothing in our excelFile
        if self.excelFile is None:
            # create a new message box
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("Sem nada para exportar:")
            msg.setInformativeText("Arquivos não carregados ou sem arquivo para exportar.")
            msg.setWindowTitle("Sem Arquivo para exportar!")
            msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
            # msg.setDetailedText("Detalhes: \nPrecisamos de um arquivo e suas colunas das planilhas selecionadas")
            # just setting the cancel button, to portuguese
            btnCancel = msg.button(QMessageBox.Cancel)
            btnCancel.setText("Cancelar")
            # execute the message
            msg.exec_()
            return
        # File Dialog Filter
        FilesFilter = "Excel 2010 (*.xlsx);; Excel 2003 (*.xls);; Todos Arquivos (*.*)"
        # a filedialog to export our file
        exportExcelFile = QFileDialog.getSaveFileName(
            parent=self,
            caption="Abrir Arquivo Excel",
            filter=FilesFilter,
            initialFilter="Excel 2010 (*.xlsx)",
        )
        if exportExcelFile[0]:
            ExportDataFrame = self.dataFrame.to_excel(exportExcelFile[0], index=False)
            # Print our current dataframe
            print("Current Data:")
            print(self.dataFrame)
            # Print that we exported our data
            print("Exported Data!!!")
            print("Data exported to: " + exportExcelFile[0])
            # Exported data message box
            # QMessageBox.showinfo(title="Arquivo Exportado!", message="O Arquivo foi exportado para: " + exportExcelFile)
            # Make our DataFrame empty again
            self.dataFrame = pd.DataFrame()
            # Make our Excelfile as None
            self.excelFile = None
            # Make our exportDataFrame as None
            ExportDataFrame = None
            # return our label to default text
            self.LabelFileStatus.setText("Esperando Arquivo...")

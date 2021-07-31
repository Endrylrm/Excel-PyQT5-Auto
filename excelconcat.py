import os
import pandas as pd
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from msgboxutils import (
    CreateInfoMessageBox,
    CreateMessageBox,
    CreateWarningMessageBox,
    DisplayMessageBox,
)

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
        # our DataFrame Dictionary to hold our dataframes and it's sheets names
        self.dataFramesDict = {}
        # our Excel Files list
        self.excelFilesList = []
        # our concatenated excel file, used by Single Sheet option
        self.ConcatenatedFile = None
        # is the Date get by it's sheet name (example: "20122021")?
        self.DatebySheetName = False
        # is this a multi sheet file?
        self.isMultiSheetFile = False
        # show a message in our QListWidget, that we are waiting a file
        self.debugMessage("Esperando Arquivos para concatenação...")

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
        # CheckBox - Multi Sheet File
        self.CheckMultiSheet = QCheckBox(self, text="Exportar para múltiplas planilhas?")
        self.CheckMultiSheet.setToolTip(
            "Em vez de exportar uma única planilha no arquivo, exportar em várias planilhas de um arquivo."
        )
        self.CheckMultiSheet.stateChanged.connect(self.MultiSheetCheck)
        # Load Excel Button
        self.LoadExcelFileButton = QPushButton(self, text="Abrir Arquivos do Excel")
        self.LoadExcelFileButton.setToolTip(
            "<b>Abrir Arquivos do Excel:</b> Escolha um ou vários arquivos para carregar."
        )
        self.LoadExcelFileButton.clicked.connect(self.LoadExcelFiles)
        # Export to Excel Button
        self.ExportExcelFileButton = QPushButton(self, text="Exportar Arquivo do Excel")
        self.ExportExcelFileButton.setToolTip("<b>Exporte Arquivos do Excel:</b> Exporte seu arquivo concatenado.")
        self.ExportExcelFileButton.clicked.connect(self.ExportExcelFile)
        # Debug message Box - List Widget
        self.listDebugMessages = QListWidget(self)
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
        # myGridLayout.setRowStretch(8, 1)
        # Label - Title
        myGridLayout.addWidget(self.LabelTitle, 0, 0, 1, 2, Qt.AlignmentFlag.AlignCenter)
        # Label - File Status
        myGridLayout.addWidget(self.LabelFileStatus, 1, 0, 1, 2, Qt.AlignmentFlag.AlignCenter)
        # CheckBox - Date by Sheet Name
        myGridLayout.addWidget(self.CheckDateBySheetName, 2, 0, 1, 2, Qt.AlignmentFlag.AlignCenter)
        # CheckBox - Multi sheet files
        myGridLayout.addWidget(self.CheckMultiSheet, 3, 0, 1, 2, Qt.AlignmentFlag.AlignCenter)
        # Button Load Excel
        myGridLayout.addWidget(self.LoadExcelFileButton, 4, 0, Qt.AlignmentFlag.AlignCenter)
        # Button Export Excel
        myGridLayout.addWidget(self.ExportExcelFileButton, 4, 1, Qt.AlignmentFlag.AlignCenter)
        # Label - DevNote
        myGridLayout.addWidget(self.LabelDevNote, 5, 0, 1, 2, Qt.AlignmentFlag.AlignCenter)
        # List Widget - Debug messages
        myGridLayout.addWidget(self.listDebugMessages, 6, 0, 1, 2)
        # Button - Application Home
        myGridLayout.addWidget(self.buttonAppHome, 7, 0, 1, 2, Qt.AlignmentFlag.AlignCenter)
        # set this widget layout to the grid layout
        self.setLayout(myGridLayout)

    # CheckBox - Date by it's sheet name - callback
    def IsDateBySheetName(self, state):
        """
        Function IsDateBySheetName(state)
        state: The current state of our CheckBox.

        Used to set our date by the sheet name, when we change our CheckBox state.
        """

        if state == Qt.Checked:
            self.DatebySheetName = True
            print("Checked")
        else:
            self.DatebySheetName = False
            print("Unchecked")

    # CheckBox - Multi Sheet File - callback
    def MultiSheetCheck(self, state):
        """
        Function MultiSheetCheck(state)
        state: The current state of our CheckBox.

        Used to set exported file as a multi sheet file, when we change our CheckBox state.
        So we concatenate every sheet into the exported file.
        """

        if state == Qt.Checked:
            self.isMultiSheetFile = True
            print("Checked")
        else:
            self.isMultiSheetFile = False
            print("Unchecked")

    def debugMessage(self, msg: str, bold: bool = False, italic: bool = False):
        """
        Function debugMessage(msg)
        msg: The message we are going to show on our QListWidget.
        bold: is the font bold? Defaults to False.
        italic: is the font italic? Defaults to False.

        Used to show a message in our QListWidget.
        """

        # the item used to show our message
        item = QListWidgetItem(msg)
        # set to no flags, so our item is unselectable
        item.setFlags(Qt.NoItemFlags)
        # Create a new font
        itemFont = QFont()
        # set the font to bold
        if bold:
            # set our item font to bold
            itemFont.setBold(True)
        # set font to italic
        if italic:
            # set our item font to italic
            itemFont.setItalic(True)
        # set our item font to the new font
        item.setFont(itemFont)
        # Add our item to our List Widget
        self.listDebugMessages.addItem(item)

    # Load Excel files from a FileDialog
    def LoadExcelFiles(self):
        """
        Function LoadMultExcelFiles()

        Load multiple excel files for our Dataframe Dictionary from a FileDialog,
        it also append/concatenate everything into our Concatenated DataFrame, so later
        we can export it to a Excel File, as a single or multi sheet file.
        """

        # Make our DataFrame Dictionary empty each time we load our files
        # To not concatenate wrong data to our DataFrame
        self.dataFramesDict = {}
        # Make our Excel files list empty, each time we load our files
        self.excelFilesList = []
        # Make concatenated excel file as None, each time we load a file
        self.ConcatenatedFile = None
        # clear our Debug Messages list
        self.listDebugMessages.clear()
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
                    print("File loaded: " + os.path.basename(File))
                    # show a message in our QListWidget, that we loaded our file
                    self.debugMessage("Arquivo carregado: " + os.path.basename(File), True)
                    # excel file in our files
                    excelFile = pd.ExcelFile(File)
                    # debug a message for each sheet in out excel file
                    for debugSheet in excelFile.sheet_names:
                        # show a message in our QListWidget, that we are concatenating the sheets
                        self.debugMessage("Concatenando a planilha: " + debugSheet)
                    # append / concatenate to our excel file list
                    self.excelFilesList.append(excelFile)
                    # Bugfix: used for files with sheets that have the same name
                    currentSheet = 0
                    # for each excel file in our excel files list
                    for xlsFile in self.excelFilesList:
                        # turn our sheets into DataFrames
                        for sheets in xlsFile.sheet_names:
                            # first parse each sheet
                            SheetsToConcatenate = pd.DataFrame(xlsFile.parse(sheets))
                            # Date by sheet name (example = "01012021")
                            if self.DatebySheetName == True:
                                # Add a column for the Dates
                                dates = SheetsToConcatenate.insert(0, "Data", sheets)
                                # format to date our sheet name
                                SheetsToConcatenate["Data"] = pd.to_datetime(
                                    SheetsToConcatenate["Data"], format="%d%m%Y"
                                )
                                # format to DD/MM/YYYY
                                SheetsToConcatenate["Data"] = SheetsToConcatenate["Data"].dt.strftime("%d/%m/%Y")
                            # Just for debug purposes
                            print(SheetsToConcatenate)
                            # Dictionaries can't have two or more Keys with the same "name"
                            if sheets in self.dataFramesDict:
                                self.dataFramesDict[sheets + str(currentSheet)] = SheetsToConcatenate
                            else:
                                self.dataFramesDict[sheets] = SheetsToConcatenate
                            # increment our currentSheet
                            currentSheet += 1
                            # print our keys
                            print(self.dataFramesDict.keys())
                # if our files paths are bigger than 0
                if len(Files[0]) > 0:
                    print("Concatenating sheets and putting to our DataFrame!")
                    # concatenate our files and sheets into a single one
                    # from our DataFrame Dictionary
                    concatExcelFiles = pd.concat(self.dataFramesDict.values(), ignore_index=True)
                    # Turn our concatenated files in a DataFrame
                    self.ConcatenatedFile = pd.DataFrame(concatExcelFiles)
                    # show that the file are ready to export
                    self.LabelFileStatus.setText("Arquivo prontos para a exportação!!!")
                    # show a message in our QListWidget, that we concatenated our file
                    self.debugMessage("Arquivos concatenados com sucesso, esperando exportação!", True)
                    # File loaded String for our MessageBox
                    filesLoadedString = ""
                    # for each path in our files
                    for pathString in Files[0]:
                        # add new lines for each path to separate each file
                        filesLoadedString += os.path.basename(pathString) + "\n"
                    # create a new message box
                    CreateInfoMessageBox(
                        msgWinTitle="Concatenação completa!",
                        msgText="Os seguintes arquivos e suas planilhas, foram carregados e concatenados:",
                        msgInfoText=filesLoadedString,
                    )
            # Error handling
            except ValueError:
                # just checking if there's a file
                if len(Files[0]) > 0:
                    self.LabelFileStatus.setText("Não é possível abrir esse arquivo: " + os.path.basename(File))
                    print("Unable to open this file: " + os.path.basename(File))
                    # show a message in our QListWidget, that we are unable to open this file
                    self.debugMessage("Não é possível abrir esse arquivo: " + os.path.basename(File))
            # Obsolete with FileDialog
            # Just here in case, we switch to a automatic directory file loading
            except FileNotFoundError:
                self.LabelFileStatus.setText("O seguinte Arquivo não foi encontrado: " + os.path.basename(File))
                print("File not found, please try another file: " + os.path.basename(File))

    # Export / Save our DataFrame Dictionary to a excel file
    def ExportExcelFile(self):
        """
        Function ExportExcelFile()

        Responsible to export / save our concatenated / appended DataFrame Dictionary
        to a single Excel File, as single or multi sheet file.
        """

        # First show a warning and return
        # in case there's nothing in our DataFrame Dictionary or Concatenated File
        if self.ConcatenatedFile is None or len(self.dataFramesDict) == 0:
            # create a new message box and display it
            # to show a warning that there isn't a file to export
            CreateWarningMessageBox(
                msgWinTitle="Sem Arquivo para exportar!",
                msgText="Sem nada para exportar:",
                msgInfoText="Arquivos não carregados ou sem arquivo para exportar.",
            )
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
            ExportDataFrame = None
            # multi sheets file Export handling
            if self.isMultiSheetFile == True:
                # create a Excel Writer for our multi sheet file
                writer = pd.ExcelWriter(exportExcelFile[0])
                # for each sheet in our DataFrame Dictionary
                for sheet in self.dataFramesDict:
                    # export to our excel file as a new sheet in our file
                    ExportDataFrame = self.dataFramesDict[sheet].to_excel(writer, sheet_name=sheet, index=False)
                # Print our current dataframe
                print("Current Data:")
                print(self.dataFramesDict)
                # Save our file to our destination
                writer.save()
            else:
                # otherwise export a single sheet file and it's concatenated sheets
                ExportDataFrame = self.ConcatenatedFile.to_excel(exportExcelFile[0], index=False)
                # Print our current dataframe
                print("Current Data:")
                print(self.ConcatenatedFile)
            # Print that we exported our data
            print("Exported Data!!!")
            # print where we exported our file
            print("Data exported to: " + exportExcelFile[0])
            # clear our Debug Messages list
            self.listDebugMessages.clear()
            # create a new message box and display it
            # to show that we successfully exported our file
            msg = CreateMessageBox(
                msgWinTitle="Exportação Completa!!!",
                msgText="O Arquivo foi exportado para:",
                msgInfoText=exportExcelFile[0],
            )
            # display our message box
            DisplayMessageBox(msg)
            # Make our DataFrame Dictionary empty again
            self.dataFramesDict = {}
            # Make our Excel files List empty
            self.excelFilesList = []
            # Make our sheet name list empty
            self.ConcatenatedFile = None
            # Make our exportDataFrame as None
            ExportDataFrame = None
            # show a message in our QListWidget, that we are waiting for a file
            self.debugMessage("Esperando Arquivos para concatenação...")
            # return our label to default text
            self.LabelFileStatus.setText("Esperando Arquivo...")

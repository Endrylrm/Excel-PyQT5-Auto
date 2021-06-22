import os
import pandas as pd
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

# Excel Automation - Columns Page/Frame Widget
class ExcelXlsAuto(QWidget):
    """
    Frame/Page Widget "Excel Files Automation - Columns":

    This Frame Module is responsible to load excel files,
    get all relevant data from .xls/.xlsx/.xlsm and export them
    into a single Excel file.
    """

    def __init__(self, parent, controller):
        """
        Frame/Page Widget "Excel Files Automation - Columns":

        Class ExcelXlsAuto(parent, controller):
        parent: This widget parent, usually the container widget.
        controller: QT Main window, to control changing pages.

        Initialization of the Excel Columns Automation Class page.
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
        # Current Selected Sheet
        self.CurrentSelectedSheet = None
        # Current Selected columns
        self.selectedColumns = []
        # Current Selected columns in our sort ListWidget
        self.sortSelectedColumns = []

    # Creation of widgets on screen
    def CreateWidgets(self, controller):
        """
        Function CreateWidgets(controller)
        controller: Our main window / controller widget.

        Used to create new widgets (labels, buttons, etc.),
        controller is used in buttons to show another page/frame widget.
        """

        # Label Title
        self.LabelTitle = QLabel(self, text="Exportação de Colunas - Excel")
        # Label Title
        self.LabelFileStatus = QLabel(self, text="Esperando Arquivo...")
        # Load Excel Button
        self.LoadExcelFileButton = QPushButton(self, text="Abrir Arquivos do Excel")
        self.LoadExcelFileButton.setToolTip(
            "<b>Abrir Arquivos do Excel:</b> Escolha um ou vários arquivos para carregar."
        )
        self.LoadExcelFileButton.clicked.connect(lambda: self.LoadExcelFiles())
        # Export to Excel Button
        self.ExportExcelFileButton = QPushButton(self, text="Exportar Arquivo do Excel")
        self.ExportExcelFileButton.clicked.connect(lambda: self.ExportExcelFile())
        # Home Page Button
        self.buttonAppHome = QPushButton(self, text="Menu Inicial")
        self.buttonAppHome.clicked.connect(lambda: controller.show_Page("AppHome"))
        # QListWidget - Excel Files ListBox
        self.LabelXlsFiles = QLabel(self, text="Arquivos")
        self.listBoxXlsFiles = QListWidget(self)
        self.listBoxXlsFiles.itemClicked.connect(self.FileSelection)
        # QListWidget - Excel Sheets ListBox
        self.LabelXlsSheets = QLabel(self, text="Planilhas")
        self.listBoxXlsSheets = QListWidget(self)
        self.listBoxXlsSheets.itemClicked.connect(self.SheetSelection)
        # QListWidget - Excel Sheets Columns ListBox
        self.LabelXlsColumns = QLabel(self, text="Colunas")
        self.listBoxXlsColumns = QListWidget(self)
        self.listBoxXlsColumns.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.listBoxXlsColumns.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.listBoxXlsColumns.itemSelectionChanged.connect(self.ColumnsSelection)
        # QListWidget - Excel Sheets Sort Columns ListBox
        self.LabelXlsSortColumns = QLabel(self, text="Ordenar Colunas")
        self.listBoxXlsSortColumns = QListWidget(self)
        self.listBoxXlsSortColumns.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.listBoxXlsSortColumns.itemSelectionChanged.connect(self.SortColumnsSelection)

    # Grid Configuration
    def GridConfigs(self):
        """
        Function GridConfigs()

        Used to configure this frame grid (columns and rows) for our widgets.
        """

        # Grid Creation
        myGridLayout = QGridLayout(self)
        # Label - Title
        myGridLayout.addWidget(self.LabelTitle, 0, 0, 1, 4, Qt.AlignmentFlag.AlignCenter)
        # Label - File Status
        myGridLayout.addWidget(self.LabelFileStatus, 1, 0, 1, 4, Qt.AlignmentFlag.AlignCenter)
        # Button Load Excel
        myGridLayout.addWidget(self.LoadExcelFileButton, 2, 0, Qt.AlignmentFlag.AlignCenter)
        # Button Export Excel
        myGridLayout.addWidget(self.ExportExcelFileButton, 2, 1, Qt.AlignmentFlag.AlignCenter)
        # Label - Excel Files
        myGridLayout.addWidget(self.LabelXlsFiles, 3, 0, Qt.AlignmentFlag.AlignCenter)
        # Label - Excel Sheets
        myGridLayout.addWidget(self.LabelXlsSheets, 3, 1, Qt.AlignmentFlag.AlignCenter)
        # Label - Excel Sheets Columns
        myGridLayout.addWidget(self.LabelXlsColumns, 3, 2, Qt.AlignmentFlag.AlignCenter)
        # Label - Excel Sheets Sort Columns
        myGridLayout.addWidget(self.LabelXlsSortColumns, 3, 3, Qt.AlignmentFlag.AlignCenter)
        # List Widget - Excel Files
        myGridLayout.addWidget(self.listBoxXlsFiles, 4, 0)
        # List Widget - Excel Sheets
        myGridLayout.addWidget(self.listBoxXlsSheets, 4, 1)
        # List Widget - Excel Sheets Columns
        myGridLayout.addWidget(self.listBoxXlsColumns, 4, 2)
        # List Widget - Excel Sort Sheets Columns
        myGridLayout.addWidget(self.listBoxXlsSortColumns, 4, 3)
        # Button - Application Home
        myGridLayout.addWidget(self.buttonAppHome, 5, 0)
        # set this widget layout to the grid layout
        self.setLayout(myGridLayout)

    # Load Excel files from a FileDialog
    def LoadExcelFiles(self):
        """
        Function LoadExcelFiles()

        Load multiple excel files for our Dataframe from a FileDialog,
        it put's every file path in a ListBox so we can select the file
        that we want, so later we can export it again to a Excel File,
        with all the data that we might need.
        """

        # Make our DataFrame empty each time we load our files
        # To not append data wrong to our DataFrame
        self.dataFrame = pd.DataFrame()
        # Make our Excelfile equals None each time we load our files
        self.excelFile = None
        # set to none our current selected Sheet
        self.CurrentSelectedSheet = None
        # let's first delete anything in our Lists
        # when we load another file.
        self.listBoxXlsFiles.clear()
        self.listBoxXlsSheets.clear()
        self.listBoxXlsColumns.clear()
        self.listBoxXlsSortColumns.clear()
        self.selectedColumns.clear()
        self.sortSelectedColumns.clear()
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
                    # create a List Item with our paths
                    item = QListWidgetItem(File)
                    # append the files paths to our ListBox
                    self.listBoxXlsFiles.addItem(item)
                # if our files paths are bigger than 0
                if len(Files[0]) > 0:
                    # set our file status label message
                    self.LabelFileStatus.setText("Arquivos carregados!!!")
                    # create a new message box
                    msg = QMessageBox()
                    msg.setWindowTitle("Arquivos carregados!")
                    msg.setIcon(QMessageBox.Information)
                    msg.setText("Os seguintes arquivos foram carregados:")
                    # File loaded String for our MessageBox
                    filesLoadedString = ""
                    # for each path in our files
                    for pathString in Files[0]:
                        # add new lines for each path to separate each file
                        filesLoadedString += pathString + "\n"
                    # set our informative text, showing our loaded files.
                    msg.setInformativeText(filesLoadedString)
                    msg.setStandardButtons(QMessageBox.Ok)
                    # execute the message box
                    msg.exec_()
            # Error handling
            except ValueError:
                if len(Files[0]) > 0:
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
            msg.setInformativeText("Colunas não selecionadas ou sem arquivo para exportar.")
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
            # When we rearrange our List
            self.dataFrame = self.dataFrame[self.selectedColumns]
            # Sorting our dataframe
            # requires atleast one column selected in our Sort column ListBox
            if len(self.sortSelectedColumns) > 0:
                sortedDF = self.dataFrame.sort_values(by=self.sortSelectedColumns, ascending=True, inplace=True)
                print("Data sorted!!!")
            print(self.dataFrame)
            # Export our data frame to a Excel file
            ExportDataFrame = self.dataFrame.to_excel(
                exportExcelFile[0], sheet_name=self.CurrentSelectedSheet, columns=self.selectedColumns, index=False
            )
            # Print our current dataframe
            print("Current Data:")
            print(self.dataFrame)
            # create a new message box
            msg = QMessageBox()
            msg.setWindowTitle("Exportação Completa!!!")
            msg.setIcon(QMessageBox.Information)
            msg.setText("O Arquivo foi exportado para:")
            # set our informative text, showing our loaded files.
            msg.setInformativeText(exportExcelFile[0])
            msg.setStandardButtons(QMessageBox.Ok)
            # execute the message box
            msg.exec_()
            # Print that we exported our data
            print("Exported Data!!!")
            print("Data exported to: " + exportExcelFile[0])
            # Exported data message box
            # QMessageBox.showinfo(title="Arquivo Exportado!", message="O Arquivo foi exportado para: " + exportExcelFile)
            # Make our DataFrame empty again
            self.dataFrame = pd.DataFrame()
            # Make our Excelfile as None
            self.excelFile = None
            # clear our ListBoxes, except the file one
            self.listBoxXlsSheets.clear()
            self.listBoxXlsColumns.clear()
            self.listBoxXlsSortColumns.clear()
            # also clear everything in our lists
            # Selected columns list
            self.selectedColumns.clear()
            # Selected columns to sort List
            self.sortSelectedColumns.clear()
            # Make our current selected Sheet as None
            self.CurrentSelectedSheet = None
            # Make our exportDataFrame as None
            ExportDataFrame = None
            sortedDF = None
            # return our label to default text
            self.LabelFileStatus.setText("Esperando Arquivo...")

    # ListBox File Selection callback
    def FileSelection(self):
        """
        Function FileSelection()

        Called when we select something in our Excel Files ListBox,
        this Load our Excel Sheets from the selected file to it's
        respective ListBox, through selection in our ListBox.
        """
        selection = self.listBoxXlsFiles.selectedItems()
        # if something is selected
        if selection:
            # first delete everything in our Sheets ListWidgets
            self.listBoxXlsSheets.clear()
            # then clear the Columns ListWidget
            self.listBoxXlsColumns.clear()
            # and clear our Sort Columns ListWidget
            self.listBoxXlsSortColumns.clear()
            # get the selected position and data
            selectedFile = selection[0].text()
            # set our excelFile
            self.excelFile = pd.ExcelFile(selectedFile, engine="openpyxl")
            # for each sheet in our excel file
            for sheet in self.excelFile.sheet_names:
                # create a ListWidgetItem
                item = QListWidgetItem(sheet)
                # and store our sheets in our ListWidget
                self.listBoxXlsSheets.addItem(item)

    # ListBox Sheet Selection callback
    def SheetSelection(self):
        """
        Function SheetSelection()

        Called when we select something in our Excel Sheets ListBox,
        this load all our Excel Sheet - Columns to it's respective ListBox
        through selection in our ListBox.
        """
        selection = self.listBoxXlsSheets.selectedItems()
        # if something is selected
        if selection:
            # first delete everything in our Column ListBox
            self.listBoxXlsColumns.clear()
            # also delete everything in our Sort Columns ListBox
            self.listBoxXlsSortColumns.clear()
            # clear our Selected Columns List
            self.selectedColumns.clear()
            self.sortSelectedColumns.clear()
            # get the selected sheet in our list
            selectedSheet = selection[0].text()
            # print the selected sheet
            print(selectedSheet)
            # the current selected sheet, we need this later to export
            self.CurrentSelectedSheet = selectedSheet
            # print the current selected sheet
            print(self.CurrentSelectedSheet)
            # parse the excel file as a dataFrame
            columnData = pd.DataFrame(self.excelFile.parse(self.CurrentSelectedSheet))
            # set our dataFrame as our selected sheet
            self.dataFrame = columnData
            # print our dataFrame
            print(self.dataFrame)
            # for columns in our selected sheet
            for column in columnData.columns:
                # create a ListWidgetItem
                item = QListWidgetItem(column)
                # Add to our ListBox
                self.listBoxXlsColumns.addItem(item)

    # ListBox Columns Selection callback
    def ColumnsSelection(self):
        """
        Function ColumnsSelection()

        Called when we select the columns in our Excel columns ListBox,
        this let's us select the columns we want by selection in our ListBox,
        so we can later export only the selected columns.

        It also add the selected columns, to the Sort Columns ListBox.
        """
        selection = self.listBoxXlsColumns.selectedItems()
        # if something is selected
        if selection:
            # first clear our list for multiselection
            self.selectedColumns.clear()
            # get every selected item
            for sel in selection:
                # first delete everything in our Sort Column ListBox
                self.listBoxXlsSortColumns.clear()
                # get the data from the current selected columns
                currentSelectedCols = sel.text()
                # Print our current selected columns
                print(currentSelectedCols)
                # append to our List
                self.selectedColumns.append(currentSelectedCols)
                for column in self.selectedColumns:
                    # create a ListWidgetItem
                    item = QListWidgetItem(column)
                    # Add to our ListBox
                    self.listBoxXlsSortColumns.addItem(item)
            # Print our selected columns List
            print(self.selectedColumns)

    # ListBox to Sort Columns Selection callback
    def SortColumnsSelection(self):
        """
        Function ColumnsSelection()

        Called when we select the columns in our Excel columns ListBox,
        this let's us select which column to sort by selection in our ListBox.
        """
        selection = self.listBoxXlsSortColumns.selectedItems()
        # if something is selected
        if selection:
            # first clear our list for multiselection
            self.sortSelectedColumns.clear()
            # get every selected item
            for sel in selection:
                # get the data from the current selected columns
                currentSelectedCols = sel.text()
                # Print our current selected columns
                print(currentSelectedCols)
                # append to our List
                self.sortSelectedColumns.append(currentSelectedCols)
            # Print our selected columns to sort List
            print(self.sortSelectedColumns)

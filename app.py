from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

from apphome import AppHome
from excelxlsauto import ExcelXlsAuto
from excelconcat import ExcelConcat

# Application Main Window
class MyApp(QMainWindow):
    """
    Main Window Widget "Application":

    This is our Main Window, it is responsible to Load, switch
    and show the frame that we want.

    Uses a old Tkinter Frame/Page Switch logic.
    """

    def __init__(self, AppTitle, AppWidth, AppHeight, AppIcon, *args, **kwargs):
        """
        Main Window Widget "Application":

        Class MyApp(AppTitle, AppWidth, AppHeight):
        AppTitle: Used to change the Application Window Title.
        AppWidth: Used to change the Application Window Width.
        AppHeight: Used to change the Application Window Height.

        Initialization of Application (Main Window) Class.
        """

        super().__init__(*args, **kwargs)
        # Window Title
        self.setWindowTitle(AppTitle)
        # window size
        self.CenterWindow(AppWidth, AppHeight)
        # the Window icon
        self.setWindowIcon(QIcon(AppIcon))
        # a container for our pages widgets
        self.container = QWidget(self)
        # it's geometry is the size of our window
        self.container.setGeometry(0, 0, AppWidth, AppHeight)
        # our container grid layout
        self.containerGrid = QGridLayout(self.container)
        # set our container layout a grid layout
        self.container.setLayout(self.containerGrid)
        # Start our pages Dictionary with nothing
        self.pages = {}
        # current Page
        self.currentPage = None
        # init our pages
        self.initPages(self.container)
        # show the home page when we open our app
        self.show_Page("AppHome")
        # set our container as our central widget
        # so our pages / frames resize with our window
        self.setCentralWidget(self.container)

    # Initialization of pages/frames
    def initPages(self, container):
        """
        Function initPages(container)
        container: Our container module, to control our widget size.
        AppWidth: The Page Width.
        AppHeight: The Page Height.

        Initialization for our Frames/Pages, requires a container.
        """

        self.pages["AppHome"] = AppHome(container, self)
        self.pages["AppHome"].hide()
        self.pages["ExcelXlsAuto"] = ExcelXlsAuto(container, self)
        self.pages["ExcelXlsAuto"].hide()
        self.pages["ExcelConcat"] = ExcelConcat(container, self)
        self.pages["ExcelConcat"].hide()

        # for each of our pages
        for page in self.pages:
            # Add our pages to Main Window container grid
            self.containerGrid.addWidget(self.pages[page], 0, 0)

    # Page/Frame to show function
    def show_Page(self, page):
        """
        Function show_Page(page)
        page: The string reference for the page in our pages Dictionary.

        Show a page for the given page name.
        """

        # hide our current page
        if self.currentPage is not None:
            self.pages[self.currentPage].hide()
        # show our new page
        newPage = self.pages[page].show()
        # set our current page as our new page
        self.currentPage = page

    # Center window function
    def CenterWindow(self, AppWidth=300, AppHeight=300):
        """
        Function CenterWindow(width, height)
        width: The width of our root/window module.
        height: the height of our root/window module.

        Center our windows on screen and controls it's size.
        """

        # set the size
        self.setGeometry(0, 0, AppWidth, AppHeight)
        # get the frame geometry (size and position)
        AppGeometry = self.frameGeometry()
        # get the center of the screen
        screen = self.screen().availableGeometry().center()
        # Move the frame to the center of the screen
        AppGeometry.moveCenter(screen)
        # Also move our widget to the center of the screen
        # based on the top left point
        self.move(AppGeometry.topLeft())

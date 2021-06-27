__author__ = "Endryl Richard Monteiro"
__author_email__ = "Endrylrm@hotmail.com"

from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

"""
Message Box Utilities:

Creation and editing of message box made faster.
"""

# Edit properties in a Message Box
def EditMessageBoxProps(
    msg: QMessageBox,
    msgWinTitle=None,
    msgText=None,
    msgInfoText=None,
    msgDetailText=None,
    msgIcon=None,
    msgButtons=None,
    msgDefaultButton=None,
    msgEscapeButton=None,
):
    """
    Function EditMessageBoxProps(msg, msgWinTitle, msgText, msgInfoText, msgIcon, msgButtons, msgDefaultButton, msgEscapeButton)

    msg: The message box created previously.
    msgWinTitle: the window title of our message box.
    msgText: the text in our message box.
    msgInfoText: the information text in our message box.
    msgDetailText: the detailed text of our message box.
    msgIcon: the icon on our message box.
    msgButtons: the button on our message box.
    msgDefaultButton: the confirmation button/key, defaults to enter.
    msgEscapeButton: the escape button/key, defaults to esc.

    Used to edit properties in our message Box.
    """
    if msg is None:
        print("no MessageBox to Edit!")
        return

    # set message box window title, if it's not none.
    if msgWinTitle is not None:
        msg.setWindowTitle(msgWinTitle)
    # set message box text, if it's not none.
    if msgText is not None:
        msg.setText(msgText)
    # set our informative text, if it's not none.
    if msgInfoText is not None:
        msg.setInformativeText(msgInfoText)
    # set our detailed text, if it's not none.
    if msgDetailText is not None:
        msg.setDetailedText(msgDetailText)
    # set message box icon, if it's not none.
    if msgIcon is not None:
        msg.setIcon(msgIcon)
    # set message box button, if it's not none.
    if msgButtons is not None:
        msg.setStandardButtons(msgButtons)
    # set message box default confirmation button/key (enter key), if it's not none.
    if msgDefaultButton is not None:
        msg.setDefaultButton(msgDefaultButton)
    # set message box escape confirmation button/key (esc key), if it's not none.
    if msgEscapeButton is not None:
        msg.setEscapeButton(msgEscapeButton)


# Display a Message Box
def DisplayMessageBox(msg: QMessageBox):
    """
    Function DisplayMessageBox(msg)
    msg: The message box you want to execute/display.

    Used to display a message box.
    """

    # execute the message box
    msg.exec_()


# Create a new Message Box
def CreateMessageBox(
    msgWinTitle=None,
    msgText=None,
    msgInfoText=None,
    msgDetailText=None,
    msgIcon=QMessageBox.Information,
    msgButtons=QMessageBox.Ok,
    msgDefaultButton=None,
    msgEscapeButton=None,
) -> QMessageBox:
    """
    Function CreateMessageBox(msgWinTitle, msgText, msgInfoText, msgDetailText, msgIcon, msgButtons, msgDefaultButton, msgEscapeButton)

    msgWinTitle: Message box window title.
    msgText: The text in our message box.
    msgInfoText: the information text in our message box.
    msgDetailText: the detailed text of our message box, defaults to None.
    msgIcon: the icon on our message box (defaults to QMessageBox.Information).
    msgButtons: the button on our message box (defaults to QMessageBox.Ok).
    msgDefaultButton: the confirmation button/key, defaults to enter.
    msgEscapeButton: the escape button/key, defaults to esc.

    Used to create a new message box, that's editable to your needs,
    returns a QMessageBox object.
    """

    # create a new message box
    msg = QMessageBox()
    # set message box window title, if it's not none.
    if msgWinTitle is not None:
        msg.setWindowTitle(msgWinTitle)
    # set message box text, if it's not none.
    if msgText is not None:
        msg.setText(msgText)
    # set our informative text, if it's not none.
    if msgInfoText is not None:
        msg.setInformativeText(msgInfoText)
    # set our detailed text, if it's not none.
    if msgDetailText is not None:
        msg.setDetailedText(msgDetailText)
    # set message box icon, if it's not none.
    if msgIcon is not None:
        msg.setIcon(msgIcon)
    # set message box button, if it's not none.
    if msgButtons is not None:
        msg.setStandardButtons(msgButtons)
    # set message box default confirmation button/key (enter key), if it's not none.
    if msgDefaultButton is not None:
        msg.setDefaultButton(msgDefaultButton)
    # set message box escape confirmation button/key (esc key), if it's not none.
    if msgEscapeButton is not None:
        msg.setEscapeButton(msgEscapeButton)
    # return our new message box
    return msg


# Create a new Information Message Box
def CreateInfoMessageBox(
    msgWinTitle,
    msgText,
    msgInfoText,
    msgDetailText=None,
    msgButtons=QMessageBox.Ok,
    msgDefaultButton=None,
    msgEscapeButton=None,
):
    """
    Function CreateInfoMessageBox(winTitle, msgText, msgInfoText, msgDetailText, msgButtons, msgDefaultButton, msgEscapeButton)
    msgWinTitle: Message box window title.
    msgText: The text in our message box.
    msgInfoText: the information text in our message box.
    msgDetailText: the detailed text of our message box, defaults to None.
    msgButtons: the button on our message box (defaults to QMessageBox.Ok).
    msgDefaultButton: the confirmation button/key, defaults to enter.
    msgEscapeButton: the escape button/key, defaults to esc.

    Used to create a new information message box and display it.
    """

    # create a new message box
    msg = CreateMessageBox(
        # set message box window title
        msgWinTitle=msgWinTitle,
        # set message box text
        msgText=msgText,
        # set our information text.
        msgInfoText=msgInfoText,
        # set our detailed text.
        msgDetailText=msgDetailText,
        # this is a information message box
        msgIcon=QMessageBox.Information,
        # set message box button
        msgButtons=msgButtons,
        # set message box default confirmation button/key (enter key)
        msgDefaultButton=msgDefaultButton,
        # set message box escape confirmation button/key (esc key)
        msgEscapeButton=msgEscapeButton,
    )
    # execute the message box
    DisplayMessageBox(msg)


# Create a new Warning Message Box
def CreateWarningMessageBox(
    msgWinTitle,
    msgText,
    msgInfoText,
    msgDetailText=None,
    msgButtons=QMessageBox.Ok,
    msgDefaultButton=None,
    msgEscapeButton=None,
):
    """
    Function CreateWarningMessageBox(winTitle, msgText, msgInfoText, msgDetailText, msgButtons, msgDefaultButton, msgEscapeButton)
    msgWinTitle: Message box window title.
    msgText: The text in our message box.
    msgInfoText: the information text in our message box.
    msgDetailText: the detailed text of our message box, defaults to None.
    msgButtons: the button on our message box (defaults to QMessageBox.Ok).
    msgDefaultButton: the confirmation button/key, defaults to enter.
    msgEscapeButton: the escape button/key, defaults to esc.

    Used to create a new Warning message box and display it.
    """

    # create a new message box
    msg = CreateMessageBox(
        # set message box window title
        msgWinTitle=msgWinTitle,
        # set message box text
        msgText=msgText,
        # set our information text.
        msgInfoText=msgInfoText,
        # set our detailed text.
        msgDetailText=msgDetailText,
        # this is a Warning message box
        msgIcon=QMessageBox.Warning,
        # set message box button
        msgButtons=msgButtons,
        # set message box default confirmation button/key (enter key)
        msgDefaultButton=msgDefaultButton,
        # set message box escape confirmation button/key (esc key)
        msgEscapeButton=msgEscapeButton,
    )
    # execute the message box
    DisplayMessageBox(msg)


# Create a new Critical/Error Message Box
def CreateErrorMessageBox(
    msgWinTitle,
    msgText,
    msgInfoText,
    msgDetailText=None,
    msgButtons=QMessageBox.Ok,
    msgDefaultButton=None,
    msgEscapeButton=None,
):
    """
    Function CreateErrorMessageBox(winTitle, msgText, msgInfoText, msgDetailText, msgButtons, msgDefaultButton, msgEscapeButton)
    msgWinTitle: Message box window title.
    msgText: The text in our message box.
    msgInfoText: the information text in our message box.
    msgDetailText: the detailed text of our message box, defaults to None.
    msgButtons: the button on our message box (defaults to QMessageBox.Ok).
    msgDefaultButton: the confirmation button/key, defaults to enter.
    msgEscapeButton: the escape button/key, defaults to esc.

    Used to create a new critical/error message box and display it.
    """

    # create a new message box
    msg = CreateMessageBox(
        # set message box window title
        msgWinTitle=msgWinTitle,
        # set message box text
        msgText=msgText,
        # set our information text.
        msgInfoText=msgInfoText,
        # set our detailed text.
        msgDetailText=msgDetailText,
        # this is a Critical/Error message box
        msgIcon=QMessageBox.Critical,
        # set message box button
        msgButtons=msgButtons,
        # set message box default confirmation button/key (enter key)
        msgDefaultButton=msgDefaultButton,
        # set message box escape confirmation button/key (esc key)
        msgEscapeButton=msgEscapeButton,
    )
    # execute the message box
    DisplayMessageBox(msg)


# Create a new Question Message Box
def CreateQuestionMessageBox(
    msgWinTitle,
    msgText,
    msgInfoText,
    msgDetailText=None,
    msgButtons=QMessageBox.Ok,
    msgDefaultButton=None,
    msgEscapeButton=None,
) -> QMessageBox:
    """
    Function CreateQuestionMessageBox(winTitle, msgText, msgInfoText, msgDetailText, msgButtons, msgDefaultButton, msgEscapeButton)
    msgWinTitle: Message box window title.
    msgText: The text in our message box.
    msgInfoText: the information text in our message box.
    msgDetailText: the detailed text of our message box, defaults to None.
    msgButtons: the button on our message box (defaults to QMessageBox.Ok).
    msgDefaultButton: the confirmation button/key, defaults to enter.
    msgEscapeButton: the escape button/key, defaults to esc.

    Used to create a new Question message box, returns a QMessageBox object.
    """

    # create a new message box
    msg = CreateMessageBox(
        # set message box window title
        msgWinTitle=msgWinTitle,
        # set message box text
        msgText=msgText,
        # set our information text.
        msgInfoText=msgInfoText,
        # set our detailed text.
        msgDetailText=msgDetailText,
        # this is a Critical/Error message box
        msgIcon=QMessageBox.Question,
        # set message box button
        msgButtons=msgButtons,
        # set message box default confirmation button/key (enter key)
        msgDefaultButton=msgDefaultButton,
        # set message box escape confirmation button/key (esc key)
        msgEscapeButton=msgEscapeButton,
    )
    # return our new Question message box
    return msg

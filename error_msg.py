from sys import exception

from PyQt6.QtWidgets import QErrorMessage


def error_msg(error):
    # creates error dialogs
    error_dialog = QErrorMessage()
    error_dialog.showMessage(error)
    error_dialog.setWindowTitle('Notification')

    # displays it
    error_dialog.exec()

    # runs an exception
    exception()

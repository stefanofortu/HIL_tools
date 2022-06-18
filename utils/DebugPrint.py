from PySide6.QtWidgets import QMessageBox


def debugPrint(*args, level=0, sep=' ', end='\n', file=None):
    if level == 0:
        print(args)
        print(args, sep=sep, end=end, file=file)



    # tag::button_clicked[]
    def button_clicked(self):

        button = QMessageBox.critical(
            self,
            "Oh dear!",
            "Told you not to press this.",
            buttons=QMessageBox.Discard | QMessageBox.NoToAll | QMessageBox.Ignore,
            defaultButton=QMessageBox.Discard,
        )

        if button == QMessageBox.Discard:
            print("Discard!")
        elif button == QMessageBox.NoToAll:
            print("No to all!")
        else:
            print("Ignore!")
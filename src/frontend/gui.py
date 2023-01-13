"""Generate GUI."""

import os
import subprocess
import sys

from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QVBoxLayout
from PyQt5.QtWidgets import QWidget


class StartWindow(QWidget):
    """A class of GUI, include components & slots etc."""

    def __init__(self):
        """Setup GUI layout & components."""
        super().__init__()
        self.setWindowTitle('propens')
        self.current_path = os.getcwd()
        layout = QVBoxLayout()
        self.choose_button = QPushButton(self)
        self.choose_button.setText('button1')
        layout.addWidget(self.choose_button)
        self.start_button = QPushButton(self)
        self.start_button.setText('button2')
        layout.addWidget(self.start_button)
        self.setLayout(layout)
        self.file_path = '.'

        # setup connections
        self.choose_button.clicked.connect(self.choose_button_clicked)
        self.start_button.clicked.connect(self.start_button_clicked)

    # slots
    def choose_button_clicked(self):
        """Save path of selected audio file."""
        file_path_type = QFileDialog.getOpenFileName(
            self,
            'choose file',
            self.current_path,
            'Audio Files (*.wav)',
        )
        self.file_path = file_path_type[0]

    def start_button_clicked(self):
        """Start rendering."""
        if self.file_path == '':
            QMessageBox.about(self, 'No Selected File', 'Please choose a audio file first.')
        else:
            opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
            subprocess.call([opener, os.path.abspath(os.path.join(self.file_path, '..'))])
            self.file_path = ''
            sys.exit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    StartWindow = StartWindow()
    StartWindow.show()
    sys.exit(app.exec_())

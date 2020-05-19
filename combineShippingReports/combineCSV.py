import glob
import os
import sys
from datetime import datetime
from os.path import expanduser

import pandas as pd
from PyQt5.QtCore import pyqtSlot, QCoreApplication
from PyQt5.QtWidgets import QApplication, QFileDialog, QWidget, QPushButton, QMessageBox


class App(QWidget):

  def __init__(self):
    super().__init__()
    self.title = 'Combine Shipping Reports'
    self.left = 10
    self.top = 10
    self.width = 320
    self.height = 200
    self.initUI()

  def initUI(self):
    self.setWindowTitle(self.title)
    self.setGeometry(self.left, self.top, self.width, self.height)

    button = QPushButton('Choose Directory', self)
    button.setToolTip('This is an example button')
    button.move(100, 70)
    button.clicked.connect(self.on_click)

    self.show()

  @pyqtSlot()
  def on_click(self):
    path = str(QFileDialog.getExistingDirectory(None, 'Select Folder', expanduser('~')))
    all_df_list = []
    for filename in glob.glob(os.path.join(path, '*.csv')):
      # with open(filename, 'r') as f:  # open in readonly mode
      excel_data = pd.read_csv(filename)
      # excel_data['Ship Date'] = ''
      # excel_data.fillna('', inplace=True)
      all_df_list.append(excel_data)

    # Merge all the dataframes in all_df_list
    # Pandas will automatically append based on similar column names
    appended_df = pd.concat(all_df_list)

    # Write the appended dataframe to an excel file
    # Add index=False parameter to not include row numbers
    report = 'CombinedReport-' + datetime.now().strftime('%m_%d_%Y') + '.xlsx'
    appended_df.to_excel(report, index=False)
    msg = QMessageBox()
    msg.setWindowTitle('Success!')
    msg.setText('New report located at ' + os.getcwd() + os.path.sep + report)
    msg.setIcon(QMessageBox.Information)
    msg.setStandardButtons(QMessageBox.Ok)
    msg.setDefaultButton(QMessageBox.Ok)
    msg.exec_()
    msg.buttonClicked.connect(self.popUpClicked)
    QCoreApplication.quit()

  def popUpClicked(self):
    print('clicked')


if __name__ == '__main__':
  app = QApplication(sys.argv)
  ex = App()
  sys.exit(app.exec_())

from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5 import uic

RADIOBUTTON_STATE_NEW_ACCOUNT = 1
RADIOBUTTON_STATE_EXSISTING_ACCOUNT = 2

def run_main_program_debug(create_new_account, csv_transactions_file, year, account_filepath):
    success = False
    message = 'asdfasdfaa asdfasdfaa asdfasdfaa asdfasdfaa asdfasdfaa asdfasdfaa asdfasdfaa'
    return success, message


class MyGUI(QMainWindow):
    main_program_function = None
    
    def __init__ (self, default_create_new_account, default_year, default_name, default_new_transactions_file_dir, default_account_location, default_exsisting_form, main_program_function):
        super(MyGUI, self).__init__()
        try:
            uic.loadUi('user_interface/main_window.ui', self)
        except:
            uic.loadUi('main_window.ui', self)
        self.show()
        
        self.main_program_function = staticmethod(main_program_function)
        
        # Create class variables
        self.default_new_transactions_file_dir = default_new_transactions_file_dir
        self.default_account_location = default_account_location
        
        # Set on-click handlers for buttons
        self.radioButton_make_new_account.clicked.connect(self.update_radiobutton_choice)
        self.radioButton_extend_exsisting_account.clicked.connect(self.update_radiobutton_choice)
        
        self.toolButton_new_transactions_file.clicked.connect(self.select_csv_transactions)
        self.toolButton_selcet_exsisting_account.clicked.connect(self.select_exsisting_account)
        self.toolButton_save_new_account_as.clicked.connect(self.save_new_account)
        
        self.pushButton_run.clicked.connect(self.attempt_to_run)
        
        # Initialize default year
        self.lineEdit_account_year.setText(str(default_year))
        
        # Initialize default location for exsisting account
        self.lineEdit_selcet_exsisting_account.setText(str(default_exsisting_form))
        
        # Initialize default radiobutton state
        if default_create_new_account:
            self.radioButton_make_new_account.nextCheckState()
        else:
            self.radioButton_extend_exsisting_account.nextCheckState()
        self.radio_button_state = None
        self.update_radiobutton_choice()
        
        self.pushButton_cancel.clicked.connect(exit)
    
    
    # Hide and show options based on radio button state
    def update_radiobutton_choice(self):
        if self.radioButton_make_new_account.isChecked():
            self.radio_button_state = RADIOBUTTON_STATE_NEW_ACCOUNT
            self.label_save_new_account_as.show()
            self.lineEdit_save_new_account_as.show()
            self.toolButton_save_new_account_as.show()
            self.label_account_year.show()
            self.lineEdit_account_year.show()
            self.lineEdit_account_name.show()
            self.label_account_name.show()
            self.label_selcet_exsisting_account.hide()
            self.lineEdit_selcet_exsisting_account.hide()
            self.toolButton_selcet_exsisting_account.hide()
        else:
            self.radio_button_state = RADIOBUTTON_STATE_EXSISTING_ACCOUNT
            self.label_save_new_account_as.hide()
            self.lineEdit_save_new_account_as.hide()
            self.toolButton_save_new_account_as.hide()
            self.label_account_year.hide()
            self.lineEdit_account_name.hide()
            self.label_account_name.hide()
            self.lineEdit_account_year.hide()
            self.label_selcet_exsisting_account.show()
            self.lineEdit_selcet_exsisting_account.show()
            self.toolButton_selcet_exsisting_account.show()
    
    # Open file explorer to select transactions exported from Sparebanken Sør
    def select_csv_transactions(self):
        default_dir = self.default_new_transactions_file_dir
        if self.lineEdit_new_transactions_file.text() != '':
            default_dir = self.lineEdit_new_transactions_file.text()
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, caption="Velg fil", directory=default_dir, filter="CSV-fil (*.csv)", options=options)
        self.lineEdit_new_transactions_file.setText(file_name)
    
    # Open file explorer to selext exsisting accout for modification
    def select_exsisting_account(self):
        default_dir = self.default_account_location
        if self.lineEdit_selcet_exsisting_account.text() != '':
            default_dir = self.lineEdit_selcet_exsisting_account.text()
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, caption="Velg regnskap", directory=default_dir, filter="Excel file (*.xlsx)", options=options)
        self.lineEdit_selcet_exsisting_account.setText(file_name)
    
    # Select location to save newly created account
    def save_new_account(self):
        options = QFileDialog.Options()
        if self.lineEdit_save_new_account_as.text() != '':
            default_dir = self.lineEdit_save_new_account_as.text()
        file_name, _ = QFileDialog.getSaveFileName(self, caption="Lagre som", directory=self.default_account_location, filter="Excel file (*.xlsx)", options=options)
        self.lineEdit_save_new_account_as.setText(file_name)
    
    # Run main program, end exit applicaion on success
    def attempt_to_run(self):
        create_new_account = (self.radio_button_state == RADIOBUTTON_STATE_NEW_ACCOUNT)
        csv_transactions_file = self.lineEdit_new_transactions_file.text()
        year = self.lineEdit_account_year.text()
        name = self.lineEdit_account_name.text()
        account_filepath = ''
        if create_new_account:
            account_filepath = self.lineEdit_save_new_account_as.text()
        else:
            account_filepath = self.lineEdit_selcet_exsisting_account.text()
        try:
            success, message, title = self.main_program_function(create_new_account, csv_transactions_file, year, name, account_filepath)
        except Exception as e:
            success = False
            title = 'Error'
            message = e
        message_box = QMessageBox()
        message_box.setText(message)
        message_box.setWindowTitle(title)
        message_box.exec_()
        if success:
            exit()

def run_GUI(default_create_new_account, default_year, default_name, default_new_transactions_file_dir, default_account_location, default_exsisting_form, main_program_function):
    app = QApplication([])
    win = MyGUI(default_create_new_account, default_year, default_name, default_new_transactions_file_dir, default_account_location, default_exsisting_form, main_program_function)
    app.exec_()

def main():
    default_create_new_account = True
    default_year = 2023
    default_name = 'FILTER'
    default_new_transactions_file_dir = 'C:\\Users\\stian\\Downloads'
    default_account_location = 'C:\\Users\\stian\\Documents'
    default_exsisting_form = 'C:\\Users\\stian\\Documents\\default_account.xlsx'
    run_GUI(default_create_new_account, default_year, default_name, default_new_transactions_file_dir, default_account_location, default_exsisting_form, run_main_program_debug)

if __name__ == "__main__":
    main()

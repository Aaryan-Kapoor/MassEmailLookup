import sys
import time
import os
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QLabel, QFileDialog,
    QVBoxLayout, QWidget, QProgressBar, QMessageBox, QHBoxLayout, QSplashScreen, QLineEdit
)
from PySide6.QtCore import QThread, Signal, Qt, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QIcon, QFont, QPixmap
import qdarkstyle
import main  # Assuming the email lookup logic is in main.py
from cryptography.fernet import Fernet

# Get the directory of the script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Path to the icons directory
icons_dir = os.path.join(script_dir, 'icons')

# Function to generate and save a key for encryption
def generate_key(key_file_path):
    key = Fernet.generate_key()
    with open(key_file_path, "wb") as key_file:
        key_file.write(key)

# Function to load the encryption key from a file
def load_key(key_file_path):
    return open(key_file_path, "rb").read()

# Function to encrypt the API key
def encrypt_api_key(api_key, key_file_path, encrypted_file_path):
    key = load_key(key_file_path)
    fernet = Fernet(key)
    encrypted_api_key = fernet.encrypt(api_key.encode())
    with open(encrypted_file_path, "wb") as encrypted_file:
        encrypted_file.write(encrypted_api_key)

# Function to decrypt the API key
def decrypt_api_key(key_file_path, encrypted_file_path):
    key = load_key(key_file_path)
    fernet = Fernet(key)
    with open(encrypted_file_path, "rb") as encrypted_file:
        encrypted_api_key = encrypted_file.read()
    return fernet.decrypt(encrypted_api_key).decode()

# Worker thread for running the API lookup in the background
class EmailLookupThread(QThread):
    progress = Signal(int)
    finished = Signal(str)

    def __init__(self, excel_file, api_key):
        super().__init__()
        self.excel_file = excel_file
        self.api_key = api_key

    def run(self):
        # Load the Excel file
        df = pd.read_excel(self.excel_file)
        people_data = df[['FirstName', 'LastName', 'Title', 'Organization']]
        total_rows = len(people_data)

        results = []

        # Process each person, updating progress
        for index, row in people_data.iterrows():
            first_name = row['FirstName']
            last_name = row['LastName']
            title = row['Title']
            organization = row['Organization']

            # Call the email lookup function from main.py
            email_data = main.lookup_email(first_name, last_name, title, organization, self.api_key)
            print(email_data)  # REMOVE THIS | TESTING
            
            # Append the result with all fields
            results.append({
                "FirstName": first_name,
                "LastName": last_name,
                "Title": title,
                "Organization": organization,
                "Email": email_data.get('email'),
                "Phone Number": email_data.get('phone_number'),  # Add phone number
                "Is Edu Email": email_data.get('edu_email', False),  # Default to False if not available
                "Source Link": email_data.get('source_link')  # Add source link
            })

            # Update progress
            self.progress.emit(int((index + 1) / total_rows * 100))

            # Cooldown for API rate limit (Handled in main.py)
            time.sleep(3)

        # Convert results to a DataFrame and save it to an output file
        output_df = pd.DataFrame(results)
        output_file_path = f"{self.excel_file}_output.xlsx"
        output_df.to_excel(output_file_path, index=False)

        # Emit finished signal when done with the output file path
        self.finished.emit(output_file_path)

# AnimatedButton class with hover effect
class AnimatedButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.default_style = "background-color: #3A3A3A; border-radius: 5px; color: white; padding: 10px;"
        self.hover_style = "background-color: #505050; border-radius: 5px; color: white; padding: 10px;"
        self.setStyleSheet(self.default_style)

    def enterEvent(self, event):
        self.setStyleSheet(self.hover_style)

    def leaveEvent(self, event):
        self.setStyleSheet(self.default_style)

# Custom Title Bar for Frameless Window
class TitleBar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent

        # Initialize startPos as None
        self.startPos = None

        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        # Title label
        self.title = QLabel("Email Lookup Tool")
        self.title.setStyleSheet("color: white; font-size: 14px;")

        # Minimize button
        self.min_btn = QPushButton("-")
        self.min_btn.setFixedSize(30, 30)
        self.min_btn.setStyleSheet("background-color: #404040; color: white; border: none;")
        self.min_btn.clicked.connect(self.parent.showMinimized)

        # Close button
        self.close_btn = QPushButton("X")
        self.close_btn.setFixedSize(30, 30)
        self.close_btn.setStyleSheet("background-color: #E81123; color: white; border: none;")
        self.close_btn.clicked.connect(self.parent.close)

        layout.addWidget(self.title)
        layout.addStretch()
        layout.addWidget(self.min_btn)
        layout.addWidget(self.close_btn)
        self.setLayout(layout)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            # Store the position of the mouse press
            self.startPos = event.pos()

    def mouseMoveEvent(self, event):
        if self.startPos:
            # Move the window according to the current mouse position and the stored position
            self.parent.move(self.parent.pos() + event.pos() - self.startPos)

    def mouseReleaseEvent(self, event):
        # Reset startPos when the mouse is released
        self.startPos = None

# Main Window Class
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Lookup Tool")
        self.setGeometry(300, 300, 650, 250)
        

        # Keep Frameless Window
        #self.setWindowFlags(Qt.FramelessWindowHint)
        self.setWindowIcon(QIcon(os.path.join(icons_dir, 'icon.png')))

        self.key_folder = 'secure'
        if not os.path.exists(self.key_folder):
            os.makedirs(self.key_folder)

        self.key_file = os.path.join(self.key_folder, "secret.key")
        self.encrypted_file = os.path.join(self.key_folder, "api_key.enc")

        # Add border to simulate window frame
        self.setStyleSheet("""
            QMainWindow {
                border: 1px solid #5A5A5A;
                background-color: #2D2D2D;
            }
        """)

        # Main layout
        main_layout = QVBoxLayout()

        # Title bar with close, minimize buttons
        self.title_bar = TitleBar(self)
        main_layout.addWidget(self.title_bar)

        # Central widget layout
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        button_layout = QHBoxLayout()

        # Label
        self.label = QLabel("Select an Excel file to begin", self)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("font-size: 12pt;")
        layout.addWidget(self.label)

        # Button to open file dialog
        self.button = AnimatedButton("Choose Excel File", self)
        self.button.setIcon(QIcon(os.path.join(icons_dir, 'file_icon.png')))
        self.button.clicked.connect(self.open_file_dialog)
        button_layout.addWidget(self.button)

        # Button to start process
        self.start_button = AnimatedButton("Start Email Lookup", self)
        self.start_button.setIcon(QIcon(os.path.join(icons_dir, 'start_icon.png')))
        self.start_button.clicked.connect(self.start_email_lookup)
        self.start_button.setEnabled(False)  # Disabled until file is selected
        button_layout.addWidget(self.start_button)

        # Add the button layout (horizontally stacked)
        layout.addLayout(button_layout)

        # Progress bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.progress_bar)

        # API Key input field
        self.api_key_input = QLineEdit(self)
        self.api_key_input.setPlaceholderText("Enter your Perplexity API Key")
        self.api_key_input.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.api_key_input)

        # Buttons for saving and loading API Key
        api_button_layout = QHBoxLayout()
        self.save_api_key_button = AnimatedButton("Save API Key", self)
        self.save_api_key_button.clicked.connect(self.save_api_key)
        api_button_layout.addWidget(self.save_api_key_button)

        self.load_api_key_button = AnimatedButton("Load API Key", self)
        self.load_api_key_button.clicked.connect(self.load_api_key)
        api_button_layout.addWidget(self.load_api_key_button)

        layout.addLayout(api_button_layout)

        # Add layout to main layout
        main_layout.addLayout(layout)

        # Set layout
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # Variable to store the selected Excel file
        self.excel_file = ""

    def save_api_key(self):
        api_key = self.api_key_input.text()
        if api_key:
            if not os.path.exists(self.key_file):
                generate_key(self.key_file)  # Generate encryption key if not already exists
            encrypt_api_key(api_key, self.key_file, self.encrypted_file)
            QMessageBox.information(self, "Success", "API Key has been saved securely!")
            # Clear the input field for security
            self.api_key_input.setText("")
        else:
            QMessageBox.warning(self, "Error", "Please enter a valid API Key.")

    def load_api_key(self):
        if os.path.exists(self.encrypted_file) and os.path.exists(self.key_file):
            try:
                api_key = decrypt_api_key(self.key_file, self.encrypted_file)
                self.api_key_input.setText(api_key)
                QMessageBox.information(self, "Success", "API Key has been loaded successfully!")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to load API Key: {str(e)}")
        else:
            QMessageBox.warning(self, "Error", "No saved API Key found.")

    def open_file_dialog(self):
        # Open file dialog to select Excel file
        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)")
        if file_path:
            self.excel_file = file_path
            self.label.setText(f"Selected File: {os.path.basename(file_path)}")
            self.start_button.setEnabled(True)  # Enable the start button after file is selected

    def start_email_lookup(self):
        if not self.excel_file:
            QMessageBox.warning(self, "No file selected", "Please select an Excel file before starting.")
            return

        api_key = self.api_key_input.text()
        if not api_key:
            QMessageBox.warning(self, "No API Key", "Please enter your API Key before starting.")
            return

        # Provide immediate feedback that processing has started
        self.label.setText("Processing started...")

        # Disable buttons while processing
        self.start_button.setEnabled(False)
        self.button.setEnabled(False)
        self.save_api_key_button.setEnabled(False)
        self.load_api_key_button.setEnabled(False)

        # Start the background thread
        self.thread = EmailLookupThread(self.excel_file, api_key)
        self.thread.progress.connect(self.update_progress)
        self.thread.finished.connect(self.email_lookup_finished)
        self.thread.start()

    def update_progress(self, value):
        # Animate the progress bar
        self.animation = QPropertyAnimation(self.progress_bar, b"value")
        self.animation.setDuration(500)
        self.animation.setStartValue(self.progress_bar.value())
        self.animation.setEndValue(value)
        self.animation.setEasingCurve(QEasingCurve.InOutQuad)
        self.animation.start()

        # Update the label to show the progress status
        self.label.setText(f"Processing... {value}% completed")

    def email_lookup_finished(self, output_file_path):
        # Enable the buttons after processing
        self.start_button.setEnabled(True)
        self.button.setEnabled(True)
        self.save_api_key_button.setEnabled(True)
        self.load_api_key_button.setEnabled(True)

        # Clear the progress bar and reset label
        self.progress_bar.setValue(0)
        self.label.setText("Select an Excel file to begin")

        # Show a message box indicating the process is complete
        QMessageBox.information(self, "Finished", f"Email lookup process completed successfully!\nOutput saved to: {output_file_path}")

# Main function to run the application
if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Set the dark theme
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyside6())

    # Before showing the main window
    splash_pix = QPixmap(os.path.join(icons_dir, 'splash.png'))
    splash = QSplashScreen(splash_pix)
    splash.show()
    app.processEvents()

    # Create and show the main window
    window = MainWindow()
    time.sleep(0.5)  # Simulate startup delay
    window.show()
    splash.finish(window)

    sys.exit(app.exec())
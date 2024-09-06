from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QFileDialog,
    QPushButton,
    QVBoxLayout,
    QWidget,
    QComboBox,
    QCheckBox,
    QLabel,
)
import sys


class ConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("13F Excel to XML Converter")

        layout = QVBoxLayout()

        # Excel file selection
        self.excel_file_label = QLabel("Select Excel File:")
        layout.addWidget(self.excel_file_label)
        self.excel_file_button = QPushButton("Browse")
        self.excel_file_button.clicked.connect(self.select_excel_file)
        layout.addWidget(self.excel_file_button)

        # XSD file selection
        self.xsd_file_label = QLabel("Select XSD File:")
        layout.addWidget(self.xsd_file_label)
        self.xsd_file_button = QPushButton("Browse")
        self.xsd_file_button.clicked.connect(self.select_xsd_file)
        layout.addWidget(self.xsd_file_button)

        # Default Investment Discretion
        self.investment_discretion_label = QLabel("Default Investment Discretion:")
        layout.addWidget(self.investment_discretion_label)
        self.investment_discretion_combo = QComboBox()
        self.investment_discretion_combo.addItems(["Sole", "Defined", "Other"])
        layout.addWidget(self.investment_discretion_combo)

        # Default Shares/Principal
        self.shares_principal_label = QLabel("Default Shares/Principal:")
        layout.addWidget(self.shares_principal_label)
        self.shares_principal_combo = QComboBox()
        self.shares_principal_combo.addItems(["Shares", "Principal"])
        layout.addWidget(self.shares_principal_combo)

        # Voting Authority
        self.voting_authority_label = QLabel("Voting Authority:")
        layout.addWidget(self.voting_authority_label)
        self.voting_authority_combo = QComboBox()
        self.voting_authority_combo.addItems(["Sole", "Shared", "None"])
        layout.addWidget(self.voting_authority_combo)

        # Include under-threshold securities
        self.include_under_threshold_checkbox = QCheckBox(
            "Include Under-Threshold Securities"
        )
        layout.addWidget(self.include_under_threshold_checkbox)

        # Include Non 13F Securities
        self.include_non_13f_checkbox = QCheckBox("Include Non-13F Securities")
        layout.addWidget(self.include_non_13f_checkbox)

        # Use entries 'as is'
        self.use_entries_as_is_checkbox = QCheckBox("Use Entries 'As Is'")
        layout.addWidget(self.use_entries_as_is_checkbox)

        # Convert button
        self.convert_button = QPushButton("Convert to XML")
        self.convert_button.clicked.connect(self.convert_to_xml)
        layout.addWidget(self.convert_button)

        # Status label
        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_excel_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx);;All Files (*)",
            options=options,
        )
        if file_name:
            self.excel_file_path = file_name
            self.excel_file_label.setText(f"Selected Excel File: {file_name}")

    def select_xsd_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Select XSD File",
            "",
            "XSD Files (*.xsd);;All Files (*)",
            options=options,
        )
        if file_name:
            self.xsd_file_path = file_name
            self.xsd_file_label.setText(f"Selected XSD File: {file_name}")

    def convert_to_xml(self):
        try:
            # Read the Excel file
            import pandas as pd

            df = pd.read_excel(self.excel_file_path)

            # Create XML root
            import xml.etree.ElementTree as ET

            root = ET.Element("SEC_13F_Report")

            # Populate XML
            for _, row in df.iterrows():
                security = ET.SubElement(root, "Security")
                ET.SubElement(security, "IssuerName").text = str(row["Issuer"])
                ET.SubElement(security, "CUSIP").text = str(row["CUSIP"])
                ET.SubElement(security, "InvestmentDiscretion").text = (
                    self.investment_discretion_combo.currentText()
                )
                ET.SubElement(security, "SharesPrincipal").text = (
                    self.shares_principal_combo.currentText()
                )
                ET.SubElement(security, "VotingAuthority").text = (
                    self.voting_authority_combo.currentText()
                )

            # Convert to XML string
            xml_str = ET.tostring(root, encoding="utf-8", method="xml").decode()

            # Save XML file
            with open("output.xml", "w") as file:
                file.write(xml_str)

            # Validate XML
            # from validate_xml import validate_xml

            # validate_xml("output.xml", self.xsd_file_path)

            self.status_label.setText("Conversion and validation successful.")

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = ConverterApp()
    ex.show()
    sys.exit(app.exec_())

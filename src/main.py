import pandas as pd
from lxml import etree
import xml.etree.ElementTree as ET
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
    QLineEdit,
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
            df = pd.read_excel(self.excel_file_path)

            # Create the root element
            root = ET.Element(
                "ns1:informationTable",
                xmlns_ns1="http://www.sec.gov/edgar/document/thirteenf/informationtable",
            )

            # Iterate over rows in the DataFrame
            for _, row in df.iterrows():
                info_table = ET.SubElement(root, "ns1:infoTable")

                # Add elements to infoTable
                ET.SubElement(info_table, "ns1:nameOfIssuer").text = row["NameOfIssuer"]
                ET.SubElement(info_table, "ns1:titleOfClass").text = row["TitleOfClass"]
                ET.SubElement(info_table, "ns1:cusip").text = row["CUSIP"]
                if pd.notna(row["FIGI"]):
                    ET.SubElement(info_table, "ns1:figi").text = row["FIGI"]
                ET.SubElement(info_table, "ns1:value").text = str(row["Value"])

                shrs_or_prn_amt = ET.SubElement(info_table, "ns1:shrsOrPrnAmt")
                ET.SubElement(shrs_or_prn_amt, "ns1:sshPrnamt").text = str(
                    row["Shares"]
                )
                ET.SubElement(shrs_or_prn_amt, "ns1:sshPrnamtType").text = row[
                    "SharesOrPrincipal"
                ]

                if pd.notna(row["PutOrCall"]):
                    ET.SubElement(info_table, "ns1:putCall").text = row["PutOrCall"]

                ET.SubElement(info_table, "ns1:investmentDiscretion").text = row[
                    "InvestmentDiscretion"
                ]

                if pd.notna(row["OtherManagers"]):
                    ET.SubElement(info_table, "ns1:otherManager").text = row[
                        "OtherManagers"
                    ]

                voting_auth = ET.SubElement(info_table, "ns1:votingAuthority")
                ET.SubElement(voting_auth, "ns1:Sole").text = str(row["Sole"])
                ET.SubElement(voting_auth, "ns1:Shared").text = str(row["Shared"])
                ET.SubElement(voting_auth, "ns1:None").text = str(row["None"])

            # Convert to XML string
            xml_str = ET.tostring(
                root, encoding="utf-8", xml_declaration=True, method="xml"
            ).decode()
            # Save XML file
            with open("output.xml", "w") as file:
                file.write(xml_str)

            # Validate XML
            from validate_xml import validate_xml

            validate_xml("output.xml", self.xsd_file_path)

            self.status_label.setText("Conversion and validation successful.")

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = ConverterApp()
    ex.show()
    sys.exit(app.exec_())

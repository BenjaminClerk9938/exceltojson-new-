import pandas as pd
from lxml import etree
import xml.dom.minidom
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

            # Define namespaces
            ns = {"ns1": "http://www.sec.gov/edgar/document/thirteenf/informationtable"}

            # Create the root element
            root = etree.Element(
                "{http://www.sec.gov/edgar/document/thirteenf/informationtable}informationTable",
                nsmap=ns,
            )

            # Iterate over rows in the DataFrame
            for _, row in df.iterrows():
                info_table = etree.SubElement(
                    root,
                    "{http://www.sec.gov/edgar/document/thirteenf/informationtable}infoTable",
                )

                # Add elements to infoTable
                etree.SubElement(
                    info_table,
                    "{http://www.sec.gov/edgar/document/thirteenf/informationtable}nameOfIssuer",
                ).text = str(row.get("NameOfIssuer", ""))
                etree.SubElement(
                    info_table,
                    "{http://www.sec.gov/edgar/document/thirteenf/informationtable}titleOfClass",
                ).text = str(row.get("TitleOfClass", ""))
                etree.SubElement(
                    info_table,
                    "{http://www.sec.gov/edgar/document/thirteenf/informationtable}cusip",
                ).text = str(row.get("CUSIP", ""))
                if pd.notna(row.get("FIGI")):
                    etree.SubElement(
                        info_table,
                        "{http://www.sec.gov/edgar/document/thirteenf/informationtable}figi",
                    ).text = str(row.get("FIGI", ""))
                etree.SubElement(
                    info_table,
                    "{http://www.sec.gov/edgar/document/thirteenf/informationtable}value",
                ).text = str(row.get("Value", ""))

                shrs_or_prn_amt = etree.SubElement(
                    info_table,
                    "{http://www.sec.gov/edgar/document/thirteenf/informationtable}shrsOrPrnAmt",
                )
                etree.SubElement(
                    shrs_or_prn_amt,
                    "{http://www.sec.gov/edgar/document/thirteenf/informationtable}sshPrnamt",
                ).text = str(row.get("Shares", ""))
                etree.SubElement(
                    shrs_or_prn_amt,
                    "{http://www.sec.gov/edgar/document/thirteenf/informationtable}sshPrnamtType",
                ).text = str(row.get("SharesOrPrincipal", ""))

                if pd.notna(row.get("PutOrCall")):
                    etree.SubElement(
                        info_table,
                        "{http://www.sec.gov/edgar/document/thirteenf/informationtable}putCall",
                    ).text = str(row.get("PutOrCall", ""))

                etree.SubElement(
                    info_table,
                    "{http://www.sec.gov/edgar/document/thirteenf/informationtable}investmentDiscretion",
                ).text = str(row.get("InvestmentDiscretion", ""))

                if pd.notna(row.get("OtherManagers")):
                    etree.SubElement(
                        info_table,
                        "{http://www.sec.gov/edgar/document/thirteenf/informationtable}otherManager",
                    ).text = str(row.get("OtherManagers", ""))

                voting_auth = etree.SubElement(
                    info_table,
                    "{http://www.sec.gov/edgar/document/thirteenf/informationtable}votingAuthority",
                )
                etree.SubElement(
                    voting_auth,
                    "{http://www.sec.gov/edgar/document/thirteenf/informationtable}Sole",
                ).text = str(row.get("Sole", ""))
                etree.SubElement(
                    voting_auth,
                    "{http://www.sec.gov/edgar/document/thirteenf/informationtable}Shared",
                ).text = str(row.get("Shared", ""))
                etree.SubElement(
                    voting_auth,
                    "{http://www.sec.gov/edgar/document/thirteenf/informationtable}None",
                ).text = str(row.get("None", ""))

            # Convert to XML string
            xml_str = etree.tostring(
                root, encoding="utf-8", xml_declaration=True, pretty_print=True
            ).decode()

            # Save XML file
            with open("output.xml", "w") as file:
                file.write(xml_str)

            # Validate XML
            from validate_xml import validate_xml

            if validate_xml("output.xml", self.xsd_file_path):
                self.status_label.setText("Conversion and validation successful.")
            else:
                self.status_label.setText("Validation failed.")

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = ConverterApp()
    ex.show()
    sys.exit(app.exec_())

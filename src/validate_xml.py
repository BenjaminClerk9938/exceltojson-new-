from lxml import etree


def validate_xml(xml_file, xsd_file):
    try:
        # Load the XSD schema
        xsd_doc = etree.parse(xsd_file)
        xsd_schema = etree.XMLSchema(xsd_doc)

        # Parse the XML file
        xml_doc = etree.parse(xml_file)

        # Validate the XML against the XSD
        is_valid = xsd_schema.validate(xml_doc)
        if not is_valid:
            raise ValueError(xsd_schema.error_log)
    except Exception as e:
        print(e)
        raise ValueError(f"Validation error: {e}")


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 3:
        print("Usage: python validate_xml.py <xml_file> <xsd_file>")
    else:
        validate_xml(sys.argv[1], sys.argv[2])

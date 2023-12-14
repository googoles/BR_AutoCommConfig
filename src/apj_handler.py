# apj_handler.py
import xml.etree.ElementTree as ET

def process_apj_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as apj_file:
            apj_content = apj_file.read()

            # Parse XML content
            root = ET.fromstring(apj_content)

            # Find the AutomationStudio element
            automation_studio_elem = root.find('.//{http://br-automation.co.at/AS/Project}AutomationStudio')

            if automation_studio_elem is not None:
                version = automation_studio_elem.get('Version', '')
                working_version = automation_studio_elem.get('WorkingVersion', '')

                return version, working_version
            else:
                return None, None
    except Exception as e:
        raise ValueError(f"Error processing APJ file: {str(e)}")

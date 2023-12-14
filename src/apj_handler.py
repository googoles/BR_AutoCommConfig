# apj_handler.py
import xml.etree.ElementTree as ET
import re

def process_apj_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as apj_file:
            apj_content = apj_file.read()

            # Extract version information using regular expressions
            match_version = re.search(r'(?<=Version=")(.*?)(?=")', apj_content)
            version = match_version.group(0) if match_version else None

            match_working_version = re.search(r'(?<=WorkingVersion=")(.*?)(?=")', apj_content)
            working_version = match_working_version.group(0) if match_working_version else None

            return version, working_version
    except Exception as e:
        raise ValueError(f"Error processing APJ file: {str(e)}")

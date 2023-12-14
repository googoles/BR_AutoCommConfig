import re

def extract_version(xml_string):
    try:
        match = re.search(r'(?<=Version=")(.*?)(?=")', xml_string)
        version = match.group(0) if match else None

        match = re.search(r'(?<=WorkingVersion=")(.*?)(?=")', xml_string)
        working_version = match.group(0) if match else None

        return version, working_version
    except Exception as e:
        print(f"Error extracting version: {str(e)}")
        return None, None

# Example XML string
xml_string = '<?AutomationStudio Version="4.12.4.107 SP" WorkingVersion="4.12"?>'

# Extract version information
version, working_version = extract_version(xml_string)

# Print the results
print(f"Version: {version}")
print(f"Working Version: {working_version}")

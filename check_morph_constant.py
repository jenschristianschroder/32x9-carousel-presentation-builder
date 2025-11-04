"""Check what constant value corresponds to Morph transition."""

import win32com.client

# Print all transition effect constants
constants = win32com.client.constants
print("Searching for Morph transition constant...")

for attr in dir(constants):
    if 'morph' in attr.lower() or 'ppEffect' in attr:
        try:
            value = getattr(constants, attr)
            print(f"{attr} = {value}")
        except:
            pass

# Verify Branding Automation in Azure

## Requirements:
  1. [Python](https://www.python.org/downloads/)
  
  2. Python Modules
      ```bash
      pip install pyodbc
      ```
## System Compatibility Requirements
  1. Registry Permissions:
  
      - Navigate to these path in Registry
        ```bash
        Computer\HKEY_LOCAL_MACHINE\SOFTWARE\ODBC
        Computer\HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI
        Computer\HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI
        ```
      Perform below mentioned steps for all the Registry folders given above.
      - Go to Permissions by right clicking on them
      - Allow `Full Control` for the current user & ALL APPLICATION PACKAGES
        ![img.png](img.png)
## Input:
  1. `DSN Name                    `- `Datasource Name (i.e. "IBM Facebook")`
  2. `MetaData Tester Path        `- `Path to Metadata tester's exe`
  3. `System Bit                  `- `32 / 64 (As per system configuration)`
  
# Usage
- To Vertify Branding
  
      python AzDoVerifyBranding.py <DSN Name> <Path to Metadata tester> <System Bits>

- Give valid path which is accepted for python as mentioned below

      python AzDoVerifyBranding.py "IBM Facebook" "C:\\Users\\fparmar\\Desktop\\Facebook\\MetaTester\\bin\\x64\\Debug" 64

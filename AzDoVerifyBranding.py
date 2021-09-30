import subprocess as sp
import pyodbc 
import os
import platform
import winreg
import json

def SetRegForErrMes(inTargetKey, inDriverBit, inDriverRegistryConfig):
    #Start Setting UseEncryptedEndpoints = 0
    index = 0
    systemBit = int(platform.architecture()[0][:2])
    try:
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
            with winreg.OpenKey(hkey, 'Software', 0, winreg.KEY_READ) as parentKey:
                if systemBit != inDriverBit:
                    parentKey = winreg.OpenKey(parentKey, 'Wow6432Node', 0, winreg.KEY_READ)
                with winreg.OpenKey(parentKey, 'ODBC', 0, winreg.KEY_READ) as odbcKey:
                    with winreg.OpenKey(odbcKey, 'ODBC.INI', 0, winreg.KEY_WRITE) as odbcIniKey:
                        with winreg.CreateKeyEx(odbcIniKey, inTargetKey, 0, winreg.KEY_ALL_ACCESS) as driverKey:                        
                            for key, value in inDriverRegistryConfig.items():
                                winreg.SetValueEx(driverKey, key, 0, winreg.REG_SZ, value)                    
    except Exception as e:
        print(f"Error: {e}")
    #End Setting UseEncryptedEndpoints = 0

def ReadErrMsgs(ErrMsgPath):
    #Start Reading Error to find Component name and Vendor name    
    file = open(ErrMsgPath, "r")
    f = file.read()
    ReadValue = list()
    cnt = 0
    temp_str = ""
    flg = 0
    for i in range(0,len(f)):
        if f[i] == "[":
            flg = 1
        elif(f[i] == "]"):
            flg = 0
            ReadValue.append(temp_str)
            temp_str = ""
        elif flg == 1:
            temp_str = temp_str + f[i]

    print("Please verify component name and vendor name from below.")
    print(ReadValue)
    #for i in ReadValue:
    #    print(i,start=",")
    #file.close()
    
if __name__ == '__main__':
    
    '''
    Lib_path = Absolute path of the driver's LIB folder
    Dll_file = Name of DLL file of the driver
    OdbcLibDriverPath = Path of the dll file getting from registery
    '''
    
    pyodbc.pooling = False
    pyodbc.autocommit = True

    
    jsonFile = open('UserInput.json',"r")
    data = json.load(jsonFile)
    MetaData_path = data["Plugin"]["Compile"][0]["MetadataTesterPath"]
    DSN_Name = data["Plugin"]["Compile"][0]["DataSourceConfiguration"]["Dsn"]
    inDriverBit = data["Plugin"]["Compile"][0]["inDriverBit"]
    
    inDSN = "DSN="+DSN_Name
    connection = pyodbc.connect(inDSN,autocommit=True)
    print('Driver name         =',connection.getinfo(6))
    print('Data Source Name    =',connection.getinfo(2))
    print('Driver version      =',connection.getinfo(7))
    
    Dll_file = connection.getinfo(6)
    

    #Start finding DLL path
    inTargetKey = DSN_Name
    index = 0
    valueName = "Driver" 
    inDriverBit = 64
    inDriverRegistryConfig = {"UseEncryptedEndpoints":"1"}
    OdbcLibDriverPath = ""
    systemBit = int(platform.architecture()[0][:2])
    try:
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as hkey:
            with winreg.OpenKey(hkey, 'Software', 0, winreg.KEY_READ) as parentKey:
                if systemBit != inDriverBit:
                    parentKey = winreg.OpenKey(parentKey, 'Wow6432Node', 0, winreg.KEY_READ)
                with winreg.OpenKey(parentKey, 'ODBC', 0, winreg.KEY_READ) as odbcKey:
                    with winreg.OpenKey(odbcKey, 'ODBC.INI', 0, winreg.KEY_WRITE) as odbcIniKey:
                        with winreg.CreateKeyEx(odbcIniKey, inTargetKey, 0, winreg.KEY_ALL_ACCESS) as driverKey:                        
                            for key, value in inDriverRegistryConfig.items():
                                winreg.SetValueEx(driverKey, key, 0, winreg.REG_SZ, value)                    
                            DriverPath = winreg.QueryValueEx(driverKey, valueName)[0]                            
                        
                    with winreg.OpenKey(odbcKey, 'ODBCINST.INI', 0, winreg.KEY_WRITE) as odbcIniKey:    
                        with winreg.CreateKeyEx(odbcIniKey, DriverPath, 0, winreg.KEY_ALL_ACCESS) as inOdbcInstKey:                                            
                            OdbcLibDriverPath = winreg.QueryValueEx(inOdbcInstKey, valueName)[0]
    except Exception as e:
        print(f"Error: {e}")
    #END finding DLL path
    Lib_path = OdbcLibDriverPath[0:OdbcLibDriverPath.rfind("\\",0) + 1]
    
    os.chdir(Lib_path)
    Dll_Version = sp.getoutput("powershell \"(Get-Item -path .\{}).VersionInfo.ProductVersion".format(Dll_file))
    print("Dll version         =",Dll_Version)
    
    #ErrorMessageFetching
    #Set UseEncryptedEndpoints to 0 to generate Error Message    
    inDriverRegistryConfig = {"UseEncryptedEndpoints":"0"}
    SetRegForErrMes(DSN_Name,inDriverBit,inDriverRegistryConfig)
    os.chdir(MetaData_path)
    #os.system("MetaTester64.exe -d \"{}\" -o ErrorMessage.txt".format(DSN_Name))
    ErrMsgStr = sp.getoutput("MetaTester64.exe -d \"{}\" -o ErrorMessage.txt".format(DSN_Name))
    #print(ErrMsgStr)
    
    #Set UseEncryptedEndpoints to 1 to it's default
    inDriverRegistryConfig = {"UseEncryptedEndpoints":"1"}
    SetRegForErrMes(DSN_Name,inDriverBit,inDriverRegistryConfig)
    
    
    
    #Start Reading ErrorMessage to find Component name and Vendor name    
    ErrMsgPath = os.path.join(MetaData_path,"ErrorMessage.txt")
    ReadErrMsgs(ErrMsgPath)
    #print(ErrMsgStr)
    

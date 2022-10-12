import win32com.client
def WMIDateStringToDate(dtmDate):
    strDateTime = ""
    if (dtmDate[4] == 0):
        strDateTime = dtmDate[5] + '/'
    else:
        strDateTime = dtmDate[4] + dtmDate[5] + '/'
    if (dtmDate[6] == 0):
        strDateTime = strDateTime + dtmDate[7] + '/'
    else:
        strDateTime = strDateTime + dtmDate[6] + dtmDate[7] + '/'
        strDateTime = strDateTime + dtmDate[0] + dtmDate[1] + dtmDate[2] + dtmDate[3] + " " + dtmDate[8] + dtmDate[9] + ":" + dtmDate[10] + dtmDate[11] +':' + dtmDate[12] + dtmDate[13]
    return strDateTime

strComputer = "."
objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_NetworkAdapter")
for objItem in colItems:
    if objItem.AdapterType != None:
        print("AdapterType:" + repr(objItem.AdapterType))
    if objItem.AdapterTypeId != None:
        print("AdapterTypeId:" + repr(objItem.AdapterTypeId))
    if objItem.AutoSense != None:
        print ("AutoSense:" + repr(objItem.AutoSense))
    if objItem.Availability != None:
        print ("Availability:" + repr(objItem.Availability))
    if objItem.Caption != None:
        print ("Caption:" + repr(objItem.Caption))
    if objItem.ConfigManagerErrorCode != None:
        print ("ConfigManagerErrorCode:" + repr(objItem.ConfigManagerErrorCode))
    if objItem.ConfigManagerUserConfig != None:
        print ("ConfigManagerUserConfig:" + repr(objItem.ConfigManagerUserConfig))
    if objItem.CreationClassName != None:
        print ("CreationClassName:" + repr(objItem.CreationClassName))
    if objItem.Description != None:
        print ("Description:" + repr(objItem.Description))
    if objItem.DeviceID != None:
        print ("DeviceID:" + repr(objItem.DeviceID))
    if objItem.ErrorCleared != None:
        print ("ErrorCleared:" + repr(objItem.ErrorCleared))
    if objItem.ErrorDescription != None:
        print ("ErrorDescription:" + repr(objItem.ErrorDescription))
    if objItem.Index != None:
        print ("Index:" + repr(objItem.Index))
    if objItem.InstallDate != None:
        print ("InstallDate:" + WMIDateStringToDate(objItem.InstallDate))
    if objItem.Installed != None:
        print ("Installed:" + repr(objItem.Installed))
    if objItem.InterfaceIndex != None:
        print ("InterfaceIndex:" + repr(objItem.InterfaceIndex))
    if objItem.LastErrorCode != None:
        print ("LastErrorCode:" + repr(objItem.LastErrorCode))
    if objItem.MACAddress != None:
        print ("MACAddress:" + repr(objItem.MACAddress))
    if objItem.Manufacturer != None:
        print ("Manufacturer:" + repr(objItem.Manufacturer))
    if objItem.MaxNumberControlled != None:
        print ("MaxNumberControlled:" + repr(objItem.MaxNumberControlled))
    if objItem.MaxSpeed != None:
        print ("MaxSpeed:" + repr(objItem.MaxSpeed))
    if objItem.Name != None:
        print ("Name:" + repr(objItem.Name))
    if objItem.NetConnectionID != None:
        print ("NetConnectionID:" + repr(objItem.NetConnectionID))
    if objItem.NetConnectionStatus != None:
        print ("NetConnectionStatus:" + repr(objItem.NetConnectionStatus))
    print ("NetworkAddresses:")
    strList = " "
    try :
        for objElem in objItem.NetworkAddresses :
            strList = strList + repr(objElem) + ","
    except:
        strList = strList + 'null'
    print(strList)
    if objItem.PermanentAddress != None:
        print
    if objItem.PNPDeviceID != None:
        print ("PNPDeviceID:" + repr(objItem.PNPDeviceID))
    print ("PowerManagementCapabilities:")
    strList = " "
    try :
        for objElem in objItem.PowerManagementCapabilities :
            strList = strList + repr(objElem) + ","
    except:
        strList = strList + 'null'
    print (strList)
    if objItem.PowerManagementSupported != None:
        print ("PowerManagementSupported:" + repr(objItem.PowerManagementSupported))
    if objItem.ProductName != None:
        print ("ProductName:" + repr(objItem.ProductName))
    if objItem.ServiceName != None:
        print ("ServiceName:" + repr(objItem.ServiceName))
    if objItem.Speed != None:
        print ("Speed:" + repr(objItem.Speed))
    if objItem.Status != None:
        print ("Status:" + repr(objItem.Status))
    if objItem.StatusInfo != None:
        print ("StatusInfo:" + repr(objItem.StatusInfo))
    if objItem.SystemCreationClassName != None:
        print ("SystemCreationClassName:" + repr(objItem.SystemCreationClassName))
    if objItem.SystemName != None:
        print ("SystemName:" + repr(objItem.SystemName))
    if objItem.TimeOfLastReset != None:
        print ("TimeOfLastReset:" + WMIDateStringToDate(objItem.TimeOfLastReset))

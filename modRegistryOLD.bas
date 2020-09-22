Attribute VB_Name = "modRegistryOLD"
'Use the following syntax for the SaveSetting statement:
'SaveSetting appname, section, Key, Value
'Place some settings in the registry.
    
Public Sub regSaveValue(ApplicationName As String, SaveSection As String, SaveKey As String, SaveValue As String)
    SaveSetting ApplicationName, SaveSection, SaveKey, SaveValue
End Sub
    
'GetSetting(appname, section, key, default)
'The following code retrieves the value of the specified key in the
'Application's Startup section

Public Function regGetValue(ApplicationName As String, GetSection As String, GetKey As String, Optional DefaultValue As Variant) As String
   regGetValue = GetSetting(ApplicationName, GetSection, GetKey, DefaultValue)
End Function

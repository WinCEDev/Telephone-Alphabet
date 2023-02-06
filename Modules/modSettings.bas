Attribute VB_Name = "Settings"
Option Explicit

Public Declare Function Settings_RegOpenKeyEx _
               Lib "Coredll" _
               Alias "RegOpenKeyExW" (ByVal hKey As Long, _
                                      ByVal lpSubKey As String, _
                                      ByVal ulOptions As Long, _
                                      ByVal samDesired As Long, _
                                      ByRef phkResult As Long) As Long

Public Declare Function Settings_RegQueryValueEx _
               Lib "Coredll" _
               Alias "RegQueryValueExW" (ByVal hKey As Long, _
                                         ByVal lpValueName As String, _
                                         ByVal lpReserved As Long, _
                                         lpType As Long, _
                                         ByVal lpData As String, _
                                         lpcbData As Long) As Long
                                         
Public Declare Function Settings_RegCreateKeyEx _
               Lib "Coredll" _
               Alias "RegCreateKeyExW" (ByVal hKey As Long, _
                                        ByVal lpSubKey As String, _
                                        ByVal Reserved As Long, _
                                        ByVal lpClass As String, _
                                        ByVal dwOptions As Long, _
                                        ByVal samDesired As Long, _
                                        ByRef lpSecurityAttributes As Long, _
                                        ByRef phkResult As Long, _
                                        ByRef lpdwDisposition As Long) As Long

Public Declare Function Settings_RegSetValueEx _
               Lib "Coredll" _
               Alias "RegSetValueExW" (ByVal hKey As Long, _
                                       ByVal lpValueName As String, _
                                       ByVal Reserved As Long, _
                                       ByVal dwType As Long, _
                                       ByVal lpData As String, _
                                       ByVal cbData As Long) As Long

Public Declare Function Settings_RegCloseKey _
               Lib "Coredll" _
               Alias "RegCloseKey" (ByVal hKey As Long) As Long

Private Const HKEY_CURRENT_USER As Long = &H80000001

Private Const ERROR_SUCCESS     As Long = 0

Private Const REG_SZ            As Long = 1

Public Function GetSetting(ByVal AppName As String, _
                           ByVal Section As String, _
                           ByVal Key As String, _
                           ByVal Default As String) As String

    Dim lngKey         As Long

    Dim lngReturnValue As Long

    Dim strValue       As String

    Dim lngValueLength As Long

    Dim lngType        As Long

    strValue = String(128, vbNullChar)

    lngValueLength = Len(strValue) * 2
    
    Dim strPath As String

    strPath = "Software\" & App.CompanyName & "\" & AppName & "\" & Section

    lngReturnValue = Settings_RegOpenKeyEx(HKEY_CURRENT_USER, strPath, 0, 0, lngKey)

    If lngReturnValue = ERROR_SUCCESS Then 'Success
        lngReturnValue = Settings_RegQueryValueEx(lngKey, Key, 0, lngType, strValue, lngValueLength)

        If lngReturnValue = ERROR_SUCCESS Then
            Settings_RegCloseKey lngKey
            GetSetting = LeftB(strValue, lngValueLength)
        Else
            GetSetting = Default
        End If

    Else
        GetSetting = Default
    End If

End Function

Public Sub SaveSetting(ByVal AppName As String, _
                       ByVal Section As String, _
                       ByVal Key As String, _
                       ByVal Setting As String)
                       
    Dim lngKey         As Long

    Dim lngReturnValue As Long
    
    lngReturnValue = Settings_RegCreateKeyEx(HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & AppName & "\" & Section, 0, 0, 0, 0, 0, lngKey, 0)
    
    If lngReturnValue <> ERROR_SUCCESS Then
        Err.Clear
        Err.Raise vbObjectError + 8799, "SaveSetting", "Error writing value."
    End If

    lngReturnValue = Settings_RegSetValueEx(lngKey, Key, 0, REG_SZ, Setting, LenB(Setting))

    If lngReturnValue <> ERROR_SUCCESS Then
        Err.Clear
        Err.Raise vbObjectError + 8799, "SaveSetting", "Error writing value."
    End If

    Settings_RegCloseKey lngKey

End Sub


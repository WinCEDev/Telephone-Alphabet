Attribute VB_Name = "SystemSound"
Option Explicit

'Private Declares are not supported for Windows CE.

Public Declare Function SystemSound_PlaySound _
               Lib "Coredll" _
               Alias "PlaySoundW" (ByVal lpszName As String, _
                                   ByVal hModule As Long, _
                                   ByVal dwFlags As Long) As Long

Private Const SND_ASYNC = &H1         'Play asynchronously

Private Const SND_ALIAS = &H10000     'Name is a WIN.INI [sounds] entry

Public Const ceSystemSoundAsterisk       As String = "SystemAsterisk"

Public Const ceSystemSoundDefault        As String = "SystemDefault"

Public Const ceSystemSoundExclamation    As String = "SystemExclamation"

Public Const ceSystemSoundSystemExit     As String = "SystemExit"

Public Const ceSystemSoundSystemHand     As String = "SystemHand"

Public Const ceSystemSoundSystemQuestion As String = "SystemQuestion"

Public Const ceSystemSoundSystemStart    As String = "SystemStart"

Public Const ceSystemSoundSystemWelcome  As String = "SystemWelcome"

Public Function SystemSound_Play(ByVal Id As String) As Long
    SystemSound_Play = SystemSound_PlaySound(Id, 0, SND_ALIAS)
End Function


Attribute VB_Name = "TextBoxHelpers"
Option Explicit

Public Declare Function TextBoxHelpers_SendMessage _
               Lib "Coredll" _
               Alias "SendMessageW" (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Long) As Long

Public Declare Function TextBoxHelpers_GetWindowLong _
               Lib "Coredll" _
               Alias "GetWindowLongW" (ByVal hwnd As Long, _
                                       ByVal nIndex As Long) As Long

Public Declare Function TextBoxHelpers_SetWindowLong _
               Lib "Coredll" _
               Alias "SetWindowLongW" (ByVal hwnd As Long, _
                                       ByVal nIndex As Long, _
                                       ByVal dwNewLong As Long) As Long
                                       
Public Const WM_UNDO = &H304

Public Const WM_CUT = &H300

Public Const WM_COPY = &H301

Public Const WM_PASTE = &H302

Public Const WM_CLEAR = &H303

Public Const GWL_STYLE = (-16)

Public Const ES_UPPERCASE = &H8&

Public Function TextBoxHelpers_Undo(ByVal TextBox As TextBox) As Long

    TextBoxHelpers_Undo = TextBoxHelpers_SendMessage(TextBox.hwnd, WM_UNDO, 0, 0)

End Function

Public Function TextBoxHelpers_Cut(ByVal TextBox As TextBox) As Long

    TextBoxHelpers_Cut = TextBoxHelpers_SendMessage(TextBox.hwnd, WM_CUT, 0, 0)

End Function

Public Function TextBoxHelpers_Copy(ByVal TextBox As TextBox) As Long

    TextBoxHelpers_Copy = TextBoxHelpers_SendMessage(TextBox.hwnd, WM_COPY, 0, 0)

End Function

Public Function TextBoxHelpers_Paste(ByVal TextBox As TextBox) As Long

    TextBoxHelpers_Paste = TextBoxHelpers_SendMessage(TextBox.hwnd, WM_PASTE, 0, 0)

End Function

Public Function TextBoxHelpers_Clear(ByVal TextBox As TextBox) As Long

    TextBoxHelpers_Clear = TextBoxHelpers_SendMessage(TextBox.hwnd, WM_CLEAR, 0, 0)

End Function

Public Sub TextBoxHelpers_SelectAll(ByVal TextBox As TextBox)

    TextBox.SelStart = 0
    TextBox.SelLength = Len(TextBox.Text)

End Sub

Public Function TextBoxHelpers_SetUpperCase(ByVal TextBox As TextBox, _
                                            ByVal Uppercase As Boolean) As Long

    If Uppercase Then
        TextBoxHelpers_SetUpperCase = TextBoxHelpers_SetWindowLong(TextBox.hwnd, GWL_STYLE, TextBoxHelpers_GetWindowLong(TextBox.hwnd, GWL_STYLE) Or ES_UPPERCASE)
    Else
        TextBoxHelpers_SetUpperCase = TextBoxHelpers_SetWindowLong(TextBox.hwnd, GWL_STYLE, TextBoxHelpers_GetWindowLong(TextBox.hwnd, GWL_STYLE) And Not ES_UPPERCASE)
    End If

End Function


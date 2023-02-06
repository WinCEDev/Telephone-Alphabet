Attribute VB_Name = "Phonetic"
Option Explicit

Public Phonetic_Language As String 'Contains the name of the loaded language.

Private Index            As Integer 'Contains Unicode character values.

Private Value            As String 'Contains phonetic word to use for the entered character.

Public Sub Phonetic_ReadFile(ByVal FileObj As File, _
                             ByVal FilePath As String, _
                             ByVal Language As String)

    FileObj.Open FilePath, fsModeInput, fsAccessRead, fsLockWrite

    Dim strContents As String

    strContents = Split(FileObj.Input(FileObj.LOF), vbNewLine)
    
    Dim lngUpperBound As Long

    lngUpperBound = UBound(strContents)
    
    Index = Empty
    Value = Empty
    
    ReDim Index(lngUpperBound)
    ReDim Value(lngUpperBound)

    Dim i As Long

    For i = 0 To UBound(strContents)
        Index(i) = CLng(Left(strContents(i), InStr(strContents(i), vbTab) - 1))
        Value(i) = Right(strContents(i), Len(strContents(i)) - InStrRev(strContents(i), vbTab))
    Next

    FileObj.Close
    
    Phonetic_Language = Language

End Sub

Public Function Phonetic_FromString(ByVal Text As String) As String

    Dim strResult As String

    Dim i         As Long

    For i = 0 To Len(Text) - 1

        strResult = strResult & Phonetic_FromChar(AscW(Mid(Text, i + 1, 1))) & " "

    Next

    Phonetic_FromString = strResult

End Function

Public Function Phonetic_FromChar(ByVal Char As Long) As String

    'Convert character to uppercase if neccesary.
    If Char >= 97 Then
        If Char <= 122 Then
            Char = Char - 32
        End If
    End If

    Dim i As Long

    For i = 0 To UBound(Index)

        If Index(i) = Char Then
            Phonetic_FromChar = Value(i)

            Exit Function

        End If

    Next

    Phonetic_FromChar = vbNullString

End Function


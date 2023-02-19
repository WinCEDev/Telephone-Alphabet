Attribute VB_Name = "Phonetic"
Option Explicit

Private Const phLanguage As Long = 0

Private Const phIndex    As Long = 1

Private Const phValue    As Long = 2

Public Function Phonetic_Create(ByRef FileObj As File, _
                                ByRef FilePath As String, _
                                ByRef Language As String)

    FileObj.Open FilePath, fsModeInput, fsAccessRead, fsLockWrite

    Dim strContents As String

    strContents = Split(FileObj.Input(FileObj.LOF), vbNewLine)
    
    Dim lngUpperBound As Long

    lngUpperBound = UBound(strContents)
    
    Dim varIndex As Variant 'Contains Unicode character values.

    Dim varValue As Variant 'Contains phonetic word to use for the entered character.
    
    ReDim varIndex(lngUpperBound)
    ReDim varValue(lngUpperBound)

    Dim i As Long
    
    For i = 0 To UBound(strContents)
        varIndex(i) = CLng(Left(strContents(i), InStr(strContents(i), vbTab) - 1))
        varValue(i) = Right(strContents(i), Len(strContents(i)) - InStrRev(strContents(i), vbTab))
    Next

    FileObj.Close
    
    Phonetic_Create = Array(Language, varIndex, varValue)

End Function

Public Sub Phonetic_Destroy(ByRef Instance As Variant)
    Instance(phLanguage) = Empty
    Erase Instance(phIndex)
    Erase Instance(phValue)
End Sub

Public Function Phonetic_FromString(ByRef Instance As String, _
                                    ByRef Text As String) As String

    Dim strWords() As String

    ReDim strWords(Len(Text) - 1)

    Dim i As Long

    For i = 0 To UBound(strWords)

        strWords(i) = Phonetic_FromChar(Instance, AscW(Mid(Text, i + 1, 1)))

    Next
    
    strWords(UBound(strWords) - 1) = strWords(UBound(strWords) - 1) & " "

    Phonetic_FromString = Join(strResult, " ")

End Function

Public Function Phonetic_FromChar(ByRef Instance As Variant, ByRef Char As Long) As String

    'Convert character to uppercase if neccesary.
    If Char >= 97 Then
        If Char <= 122 Then
            Char = Char - 32
        End If
    End If

    Dim i As Long

    For i = 0 To UBound(Instance(phIndex))

        If Instance(phIndex)(i) = Char Then
            Phonetic_FromChar = Instance(phValue)(i)

            Exit Function

        End If

    Next

    Phonetic_FromChar = vbNullString

End Function




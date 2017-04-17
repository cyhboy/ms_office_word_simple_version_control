Private Sub Document_Open()
    theDrive = GetTheDrive
    theUser = Environ$("username")
    Dim tempAuthor As String
    tempAuthor = ActiveDocument.BuiltInDocumentProperties("Author").Value

    ' TODO:
    ' 1.Remove current owner in the owner list if have
    ' 2.Close the file directly without saving
    
    Dim Arr

    Dim idx As Integer
    
    If InStr(tempAuthor, theUser) > 0 Then
        Arr = Split(tempAuthor, ";")
        For i = 0 To UBound(Arr)
            If InStr(Arr(i), theUser) Then
                idx = i
                Exit For
            End If
            
        Next
    
        If idx = UBound(Arr) Then
            ActiveDocument.BuiltInDocumentProperties("Author").Value = Replace(tempAuthor, ";" & Arr(idx), "")
        Else
            ActiveDocument.BuiltInDocumentProperties("Author").Value = Replace(tempAuthor, Arr(idx) & ";", "")
        End If

    End If
    
    If CheckSynWithM_OPEN Then
        Me.Saved = True
        ActiveDocument.Close
    Else
        ActiveDocument.BuiltInDocumentProperties("Author").Value = tempAuthor
        On Error Resume Next
        ActiveWindow.Caption = ActiveDocument.FullName
    End If
End Sub

Public Sub FileSave()
    ActiveDocument.Save
    CheckSynWithAllAvailableDrive
End Sub

Public Sub Document_Close()
    Me.Saved = True
End Sub

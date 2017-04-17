Option Explicit

Private Sub Document_Open()
    theDrive = GetTheDrive
    theUser = Environ$("username")
    Dim tempAuthor As String
    tempAuthor = ActiveDocument.BuiltInDocumentProperties("Author").Value

    ' TODO:
    ' 1.Remove current owner in the owner list if have
    ' 2.Close the file directly without saving
    
    Dim Arr As Variant

    Dim idx As Integer
    Dim i As Integer
    
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
        'On Error Resume Next
        ActiveDocument.ActiveWindow.Caption = ActiveDocument.FullName
    End If
End Sub

Public Sub FileSave()
    ActiveDocument.Save
    CheckSynWithAllAvailableDrive
End Sub

Public Sub Document_Close()
    Me.Saved = True
End Sub

Public Function CheckSynWithM_OPEN()
    If testing Then Exit Function
    
    Dim closeFlag As Boolean
    closeFlag = False
    
    Dim activeName As String
    activeName = ActiveDocument.FullName
    
    If InStr(activeName, "http") = 1 Then
        CheckSynWithM_OPEN = closeFlag
        Exit Function
    End If
    
    Dim fso As Object
    Dim oFileObj, cFileObj, mFileObj As Object
    
    Dim cPath, mPath As String
    Dim iRet As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFileObj = fso.GetFile(activeName)
    
    Dim path As String
    Dim parameter As String
    
    If InStr(activeName, "C:") > 0 Then
        mPath = Replace(activeName, "C:", theDrive)
        
        If Dir(mPath) <> "" Then
            Set mFileObj = fso.GetFile(mPath)
            
            If oFileObj.DateLastModified < mFileObj.DateLastModified Then
                If InStr(ActiveDocument.BuiltInDocumentProperties("Author").Value, theUser) = 0 Then
                    MyQuestionBox "You are not the author of this document and found another updated verion in share drive, Do U want to override from " & theDrive & " after close? Last modifier is " & DocumentProperties(mPath, "Last Author"), "Yes", "No", 10
                    If confirmation = "Yes" Then
                        nexttime = Now() + TimeSerial(0, 0, 5)
                        Application.OnTime nexttime, "'CpFil2FilBk """ & mPath & """, """ & activeName & """, True'"
                        
                        closeFlag = True
                    Else
                        closeFlag = False
                    End If
                Else
                    closeFlag = False
                End If
            Else
                closeFlag = False
            End If
        Else
            closeFlag = False
        End If
    ElseIf InStr(activeName, theDrive) > 0 Then
        cPath = Replace(activeName, theDrive, "C:")
        
        If Dir(cPath) <> "" Then
            Set cFileObj = fso.GetFile(cPath)
            
            If oFileObj.DateLastModified < cFileObj.DateLastModified Then
                If InStr(ActiveDocument.BuiltInDocumentProperties("Author").Value, theUser) = 0 Then
                    MyQuestionBox "You are not the author of this document and found another updated verion in C drive, Do U want to override from C: after close? Last modifier is " & DocumentProperties(cPath, "Last Author"), "Yes", "No", 10
                    If confirmation = "Yes" Then
                        nexttime = Now() + TimeSerial(0, 0, 5)
                        Application.OnTime nexttime, "'CpFil2FilBk """ & cPath & """, """ & activeName & """, True'"
                        
                        closeFlag = True
                    Else
                        closeFlag = False
                    End If
                Else
                    closeFlag = False
                End If
            Else
                closeFlag = False
            End If
        Else
            closeFlag = False
        End If
        
    End If
    
    Set fso = Nothing
    
    CheckSynWithM_OPEN = closeFlag
End Function

Public Function CheckSynWithAllAvailableDrive_OPEN()
    If testing Then Exit Function
    
    Dim closeFlag As Boolean
    
    Dim activeName As String
    activeName = ActiveDocument.FullName
    
    Dim fso As Object
    Dim oFileObj, cFileObj, mFileObj As Object
    
    Dim cPath, mPath As String
    Dim iRet As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFileObj = fso.GetFile(activeName)
    
    Dim path As String
    Dim parameter As String
    
    For Each obj In fso.Drives
        
        If InStr(activeName, "C:") > 0 Then
            mPath = Replace(activeName, "C:", obj.path)
            
            If Dir(mPath) <> "" Then
                Set mFileObj = fso.GetFile(mPath)
                
                If oFileObj.DateLastModified < mFileObj.DateLastModified Then
                    If InStr(ActiveDocument.BuiltInDocumentProperties("Author").Value, theUser) = 0 Then
                        MyQuestionBox "You are not the author of this document and found another updated verion in share drive, Do U want to override from " & theDrive & " after close? Last modifier is " & DocumentProperties(mPath, "Last Author"), "Yes", "No", 10
                        If confirmation = "Yes" Then
                            nexttime = Now() + TimeSerial(0, 0, 5)
                            Application.OnTime nexttime, "'CpFil2FilBk """ & mPath & """, """ & activeName & """, True'"
                            
                            closeFlag = True
                        Else
                            closeFlag = False
                        End If
                    Else
                        closeFlag = False
                    End If
                Else
                    closeFlag = False
                End If
            Else
                closeFlag = False
            End If
        ElseIf InStr(activeName, obj.path) > 0 Then
            cPath = Replace(activeName, obj.path, "C:")
            
            If Dir(cPath) <> "" Then
                Set cFileObj = fso.GetFile(cPath)
                
                If oFileObj.DateLastModified < cFileObj.DateLastModified Then
                    If InStr(ActiveDocument.BuiltInDocumentProperties("Author").Value, theUser) < 0 Then
                        MyQuestionBox "You are not the author of this document and found another updated verion in C drive, Do U want to override from C: after close? Last modifier is " & DocumentProperties(cPath, "Last Author"), "Yes", "No", 10
                        If confirmation = "Yes" Then
                            nexttime = Now() + TimeSerial(0, 0, 5)
                            Application.OnTime nexttime, "'CpFil2FilBk """ & cPath & """, """ & activeName & """, True'"
                            
                            closeFlag = True
                        Else
                            closeFlag = False
                        End If
                    Else
                        closeFlag = False
                    End If
                Else
                    closeFlag = False
                End If
            Else
                closeFlag = False
            End If
            
        End If
        
        
    Next
    
    Set fso = Nothing
    
    CheckSynWithAllAvailableDrive_OPEN = closeFlag
End Function


Public Sub CheckSynWithAllAvailableDrive()
    If testing Then Exit Sub
    
    Dim activeName As String
    activeName = ActiveDocument.FullName
    
    Dim fso As Object
    Dim oFileObj, cFileObj, mFileObj As Object
    
    Dim cPath, mPath As String
    Dim iRet As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFileObj = fso.GetFile(activeName)
    
    Dim obj As Object
    
    For Each obj In fso.Drives
        
        On Error GoTo ErrorHandler
        If obj.DriveType = 3 Then
        'MsgBox obj.path
        If InStr(activeName, "C:") > 0 Then
            mPath = Replace(activeName, "C:", obj.path)
            
            If Dir(mPath) <> "" Then
                Set mFileObj = fso.GetFile(mPath)
                
                If oFileObj.DateLastModified > mFileObj.DateLastModified Then
                    
                    If InStr(ActiveDocument.BuiltInDocumentProperties("Author").Value, theUser) > 0 Then
                        MyQuestionBox "You are the author of this document, Do U want to update " & obj.path & " as well? ", "Yes", "No", 10
                        If confirmation = "Yes" Then
                            fso.copyfile activeName, mPath, True
                        End If
                    End If
                    
                End If
                
            End If
        ElseIf InStr(activeName, theDrive) > 0 Then
            cPath = Replace(activeName, theDrive, "C:")
            
            If Dir(cPath) <> "" Then
                Set cFileObj = fso.GetFile(cPath)
                
                If oFileObj.DateLastModified > cFileObj.DateLastModified Then
                    
                    If InStr(ActiveDocument.BuiltInDocumentProperties("Author").Value, theUser) = 0 Then
                        MyQuestionBox "You are not the author of this document, Do U want to update C: as well? ", "Yes", "No", 10
                        If confirmation = "Yes" Then
                            fso.copyfile activeName, cPath, True
                        End If
                    End If
                    
                End If
                
            End If
            
        End If
        End If
ErrorHandler:
        If Err.Number <> 0 Then
            MyMsgBox Err.Number & " " & Err.Description, 30
        End If
    Next
    
    Set fso = Nothing
End Sub

Public Sub SynMZ()
    If testing Then Exit Function
    
    Dim activeName As String
    activeName = ActiveDocument.FullName
    
    Dim fso As Object
    Dim oFileObj, cFileObj, mFileObj As Object
    
    Dim cPath, mPath As String
    Dim iRet As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFileObj = fso.GetFile(activeName)
    
    For Each obj In fso.Drives
        If obj.path = theDrive Or obj.path = "Z:" Then
            
            If InStr(activeName, "C:") > 0 Then
                mPath = Replace(activeName, "C:", obj.path)
                
                If fso.FileExists(mPath) Then
                    
                    Set mFileObj = fso.GetFile(mPath)
                    
                    If oFileObj.DateLastModified > mFileObj.DateLastModified Then
                        If InStr(ActiveDocument.BuiltInDocumentProperties("Author").Value, theUser) = 0 Then
                            iRet = MsgBox("You are NOT the author of this document, Do U want to manually update " & obj.path & " as well? ", vbYesNo, "Question")
                            If iRet = vbYes Then
                                fso.copyfile activeName, mPath, True
                            End If
                        End If
                    End If
                Else
                    If InStr(ActiveDocument.BuiltInDocumentProperties("Author").Value, theUser) > 0 Then
                        iRet = MsgBox("You are the author of this document, Do U want to manually append " & obj.path & " as well? ", vbYesNo, "Question")
                        If iRet = vbYes Then
                            fso.copyfile activeName, mPath, True
                        End If
                    End If
                End If
            ElseIf InStr(activeName, obj.path) > 0 Then
                cPath = Replace(activeName, obj.path, "C:")
                
                If fso.FileExists(cPath) Then
                    Set cFileObj = fso.GetFile(cPath)
                    
                    If oFileObj.DateLastModified > cFileObj.DateLastModified Then
                        If InStr(ActiveDocument.BuiltInDocumentProperties("Author").Value, theUser) > 0 Then
                            iRet = MsgBox("You are the author of this document, Do U want to manually update C: as well? ", vbYesNo, "Question")
                            If iRet = vbYes Then
                                fso.copyfile activeName, cPath, True
                            End If
                        End If
                    End If
                Else
                    If InStr(ActiveDocument.BuiltInDocumentProperties("Author").Value, theUser) = 0 Then
                        iRet = MsgBox("You are NOT the author of this document, Do U want to manually append C: as well? ", vbYesNo, "Question")
                        If iRet = vbYes Then
                            fso.copyfile activeName, cPath, True
                        End If
                    End If
                End If
                
            End If
            
        End If
    Next
    
    Set fso = Nothing
End Sub

Public Sub DspEnv()
    WScript.Echo theUser & " is connecting " & theDrive
End Sub

Public Function GetTheDrive()
    Dim mDrive As String
    mDrive = ""
    Dim fso As Object
    Dim obj As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim i As Integer
    For i = Asc("M") To Asc("Z")
        If mDrive <> "" Then
            Exit For
        End If
        For Each obj In fso.Drives
            If obj.path = Chr(i) & ":" Then
                If Dir(obj.path & "\AppFiles\SupportSetup\Justacro.xlam") <> "" Then
                    mDrive = obj.path
                    Exit For
                End If
            End If
        Next
    Next
    Set fso = Nothing
    GetTheDrive = mDrive
End Function

Public Sub CpFil2FilBk(filePath1 As String, filePath2 As String, displayFlag As Boolean)
    If testing Then Exit Sub
    On Error GoTo ErrorHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim result As String
    result = fso.copyfile(filePath1, filePath2)
    Set fso = Nothing
    
    If displayFlag Then
        If result = "" Then
            MyMsgBox filePath1 & " to " & filePath2 & " copied", 5
        End If
    End If
    Application.Workbooks.Open filePath2

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

Public Function DocumentProperties(filePath, propName) As String
    Dim retValue As String
    Dim appOffice As New Application
    Dim richFile As Document
    Set richFile = appOffice.Documents.Open(filePath)
    retValue = richFile.BuiltInDocumentProperties(propName)
    richFile.Saved = True
    richFile.Close
    appOffice.Quit
    Set appOffice = Nothing
    DocumentProperties = retValue
End Function

Public Sub MyMsgBox(detail As String, duration As Long)
    If testing Then Exit Sub
    nexttime = Now() + TimeSerial(0, 0, duration)
    Application.OnTime nexttime, "MyMsgBoxHide"
    
    Set uf1 = New UserForm1
    uf1.TextBox1.Text = detail
    uf1.TextBox1.SetFocus
    uf1.Show
End Sub

Public Sub MyQuestionBox(detail As String, answer1 As String, answer2 As String, duration As Long)
    If testing Then Exit Sub
    nexttime = Now() + TimeSerial(0, 0, duration)
    Application.OnTime nexttime, "MyQuestionBoxHide"
    confirmation = ""
    
    Set uf2 = New UserForm2
    uf2.CommandButton1.Caption = answer1
    uf2.CommandButton2.Caption = answer2
    uf2.TextBox1.Text = detail
    uf2.TextBox1.SetFocus
    uf2.Show
End Sub

Public Sub MyQuestionBoxHide()
    If testing Then Exit Sub
    confirmation = uf2.CommandButton1.Caption
    uf2.Hide
    Set uf2 = Nothing
End Sub

Public Sub MyMsgBoxHide()
    If testing Then Exit Sub
    uf1.Hide
    Set uf1 = Nothing
End Sub

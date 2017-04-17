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
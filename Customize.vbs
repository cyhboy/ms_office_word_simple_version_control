Option Explicit

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
		
		On Error Goto ErrorHandler
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
	If testing Then Exit Sub
	MsgBox theUser & " is connecting " & theDrive
End Sub

Public Sub CpFil2FilBk(filePath1 As String, filePath2 As String, displayFlag As Boolean)
	If testing Then Exit Sub
	On Error Goto ErrorHandler
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
	Application.Documents.Open filePath2
	
	ErrorHandler:
	If Err.Number <> 0 Then
		MyMsgBox Err.Number & " " & Err.Description, 30
	End If
End Sub

Public Sub MyMsgBox(detail As String, duration As Long)
	If testing Then Exit Sub
	nexttime = Now() + TimeSerial(0, 0, duration)
	Application.OnTime nexttime, "MyMsgBoxHide"
	
	Set uf1 = New UserForm1
	uf1.TextBox1.text = detail
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
	uf2.TextBox1.text = detail
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

Public Sub TestVBA()
	testing = True
	On Error Goto ErrorHandler
	Dim objProject As Object
	
	Dim objComponent As Object
	
	Dim objCode As Object
	
	' Declare other miscellaneous variables.
	Dim iLine As Integer
	Dim sProcName As String
	Dim pk As VBIDE.vbext_ProcKind
	
	
	Dim i As Integer
	Dim comm As String
	Dim codeOfLine As String
	
	Set objProject = NormalTemplate.VBProject
	
	Dim subCount0 As Integer
	Dim subCount1 As Integer
	Dim subCount2 As Integer
	Dim subCount3 As Integer
	Dim subCount4 As Integer
	Dim subCount5 As Integer
	Dim subCount6 As Integer
	Dim subCount7 As Integer
	Dim subCount8 As Integer
	Dim subCountX As Integer
	
	Dim funcCount0 As Integer
	Dim funcCount1 As Integer
	Dim funcCount2 As Integer
	Dim funcCount3 As Integer
	Dim funcCount4 As Integer
	Dim funcCount5 As Integer
	Dim funcCount6 As Integer
	Dim funcCount7 As Integer
	Dim funcCount8 As Integer
	Dim funcCountX As Integer
	
	Dim xObj As Variant
	'Iterate through each component in the project.
	For Each objComponent In objProject.VBComponents
		'If InStr(objComponent.Name, "All") > 0 Or InStr(objComponent.Name, "SubParam") > 0 Or InStr(objComponent.Name, "FuncNoParam") > 0 Or InStr(objComponent.Name, "FuncParam") Then
		'Find the code module for the project.
		Set objCode = objComponent.CodeModule
		'Scan through the code module, looking for procedures.
		iLine = 1
		Do While iLine < objCode.CountOfLines
			
			codeOfLine = objCode.Lines(iLine, 1)
			If Trim(codeOfLine) <> "" And Not StartsWith(Trim(codeOfLine), "'") Then
				sProcName = objCode.ProcOfLine(iLine, pk)
				If sProcName <> "" And sProcName <> "Ver" And sProcName <> "Test" And sProcName <> "TestVBA" And sProcName <> "CountRegx" And sProcName <> "ListNodes" And sProcName <> "CntOfficeUI" And sProcName <> "TestCall" And sProcName <> "StartsWith" And sProcName <> "EndsWith" Then
					comm = ""
					If testing Then
						If InStr(Trim(codeOfLine), "Public Sub " & sProcName & "()") > 0 Then
							'MsgBox objComponent.Name & "." & sProcName & " " & Trim(codeOfLine)
							'RobotRunByParam objComponent.Name & "." & sProcName
							comm = objComponent.Name & "." & sProcName
							Application.Run comm
							subCount0 = subCount0 + 1
						Else
							If InStr(Trim(codeOfLine), "Public Sub " & sProcName) > 0 Then
								'MsgBox objComponent.Name & "." & sProcName & " " & Trim(codeOfLine)
								comm = objComponent.Name & "." & sProcName
								If CountRegx(Trim(codeOfLine), ", ") = 0 Then
									Application.Run comm, "0"
									subCount1 = subCount1 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 1 Then
									Application.Run comm, "0", "0"
									subCount2 = subCount2 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 2 Then
									Application.Run comm, "0", "0", "0"
									subCount3 = subCount3 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 3 Then
									Application.Run comm, "0", "0", "0", "0"
									subCount4 = subCount4 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 4 Then
									Application.Run comm, "0", "0", "0", "0", "0"
									subCount5 = subCount5 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 5 Then
									Application.Run comm, "0", "0", "0", "0", "0", "0"
									subCount6 = subCount6 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 6 Then
									Application.Run comm, "0", "0", "0", "0", "0", "0", "0"
									subCount7 = subCount7 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 7 Then
									Application.Run comm, "0", "0", "0", "0", "0", "0", "0", "0"
									subCount8 = subCount8 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") >= 8 Then
									'Application.Run comm, "0", "0", "0", "0", "0", "0", "0", "0", "0"
									MsgBox objComponent.Name & "." & sProcName & " " & Trim(codeOfLine) & " Not Test As Too Many Param"
									subCountX = subCountX + 1
								End If
							End If
						End If
						
						If InStr(Trim(codeOfLine), "Public Function " & sProcName & "()") > 0 Then
							'MsgBox objComponent.Name & "." & sProcName & " " & Trim(codeOfLine)
							'RobotRunByParam objComponent.Name & "." & sProcName
							comm = objComponent.Name & "." & sProcName
							Application.Run comm
							funcCount0 = funcCount0 + 1
						Else
							If InStr(Trim(codeOfLine), "Public Function " & sProcName) > 0 Then
								'MsgBox objComponent.Name & "." & sProcName & " " & Trim(codeOfLine)
								comm = objComponent.Name & "." & sProcName
								If CountRegx(Trim(codeOfLine), ", ") = 0 Then
									Application.Run comm, xObj
									funcCount1 = funcCount1 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 1 Then
									Application.Run comm, xObj, xObj
									funcCount2 = funcCount2 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 2 Then
									Application.Run comm, xObj, xObj, xObj
									funcCount3 = funcCount3 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 3 Then
									Application.Run comm, xObj, xObj, xObj, xObj
									funcCount4 = funcCount4 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 4 Then
									Application.Run comm, xObj, xObj, xObj, xObj, xObj
									funcCount5 = funcCount5 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 5 Then
									Application.Run comm, xObj, xObj, xObj, xObj, xObj, xObj
									funcCount6 = funcCount6 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 6 Then
									Application.Run comm, xObj, xObj, xObj, xObj, xObj, xObj, xObj
									funcCount7 = funcCount7 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") = 7 Then
									Application.Run comm, xObj, xObj, xObj, xObj, xObj, xObj, xObj, xObj
									funcCount8 = funcCount8 + 1
								End If
								If CountRegx(Trim(codeOfLine), ", ") >= 8 Then
									'Application.Run comm, xObj, xObj, xObj, xObj, xObj, xObj, xObj, xObj, xObj
									MsgBox objComponent.Name & "." & sProcName & " " & Trim(codeOfLine) & " Not Test As Too Many Param"
									funcCountX = funcCountX + 1
								End If
							End If
						End If
						
						'Exit Sub
					End If
					
				End If
				'iLine = iLine + objCode.ProcCountLines(sProcName, pk) - 2
			End If
			iLine = iLine + 1
		Loop
		Set objCode = Nothing
		Set objComponent = Nothing
		'End If
	Next
	Set objProject = Nothing
	Dim resultStr As String
	resultStr = resultStr & "Total Sub0 Testing Count " & subCount0 & vbCrLf
	resultStr = resultStr & "Total Sub1 Testing Count " & subCount1 & vbCrLf
	resultStr = resultStr & "Total Sub2 Testing Count " & subCount2 & vbCrLf
	resultStr = resultStr & "Total Sub3 Testing Count " & subCount3 & vbCrLf
	resultStr = resultStr & "Total Sub4 Testing Count " & subCount4 & vbCrLf
	resultStr = resultStr & "Total Sub5 Testing Count " & subCount5 & vbCrLf
	resultStr = resultStr & "Total Sub6 Testing Count " & subCount6 & vbCrLf
	resultStr = resultStr & "Total Sub7 Testing Count " & subCount7 & vbCrLf
	resultStr = resultStr & "Total Sub8 Testing Count " & subCount8 & vbCrLf
	resultStr = resultStr & "Total SubX Not Testing Count " & subCountX & vbCrLf
	
	resultStr = resultStr & "Total Sub Count " & (subCount0 + subCount1 + subCount2 + subCount3 + subCount4 + subCount5 + subCount6 + subCount7 + subCount8 + subCountX) & vbCrLf & vbCrLf
	
	resultStr = resultStr & "Total Func0 Testing Count " & funcCount0 & vbCrLf
	resultStr = resultStr & "Total Func1 Testing Count " & funcCount1 & vbCrLf
	resultStr = resultStr & "Total Func2 Testing Count " & funcCount2 & vbCrLf
	resultStr = resultStr & "Total Func3 Testing Count " & funcCount3 & vbCrLf
	resultStr = resultStr & "Total Func4 Testing Count " & funcCount4 & vbCrLf
	resultStr = resultStr & "Total Func5 Testing Count " & funcCount5 & vbCrLf
	resultStr = resultStr & "Total Func6 Testing Count " & funcCount6 & vbCrLf
	resultStr = resultStr & "Total Func7 Testing Count " & funcCount7 & vbCrLf
	resultStr = resultStr & "Total Func8 Testing Count " & funcCount8 & vbCrLf
	resultStr = resultStr & "Total FuncX Not Testing Count " & funcCountX & vbCrLf
	
	resultStr = resultStr & "Total Func Count " & (funcCount0 + funcCount1 + funcCount2 + funcCount3 + funcCount4 + funcCount5 + funcCount6 + funcCount7 + funcCount8 + funcCountX) & vbCrLf & vbCrLf
	
	MsgBox resultStr
	ErrorHandler:
	If Err.Number <> 0 Then
		MyMsgBox Err.Number & " " & Err.Description & " " & objComponent.Name & "." & sProcName, 30
	End If
	testing = False
End Sub



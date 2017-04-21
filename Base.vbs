Option Explicit

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
	
	Dim obj As Object
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

Public Function GetTheDrive()
	If testing Then Exit Function
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

Public Function DocumentProperties(filePath, propName) As String
	If testing Then Exit Function
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

Public Function CountRegx(text As String, patt As String) As Long
	On Error Goto ErrorHandler
	Dim RE As New RegExp
	RE.Pattern = patt
	RE.Global = True
	RE.IgnoreCase = False
	RE.MultiLine = True
	'Retrieve all matches
	Dim Matches As MatchCollection
	Set Matches = RE.Execute(text)
	'Return the corrected count of matches
	CountRegx = Matches.Count
	ErrorHandler:
	If Err.Number <> 0 Then
		MyMsgBox Err.Number & " " & Err.Description, 30
	End If
End Function

Public Function EndsWith(str As String, ending As String) As Boolean
	'If testing Then Exit Function
	Dim endingLen As Integer
	endingLen = Len(ending)
	EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function

Public Function StartsWith(str As String, start As String) As Boolean
	'If testing Then Exit Function
	Dim startLen As Integer
	startLen = Len(start)
	StartsWith = (Left(Trim(UCase(str)), startLen) = UCase(start))
End Function

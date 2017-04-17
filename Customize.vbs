Public Sub DspEnv()
	WScript.Echo theUser & " is connecting " & theDrive
End Sub

Public Function GetTheDrive()
	Dim mDrive As String
	mDrive = ""
	Dim fso As Object
	Dim obj As Object
	Set fso = CreateObject("Scripting.FileSystemObject")
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
    Dim result
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

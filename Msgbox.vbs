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
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


Public Sub DspEnv()
    Wscript.Echo theUser & " is connecting " & theDrive
End Sub

Public Function GetTheDrive()
    Dim mDrive
    mDrive = ""
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
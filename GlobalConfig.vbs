Option Explicit

Public theDrive:theDrive="M:" 'Must Be Defined In Module, the default share drive value is M:, M: means moblile.
Public theUser:theUser="Guest" 'Must Be Defined In Module, user have 2 mode, standalone and active directory

Public Sub DspEnv()
Wscript.Echo theUser & " is connecting " & theDrive
End Sub

Call DspEnv 'For testing only
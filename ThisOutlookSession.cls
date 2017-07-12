VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Application_Reminder(ByVal Item As Object)

'Tries to ensure that it only runs once a day and doesn't dive into a deeper module if not needed
    setting = GetSetting("LastRun", "Automation", "Date", DateValue(Now) - 1)
    
    If DateValue(setting) <> DateValue(Now) Then
        BlockTime "Email"
        BlockTime "Lunch"
        BlockTime "Exercise"
        
        SaveSetting "LastRun", "Automation", "Date", Now
    End If
    
End Sub


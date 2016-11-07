Attribute VB_Name = "Automation"
Sub Test()
'BlockTime "Email"
BlockTime "Lunch"
End Sub
Sub None15()
    MakeAppointment 15, "", ""
End Sub
Sub None30()
    MakeAppointment 30, "", ""
End Sub
Sub None60()
    MakeAppointment 60, "", ""
End Sub
Sub Done15()
    MakeAppointment 15, "Done", ""
End Sub
Sub Done30()
    MakeAppointment 30, "Done", ""
End Sub
Sub Done60()
    MakeAppointment 60, "Done", ""
End Sub


Sub BlockTime(strType As String)
    On Error GoTo ext:
    InProgress = GetSetting("InProgress", "Automation", "InProgress", "False") 'Need to make sure multiple requests don't kick off
    If InProgress = "False" Then
        SaveSetting "InProgress", "Automation", "InProgress", "True"
        
        Dim oCurrentUser As ExchangeUser
        Dim FreeBusy As String
        Dim BusySlot As Long
        Dim DateBusySlot As Date
        Dim aryDateBusySlot(50) As Date
        Dim SlotLength As Integer
        Dim i As Long
        Set myOlApp = CreateObject("Outlook.Application")
        Set myNameSpace = myOlApp.GetNamespace("MAPI")
        Set MyFolder = myNameSpace.GetDefaultFolder(olFolderCalendar)
        Dim Scheduled As Boolean
        Dim X As Integer
        Dim setting As String
        Dim RandNum As Integer
        Dim intDateDiff As Integer
        
        If strType = "Email" Then
            setting = GetSetting("LastScheduleEmail", "Automation", "Date", DateValue(Now) - 6)
            
        ElseIf strType = "Lunch" Then
            setting = GetSetting("LastBlockLunch", "Automation", "Date", DateValue(Now) - 6)
        End If
        'setting = DateAdd("d", -5, Now)
         
         'Compare Values for Daily Run
         If DateValue(setting) <> DateValue(Now) Then
         
            'Get the DateDiff
            intDateDiff = DateDiff("d", DateValue(setting), DateValue(Now))
            If intDateDiff >= 7 Then 'don't want to do the past
                intDateDiff = 6
            End If
                
            'Set Variables
            Scheduled = False
            
    
            'Get ExchangeUser for CurrentUser
            If Application.Session.CurrentUser.AddressEntry.Type = "EX" Then
                Set oCurrentUser = Application.Session.CurrentUser.AddressEntry.GetExchangeUser
                
                'Loop through each date that was missed
                For Y = 1 To intDateDiff
                    X = 1
                    Erase aryDateBusySlot
                    
                    'TODO
                    'Reset arryDateSlot
                    'Randomize SlotLength
                    
                    'Randomize the SlotLength
                    Randomize
                    RandNum = (2 * Rand) + 1
                    If RandNum = 1 Then
                        SlotLength = 60
                    Else
                        SlotLength = 30
                    End If
a:
                    'Reset FreeBusy
                    FreeBusy = oCurrentUser.GetFreeBusy(Now, SlotLength)
                    
                    For i = 1 To Len(FreeBusy)
                       If CLng(Mid(FreeBusy, i, 1)) = 0 Then
                          'get the number of minutes into the day for free interval
                          BusySlot = (i - 1) * SlotLength
                          
                          'get an actual date/time
                          DateBusySlot = DateAdd("n", BusySlot, Date)
                          
                          'Morning 9:30-11:30
                          'Afternoon 1-5
                          'Single Day
                          'Not Weekend
                          If strType = "Email" Then
                              If ((TimeValue(DateBusySlot) >= TimeValue(#9:30:00 AM#) And TimeValue(DateBusySlot) <= TimeValue(#11:30:00 AM#)) Or _
                                   (TimeValue(DateBusySlot) >= TimeValue(#1:00:00 PM#) And TimeValue(DateBusySlot) < TimeValue(#5:00:00 PM#))) And _
                                   DateValue(DateBusySlot) >= DateAdd("d", 7 - (Y - 1), DateValue(Now)) And DateValue(DateBusySlot) < DateAdd("d", 8 - (Y - 1), DateValue(Now)) And _
                                   Weekday(DateBusySlot) <> 7 And Weekday(DateBusySlot) <> 1 Then
                                   
                                   aryDateBusySlot(X) = DateBusySlot
                                   X = X + 1
                                   'MsgBox (oCurrentUser.name & " first open interval:" & vbCrLf & Format$(DateBusySlot, "dddd, mmm d yyyy hh:mm AMPM"))
                                   
                                   'Scheduled = True
                                 'Exit For
                              End If
                           ElseIf strType = "Lunch" Then
                               If TimeValue(DateBusySlot) >= TimeValue(#11:30:00 AM#) And TimeValue(DateBusySlot) < TimeValue(#1:00:00 PM#) And _
                                   DateValue(DateBusySlot) >= DateAdd("d", 7 - (Y - 1), DateValue(Now)) And DateValue(DateBusySlot) < DateAdd("d", 8 - (Y - 1), DateValue(Now)) And _
                                   Weekday(DateBusySlot) <> 7 And Weekday(DateBusySlot) <> 1 Then
                                   
                                   aryDateBusySlot(X) = DateBusySlot
                                   X = X + 1
                                   'MsgBox (oCurrentUser.name & " first open interval:" & vbCrLf & Format$(DateBusySlot, "dddd, mmm d yyyy hh:mm AMPM"))
                                   
                                   'Scheduled = True
                                 'Exit For
                              End If
                           End If
                       End If
                    Next i
                    
                    'If Array is empty and timeslot = 60 then try 30, if 30 then try 15
                    If aryDateBusySlot(1) = "12:00:00 AM" Then
                        If SlotLength = 60 Then
                            SlotLength = 30
                            GoTo a:
                        ElseIf SlotLength = 30 Then
                            SlotLength = 15
                            GoTo a:
                        ElseIf SlotLength = 15 Then
                            GoTo b:
                        End If
                    End If
                    
                    'X will always be 1 higher than the count, subtract 1 to get right count
                    X = X - 1
                    
                    'Generate a random number between 1 and number of slots to pick a timeframe
                    Dim finDateBusySlot As Date
                    
                    If X > 1 Then
                        Randomize
                        RandNum = Int(X * Rnd) + 1
                        finDateBusySlot = aryDateBusySlot(RandNum)
                        Scheduled = True
                    ElseIf X = 1 Or X = 0 Then
                        finDateBusySlot = aryDateBusySlot(1)
                        Scheduled = True
                    End If
                    
                    'Schedule appointment
                    If Scheduled = True Then
                        If strType = "Email" Then
                            'Set the Appointment
                            Set olAppt = MyFolder.Items.Add(olAppointmentItem)
                            
                            
                            
                            With olAppt
                                'Define calendar item properties
                                .Start = finDateBusySlot
                                '.MeetingStatus = olMeeting
                                .End = DateAdd("n", SlotLength, finDateBusySlot)
                                .Subject = "Email"
                                .Location = "Office"
                                .Body = "Created: " & Now
                                .BusyStatus = olBusy
                                .ReminderMinutesBeforeStart = 10
                                .ReminderSet = True
                                .Categories = "Self"
                                
                                'Set myRequiredAttendee = .Recipients.Add("[personal email]")
                                'myRequiredAttendee.Type = olRequired
                                
                                .Save
                                '.Send
                            End With
                            
                            
                            
                            'Save date to registry so it doesn't run again
                            SaveSetting "LastScheduleEmail", "Automation", "Date", Now
                        ElseIf strType = "Lunch" Then
                            'Set the Appointment
                            Set olAppt = MyFolder.Items.Add(olAppointmentItem)
                            With olAppt
                                'Define calendar item properties
                                .Start = finDateBusySlot
				.MeetingStatus = olMeeting
                                .End = DateAdd("n", SlotLength, finDateBusySlot)
                                .Subject = "Lunch"
                                .Location = ""
                                .Body = "Created: " & Now
                                .BusyStatus = olBusy
                                .ReminderMinutesBeforeStart = 10
                                .ReminderSet = True
                                .Categories = "Self"
                                
				Set myRequiredAttendee = .Recipients.Add("[personal email]")
                                myRequiredAttendee.Type = olRequired
                                
                                .Save
                                .Send
                            End With
                            
                            'Save date to registry so it doesn't run again
                            SaveSetting "LastBlockLunch", "Automation", "Date", Now
                        End If
                    End If
b:
                Next Y
            End If
         End If
    End If
ext:
SaveSetting "InProgress", "Automation", "InProgress", "False"
End Sub

Sub MakeAppointment(SlotLength As Integer, strFolder As String, var As String)
    Dim oCurrentUser As ExchangeUser
    Dim FreeBusy As String
    Dim BusySlot As Long
    Dim DateBusySlot As Date
    Dim i As Long
    Set myOlApp = CreateObject("Outlook.Application")
    Set myNameSpace = myOlApp.GetNamespace("MAPI")
    Set MyFolder = myNameSpace.GetDefaultFolder(olFolderCalendar)
    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objSourceFolder As Outlook.MAPIFolder
    Dim objDestFolder As Outlook.MAPIFolder
    Dim Scheduled As Boolean

    'Set Variables
    Scheduled = False
    
    Set objItem = GetCurrentItem()
    If objItem Is Nothing Then
        MsgBox ("Please select an Email")
    Else
        'Get ExchangeUser for CurrentUser
        If Application.Session.CurrentUser.AddressEntry.Type = "EX" Then
            Set oCurrentUser = Application.Session.CurrentUser.AddressEntry.GetExchangeUser
            FreeBusy = oCurrentUser.GetFreeBusy(Now, SlotLength)
            
            'Loop through each date that was missed
            For i = 1 To Len(FreeBusy)
               If CLng(Mid(FreeBusy, i, 1)) = 0 Then
                  'get the number of minutes into the day for free interval
                  BusySlot = (i - 1) * SlotLength
                  
                  'get an actual date/time
                  DateBusySlot = DateAdd("n", BusySlot, Date)
                  
                  'Friday
                  'Afternoon 2-5
                  'Greater than 3 hours from now
                  If var = "Friday" Then
                      If (TimeValue(DateBusySlot) >= TimeValue(#2:00:00 PM#) And TimeValue(DateBusySlot) < TimeValue(#5:00:00 PM#)) And _
                           DateBusySlot >= DateAdd("h", 3, Now) And Weekday(DateBusySlot) = 6 Then
    
                           'MsgBox (oCurrentUser.name & " first open interval:" & vbCrLf & Format$(DateBusySlot, "dddd, mmm d yyyy hh:mm AMPM"))
                                   'Set the Appointment
                            Set olAppt = MyFolder.Items.Add(olAppointmentItem)
                            With olAppt
                                'Define calendar item properties
                                .Start = DateBusySlot
                                .End = DateAdd("n", SlotLength, DateBusySlot)
                                .Subject = objItem.Subject
                                .Location = "Office"
                                .Attachments.Add (objItem)
                                .Body = "Created: " & Now
                                .BusyStatus = olBusy
                                .ReminderMinutesBeforeStart = 10
                                .ReminderSet = True
                                .Categories = "Self"
                                .Save
                            End With
                            
                            If strFolder <> "" Then
                              'Move the item to Input or Files/Forwards
                              Set objOutlook = Application
                              Set objNamespace = objOutlook.GetNamespace("MAPI")
                              Set objSourceFolder = objNamespace.GetDefaultFolder(olFolderDrafts)
                              Set objDestFolder = objNamespace.Folders("[work email]").Folders("Inbox").Folders(strFolder)
        
                              objItem.Move objDestFolder
                            End If
                            Scheduled = True
                            Exit For
                        End If
                  Else
                        'Morning 9:30-11:30
                        'Afternoon 1-5
                        'Greater than 3 hours from now
                        'Not Weekend
                      If ((TimeValue(DateBusySlot) >= TimeValue(#9:30:00 AM#) And TimeValue(DateBusySlot) <= TimeValue(#11:30:00 AM#)) Or _
                           (TimeValue(DateBusySlot) >= TimeValue(#1:00:00 PM#) And TimeValue(DateBusySlot) < TimeValue(#5:00:00 PM#))) And _
                           DateBusySlot >= DateAdd("h", 3, Now) And Weekday(DateBusySlot) <> 7 And Weekday(DateBusySlot) <> 1 Then
    
                           'MsgBox (oCurrentUser.name & " first open interval:" & vbCrLf & Format$(DateBusySlot, "dddd, mmm d yyyy hh:mm AMPM"))
                                   'Set the Appointment
                            Set olAppt = MyFolder.Items.Add(olAppointmentItem)
                            With olAppt
                                'Define calendar item properties
                                .Start = DateBusySlot
                                .End = DateAdd("n", SlotLength, DateBusySlot)
                                .Subject = objItem.Subject
                                .Location = "Office"
                                .Attachments.Add (objItem)
                                .Body = "Created: " & Now
                                .BusyStatus = olBusy
                                .ReminderMinutesBeforeStart = 10
                                .ReminderSet = True
                                .Categories = "Self"
                                .Save
                            End With
                            
                            If strFolder <> "" Then
                              'Move the item to Input or Files/Forwards
                              Set objOutlook = Application
                              Set objNamespace = objOutlook.GetNamespace("MAPI")
                              Set objSourceFolder = objNamespace.GetDefaultFolder(olFolderDrafts)
                              Set objDestFolder = objNamespace.Folders("[work email]").Folders("Inbox").Folders(strFolder)
        
                              objItem.Move objDestFolder
                            End If
                            Scheduled = True
                         Exit For
                      End If
                  End If
               End If
            Next
        End If
    End If
    If Scheduled = False Then
        MsgBox ("There was an Error, nothing was scheduled.")
    Else
        MsgBox ("Done")
    End If
End Sub

Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application
            
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
        
    Set objApp = Nothing
End Function

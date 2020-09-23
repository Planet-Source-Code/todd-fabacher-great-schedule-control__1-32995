Attribute VB_Name = "MGetNextTime"
Option Explicit

Private mdtLastDateTime As Date
Private mxmlSchedule As IXMLDOMElement

Public Type ScheduleData
    InActive As Boolean
    LastActionDateTime As Date
    NextActionDateTime As Date
    
    '*************************
    'Occurance
    Occurrence As schd_Occurs
    
    '##Daily Options
    Occurs_DailyNum As Integer
    
    
    '##Weekly
    Occurs_WeeklyNum As Integer
    Occurs_WeeklyWeekday(7) As Integer


    '##Monthly
    Occurs_MonthlyOption As schd_OccursMonthly
    
    'Each option
    Occurs_Monthly_Each_Day As Integer
    Occurs_Monthly_Each_Num As Integer
    
    'Every option
    Occurs_Monthly_Every_Week As schd_MonthlyWeek
    Occurs_Monthly_Every_WeekDay As schd_DayofWeek
    Occurs_Monthly_Every_Num As Integer
    

    '*************************
    'Frequency
    Freq_Option As schd_Frequency
    
    '##Frequency Once
    Freq_Once_Time As Date
    Freq_Once_Hr As Integer
    Freq_Once_Min As Integer
    Freq_Once_AMPM As schd_ampm
    
    
    '##Frequency Every
    Freq_Every_Interval As Integer
    Freq_Every_Interval_HrMin As schd_FreqInterval
    
    Freq_Every_StartTime As Date
    Freq_Every_StartHr As Integer
    Freq_Every_StartMin As Integer
    Freq_Every_Startampm As schd_ampm
    
    Freq_Every_EndTime As Date
    Freq_Every_EndHr As Integer
    Freq_Every_EndMin As Integer
    Freq_Every_Endampm As schd_ampm
    
    
    Freq_BuildStartTime As Date
    Freq_BuildFinished As Boolean
    
    '*************************
    'Duration
    Duration_StartDate As Date
    Duration_Start_Month As Integer
    Duration_Start_Day As Integer
    Duration_Start_Year As Integer
    
    Duration_End_Option As Integer
    
    Duration_End_Date As Date
    Duration_End_Month As Integer
    Duration_End_Day As Integer
    Duration_End_Year As Integer
End Type
Private TScheduleData As ScheduleData

Public Enum schd_Occurs
    schd_Daily = 0
    Schd_Weekly = 1
    Schd_Monthly = 2
End Enum

Public Enum schd_OccursMonthly
    schd_OccursMonthlyDay = 0
    schd_OccursMonthlyEvery = 1
End Enum

Public Enum schd_MonthlyWeek
    schdWeek_1st = 0
    schdWeek_2nd = 1
    schdWeek_3rd = 2
    schdWeek_4th = 3
    schdWeek_Last = 4
End Enum

Public Enum schd_DayofWeek
    ' Declare constants for the week days
    schd_Sunday = 1
    schd_Monday = 2
    schd_Tuesday = 3
    schd_Wednesday = 4
    schd_Thursday = 5
    schd_Friday = 6
    schd_Saturday = 7
    schd_Day = 8
    schd_Weekday = 9
    schd_Weekend = 10
End Enum

Public Enum schd_Frequency
    schdFreq_Once = 0
    SchdFreq_Interval = 1
End Enum

Public Enum schd_ampm
    schd_am = 0
    schd_pm = 1
End Enum

Public Enum schd_FreqInterval
    schdFreq_Hour = 0
    SchdFreq_Min = 1
End Enum
Public Function GetSchNextAction(Optional xmlSchedule As IXMLDOMElement) As Date
'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: This is the only public Logic function. It starts from here
'             and calls the other functions when needed.
'      Notes:
' Parameters:
'    Returns: Date of the Next Time for the Schedule.
'---------------------------------------------------------------------
'Revision History
'Date       Author  Change
'03/01/2002 Todd    Initial Design
'03/23/2002 Todd    Added option to calculate based on current time. So
'                   not keep rescheduling every time called.
'=====================================================================
    
    'Set the Module level variable
    If Not xmlSchedule Is Nothing Then
        Set mxmlSchedule = xmlSchedule
        TScheduleData = ReadXML(mxmlSchedule)
    End If
    
    With TScheduleData
        'Check if there is an end Duration date
        If .Duration_End_Option <> 1 Then
            '**We need to add time to the last day so it is 11:59
            If .NextActionDateTime > .Duration_End_Date + TimeSerial(23, 59, 0) Then
                'Set Next Action Time to less than Now so they know we are finished
                .NextActionDateTime = Now - TimeSerial(1, 0, 0)
                GoTo GetSchNextAction_Exit
            End If
        End If

        'Check if the Date is less than the start Date
        If .Duration_StartDate > .NextActionDateTime Then
            .NextActionDateTime = .Duration_StartDate + .Freq_BuildStartTime
        Else
            'Get the Next starting time.
            '**Pluse if we are finished the frequency(StartTime to EndTime)
            '  The will also be set in GetStartingTime. If we are done then
            '  we can find the next Occurs(Day,Week, Month)
            Call GetStartingTime
    
            'If we are finished the frequency, that add the next Occurs
            '***If you do NOT want to calculate into the future then
            'turn this condition on and it will ONLY calculate based on the
            'current Time and not keep adding. Uncomment the If and see!!
            'If .NextActionDateTime < Now Then
                If .Freq_BuildFinished Then
                    Select Case .Occurrence
                        Case schd_Daily
                            'Just add Day(s)
                            .NextActionDateTime = DateAdd("d", .Occurs_DailyNum, .NextActionDateTime)
                        Case Schd_Weekly
                            'Add Week(s) to the Next Action
                            Call AddWeek
                        Case Schd_Monthly
                            'Add Months(s) to the Next Action
                            Call AddMonth
                    End Select
                End If
            'End If
        End If

        'Check to see if the NextActionDateTime is over the Duration limit
        If .Duration_End_Option <> 1 Then
            '**We need to add time to the last day so it is 11:59
            If .NextActionDateTime > .Duration_End_Date + TimeSerial(23, 59, 0) Then
                'Set Next Action Time to less than Now so they know we are finished
                .NextActionDateTime = Now - TimeSerial(1, 0, 0)
                GoTo GetSchNextAction_Exit
            End If
        End If
    End With

GetSchNextAction_Exit:

    GetSchNextAction = TScheduleData.NextActionDateTime
    
End Function

Private Sub AddWeek()
'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: Adds x Weeks to the NextActionDateTime
'      Notes: This is not so easy as it may seem. Before we add a week
'             we need to first check if there are any days left for the
'             current week. If there are none, then get the first day
'             of x week(s) where x=Occurs_EveryWeek.
' Parameters:
'    Returns: **Sets the class property-NextActionDateTime.
'---------------------------------------------------------------------
'Revision History
'Date       Author  Change
'03/01/2002 Todd   Initial Design
'=====================================================================

Const nFinishedWeek = -100
Dim nDay As Integer
Dim nNextCheckedDayDiff As Integer
Dim nCurrentDay As Integer

With TScheduleData
    
    'Get the Current Day of the week
    nCurrentDay = Weekday(.NextActionDateTime)
    
    '**We need to set the NextCheckedDay to -1, so that after
    'we loop thru the days, if it is still -1 then we know
    'that we are finished for the current week
    nNextCheckedDayDiff = nFinishedWeek
    'Check if there any days left in the current week
    For nDay = (nCurrentDay + 1) To schd_Saturday
        If .Occurs_WeeklyWeekday(nDay) = 1 Then
             nNextCheckedDayDiff = (nDay - nCurrentDay)
             .NextActionDateTime = DateAdd("d", nNextCheckedDayDiff, .NextActionDateTime)
             Exit For
        End If
    Next
 
    'If we are Finished for the Week, than use the
    'first day of the week selected, and add the
    'week number of weeks to the Next Action Time.
    If nNextCheckedDayDiff = nFinishedWeek Then
         'Now we need to get the first day checked
         For nDay = schd_Sunday To schd_Saturday
            If .Occurs_WeeklyWeekday(nDay) = 1 Then
                '****IMPORTANT****
                'We are going back in time here, so we take
                'the difference between days and force the
                'to a negative so is correct in DateAdd function.
                nNextCheckedDayDiff = (nCurrentDay - nDay) * -1
         
                '********We need to add the week(s) because we are Finished for the
                '        the current week time period.
                'First set the day of the week. Above we have set if the diference
                'is positive or negative.
                .NextActionDateTime = DateAdd("d", nNextCheckedDayDiff, .NextActionDateTime)
                'Add the Weeks
                .NextActionDateTime = DateAdd("ww", .Occurs_WeeklyNum, .NextActionDateTime)
                Exit For
            End If
        Next
    End If

End With

End Sub


Private Sub AddMonth()
'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: Adds x Months to the NextActionDateTime.
'      Notes: x = Occurs_MonthlyEvery from the schedule class
'             **Also it is important to know that to add a Month we
'               first change the date to the first of the Month. This
'               might seem strange, but since we are only adding months
'               the current date is of no concern to us because we it
'               will change.
' Parameters:
'    Returns: **Sets the class property-NextActionDateTime.
'---------------------------------------------------------------------
'Revision History
'Date       Author  Change
'03/01/2002 Todd    Initial Design
'=====================================================================
Dim nDaysDiff As Integer
    
With TScheduleData
    'Make the Date the First on the Month
    .NextActionDateTime = (.NextActionDateTime - Day(.NextActionDateTime)) + 1
    
       
    'check if the item occur on a single day of the Month
    If .Occurs_MonthlyOption = schd_OccursMonthlyDay Then
        'Add the Number of Month(s), then figure out the Day after. We need to do this
        'because the day of the Month depends on which month we want.
        .NextActionDateTime = DateAdd("m", .Occurs_Monthly_Each_Num, .NextActionDateTime)
    
        'Yes-So just set the Next Date
        'The actual date of the month = Occurs_MonthlyDays
        nDaysDiff = .Occurs_Monthly_Each_Day - Day(.NextActionDateTime)
        .NextActionDateTime = DateAdd("d", nDaysDiff, .NextActionDateTime)
    Else
        'Add the Number of Month(s), then figure out the Day after. We need to do this
        'because the day of the Month depends on which month we want.
        .NextActionDateTime = DateAdd("m", .Occurs_Monthly_Every_Num, .NextActionDateTime)
    
        'Not so easy here. We now need to calculate the day of the month
        Call GetDayofMonth
    End If
End With

End Sub

Private Sub GetStartingTime()
'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose:
'      Notes:
' Parameters:
'    Returns:
'---------------------------------------------------------------------
'Revision History
'Date       Author  Change
'03/01/2002 Todd    Initial Design
'=====================================================================

Dim dtNextDateOnly As Date
Dim dtFreqStart As Date
Dim dtFreqEnd As Date
Dim dtOnceStartTime As Date
Dim dtNewNextTime As Date

With TScheduleData
    'This Takes out the Time and leaves ONLY the Date
    dtNextDateOnly = CVDate(Format(.NextActionDateTime, "Short Date"))
    
    'Default The FinishedFrequency to False
    .Freq_BuildFinished = False
    
    'If we ONLY run once a Frequency then we are done
    'for Today and get the next starting time
    If .Freq_Option = schdFreq_Once Then
        .Freq_BuildFinished = True
        
        If .Freq_Once_AMPM = schd_am Then
            dtOnceStartTime = TimeSerial(.Freq_Once_Hr, .Freq_Once_Min, 0)
        Else
            'PM so all 12 Hours
            dtOnceStartTime = TimeSerial(.Freq_Once_Hr + 12, .Freq_Once_Min, 0)
        End If
        
        .NextActionDateTime = dtOnceStartTime + dtNextDateOnly
    Else
        'A:No. Q:Is the frequency by Hour or Min.?
        If .Freq_Every_Interval_HrMin = schdFreq_Hour Then
            'A:By Hour. Add an hour(s) to the last Activity Time
            dtNewNextTime = DateAdd("h", .Freq_Every_Interval, .NextActionDateTime)
        Else
            'A:By Min.  Add an mins. to the last Activity Time
            dtNewNextTime = DateAdd("n", .Freq_Every_Interval, .NextActionDateTime)
        End If

        'We are finished for today and get the next starting
        'Time from the .Freq_IntervalStart value
        If .Freq_Every_Startampm = schd_am Then
            dtFreqStart = TimeSerial(.Freq_Every_StartHr, .Freq_Every_StartMin, 0)
        Else
            dtFreqStart = TimeSerial(.Freq_Every_StartHr + 12, .Freq_Every_StartMin, 0)
        End If
        
        If .Freq_Every_Endampm = schd_am Then
            dtFreqEnd = TimeSerial(.Freq_Every_EndHr, .Freq_Every_EndMin, 0)
        Else
            dtFreqEnd = TimeSerial(.Freq_Every_EndHr + 12, .Freq_Every_EndMin, 0)
        End If
        
        If dtNewNextTime >= (dtFreqEnd + dtNextDateOnly) Then
            .Freq_BuildFinished = True
            dtNewNextTime = dtFreqStart + dtNextDateOnly
        End If
        
        .NextActionDateTime = dtNewNextTime
    End If

    
End With

    

End Sub

Private Sub GetDayofMonth()
'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: To get the day of the Month. Rember the NextActionDateTime
'             is going to be the 1st of the Month when it is passed here.
'             We need to get the 1st, 2nd, 3rd 4th or Last week + Sun~Sat,
'             Day, Weekday or Weekend day.
'
'      Notes: It does not look like a lot of code but it does a lot so
'             be very careful with changes. There was hundreds line of
'             code on the first try & now its way down.
' Parameters:
'    Returns: **Sets the class property-NextActionDateTime.
'---------------------------------------------------------------------
'Revision History
'Date       Author  Change
'03/01/2002 Todd    Initial Design
'03/15/2002 Todd    Made code much more simple.
'=====================================================================
Dim nDaysDiff As Integer

With TScheduleData

    'Select the Basic Day of the Month
    Select Case .Occurs_Monthly_Every_Week
        Case schdWeek_1st 'First
            'Already set since the next Action
            'DateTime starts on the 1st of the month
        Case schdWeek_Last 'Last
            'Get the last day of the Month by gettint the first
            'day of the Month for next Month and then going back
            'a day - quick and easy!!!
            .NextActionDateTime = DateAdd("m", 1, .NextActionDateTime) - TimeSerial(24, 0, 0)
            
            '**NOTE: Since we are looking for the last week, If the Last weekday is
            'less than the weekday desired, go back a week which will still be the
            'last week of the month with that weekday.
            If (Weekday(.NextActionDateTime) - 1) < .Occurs_Monthly_Every_WeekDay Then
              .NextActionDateTime = DateAdd("ww", -1, .NextActionDateTime)
            End If
        Case Else '2-4 Weeks
            'just take the first Day and multiply
            'the number of weeks out.
    
            '**NOTE:We need to add a 1 because
            '  schdWeek_2nd = 1 and not 2
            .NextActionDateTime = DateAdd("d", (.Occurs_Monthly_Every_Week * 7), .NextActionDateTime)
    End Select

    'Once we have the Day picked out, now we need only
    'to match the weekday selected Sun-Sat.
    If .Occurs_Monthly_Every_WeekDay <> schd_Day Then
        'Get the difference between days and use the DateAdd function
        'to select the correct Weekday. Even if it is negative, DateAdd
        'will do the proper logic to figure out the correct date

        'If the First day of the Month is after our day, then add a week.
        'we need to do this because if we go back then we will move into
        'the previous month.
        If (Weekday(.NextActionDateTime) - 1) > .Occurs_Monthly_Every_WeekDay Then
          .NextActionDateTime = DateAdd("ww", 1, .NextActionDateTime)
        End If

        '**Remember that our Occurs_MonthlyWeekDay is one less then
        'VB's Weekday - so take out a day
        nDaysDiff = .Occurs_Monthly_Every_WeekDay - (Weekday(.NextActionDateTime) - 1)
        .NextActionDateTime = DateAdd("d", nDaysDiff, .NextActionDateTime)

    End If
End With

End Sub


Public Function ReadXML(xmlSch As IXMLDOMElement) As ScheduleData

'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: Read the Schedule XML into the Type Structure
'      Notes:
' Parameters:
'    Returns: The Schedule based on the XML Elements.
'---------------------------------------------------------------------
'Revision History
'Date       Author  Change
'03/01/2002 Todd    Initial Design
'=====================================================================

Dim tSch As ScheduleData

With tSch
    .LastActionDateTime = CVDate(xmlAttr(xmlSch, "LastActionDateTime", Format(Now, "General Date")))
    .NextActionDateTime = CVDate(xmlAttr(xmlSch, "NextActionDateTime", Format(Now, "General Date")))
    
    If .NextActionDateTime < .LastActionDateTime Then
        .NextActionDateTime = .LastActionDateTime
    End If
    
    '******************************
    'Occurance
    .Occurrence = Val(xmlAttr(xmlSch, "Occurrence"))
    
    '##Daily
    .Occurs_DailyNum = Val(xmlAttr(xmlSch, "Occurs_DailyNum", "1"))
    
    '##Weekly
    .Occurs_WeeklyNum = Val(xmlAttr(xmlSch, "Occurs_WeeklyNum"))
    .Occurs_WeeklyWeekday(1) = Val(xmlAttr(xmlSch, "Sunday"))
    .Occurs_WeeklyWeekday(2) = Val(xmlAttr(xmlSch, "Monday"))
    .Occurs_WeeklyWeekday(3) = Val(xmlAttr(xmlSch, "Tuesday"))
    .Occurs_WeeklyWeekday(4) = Val(xmlAttr(xmlSch, "Wednesday"))
    .Occurs_WeeklyWeekday(5) = Val(xmlAttr(xmlSch, "Thursday"))
    .Occurs_WeeklyWeekday(6) = Val(xmlAttr(xmlSch, "Friday"))
    .Occurs_WeeklyWeekday(7) = Val(xmlAttr(xmlSch, "Saturday"))
    
    '##Monthly
    .Occurs_MonthlyOption = Val(xmlAttr(xmlSch, "Occurs_MonthlyOption"))
    
    'Each
    .Occurs_Monthly_Each_Day = Val(xmlAttr(xmlSch, "Occurs_Monthly_Each_Day"))
    .Occurs_Monthly_Each_Num = Val(xmlAttr(xmlSch, "Occurs_Monthly_Each_Num"))
   
    
    'Every Month
    .Occurs_Monthly_Every_Num = Val(xmlAttr(xmlSch, "Occurs_Monthly_Every_Num"))
    .Occurs_Monthly_Every_Week = Val(xmlAttr(xmlSch, "Occurs_Monthly_Every_Week"))
    .Occurs_Monthly_Every_WeekDay = Val(xmlAttr(xmlSch, "Occurs_Monthly_Every_WeekDay"))
        
    
    '*****************************************
    'Frequency
    .Freq_Every_Interval = Val(xmlAttr(xmlSch, "Freq_Every_Interval"))
    .Freq_Every_Interval_HrMin = Val(xmlAttr(xmlSch, "Freq_Every_Interval_HrMin"))
    
    .Freq_Once_Hr = Val(xmlAttr(xmlSch, "Freq_Once_Hr"))
    .Freq_Once_Min = Val(xmlAttr(xmlSch, "Freq_Once_Min"))
    .Freq_Once_AMPM = Val(xmlAttr(xmlSch, "Freq_Once_AMPM"))
    
    .Freq_Every_StartHr = Val(xmlAttr(xmlSch, "Freq_Every_StartHr"))
    .Freq_Every_StartMin = Val(xmlAttr(xmlSch, "Freq_Every_StartMin"))
    .Freq_Every_Startampm = Val(xmlAttr(xmlSch, "Freq_Every_Startampm"))
    If .Freq_Every_Startampm = schd_am Then
        .Freq_Every_StartTime = TimeSerial(.Freq_Every_StartHr, .Freq_Every_StartMin, 0)
    Else
        .Freq_Every_StartTime = TimeSerial(.Freq_Every_StartHr + 12, .Freq_Every_StartMin, 0)
    End If
    
    'Set the Frequency to Start at
    .Freq_Option = Val(xmlAttr(xmlSch, "Freq_Option"))
    If .Freq_Option = schdFreq_Once Then
        If .Freq_Once_AMPM = schd_am Then
            .Freq_BuildStartTime = TimeSerial(.Freq_Once_Hr, .Freq_Once_Min, 0)
        Else
            .Freq_BuildStartTime = TimeSerial(.Freq_Once_Hr + 12, .Freq_Once_Min, 0)
        End If
        .Freq_Once_Time = .Freq_BuildStartTime
    Else
        .Freq_BuildStartTime = .Freq_Every_StartTime
    End If
    
    .Freq_Every_EndHr = Val(xmlAttr(xmlSch, "Freq_Every_EndHr"))
    .Freq_Every_EndMin = Val(xmlAttr(xmlSch, "Freq_Every_EndMin"))
    .Freq_Every_Endampm = Val(xmlAttr(xmlSch, "Freq_Every_Endampm"))
    If .Freq_Every_Endampm = schd_am Then
        .Freq_Every_EndTime = TimeSerial(.Freq_Every_EndHr, .Freq_Every_EndMin, 0)
    Else
        .Freq_Every_EndTime = TimeSerial(.Freq_Every_EndHr + 12, .Freq_Every_EndMin, 0)
    End If

    
    '*****************
    'Duration
    .Duration_Start_Month = Val(xmlAttr(xmlSch, "Duration_Start_Month"))
    .Duration_Start_Day = Val(xmlAttr(xmlSch, "Duration_Start_Day"))
    .Duration_Start_Year = Val(xmlAttr(xmlSch, "Duration_Start_Year"))
    .Duration_StartDate = DateSerial(.Duration_Start_Year, .Duration_Start_Month, .Duration_Start_Day)
    
    .Duration_End_Option = Val(xmlAttr(xmlSch, "Duration_End_Option"))
    .Duration_End_Month = Val(xmlAttr(xmlSch, "Duration_End_Month"))
    .Duration_End_Day = Val(xmlAttr(xmlSch, "Duration_End_Day"))
    .Duration_End_Year = Val(xmlAttr(xmlSch, "Duration_End_Year"))
    .Duration_End_Date = DateSerial(.Duration_End_Year, .Duration_End_Month, .Duration_End_Day)
       
End With

    ReadXML = tSch
    
End Function

Public Function WriteXML(tSch As ScheduleData) As IXMLDOMElement

'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: Read the the Type Structure into Schedule XML
'      Notes:
' Parameters:
'    Returns: The XML Element representing the Schedule.
'---------------------------------------------------------------------
'Revision History
'Date       Author  Change
'03/01/2002 Todd    Initial Design
'=====================================================================

Dim xmlSch As IXMLDOMElement

    TScheduleData = tSch
    Set xmlSch = LoadXML("<schedule/>")
    
With TScheduleData
        
    'Get the Next Action
    xmlSch.setAttribute "NextActionDateTime", Format(GetSchNextAction(), "General Date")
        
    '******************************
    'Occurance
    xmlSch.setAttribute "Occurrence", .Occurrence
    
    '##Daily
    xmlSch.setAttribute "Occurs_DailyNum", .Occurs_DailyNum
     
    '##Weekly
    xmlSch.setAttribute "Occurs_WeeklyNum", .Occurs_WeeklyNum
    
    xmlSch.setAttribute "Sunday", .Occurs_WeeklyWeekday(1)
    xmlSch.setAttribute "Monday", .Occurs_WeeklyWeekday(2)
    xmlSch.setAttribute "Tuesday", .Occurs_WeeklyWeekday(3)
    xmlSch.setAttribute "Wednesday", .Occurs_WeeklyWeekday(4)
    xmlSch.setAttribute "Thursday", .Occurs_WeeklyWeekday(5)
    xmlSch.setAttribute "Friday", .Occurs_WeeklyWeekday(6)
    xmlSch.setAttribute "Saturday", .Occurs_WeeklyWeekday(7)
    
    
    '##Monthly
    xmlSch.setAttribute "Occurs_MonthlyOption", .Occurs_MonthlyOption
    
    'Each Month
    xmlSch.setAttribute "Occurs_Monthly_Each_Day", .Occurs_Monthly_Each_Day
    xmlSch.setAttribute "Occurs_Monthly_Each_Num", .Occurs_Monthly_Each_Num
        
    'Every Month
    xmlSch.setAttribute "Occurs_Monthly_Every_Num", .Occurs_Monthly_Every_Num
    xmlSch.setAttribute "Occurs_Monthly_Every_Week", .Occurs_Monthly_Every_Week
    xmlSch.setAttribute "Occurs_Monthly_Every_WeekDay", .Occurs_Monthly_Every_WeekDay
    
    '*****************************************
    'Frequency
    xmlSch.setAttribute "Freq_Option", .Freq_Option
    
    '##Once
    xmlSch.setAttribute "Freq_Once_Hr", .Freq_Once_Hr
    xmlSch.setAttribute "Freq_Once_Min", .Freq_Once_Min
    xmlSch.setAttribute "Freq_Once_AMPM", .Freq_Once_AMPM
    
    '##Every
    xmlSch.setAttribute "Freq_Every_Interval", .Freq_Every_Interval
    xmlSch.setAttribute "Freq_Every_Interval_HrMin", .Freq_Every_Interval_HrMin
    
    xmlSch.setAttribute "Freq_Every_StartHr", .Freq_Every_StartHr
    xmlSch.setAttribute "Freq_Every_StartMin", .Freq_Every_StartMin
    xmlSch.setAttribute "Freq_Every_Startampm", .Freq_Every_Startampm
    
    xmlSch.setAttribute "Freq_Every_EndHr", .Freq_Every_EndHr
    xmlSch.setAttribute "Freq_Every_EndMin", .Freq_Every_EndMin
    xmlSch.setAttribute "Freq_Every_Endampm", .Freq_Every_Endampm
    
    
    If .Freq_Every_Endampm = 0 Then
        xmlSch.setAttribute "Freq_Every_StartTime", TimeSerial(.Freq_Every_StartHr, .Freq_Every_StartMin, 0)
    Else
        xmlSch.setAttribute "Freq_Every_StartTime", TimeSerial(.Freq_Every_StartHr + 12, .Freq_Every_StartMin, 0)
    End If
    
    'Set the Start Time
    xmlSch.setAttribute "Freq_BuildStartTime", Format(.Freq_BuildStartTime, "General Date")
    
    
    '*****************
    'Duration
    xmlSch.setAttribute "Duration_Start_Month", .Duration_Start_Month
    xmlSch.setAttribute "Duration_Start_Day", .Duration_Start_Day
    xmlSch.setAttribute "Duration_Start_Year", .Duration_Start_Year
    
    'End Option
    xmlSch.setAttribute "Duration_End_Option", .Duration_End_Option
    
    xmlSch.setAttribute "Duration_End_Month", .Duration_End_Month
    xmlSch.setAttribute "Duration_End_Day", .Duration_End_Day
    xmlSch.setAttribute "Duration_End_Year", .Duration_End_Year
    
End With

    Set WriteXML = xmlSch
    
End Function

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ouSchedule 
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6585
   ScaleHeight     =   4545
   ScaleWidth      =   6585
   Begin VB.Frame frameOccur 
      Caption         =   "Occurs"
      Height          =   1560
      Index           =   0
      Left            =   90
      TabIndex        =   16
      Top             =   90
      Width           =   1305
      Begin VB.OptionButton optOccurs 
         Caption         =   "&Monthly"
         Height          =   315
         Index           =   2
         Left            =   165
         TabIndex        =   19
         Top             =   1080
         Width           =   1080
      End
      Begin VB.OptionButton optOccurs 
         Caption         =   "&Weekly"
         Height          =   315
         Index           =   1
         Left            =   150
         TabIndex        =   18
         Top             =   690
         Width           =   1080
      End
      Begin VB.OptionButton optOccurs 
         Caption         =   "&Daily"
         Height          =   315
         Index           =   0
         Left            =   165
         TabIndex        =   17
         Top             =   285
         Width           =   795
      End
   End
   Begin VB.Frame frameDailyFrequency 
      Caption         =   "Daily Frequency"
      Height          =   1725
      Left            =   90
      TabIndex        =   6
      Top             =   1680
      Width           =   6435
      Begin VB.OptionButton optFreq 
         Caption         =   "Occures Once at:"
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   375
         Width           =   1635
      End
      Begin VB.OptionButton optFreq 
         Caption         =   "Occures Every:"
         Height          =   300
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   945
         Width           =   1530
      End
      Begin VB.ComboBox clstFreqInterval 
         Height          =   315
         Left            =   2700
         TabIndex        =   8
         Text            =   "clstFreqInterval"
         Top             =   960
         Width           =   915
      End
      Begin VB.ComboBox lstFreqEvery 
         CausesValidation=   0   'False
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Text            =   "lstFreqEvery"
         Top             =   960
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpFreqTime 
         Height          =   345
         Left            =   1770
         TabIndex        =   10
         Top             =   360
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         _Version        =   393216
         Format          =   22544386
         CurrentDate     =   36742
      End
      Begin MSComCtl2.DTPicker dtpFreqIntervalStart 
         Height          =   345
         Left            =   4695
         TabIndex        =   12
         Top             =   870
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         _Version        =   393216
         Format          =   22544386
         CurrentDate     =   36742
      End
      Begin MSComCtl2.DTPicker dtpFreqIntervalEnd 
         Height          =   345
         Left            =   4695
         TabIndex        =   13
         Top             =   1260
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         _Version        =   393216
         Format          =   22544386
         CurrentDate     =   36742
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Starting at:"
         Height          =   195
         Index           =   5
         Left            =   3840
         TabIndex        =   15
         Top             =   960
         Width           =   765
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ending at:"
         Height          =   195
         Index           =   4
         Left            =   3885
         TabIndex        =   14
         Top             =   1320
         Width           =   720
      End
   End
   Begin VB.Frame frameDuration 
      Caption         =   "Duration"
      Height          =   990
      Left            =   60
      TabIndex        =   0
      Top             =   3495
      Width           =   6435
      Begin VB.OptionButton opDurationEnd 
         Caption         =   "&End Date:"
         Height          =   225
         Index           =   0
         Left            =   3855
         TabIndex        =   3
         Top             =   345
         Width           =   1065
      End
      Begin VB.OptionButton opDurationEnd 
         Caption         =   "No End D&ate:"
         Height          =   225
         Index           =   1
         Left            =   3855
         TabIndex        =   1
         Top             =   690
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker dtpDurationEnd 
         Height          =   345
         Left            =   4965
         TabIndex        =   2
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   609
         _Version        =   393216
         Format          =   22544385
         CurrentDate     =   36742
      End
      Begin MSComCtl2.DTPicker dtpDurationStart 
         Height          =   345
         Left            =   840
         TabIndex        =   4
         Top             =   300
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   609
         _Version        =   393216
         Format          =   22544384
         CurrentDate     =   36742
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Start Date:"
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   5
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame frameOccurs 
      Caption         =   "Daily"
      Height          =   1560
      Index           =   0
      Left            =   1440
      TabIndex        =   42
      Top             =   90
      Width           =   5115
      Begin VB.TextBox txtNumofDays 
         Height          =   315
         Left            =   825
         TabIndex        =   43
         Text            =   "1"
         Top             =   645
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day(s)"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   45
         Top             =   690
         Width           =   450
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Every"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   44
         Top             =   675
         Width           =   405
      End
   End
   Begin VB.Frame frameOccurs 
      Caption         =   "Weekly"
      Height          =   1560
      Index           =   1
      Left            =   1440
      TabIndex        =   31
      Top             =   90
      Width           =   5115
      Begin VB.CheckBox chkDay 
         Caption         =   "Sat"
         Height          =   225
         Index           =   7
         Left            =   90
         TabIndex        =   39
         Top             =   1110
         Width           =   660
      End
      Begin VB.CheckBox chkDay 
         Caption         =   "Fri"
         Height          =   225
         Index           =   6
         Left            =   3000
         TabIndex        =   38
         Top             =   675
         Width           =   660
      End
      Begin VB.CheckBox chkDay 
         Caption         =   "Thur"
         Height          =   225
         Index           =   5
         Left            =   2280
         TabIndex        =   37
         Top             =   690
         Width           =   660
      End
      Begin VB.CheckBox chkDay 
         Caption         =   "Wed"
         Height          =   225
         Index           =   4
         Left            =   1440
         TabIndex        =   36
         Top             =   690
         Width           =   660
      End
      Begin VB.CheckBox chkDay 
         Caption         =   "Tue"
         Height          =   225
         Index           =   3
         Left            =   780
         TabIndex        =   35
         Top             =   705
         Width           =   660
      End
      Begin VB.CheckBox chkDay 
         Caption         =   "Mon"
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   34
         Top             =   705
         Width           =   660
      End
      Begin VB.CheckBox chkDay 
         Caption         =   "Sun"
         Height          =   225
         Index           =   1
         Left            =   780
         TabIndex        =   33
         Top             =   1110
         Width           =   795
      End
      Begin VB.TextBox txtOccursWeek 
         Height          =   315
         Left            =   570
         TabIndex        =   32
         Text            =   "1"
         Top             =   285
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week(s) on:"
         Height          =   195
         Index           =   3
         Left            =   1110
         TabIndex        =   41
         Top             =   330
         Width           =   870
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Every"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   40
         Top             =   315
         Width           =   405
      End
   End
   Begin VB.Frame frameOccurs 
      Caption         =   "Monthly"
      Height          =   1560
      Index           =   2
      Left            =   1440
      TabIndex        =   20
      Top             =   90
      Width           =   5115
      Begin VB.ComboBox clstMonthlyWeek 
         Height          =   315
         Left            =   705
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   900
         Width           =   825
      End
      Begin VB.ComboBox clstMonthlyWeekDay 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   900
         Width           =   1500
      End
      Begin VB.ComboBox lstMonthlyDays 
         Height          =   315
         Left            =   840
         TabIndex        =   23
         Text            =   "lstMonthlyDays"
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox lstMonthlyEvery 
         Height          =   315
         Left            =   3720
         TabIndex        =   22
         Text            =   "lstMonthlyEvery"
         Top             =   900
         Width           =   615
      End
      Begin VB.TextBox txtNumofMonths 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2025
         TabIndex        =   21
         Text            =   "1"
         Top             =   360
         Width           =   510
      End
      Begin VB.OptionButton optMonthly 
         Caption         =   "The"
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   885
         Width           =   645
      End
      Begin VB.OptionButton optMonthly 
         Caption         =   "Each"
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   750
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "of Every"
         Height          =   195
         Index           =   9
         Left            =   3120
         TabIndex        =   30
         Top             =   960
         Width           =   585
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day Of             Month(s)"
         Height          =   195
         Index           =   7
         Left            =   1485
         TabIndex        =   29
         Top             =   405
         Width           =   1695
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month(s)"
         Height          =   195
         Index           =   8
         Left            =   4365
         TabIndex        =   28
         Top             =   960
         Width           =   615
      End
   End
End
Attribute VB_Name = "ouSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mbSaveValues As Boolean
Private mbUserCreatedName As Boolean
Private mbInitForm As Boolean
Private mCurrentSchedule As ScheduleData
Private Const msMODULE_NAME = "frmDataEntry"

Private mScheduleXML As IXMLDOMElement
Private Sub clstFreqInterval_Change()
Dim nC As Integer
'set value of the interval list box
'set values for hours
If clstFreqInterval.Text = "Hour" Then
    For nC = 0 To 12
        lstFreqEvery.AddItem CStr(nC)
        
    Next
Else
'set values for min
    For nC = 0 To 60
        lstFreqEvery.AddItem CStr(nC)
        
    Next
End If
End Sub








Private Sub opDurationEnd_Click(Index As Integer)

    If Index = 0 Then
        dtpDurationEnd.Enabled = True
    Else
        dtpDurationEnd.Enabled = False
    End If

End Sub

Private Sub optFreq_Click(Index As Integer)

Dim nOnce%, nEvery%

    If Index = 0 Then
        nOnce = True
        nEvery = False
    Else
        nOnce = False
        nEvery = True
    End If

    dtpFreqTime.Enabled = nOnce
    lstFreqEvery.Enabled = nEvery
    clstFreqInterval.Enabled = nEvery
    dtpFreqIntervalStart.Enabled = nEvery
    dtpFreqIntervalEnd.Enabled = nEvery
End Sub





Private Sub optMonthly_Click(Index As Integer)

Dim OpEach As Boolean
Dim OpEvery As Boolean


    If Index = 0 Then
        OpEach = True
        OpEvery = False
    Else
        OpEach = False
        OpEvery = True
    End If
    
    'Set the Controls
    lstMonthlyDays.Enabled = OpEach
    lstMonthlyDays.Enabled = OpEach
    
  
    clstMonthlyWeek.Enabled = OpEvery
    clstMonthlyWeekDay.Enabled = OpEvery
    lstMonthlyEvery.Enabled = OpEvery

End Sub

Private Sub optOccurs_Click(Index As Integer)

'Show the Proper Frame
frameOccurs(Index).ZOrder 0


End Sub




Private Sub txtNumofDays_KeyPress(KeyAscii As Integer)

'ONLY allow Numbers & operators
If KeyAscii > 25 Then
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End If

End Sub


Private Sub InitControls()

'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: Set the default values for controls.
'      Notes:
' Parameters:
'    Returns:
'---------------------------------------------------------------------
'Revision History
'Date       Author  Change
'03/01/2002 Todd    Initial Design
'=====================================================================

On Error Resume Next

Dim nIndex As Integer
Dim nDays As Integer
Dim nMonth As Integer

    'Clear the Text
    lstFreqEvery.Text = ""
    clstFreqInterval.Text = ""
    lstMonthlyDays.Text = ""
    lstMonthlyEvery.Text = ""
    
    
    clstMonthlyWeek.AddItem "1st"
    clstMonthlyWeek.AddItem "2nd"
    clstMonthlyWeek.AddItem "3rd"
    clstMonthlyWeek.AddItem "4th"
    clstMonthlyWeek.AddItem "Last"
    
    clstMonthlyWeekDay.AddItem "Sunday"
    clstMonthlyWeekDay.AddItem "Monday"
    clstMonthlyWeekDay.AddItem "Tuesday"
    clstMonthlyWeekDay.AddItem "Wednesday"
    clstMonthlyWeekDay.AddItem "Thursday"
    clstMonthlyWeekDay.AddItem "Friday"
    clstMonthlyWeekDay.AddItem "Saturday"
    clstMonthlyWeekDay.AddItem "Day"
    clstMonthlyWeekDay.AddItem "Weekday"
    clstMonthlyWeekDay.AddItem "Weekend"
    
    clstFreqInterval.AddItem "Hour"
    clstFreqInterval.AddItem "Min"
    For nIndex = 1 To 12
        lstFreqEvery.AddItem CStr(nIndex)
        lstMonthlyEvery.AddItem CStr(nIndex)
    Next
    
    For nDays = 1 To 28
        lstMonthlyDays.AddItem CStr(nDays)
    Next
    
    
    'Set the Defaults
    optOccurs(0).Value = True
    optFreq(0).Value = True
    opDurationEnd(1).Value = True
    optMonthly(0).Value = True
    
    dtpDurationStart.Value = Now
    dtpDurationEnd.Value = Now + 1
    
End Sub




Private Sub Form_Write()

'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: Just reads the user inputed values from the form and puts
'             then in the class.
'      Notes:
' Parameters:
'    Returns:
'---------------------------------------------------------------------
'Revision History
'Date       Author  Change
'03/01/2002 Todd    Initial Design
'03/15/2002 Todd    Split Time & Date Segments for upcomming Web Version.
'=====================================================================

Dim nDay As schd_DayofWeek

    'Step One Set the Occurs
    '******************
    'Daily Schedult
    If optOccurs(schd_Daily).Value Then
        'Set the Occurs Option as Daily
        mCurrentSchedule.Occurrence = schd_Daily
        
        'Save the # of Days
        mCurrentSchedule.Occurs_DailyNum = Val(txtNumofDays.Text)
    End If
    
    '******************
    'Weekly Schedule
    If optOccurs(Schd_Weekly).Value Then
        'Set the Occurs Option as Weekly
        mCurrentSchedule.Occurrence = Schd_Weekly
        
        'Save: Every x Week(s)
        mCurrentSchedule.Occurs_WeeklyNum = Val(txtOccursWeek.Text)
        
        'Save Each Day it Occurs
        For nDay = schd_Sunday To schd_Saturday
            mCurrentSchedule.Occurs_WeeklyWeekday(nDay) = Abs(chkDay(nDay).Value)
        Next
    End If
    
    '******************
    'Mothly Schedule
    If optOccurs(Schd_Monthly).Value Then
        'Set the Occurs Option as Monthly
        mCurrentSchedule.Occurrence = Schd_Monthly
        
        'Set if it is scheduled Ever Day or Every Week
        If optMonthly(0).Value = True Then
            mCurrentSchedule.Occurs_MonthlyOption = schd_OccursMonthlyDay
            mCurrentSchedule.Occurs_Monthly_Each_Day = Val(lstMonthlyDays.Text)
            mCurrentSchedule.Occurs_Monthly_Each_Num = Val(txtNumofMonths)
        Else
            mCurrentSchedule.Occurs_MonthlyOption = schd_OccursMonthlyEvery
            
            'Save the User Options
            mCurrentSchedule.Occurs_Monthly_Every_Week = clstMonthlyWeek.ListIndex
            mCurrentSchedule.Occurs_Monthly_Every_WeekDay = clstMonthlyWeekDay.ListIndex
            mCurrentSchedule.Occurs_Monthly_Every_Num = lstMonthlyEvery.Text
        End If
    End If
    
    
    '******************
    'Frequency
    If optFreq(schdFreq_Once).Value = True Then
        mCurrentSchedule.Freq_Option = schdFreq_Once
        
        'Since the user has selected to run this only once
        'get the time of day
        mCurrentSchedule.Freq_Once_Time = dtpFreqTime.Value
        
        'Build out the Time for the HTML Interface
        mCurrentSchedule.Freq_Once_Min = dtpFreqTime.Minute
        If dtpFreqTime.Hour > 12 Then
            mCurrentSchedule.Freq_Once_AMPM = schd_pm
            mCurrentSchedule.Freq_Once_Hr = dtpFreqTime.Hour - 12
        Else
            mCurrentSchedule.Freq_Once_AMPM = schd_am
            mCurrentSchedule.Freq_Once_Hr = dtpFreqTime.Hour
        End If
    Else
        mCurrentSchedule.Freq_Option = SchdFreq_Interval
        
        '##Interval
        mCurrentSchedule.Freq_Every_Interval = Val(lstFreqEvery.Text)
        mCurrentSchedule.Freq_Every_Interval_HrMin = clstFreqInterval.ListIndex
        
        '##Start Time
        mCurrentSchedule.Freq_Every_StartTime = dtpFreqIntervalStart.Value
        
        'Set the HTML Values
        mCurrentSchedule.Freq_Every_StartMin = dtpFreqIntervalStart.Minute
        If dtpFreqIntervalStart.Hour > 12 Then
            mCurrentSchedule.Freq_Every_Startampm = schd_pm
            mCurrentSchedule.Freq_Every_StartHr = dtpFreqIntervalStart.Hour - 12
        Else
            mCurrentSchedule.Freq_Every_Startampm = schd_am
            mCurrentSchedule.Freq_Every_StartHr = dtpFreqIntervalStart.Hour
        End If
        
        
        '##End Time
        mCurrentSchedule.Freq_Every_EndTime = dtpFreqIntervalEnd.Value
        mCurrentSchedule.Freq_Every_EndMin = dtpFreqIntervalEnd.Minute
        If dtpFreqIntervalEnd.Hour > 12 Then
            mCurrentSchedule.Freq_Every_Endampm = schd_pm
            mCurrentSchedule.Freq_Every_EndHr = dtpFreqIntervalEnd.Hour - 12
        Else
            mCurrentSchedule.Freq_Every_Endampm = schd_am
            mCurrentSchedule.Freq_Every_EndHr = dtpFreqIntervalEnd.Hour
        End If
    End If
       
    
    '******************
    'Duration
    '##Start
    mCurrentSchedule.Duration_StartDate = dtpDurationStart.Value
    
    'Build out for the HTML
    mCurrentSchedule.Duration_Start_Day = dtpDurationStart.Day
    mCurrentSchedule.Duration_Start_Month = dtpDurationStart.Month
    mCurrentSchedule.Duration_Start_Year = dtpDurationStart.Year
    
    '##End
    If opDurationEnd(1).Value Then
        mCurrentSchedule.Duration_End_Option = 1
    Else
        mCurrentSchedule.Duration_End_Option = False
    End If
    mCurrentSchedule.Duration_End_Date = dtpDurationEnd.Value
    'Build out for the HTML
    mCurrentSchedule.Duration_End_Day = dtpDurationEnd.Day
    mCurrentSchedule.Duration_End_Month = dtpDurationEnd.Month
    mCurrentSchedule.Duration_End_Year = dtpDurationEnd.Year
     
End Sub

Private Sub Form_Read()

'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: Just shows the values of the class on the form.
'      Notes:
' Parameters:
'    Returns:
'---------------------------------------------------------------------
'Revision History
'Date       Author  Change
'03/01/2002  Todd    Initial Design
'=====================================================================

Dim nDay As schd_DayofWeek
      
    'Step One Set the Occurs
    optOccurs(mCurrentSchedule.Occurrence).Value = True

    '###################################################
    'IMPORTANT - we want to set all of these values
    'even if they are not selected, because they put
    'in the defaults
    '###################################################
    
    '******************
    'Daily Schedult
    txtNumofDays = mCurrentSchedule.Occurs_DailyNum
    
    '******************
    'Weekly Schedule
    txtOccursWeek = mCurrentSchedule.Occurs_WeeklyNum
        
    For nDay = schd_Sunday To schd_Saturday
        chkDay(nDay).Value = mCurrentSchedule.Occurs_WeeklyWeekday(nDay)
    Next
    
    '******************
    'Mothly Schedule
    If mCurrentSchedule.Occurs_MonthlyOption = schd_OccursMonthlyDay Then
        optMonthly(0).Value = True
        lstMonthlyDays.Text = mCurrentSchedule.Occurs_Monthly_Each_Day
        txtNumofMonths.Text = mCurrentSchedule.Occurs_Monthly_Each_Num
    Else
        optMonthly(1).Value = True
        clstMonthlyWeek.ListIndex = mCurrentSchedule.Occurs_Monthly_Every_Week
        clstMonthlyWeekDay.ListIndex = mCurrentSchedule.Occurs_Monthly_Every_WeekDay
        lstMonthlyEvery.Text = mCurrentSchedule.Occurs_Monthly_Every_Num
    End If
    
        
    '******************
    'Daily Frequency
    optFreq(mCurrentSchedule.Freq_Option).Value = True
    
    dtpFreqTime.Value = mCurrentSchedule.Freq_Once_Time
    
    lstFreqEvery.Text = mCurrentSchedule.Freq_Every_Interval
    clstFreqInterval.ListIndex = mCurrentSchedule.Freq_Every_Interval_HrMin
    dtpFreqIntervalStart.Value = mCurrentSchedule.Freq_Every_StartTime
    dtpFreqIntervalEnd.Value = mCurrentSchedule.Freq_Every_EndTime
    
    '******************
    'Duration
    dtpDurationStart.Value = mCurrentSchedule.Duration_StartDate
    dtpDurationEnd.Value = mCurrentSchedule.Duration_End_Date
    If mCurrentSchedule.Duration_End_Option Then
        opDurationEnd(1).Value = True
    Else
        opDurationEnd(0).Value = True
    End If

End Sub

Private Sub txtNumofMonths_KeyPress(KeyAscii As Integer)

'ONLY allow Numbers & operators
If KeyAscii > 25 Then
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End If

End Sub


Private Sub txtOccursWeek_KeyPress(KeyAscii As Integer)

'ONLY allow Numbers & operators
If KeyAscii > 25 Then
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End If

End Sub




Private Sub UserControl_Initialize()
    
    InitControls
    
End Sub





Public Sub ScheduleXML_Open(xmlNewValue As IXMLDOMElement)

'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: Converts the Schedule XML to out Type Structure and
'             reads it into the Controls
'      Notes:
' Parameters:
'    Returns:
'---------------------------------------------------------------------
'Revision History
'Date       Author  Change
'03/01/2002 Todd    Initial Design
'=====================================================================

     Set mScheduleXML = xmlNewValue
     
     'Convert the XML to out Table Structure
     mCurrentSchedule = ReadXML(mScheduleXML)
     
     'Load the Form
     Form_Read
     
End Sub

Public Function ScheduleXML_Save() As IXMLDOMElement

'=====================================================================
'OfficeUtilities.com
'---------------------------------------------------------------------
'    Purpose: Converts the Schedule XML to out Type Structure and
'             reads it into the Controls
'      Notes:
' Parameters:
'    Returns:
'---------------------------------------------------------------------
'Revision History
'Date       Author  Change
'03/01/2002 Todd    Initial Design
'=====================================================================

    Form_Write
    Set ScheduleXML_Save = WriteXML(mCurrentSchedule)
    
End Function

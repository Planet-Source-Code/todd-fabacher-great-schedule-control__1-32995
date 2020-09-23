VERSION 5.00
Begin VB.Form frmSchTest 
   Caption         =   "OfficeUtilities.com Schedule Control"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin ScheduleTest.ouSchedule ouSchedule 
      Height          =   4695
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8281
   End
   Begin VB.ListBox lstFuture 
      Height          =   1425
      Left            =   3000
      TabIndex        =   6
      Top             =   6360
      Width           =   3495
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Next Time ->"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtNextSchedule 
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Top             =   5520
      Width           =   2175
   End
   Begin VB.TextBox txtDateTime 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calculated another 15 time to show the progress"
      Height          =   195
      Index           =   3
      Left            =   3000
      TabIndex        =   7
      Top             =   6120
      Width           =   3405
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next Scheuled Time:"
      Height          =   195
      Index           =   2
      Left            =   4320
      TabIndex        =   3
      Top             =   5280
      Width           =   1485
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valid Date/Time"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   1155
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set a Sechedule and calculate the next time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5400
   End
End
Attribute VB_Name = "frmSchTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculate_Click()

Dim xmlSchedule As IXMLDOMElement
Dim sNextSchedule As String
Dim iFuture As Integer

  sNextSchedule = txtDateTime.Text
  
  If IsDate(sNextSchedule) Then
    'Build the Schedule XML from the Screen
    Set xmlSchedule = ouSchedule.ScheduleXML_Save()
    
    'Set the "LastActionDateTime" Time from the Form
    xmlSchedule.setAttribute "LastActionDateTime", sNextSchedule
  
    'Just The Next Scheduled Time
    sNextSchedule = Format(GetSchNextAction(xmlSchedule), "General Date")
    txtNextSchedule.Text = sNextSchedule
    
    'Shoe the Next 15 Future Scheduled Times
    lstFuture.Clear
    For iFuture = 1 To 15
      xmlSchedule.setAttribute "LastActionDateTime", sNextSchedule
      
      sNextSchedule = Format(GetSchNextAction(xmlSchedule), "General Date")
      lstFuture.AddItem sNextSchedule
    Next
  End If
  
  
End Sub


Private Sub Form_Load()

  txtDateTime.Text = Format(Now(), "General Date")
  
End Sub



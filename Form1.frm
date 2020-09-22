VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SUS Client Checker"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0442
   ScaleHeight     =   7035
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt2ndOctet 
      Height          =   285
      Left            =   780
      TabIndex        =   1
      Text            =   "168"
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox txt1stOctet 
      Height          =   285
      Left            =   300
      TabIndex        =   0
      Text            =   "192"
      Top             =   3480
      Width           =   375
   End
   Begin VB.ListBox lstResults 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   5940
      Left            =   3120
      TabIndex        =   6
      ToolTipText     =   "Click an IP address to see more details"
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton cmdPING 
      Caption         =   "&Start Check"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox txt_EndAddr 
      Height          =   285
      Left            =   2460
      TabIndex        =   4
      Text            =   "255"
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox txt_StartAddr 
      Height          =   285
      Left            =   1740
      TabIndex        =   3
      Text            =   "1"
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox txt_Subnet 
      Height          =   285
      Left            =   1260
      TabIndex        =   2
      Text            =   "1"
      Top             =   3480
      Width           =   375
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4200
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   3000
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   3000
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP Scan Range"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "alessandro.panzetta@email.it"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   405
      MouseIcon       =   "Form1.frx":0884
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   "Click to send me an email"
      Top             =   6720
      Width           =   2145
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To see the SUS Client settings for the discovered hosts, click its IP address in the list."
      ForeColor       =   &H80000005&
      Height          =   615
      Index           =   1
      Left            =   200
      TabIndex        =   9
      Top             =   1680
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "This programs scans the given subnet and reports the alive hosts that have the SUS Client enabled."
      ForeColor       =   &H80000005&
      Height          =   615
      Index           =   0
      Left            =   200
      TabIndex        =   8
      Top             =   960
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   0
      Picture         =   "Form1.frx":0CC6
      Top             =   0
      Width           =   6000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   2220
      TabIndex        =   7
      Top             =   3525
      Width           =   180
   End
   Begin VB.Image Image2 
      Height          =   7305
      Left            =   0
      Picture         =   "Form1.frx":1FDE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6060
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPING_Click()
lstResults.Clear
MousePointer = vbHourglass
For i = txt_StartAddr.Text To txt_EndAddr.Text
    Me.Caption = "Checking : " & txt1stOctet.Text & "." & txt2ndOctet.Text & "." & txt_Subnet.Text & "." & i
    EchoIt (txt1stOctet.Text & "." & txt2ndOctet.Text & "." & txt_Subnet.Text & "." & i)
    If InStr(Host_Status, "Online") > 0 Then
        CheckSUS (txt1stOctet.Text & "." & txt2ndOctet.Text & "." & txt_Subnet.Text & "." & i)
    End If
Next
MousePointer = vbDefault
Me.Caption = "Checked " & txt_EndAddr.Text - txt_StartAddr.Text + 1 & " hosts"
End Sub

Private Sub EchoIt(ByVal IP As String, Optional ListPos)
   If Len(IP) < 7 Or InStr(1, IP, ".") = 0 Then Exit Sub
   lblStatus = IP
   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Integer, StatusCode As String
   Call Ping(Trim(IP), ECHO)
   StatusCode = GetStatusCode(ECHO.status)
End Sub
Private Sub CheckSUS(ComputerIP As String)
Set Reg = New RegistryRoutines
Select Case Reg.ReadRemoteRegistryValue(ComputerIP, HKEY_LOCAL_MACHINE, "UseWUServer", "Software\Policies\Microsoft\Windows\WindowsUpdate\AU")
    Case "1"
    lstResults.AddItem ComputerIP
End Select
End Sub

Private Sub Label3_Click()
    Dim lRet As Long
    Dim sText As String
    sText = "mailto:alessandro.panzetta@email.it?subject=SUS Client Checker: Version " & App.Major & "." & App.Minor & "." & App.Revision
    lRet = shellexecute(hwnd, "open", sText, vbNull, vbNull, SW_SHOWNORMAL)
    If lRet >= 0 And lRet <= 32 Then
        MsgBox "Error!! Can't open Your Mail program!"
    End If
End Sub

Private Sub lstResults_Click()
MousePointer = vbHourglass
Set Reg = New RegistryRoutines
AUOptions = Reg.ReadRemoteRegistryValue(lstResults.Text, HKEY_LOCAL_MACHINE, "AUOptions", "Software\Policies\Microsoft\Windows\WindowsUpdate\AU")
RescheduleWaitTime = Reg.ReadRemoteRegistryValue(lstResults.Text, HKEY_LOCAL_MACHINE, "RescheduleWaitTime", "Software\Policies\Microsoft\Windows\WindowsUpdate\AU")
ScheduleInstallDay = Reg.ReadRemoteRegistryValue(lstResults.Text, HKEY_LOCAL_MACHINE, "ScheduleInstallDay", "Software\Policies\Microsoft\Windows\WindowsUpdate\AU")
ScheduleInstallTime = Reg.ReadRemoteRegistryValue(lstResults.Text, HKEY_LOCAL_MACHINE, "ScheduleInstallTime", "Software\Policies\Microsoft\Windows\WindowsUpdate\AU")
SUSServer = Reg.ReadRemoteRegistryValue(lstResults.Text, HKEY_LOCAL_MACHINE, "WUServer", "Software\Policies\Microsoft\Windows\WindowsUpdate")
NoAutoRebootWithLoggedOnUsers = Reg.ReadRemoteRegistryValue(lstResults.Text, HKEY_LOCAL_MACHINE, "NoAutoRebootWithLoggedOnUsers", "Software\Policies\Microsoft\Windows\WindowsUpdate\AU")
Select Case ScheduleInstallDay
    Case "0"
    SDay = "Every day"
    Case "1"
    SDay = "Sunday"
    Case "2"
    SDay = "Monday"
    Case "3"
    SDay = "Saturday"
    Case "4"
    SDay = "Wednesday"
    Case "5"
    SDay = "Thursday"
    Case "6"
    SDay = "Friday"
    Case "7"
    SDay = "Saturday"
End Select

Select Case NoAutoRebootWithLoggedOnUsers
    Case 1
    Reboot = "Yes"
    Case 0
    Reboot = "No"
End Select

Form2.Caption = "[" & lstResults.Text & "] SUS Client settings"
Form2.txt_susserver.Text = SUSServer
Form2.txt_susoptions.Text = AUOptions
Form2.txt_waittime.Text = RescheduleWaitTime
Form2.txt_installtime.Text = ScheduleInstallTime
Form2.txt_installday.Text = SDay
Form2.txt_reboot.Text = Reboot
Form2.Show
MousePointer = vbDefault
End Sub

Private Sub txt_EndAddr_Click()
txt_EndAddr.Text = ""
End Sub

Private Sub txt_StartAddr_Click()
txt_StartAddr.Text = ""
End Sub

Private Sub txt_Subnet_Click()
txt_Subnet.Text = ""
End Sub

Private Sub txt1stOctet_Click()
txt1stOctet.Text = ""
End Sub

Private Sub txt2ndOctet_Click()
txt2ndOctet.Text = ""
End Sub

Public Function TheHost() As String
TheHost = lstResults.Text
End Function


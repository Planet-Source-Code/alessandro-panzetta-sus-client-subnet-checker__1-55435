VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmb_forcereboot 
      ForeColor       =   &H80000001&
      Height          =   315
      ItemData        =   "Form2.frx":0442
      Left            =   6688
      List            =   "Form2.frx":044C
      TabIndex        =   31
      Text            =   "Force Reboot"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txt_reboot 
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   2415
      TabIndex        =   29
      Top             =   2880
      Width           =   375
   End
   Begin VB.ComboBox txt_newinstallday 
      ForeColor       =   &H80000001&
      Height          =   315
      ItemData        =   "Form2.frx":0459
      Left            =   6688
      List            =   "Form2.frx":0475
      TabIndex        =   27
      Text            =   "Run every.."
      Top             =   1960
      Width           =   1575
   End
   Begin VB.TextBox txt_installtime 
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   2415
      TabIndex        =   24
      Top             =   2256
      Width           =   495
   End
   Begin VB.TextBox txt_installday 
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   2415
      TabIndex        =   23
      Top             =   1944
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply new settings"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3623
      TabIndex        =   21
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox txt_newsusserver 
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   6688
      TabIndex        =   15
      Text            =   "SUS Server"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txt_newwaittime 
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   6688
      TabIndex        =   14
      Text            =   "Mins"
      Top             =   1650
      Width           =   495
   End
   Begin VB.TextBox txt_newinstalltime 
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   6688
      TabIndex        =   13
      Text            =   "Time"
      Top             =   2310
      Width           =   495
   End
   Begin VB.ComboBox txt_newsusoptions 
      ForeColor       =   &H80000001&
      Height          =   315
      ItemData        =   "Form2.frx":04C4
      Left            =   6688
      List            =   "Form2.frx":04D1
      TabIndex        =   7
      Text            =   "Choose Option"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txt_susoptions 
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   2415
      TabIndex        =   6
      Top             =   2568
      Width           =   375
   End
   Begin VB.TextBox txt_waittime 
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   2415
      TabIndex        =   5
      Top             =   1632
      Width           =   495
   End
   Begin VB.TextBox txt_susserver 
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   2415
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forced reboot"
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
      Index           =   14
      Left            =   5415
      TabIndex        =   30
      Top             =   3060
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forced reboot"
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
      Index           =   13
      Left            =   1170
      TabIndex        =   28
      Top             =   2925
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scheduled install day"
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
      Index           =   12
      Left            =   4843
      TabIndex        =   26
      Top             =   2025
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scheduled install time"
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
      Left            =   495
      TabIndex        =   25
      Top             =   2295
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scheduled install day"
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
      Index           =   11
      Left            =   570
      TabIndex        =   22
      Top             =   1995
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..:: New Settings ::.."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   240
      Index           =   10
      Left            =   5728
      TabIndex        =   20
      Top             =   960
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      Height          =   2295
      Left            =   4590
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUS Server"
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
      Index           =   9
      Left            =   5683
      TabIndex        =   19
      Top             =   1365
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wait Time"
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
      Index           =   8
      Left            =   5773
      TabIndex        =   18
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUS Options"
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
      Index           =   7
      Left            =   5608
      TabIndex        =   17
      Top             =   2685
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scheduled install time"
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
      Index           =   6
      Left            =   4768
      TabIndex        =   16
      Top             =   2370
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..:: Actual Settings ::.."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   240
      Index           =   5
      Left            =   1375
      TabIndex        =   12
      Top             =   960
      Width           =   2160
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   2295
      Left            =   375
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   1335
      Left            =   150
      Top             =   4680
      Width           =   8895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..:: SUS Options ::.."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   240
      Index           =   4
      Left            =   3675
      TabIndex        =   11
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form2.frx":04DE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Index           =   2
      Left            =   270
      TabIndex        =   10
      Top             =   5520
      Width           =   8535
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "3 - Automatically downloads updates and notify Admin-priv user of pending installation. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Index           =   1
      Left            =   270
      TabIndex        =   9
      Top             =   5160
      Width           =   8535
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2 - Notify Admin-priv user of a pending update waiting to be downloaded. User will initate the download and installation. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   8
      Top             =   4800
      Width           =   8535
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUS Options"
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
      Index           =   3
      Left            =   1335
      TabIndex        =   3
      Top             =   2610
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wait Time"
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
      Index           =   1
      Left            =   1500
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUS Server"
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
      Index           =   0
      Left            =   1410
      TabIndex        =   1
      Top             =   1365
      Width           =   945
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
      Left            =   3525
      MouseIcon       =   "Form2.frx":05CD
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Click to send me an email"
      Top             =   6120
      Width           =   2145
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   1590
      Picture         =   "Form2.frx":0A0F
      Top             =   0
      Width           =   6000
   End
   Begin VB.Image Image2 
      Height          =   6585
      Left            =   0
      Picture         =   "Form2.frx":1D27
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9420
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If txt_newsusserver.Text = "SUS Server" Or txt_newsusserver.Text = "" Then
    MsgBox "Please type a valid SUS Server!", vbOKOnly + vbCritical, "No SUS Server specified"
    Exit Sub
ElseIf InStr(txt_newsusserver.Text, "http://") <> 1 Then
    MsgBox "SUS Server must start with http://", vbOKOnly + vbCritical, "No SUS Server specified"
    txt_newsusserver.Text = ""
    Exit Sub
ElseIf txt_newwaittime.Text = "Mins" Or txt_newwaittime.Text = "" Then
    MsgBox "Please type a valid wait time!", vbOKOnly + vbCritical, "No wait time specified"
    Exit Sub
ElseIf txt_newinstallday.Text = "Run every.." Or txt_newinstallday.Text = "" Then
    MsgBox "Please choose a valid install day!", vbOKOnly + vbCritical, "No install day specified"
    Exit Sub
ElseIf IsNumeric(txt_newinstalltime.Text) = False Then
    MsgBox "Please type a valid install time!", vbOKOnly + vbCritical, "No install time specified"
    Exit Sub
ElseIf IsNumeric(txt_newwaittime.Text) = False Then
    MsgBox "Please type a valid wait time!", vbOKOnly + vbCritical, "No wait time specified"
    Exit Sub
ElseIf txt_newinstalltime.Text = "Time" Or txt_newinstalltime.Text = "" Then
    MsgBox "Please type a valid install time!", vbOKOnly + vbCritical, "No install time specified"
    Exit Sub
ElseIf txt_newsusoptions.Text = "Choose Option" Or txt_newsusoptions.Text = "" Then
    MsgBox "Please choose a valid SUS Option!", vbOKOnly + vbCritical, "No SUS Option specified"
    Exit Sub
ElseIf cmb_forcereboot.Text = "Force Reboot" Or cmb_forcereboot.Text = "" Then
    MsgBox "Please choose a valid force reboot option!", vbOKOnly + vbCritical, "No force reboot specified"
    Exit Sub
Else
    Call ApplySettings
End If
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



Private Sub txt_newinstalltime_Click()
txt_newinstalltime.Text = ""
End Sub

Private Sub txt_newsusserver_Click()
txt_newsusserver.Text = ""
End Sub

Private Sub txt_newwaittime_Click()
txt_newwaittime.Text = ""
End Sub

Sub ApplySettings()
MousePointer = vbHourglass
Set Reg = New RegistryRoutines
HOST = Form1.TheHost

If cmb_forcereboot.Text = "Yes" Then
    forcereboot = "00000001"
ElseIf cmb_forcereboot.Text = "No" Then
    forcereboot = "00000000"
End If

Select Case txt_newinstallday.Text
    Case "Every day"
    SDay = "0"
    Case "Sunday"
    SDay = "1"
    Case "Monday"
    SDay = "2"
    Case "Saturday"
    SDay = "3"
    Case "Wednesday"
    SDay = "4"
    Case "Thursday"
    SDay = "5"
    Case "Friday"
    SDay = "6"
    Case "Saturday"
    SDay = "7"
End Select
AskConfirmation = MsgBox("Do you want to apply the follwogin settings?" & vbCrLf & vbCrLf _
        & "SUS Server: " & vbTab & txt_newsusserver.Text & vbCrLf _
        & "Scheduled Day: " & vbTab & txt_newinstallday.Text & vbCrLf _
        & "Scheduled Time: " & vbTab & txt_newinstalltime.Text & vbCrLf _
        & "Wait time: " & vbTab & txt_newwaittime.Text & vbCrLf _
        & "SUS Options: " & vbTab & txt_newsusoptions.Text & vbCrLf _
        & "Force Reboot: " & vbTab & cmb_forcereboot.Text & vbCrLf _
        , vbYesNo + vbQuestion, "Confirm SUS Client settings?")
        
If AskConfirmation = 6 Then
    Reg.WriteRemoteRegistryValue HOST, HKEY_LOCAL_MACHINE, "AUOptions", txt_newsusoptions.Text, REG_DWORD, "Software\Policies\Microsoft\Windows\WindowsUpdate\AU\"
    Reg.WriteRemoteRegistryValue HOST, HKEY_LOCAL_MACHINE, "RescheduleWaitTime", txt_newwaittime.Text, REG_DWORD, "Software\Policies\Microsoft\Windows\WindowsUpdate\AU\"
    Reg.WriteRemoteRegistryValue HOST, HKEY_LOCAL_MACHINE, "ScheduleInstallDay", SDay, REG_DWORD, "Software\Policies\Microsoft\Windows\WindowsUpdate\AU\"
    Reg.WriteRemoteRegistryValue HOST, HKEY_LOCAL_MACHINE, "ScheduleInstallTime", txt_newinstalltime, REG_DWORD, "Software\Policies\Microsoft\Windows\WindowsUpdate\AU\"
    Reg.WriteRemoteRegistryValue HOST, HKEY_LOCAL_MACHINE, "WUServer", txt_newsusserver.Text, REG_SZ, "Software\Policies\Microsoft\Windows\WindowsUpdate\"
    Reg.WriteRemoteRegistryValue HOST, HKEY_LOCAL_MACHINE, "WUStatusServer", txt_newsusserver.Text, REG_SZ, "Software\Policies\Microsoft\Windows\WindowsUpdate\"
    Reg.WriteRemoteRegistryValue HOST, HKEY_LOCAL_MACHINE, "NoAutoRebootWithLoggedOnUsers", forcereboot, REG_DWORD, "Software\Policies\Microsoft\Windows\WindowsUpdate\AU\"
    MsgBox "Settings applied succesfully on " & HOST, vbOKOnly + vbInformation, HOST & " new settings applied"
End If
MousePointer = vbDefault
End Sub

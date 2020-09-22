VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{1C1CD033-D017-11D2-B467-00A0C9DC0C41}#1.0#0"; "systray.ocx"
Begin VB.Form frmMain 
   Caption         =   "IP Tracker"
   ClientHeight    =   1845
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5250
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   240
      Top             =   720
   End
   Begin SysTrayCtl.cSysTray cSysTray1 
      Left            =   1560
      Top             =   720
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "frmMain.frx":0442
      TrayTip         =   ""
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   840
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Menu mnuMainMenu 
      Caption         =   "Main"
      Begin VB.Menu mnuMain 
         Caption         =   "Refresh"
         Index           =   0
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Change URL"
         Index           =   1
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Change Refresh Rate"
         Index           =   2
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Exit"
         Index           =   3
      End
      Begin VB.Menu mnuMain 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Close Menu"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strURL As String
Dim strTime As String

Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)
    If Button = 2 Then
        PopupMenu mnuMainMenu
    End If
End Sub

Private Sub Form_Load()
    cSysTray1.InTray = True
    strURL = regGetValue("IPTracker", "Setup", "URL", "http://www.berard.cc/getip.php")
    sRefreshIP
    Timer1_Timer
End Sub

Sub sRefreshIP()
    Dim strReturn As String
    strReturn = Inet1.OpenURL(strURL, 0)
    If Len(strReturn) > 15 Then strReturn = "N/A"
    cSysTray1.TrayTip = "Your External IP Address is: " & strReturn
    Inet1.Cancel
    
    If strReturn <> regGetValue("IPTracker", "Setup", "Address", "") Then
        MsgBox "Your external IP address has changed to: " & strReturn, vbInformation, Me.Caption
        regSaveValue "IPTracker", "Setup", "Address", strReturn
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cSysTray1.InTray = False
End Sub

Private Sub mnuMain_Click(Index As Integer)
    Select Case Index
        Case 0
            sRefreshIP
        Case 1
            Dim strTempURL As String
            strTempURL = InputBox("Enter server file to obtain IP from.", Me.Caption, strURL)
            If strTempURL <> "" Then
                strURL = strTempURL
                sRefreshIP
                regSaveValue "IPTracker", "Setup", "URL", strURL
            Else
                MsgBox "URL not changed.", vbInformation, Me.Caption
            End If
        Case 2
            Dim strMinutes As String
            strMinutes = InputBox("Enter refresh rate in minutes.", Me.Caption, regGetValue("IPTracker", "Setup", "RefreshRate", 15))
            If strMinutes <> "" Then
                regSaveValue "IPTracker", "Setup", "RefreshRate", strMinutes
                Timer1_Timer
            Else
                MsgBox "Refresh rate not changed.", vbInformation, Me.Caption
            End If
        Case 3 '// Exit
            Unload Me
    End Select
End Sub

Private Sub Timer1_Timer()
    If strTime = "" Then strTime = Now
    
    If DateDiff("n", strTime, Now) = regGetValue("IPTracker", "Setup", "RefreshRate", 15) Then
        sRefreshIP
        strTime = Now
    End If
End Sub

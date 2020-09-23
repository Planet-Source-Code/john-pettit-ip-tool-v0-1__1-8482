VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "         Internet Protocol Tool 0.1"
   ClientHeight    =   1935
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4350
   BeginProperty Font 
      Name            =   "OCR A Extended"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "DISMISS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdSendIPTwo 
      Caption         =   "INTERNAL > CLIPBOARD"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdSendIPOne 
      Caption         =   "EXTERNAL > CLIPBOARD"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   40
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   0
      TabIndex        =   1
      Top             =   -55
      Width           =   2220
      Begin VB.TextBox txtIPTwo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   3
         ToolTipText     =   "This IP address can be use to access your computer ONLY from within your Local Network."
         Top             =   1080
         Width           =   1950
      End
      Begin VB.TextBox txtIPOne 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   2
         ToolTipText     =   "This is the IP address people on the Internet can use to connect to your computer."
         Top             =   480
         Width           =   1950
      End
      Begin VB.Label lblIPOne 
         Alignment       =   2  'Center
         Caption         =   "EXTERNAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblIPTwo 
         Alignment       =   2  'Center
         Caption         =   "INTERNAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdQueryIP 
      Caption         =   "REFRESH"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "http://webone.com.au/~jpettit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   1780
      Width           =   2055
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      Caption         =   "masteryoda@webone.com.au"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   1640
      Width           =   2055
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      Caption         =   "Master Yoda - 31 May 2000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   1500
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClose_Click()
' Close down Winsock mod
SocketsCleanup

' End program short and simple
End
End Sub

Private Sub cmdQueryIP_Click()

Call UpdateIPs
End Sub


Private Sub cmdSendIPOne_Click()
' Copy External IP to the clipboard
Clipboard.SetText txtIPOne.Text

End Sub

Private Sub cmdSendIPTwo_Click()
' Copy Internal IP to the clipboard
Clipboard.SetText txtIPTwo.Text

End Sub

Private Sub Form_Load()

SocketsInitialize ' Initialize Winsock module

' Lock the textboxes to stop them being changed
txtIPOne.Locked = True
txtIPTwo.Locked = True

Call UpdateIPs

' check that Program has not been run before displaying the form.
If App.PrevInstance = True Then
    Call MsgBox("IP Tool is already running!", vbExclamation)
    End
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
' Close down Winsock mod
SocketsCleanup
End Sub




Public Sub UpdateIPs()
' Fill the textbxes with IP values
If CurrentIP(False) = "" Then
    If CurrentIP(True) = "" Then
        ' No IPs preset
        lblIPOne.Caption = "NO IPs PRESENT"
        lblIPTwo.Caption = ""
        txtIPOne.Text = ""
        txtIPTwo.Text = ""
        txtIPOne.Enabled = False
        txtIPTwo.Enabled = False
        
        cmdSendIPOne.Caption = ""
        cmdSendIPTwo.Caption = ""
        cmdSendIPOne.Visible = False
        cmdSendIPTwo.Visible = False
        
    Else
        ' No Internal IP - The External IP Refers to your LAN
        ' since no Internet connection is preset
        lblIPOne.Caption = "INTERNAL"
        lblIPTwo.Caption = ""
        txtIPOne.Text = CurrentIP(True)
        txtIPTwo.Text = ""
        txtIPOne.Enabled = True
        txtIPTwo.Enabled = False
        
        cmdSendIPOne.Caption = "INTERNAL > CLIPBOARD"
        cmdSendIPTwo.Caption = ""
        cmdSendIPOne.Visible = True
        cmdSendIPTwo.Visible = False
    End If
Else
    If CurrentIP(True) = "" Then
    ' Has External IP, but no Internal IP
    ' I don't think this condition will be met
    ' but it's just a safeguard.
        lblIPOne.Caption = "INTERNAL"
        lblIPTwo.Caption = ""
        txtIPOne.Text = CurrentIP(False)
        txtIPTwo.Text = ""
        txtIPOne.Enabled = True
        txtIPTwo.Enabled = False
        
        cmdSendIPOne.Caption = "INTERNAL > CLIPBOARD"
        cmdSendIPTwo.Caption = ""
        cmdSendIPOne.Visible = True
        cmdSendIPTwo.Visible = False
    Else
    ' Both LAN and External IP detected
        lblIPOne.Caption = "EXTERNAL"
        lblIPTwo.Caption = "INTERNAL"
        txtIPOne.Text = CurrentIP(True)
        txtIPTwo.Text = CurrentIP(False)
        txtIPOne.Enabled = True
        txtIPTwo.Enabled = True
        
        cmdSendIPOne.Caption = "EXTERNAL > CLIPBOARD"
        cmdSendIPTwo.Caption = "INTERNAL > CLIPBOARD"
        cmdSendIPOne.Visible = True
        cmdSendIPTwo.Visible = True
    End If
End If


End Sub

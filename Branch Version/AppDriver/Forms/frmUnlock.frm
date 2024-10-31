VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9.ocx"
Begin VB.Form frmUnlock 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmUnlock.frx":0000
   ScaleHeight     =   2895
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Top             =   1545
      Width           =   2460
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2100
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1890
      Width           =   2460
   End
   Begin xrControl.xrButton xrButton 
      Height          =   375
      Index           =   0
      Left            =   2085
      TabIndex        =   4
      Top             =   2325
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&OK"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8438015
      BackColorDown   =   8438015
      BorderColorFocus=   33023
      BorderColorHover=   8438015
   End
   Begin xrControl.xrButton xrButton 
      Height          =   375
      Index           =   1
      Left            =   3330
      TabIndex        =   5
      Top             =   2325
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8438015
      BackColorDown   =   8438015
      BorderColorFocus=   33023
      BorderColorHover=   8438015
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash2 
      Height          =   1185
      Left            =   165
      TabIndex        =   12
      Top             =   165
      Width           =   1290
      _cx             =   2275
      _cy             =   2090
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rosalyn Lazo Descallar"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   11
      Top             =   2625
      Width           =   1665
   End
   Begin VB.Label lblfield 
      BackStyle       =   0  'Transparent
      Caption         =   "GMC-SEG."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   2325
      TabIndex        =   10
      Top             =   1080
      Width           =   1920
   End
   Begin VB.Label lblfield 
      BackStyle       =   0  'Transparent
      Caption         =   "System Administrator's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   6
      Top             =   885
      Width           =   1920
   End
   Begin VB.Label lblfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   1275
      TabIndex        =   3
      Top             =   1605
      Width           =   720
   End
   Begin VB.Label lblfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   1275
      TabIndex        =   2
      Top             =   1950
      Width           =   690
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unlock User"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1095
      TabIndex        =   8
      Top             =   150
      Width           =   3855
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unlock User"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1035
      TabIndex        =   9
      Top             =   180
      Width           =   4035
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblfield 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmUnlock.frx":1794
      ForeColor       =   &H00000000&
      Height          =   825
      Index           =   2
      Left            =   1680
      TabIndex        =   7
      Top             =   495
      Width           =   2790
   End
End
Attribute VB_Name = "frmUnlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pbCancel As Boolean
Private psAppPath As String
Private poMod As New MainModules

Property Let AppPath(ByVal Value As String)
10       psAppPath = Value
End Property

Property Get Cancel() As Boolean
10       Cancel = pbCancel
End Property

Private Sub Form_Initialize()
10       pbCancel = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10       Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
20          Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
30             poMod.SetNextFocus
40          Case vbKeyUp
50             poMod.SetPreviousFocus
60          End Select
70       End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
10       Set poMod = Nothing
End Sub

Private Sub xrButton_Click(Index As Integer)
10       pbCancel = Index = 1
20       Me.Hide
End Sub

Private Sub Form_Load()
10        ShockwaveFlash1.Movie = psAppPath & "\Images\hand_lock.swf"
20        ShockwaveFlash1.Play
End Sub



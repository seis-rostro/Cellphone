VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProgress 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmProgress.frx":0000
   ScaleHeight     =   2670
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox shpProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H007A3A14&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   0
      Left            =   2625
      ScaleHeight     =   150
      ScaleWidth      =   3885
      TabIndex        =   5
      Top             =   930
      Width           =   3885
   End
   Begin VB.PictureBox shpProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H007A3A14&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   1
      Left            =   2625
      ScaleHeight     =   150
      ScaleWidth      =   3885
      TabIndex        =   4
      Top             =   1545
      Width           =   3885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   3885
      TabIndex        =   1
      Top             =   2130
      Width           =   1125
   End
   Begin MSComCtl2.Animation aniPiston 
      Height          =   975
      Left            =   105
      TabIndex        =   0
      Top             =   645
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1720
      _Version        =   393216
      BackColor       =   1842204
      FullWidth       =   80
      FullHeight      =   65
   End
   Begin VB.Label lblRemarks 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   0
      Left            =   2610
      TabIndex        =   7
      Top             =   1065
      Width           =   3885
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRemarks 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   1
      Left            =   2610
      TabIndex        =   6
      Top             =   1695
      Width           =   3885
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00800000&
      Height          =   180
      Left            =   2610
      Top             =   915
      Width           =   3915
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   180
      Left            =   2610
      Top             =   1530
      Width           =   3915
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2640
      X2              =   6555
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   2640
      X2              =   6510
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Label lblProcess 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Processing..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   180
      Width           =   1815
   End
   Begin VB.Label lblProcess 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Processing..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   2610
      TabIndex        =   3
      Top             =   225
      Width           =   1815
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
                           
' Used to support captionless drag
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const HWND_TOPMOST = -&H1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private p_nPriMaxValue As Long
Private p_nSecMaxValue As Long
Private p_bCancelled As Boolean

Private pnPriInterval As Long
Private pnSecInterval As Long
Private pnPriProgress As Long
Private pnSecProgress As Long

Private pnCtr As Long
Private Const MaxProgress = 3915

Private Sub Form_Load()
10       pnPriProgress = 0
20       pnSecProgress = 0
   
30       shpProgress(0).Width = 0
40       shpProgress(1).Width = 0

50       aniPiston.Open App.Path & "\piston.avi"
60       aniPiston.Play
End Sub

Function MoveProgress()
10       pnSecProgress = pnSecProgress + 1
   
20       shpProgress(0).Width = Fix(pnSecProgress / p_nSecMaxValue * MaxProgress)
30       DoEvents
   
40       If pnSecProgress = p_nSecMaxValue Then
50          If pnPriProgress < p_nPriMaxValue Then
60             pnPriProgress = pnPriProgress + 1
70             DoEvents

80             shpProgress(1).Width = Fix(pnPriProgress / p_nPriMaxValue * MaxProgress)
90             DoEvents
100         End If
110         pnSecProgress = 0
120      End If
   
130      DoEvents
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' Automatically allow user to drag using any portion of form, not just titlebar,
   '  when user depresses left mousebutton. Useful for captionless forms.
10       If Button = vbLeftButton Then
20          DoEvents
30          Call ReleaseCapture
40          DoEvents
50          Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
60          DoEvents
70       End If
End Sub

Property Get SecondaryMaxValue() As Long
10        SecondaryMaxValue = p_nSecMaxValue
End Property

Property Let SecondaryMaxValue(ByVal Value As Long)
10       p_nSecMaxValue = Value
20       pnSecProgress = 0
End Property

Property Get PrimaryMaxValue() As Long
10       PrimaryMaxValue = p_nPriMaxValue
End Property

Property Let PrimaryMaxValue(ByVal Value As Long)
10       p_nPriMaxValue = Value
20       pnPriProgress = 0
End Property

Property Let ProgressStatus(ByVal Value As String)
10       lblProcess(0).Caption = Value
20       lblProcess(1).Caption = Value
End Property

Property Let PrimaryRemarks(ByVal Value As String)
10       lblRemarks(1).Caption = Value
End Property

Property Let SecondaryRemarks(ByVal Value As String)
10       lblRemarks(0).Caption = Value
End Property

Property Get Cancelled() As Boolean
10       Cancelled = p_bCancelled
End Property

Private Sub cmdCancel_Click()
10       p_bCancelled = True
20       Me.Hide
End Sub

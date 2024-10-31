VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4770
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3360
      Width           =   2460
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4770
      TabIndex        =   1
      Top             =   3000
      Width           =   2460
   End
   Begin xrControl.xrButton xrButton 
      Height          =   375
      Index           =   0
      Left            =   4770
      TabIndex        =   6
      Top             =   4515
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
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3915
      TabIndex        =   5
      Top             =   4005
      Width           =   3315
   End
   Begin xrControl.xrButton xrButton 
      Height          =   375
      Index           =   1
      Left            =   6015
      TabIndex        =   7
      Top             =   4515
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warning"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   3090
      Width           =   3450
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTelNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel No: (075) 522 1085; 522 1097"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      TabIndex        =   16
      Top             =   1665
      Width           =   5100
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Guanzon Bldg., Perez Blvd., Dagupan City"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      TabIndex        =   15
      Top             =   1425
      Width           =   5100
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  'Transparent
      Caption         =   "Guanzon Merchandising Corporation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      TabIndex        =   14
      Top             =   1185
      Width           =   5100
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000A&
      BorderColor     =   &H00FFFFFF&
      Height          =   1875
      Left            =   120
      Top             =   3000
      Width           =   3645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © - Rosalyn Lazo Descallar"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   4710
      TabIndex        =   13
      Top             =   5685
      Width           =   2640
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Authentication"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   630
      Index           =   1
      Left            =   3960
      TabIndex        =   12
      Top             =   2205
      Width           =   3240
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLogin.frx":4388
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1260
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   3315
      Width           =   3450
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblOwner 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guanzon Group of Companies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   0
      Left            =   660
      TabIndex        =   10
      Top             =   5415
      Width           =   2565
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      X1              =   120
      X2              =   3720
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "This product is owned by:"
      ForeColor       =   &H000040C0&
      Height          =   675
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   3645
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Authentication"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   630
      Index           =   0
      Left            =   3990
      TabIndex        =   8
      Top             =   2220
      Width           =   3240
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Application"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   3945
      TabIndex        =   4
      Top             =   3795
      Width           =   780
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   3945
      TabIndex        =   2
      Top             =   3405
      Width           =   690
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   3945
      TabIndex        =   0
      Top             =   3030
      Width           =   720
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbCancel As Boolean
Private pbFocus As Boolean
Private poMod As New MainModules

Property Get Cancel() As Boolean
10       Cancel = pbCancel
End Property

Private Sub Combo1_GotFocus()
10       pbFocus = True
End Sub

Private Sub Combo1_LostFocus()
10       pbFocus = False
End Sub

Private Sub Form_Initialize()
10       pbCancel = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10       Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
20          If KeyCode <> vbKeyReturn And pbFocus Then Exit Sub
30          Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
40             poMod.SetNextFocus
50          Case vbKeyUp
60             poMod.SetPreviousFocus
70          End Select
80       End Select
End Sub

Private Sub Form_Terminate()
10       Set poMod = Nothing
End Sub

Private Sub xrButton_Click(Index As Integer)
10       pbCancel = Index = 1
20       Me.Hide
End Sub

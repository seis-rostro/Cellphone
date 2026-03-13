VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmDate 
   BorderStyle     =   0  'None
   Caption         =   "Date Entry"
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "et0;eb0;et0;bc2"
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1890
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   3334
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1335
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   870
         Width           =   2565
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1335
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   525
         Width           =   2565
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         Height          =   195
         Index           =   1
         Left            =   465
         TabIndex        =   2
         Top             =   915
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   0
         Top             =   570
         Width           =   735
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   4440
      TabIndex        =   5
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmDate.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   4440
      TabIndex        =   4
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Ok"
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
      Picture         =   "frmDate.frx":077A
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private p_oSkin As clsFormSkin
Private p_oAppDrivr As clsAppDriver

Private p_dDateFrom As Date

Dim pbCancelled As Boolean

Property Let AppDriver(ByVal oValue As clsAppDriver)
   Set p_oAppDrivr = oValue
End Property

Property Let DateFrom(ByVal Value As Date)
   p_dDateFrom = Value
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancelled
End Property

Property Get DateEntry() As Date
   DateEntry = Format(txtField(1), "MM/DD/YYYY")
End Property

Private Sub cmdButton_Click(Index As Integer)
   pbCancelled = Index = 1
   Me.Hide
End Sub

Private Sub Form_Load()
   Set p_oSkin = New clsFormSkin
   Set p_oSkin.AppDriver = p_oAppDrivr
   Set p_oSkin.Form = Me
   p_oSkin.ApplySkin xeFormTransDetail
   
   pbCancelled = True
   txtField(0) = Format(p_dDateFrom, "MMMM DD, YYYY")
   txtField(1) = Format(p_oAppDrivr.ServerDate, "MMMM DD, YYYY")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set p_oSkin = Nothing
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   txtField(Index) = Format(txtField(Index), "MM/DD/YY")
   If txtField(Index).Text <> "" Then
      txtField(Index).SelStart = 0
      txtField(Index).SelLength = Len(txtField(Index).Text)
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   If Not IsDate(txtField(Index).Text) Then txtField(Index) = p_oAppDrivr.ServerDate
   txtField(Index) = Format(txtField(Index), "MMMM DD, YYYY")
End Sub

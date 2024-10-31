VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmJOPayment 
   BorderStyle     =   0  'None
   Caption         =   "Payment Entry"
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3645
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "et0;eb0;et0;bc2"
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2985
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   5265
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   2160
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   210
         Width           =   1770
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   2160
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   540
         Width           =   1770
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2175
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1785
         Width           =   1770
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2175
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1440
         Width           =   1770
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   2175
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1095
         Width           =   1770
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DATE PAYMENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   0
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SALES INVOICE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   2
         Top             =   585
         Width           =   1440
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1410
         X2              =   2010
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   1410
         X2              =   2010
         Y1              =   1575
         Y2              =   1575
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1425
         X2              =   2025
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "OTHERS"
         Height          =   195
         Index           =   3
         Left            =   30
         TabIndex        =   8
         Top             =   1830
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GRAND TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -45
         TabIndex        =   13
         Top             =   2295
         Width           =   2340
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   2445
         TabIndex        =   12
         Tag             =   "ht0;hb0"
         Top             =   2265
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PARTS"
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   6
         Top             =   1485
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LABOR"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   1140
         Width           =   1245
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   4455
      TabIndex        =   11
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
      Picture         =   "frmJOPayment.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   600
      Index           =   0
      Left            =   4455
      TabIndex        =   10
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
      Picture         =   "frmJOPayment.frx":077A
   End
End
Attribute VB_Name = "frmJOPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private p_oSkin As clsFormSkin
Private p_oAppDrivr As clsAppDriver

Dim pbCancelled As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancelled
End Property

Property Get DatePayment() As Date
   DatePayment = txtField(4).Text
End Property

Property Get SalesInvoice() As String
   SalesInvoice = txtField(3).Text
End Property

Property Get Labor() As Double
   Labor = txtField(0).Text
End Property

Property Get Parts() As Double
   Parts = txtField(1).Text
End Property

Property Get Others() As Double
   Others = txtField(2).Text
End Property

Property Get GrandTotal() As Double
   GrandTotal = lblTotal.Caption
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set p_oSkin = Nothing
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
   Select Case Index
   Case 0
      If Not IsNumeric(txtField(0).Text) Then txtField(0) = 0#
      txtField(0).Text = txtField(0).Text 'Format(txtField(Index).Text, 0#)
      lblTotal.Caption = ""
      lblTotal.Caption = Format(txtField(0) + txtField(1) + txtField(2), 0#)
   Case 3
      txtField(Index).Text = UCase(txtField(Index).Text)
   Case 4
      If Not IsDate(txtField(Index).Text) Then txtField(Index).Text = oApp.ServerDate
      txtField(Index).Text = Format(txtField(Index).Text, "MMM-DD-YYYY")
   End Select
      
End Sub

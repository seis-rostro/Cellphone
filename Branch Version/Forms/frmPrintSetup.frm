VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmPrintSetup 
   BorderStyle     =   0  'None
   Caption         =   "Print"
   ClientHeight    =   4275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   Icon            =   "frmPrintSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrButton xrButton1 
      Height          =   420
      Index           =   0
      Left            =   5160
      TabIndex        =   0
      Top             =   525
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   741
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
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   690
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   3450
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   1217
      BackColor       =   12632256
      Begin VB.CheckBox Check1 
         Caption         =   "Co&llate"
         Height          =   210
         Left            =   3210
         TabIndex        =   13
         Tag             =   "et0;fb0"
         Top             =   255
         Width           =   1275
      End
      Begin VB.TextBox txtCopy 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1995
         TabIndex        =   1
         Text            =   "1"
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Copies"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   45
         TabIndex        =   12
         Top             =   30
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No of Copies"
         Height          =   255
         Index           =   1
         Left            =   1020
         TabIndex        =   2
         Top             =   240
         Width           =   1485
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1635
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1020
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   2884
      BackColor       =   12632256
      Begin VB.TextBox txtRange 
         Height          =   315
         Left            =   1095
         TabIndex        =   5
         Top             =   750
         Width           =   3180
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&All"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Tag             =   "et0;fb0"
         Top             =   450
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Range"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   6
         Tag             =   "et0;fb0"
         Top             =   810
         Width           =   1365
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Print Range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   1155
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter page numbers and/or page ranges separated by commas. For example: 1, 6, 3-6"
         Height          =   390
         Index           =   0
         Left            =   1110
         TabIndex        =   7
         Top             =   1125
         Width           =   3465
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   690
      Index           =   2
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   2715
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   1217
      BackColor       =   12632256
      Begin VB.OptionButton Option1 
         Caption         =   "Landscape"
         Height          =   255
         Index           =   1
         Left            =   2955
         TabIndex        =   10
         Tag             =   "et0;fb0"
         Top             =   255
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Portrait"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   9
         Tag             =   "et0;fb0"
         Top             =   255
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Orientation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   45
         TabIndex        =   11
         Top             =   30
         Width           =   975
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   435
      Index           =   3
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   767
      BackColor       =   12632256
      Begin VB.Label lblPrinter 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   765
         TabIndex        =   3
         Top             =   105
         Width           =   3810
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Printer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   75
         TabIndex        =   16
         Top             =   90
         Width           =   645
      End
   End
   Begin xrControl.xrButton xrButton1 
      Height          =   420
      Index           =   1
      Left            =   5160
      TabIndex        =   8
      Top             =   990
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   741
      Caption         =   "Set&up"
      AccessKey       =   "u"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin xrControl.xrButton xrButton1 
      Height          =   420
      Index           =   2
      Left            =   5160
      TabIndex        =   15
      Top             =   1455
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   741
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
   End
End
Attribute VB_Name = "frmPrintSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmPrintSetup"

Private p_oAppDrivr As clsAppDriver
Private p_oSkin As clsFormSkin
Private p_oReport As Report
Private p_bCancel As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Set Report(Value As Report)
   Set p_oReport = Value
End Property

Property Get Collate() As Boolean
   Collate = Check1.Value
End Property

Property Let Collate(ByVal Value As Boolean)
   Check1.Value = IIf(Value, 1, 0)
End Property

Property Get Copies() As Integer
   Copies = txtCopy.Text
End Property

Property Let Copies(ByVal Value As Integer)
   txtCopy.Text = Value
End Property

Property Get PageRange() As String
   If Option2(0).Value Then
      PageRange = "xxx"
   Else
      PageRange = IIf(txtRange.Text = Empty, "xxx", txtRange.Text)
   End If
End Property

Property Let PageRange(ByVal Value As String)
   Option2(1).Value = True
   txtRange.Text = Value
End Property

Property Get Orientation() As Integer
   Orientation = IIf(Option1(0).Value, 0, 1)
End Property

Property Let Orientation(ByVal Value As Integer)
   Select Case Value
   Case 0, 1
      If Value = 0 Then
         Option1(0).Value = 1
         Option1(1).Value = 0
      Else
         Option1(1).Value = 1
         Option1(0).Value = 0
      End If
   End Select
End Property

Property Get Cancelled() As Boolean
   Cancelled = p_bCancel
End Property

Private Sub Form_Load()
   If p_oReport Is Nothing Then Unload Me

   Set p_oSkin = New clsFormSkin
   Set p_oSkin.AppDriver = p_oAppDrivr
   Set p_oSkin.Form = Me
   p_oSkin.ApplySkin xeFormTransDetail
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set p_oSkin = Nothing
End Sub

Private Sub txtRange_Validate(Cancel As Boolean)
   Option2(0).Value = txtRange.Text = Empty
   Option2(1).Value = txtRange.Text <> Empty
End Sub

Private Sub xrButton1_Click(Index As Integer)
   If Index = 1 Then
      p_oReport.PrinterSetup 0
   Else
      p_bCancel = Index = 2
      Me.Hide
   End If
End Sub

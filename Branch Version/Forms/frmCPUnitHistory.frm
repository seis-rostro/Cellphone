VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPUnitHistory 
   BorderStyle     =   0  'None
   Caption         =   "Order History"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Left            =   4665
      TabIndex        =   0
      Top             =   525
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Close"
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
      Picture         =   "frmCPUnitHistory.frx":0000
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3015
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1500
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   5318
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtMonth 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1755
      End
      Begin VB.TextBox txtMonth 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1395
         Width           =   1755
      End
      Begin VB.TextBox txtMonth 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1770
         Width           =   1755
      End
      Begin VB.TextBox txtMonth 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2145
         Width           =   1755
      End
      Begin VB.TextBox txtMonth 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1755
      End
      Begin VB.TextBox txtMonth 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   645
         Width           =   1755
      End
      Begin VB.TextBox txtDemand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   645
         Width           =   1020
      End
      Begin VB.TextBox txtDemand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1020
      End
      Begin VB.TextBox txtDemand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1395
         Width           =   1020
      End
      Begin VB.TextBox txtDemand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1770
         Width           =   1020
      End
      Begin VB.TextBox txtDemand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2145
         Width           =   1020
      End
      Begin VB.TextBox txtDemand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1020
      End
      Begin VB.TextBox txtOrder 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   3165
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1020
      End
      Begin VB.TextBox txtOrder 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   3165
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2145
         Width           =   1020
      End
      Begin VB.TextBox txtOrder 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3165
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1770
         Width           =   1020
      End
      Begin VB.TextBox txtOrder 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3165
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1395
         Width           =   1020
      End
      Begin VB.TextBox txtOrder 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3165
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1020
      End
      Begin VB.TextBox txtOrder 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3165
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   645
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "MONTH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   480
         TabIndex        =   21
         Tag             =   "wt0;fb0"
         Top             =   195
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "DEMAND"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   2048
         TabIndex        =   20
         Tag             =   "wt0;fb0"
         Top             =   195
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "ORDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3203
         TabIndex        =   19
         Tag             =   "wt0;fb0"
         Top             =   195
         Width           =   945
      End
      Begin VB.Line Line1 
         X1              =   1920
         X2              =   1920
         Y1              =   660
         Y2              =   2835
      End
      Begin VB.Line Line2 
         X1              =   3090
         X2              =   3090
         Y1              =   675
         Y2              =   2850
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   945
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   1667
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   90
         Width           =   3255
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   465
         Width           =   3255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   158
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   533
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmCPUnitHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oSkin As clsFormSkin
Private p_sBrandNme As String
Private p_sModelNme As String
Private p_vHistory As Variant
Private p_bIsCPUnit As Boolean

Public Property Let Brand(ByVal value As String)
   p_sBrandNme = value
End Property

Public Property Let Model(ByVal value As String)
   p_sModelNme = value
End Property

Public Property Let History(ByVal value As Variant)
   p_vHistory = value
End Property

Property Let IsCPUnit(ByVal value As Boolean)
   p_bIsCPUnit = value
End Property

Private Sub cmdButton_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   
   If p_bIsCPUnit Then
      Label1(12).Caption = "WEEK"
   Else
      Label1(12).Caption = "MONTH"
   End If
   
   Call loadHistory
End Sub

Private Sub loadHistory()
   Dim lnCtr As Integer
   
   On Error Resume Next
   txtField(0) = p_sBrandNme
   txtField(1) = p_sModelNme
   
   For lnCtr = 0 To UBound(p_vHistory) - 1
      If Left(p_vHistory(lnCtr, 0), 4) = "" Then GoTo nextEntry
      txtMonth(lnCtr) = Left(p_vHistory(lnCtr, 0), 4) & _
                           ", " & Right(p_vHistory(lnCtr, 0), 2)
      txtDemand(lnCtr) = p_vHistory(lnCtr, 1)
      txtOrder(lnCtr) = p_vHistory(lnCtr, 2)
nextEntry:
   Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   p_sBrandNme = ""
   p_sModelNme = ""
   p_vHistory = ""
End Sub

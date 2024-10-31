VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmJOMenu 
   BorderStyle     =   0  'None
   Caption         =   "Job Order Menu"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1560
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   2752
      ClipControls    =   0   'False
      Begin VB.OptionButton Option1 
         Caption         =   "Job Order Register"
         Height          =   195
         Index           =   2
         Left            =   540
         TabIndex        =   2
         Tag             =   "et0;fb0"
         Top             =   1005
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "New Job Order "
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   0
         Tag             =   "et0;fb0"
         Top             =   420
         Width           =   1725
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Job Order Release"
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   1
         Tag             =   "et0;fb0"
         Top             =   720
         Width           =   1935
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   60
         TabIndex        =   5
         Top             =   2175
         Visible         =   0   'False
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Shape Shape1 
         Height          =   1215
         Index           =   0
         Left            =   150
         Top             =   180
         Width           =   2610
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Left            =   165
         TabIndex        =   7
         Tag             =   "et0;fb0"
         Top             =   810
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Date"
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
         Index           =   3
         Left            =   210
         TabIndex        =   6
         Tag             =   "et0;fb0"
         Top             =   810
         Width           =   390
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   3240
      TabIndex        =   3
      Top             =   1275
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Open"
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
      Picture         =   "frmJOMenu.frx":0000
      CaptionAlign    =   0
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   1695
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJOMenu.frx":077A
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmJOMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   If bLoaded = False Then
      bLoaded = True
      Option1(0).Value = True
   End If
End Sub

Private Sub cmdButton_Click(Index As Integer)

   Select Case Index
      Case 0 'OK
         If Option1(0).Value = True Then
               frmJobOrder.Show
         ElseIf Option1(1).Value = True Then
            frmJobOrder_Release.Show
         ElseIf Option1(2).Value = True Then
            frmJobOrder_Register.Show
         End If
         Unload Me
      Case 1 'Cancel
         Unload Me
   End Select
   
End Sub

Private Sub Form_Load()

   CenterChildForm mdiMain, Me
   bLoaded = False
   
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransDetail
                  
End Sub


VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPOS_Register 
   BorderStyle     =   0  'None
   Caption         =   "POS Register"
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13995
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   555
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   5070
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   979
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   5
         Left            =   9960
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "ht0;ft0"
         Text            =   "Total Amount"
         Top             =   120
         Width           =   2235
      End
      Begin VB.TextBox txtfield 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D9FCFD&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   1260
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   120
         Width           =   1940
      End
      Begin VB.TextBox txtfield 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D9FCFD&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   5025
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "ht0;eb0"
         Text            =   "Text1"
         Top             =   120
         Width           =   1940
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   100
         Left            =   9240
         TabIndex        =   17
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Given"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         Left            =   150
         TabIndex        =   13
         Top             =   135
         Width           =   1035
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Mode of Payment"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   3435
         TabIndex        =   15
         Top             =   135
         Width           =   1560
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3375
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1605
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   5953
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3225
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   12165
         _ExtentX        =   21458
         _ExtentY        =   5689
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4194304
         BackColorSel    =   16777215
         ForeColorSel    =   4194304
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1050
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   1852
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   8
         Left            =   1440
         MaxLength       =   120
         TabIndex        =   9
         Text            =   "Remarks"
         Top             =   645
         Width           =   10725
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   1440
         TabIndex        =   7
         Text            =   "Customer Name"
         Top             =   390
         Width           =   6045
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   1440
         TabIndex        =   1
         Text            =   "Invoice No."
         Top             =   135
         Width           =   2220
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   5265
         TabIndex        =   3
         Text            =   "Transaction Date"
         Top             =   135
         Width           =   2220
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   9030
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Transaction Number"
         Top             =   135
         Width           =   3135
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   9030
         TabIndex        =   11
         Text            =   "Sales Person"
         Top             =   390
         Width           =   3135
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   5
         Left            =   105
         TabIndex        =   8
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   105
         TabIndex        =   6
         Top             =   405
         Width           =   1305
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Number"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   3
         Left            =   105
         TabIndex        =   0
         Top             =   135
         Width           =   1350
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   7
         Left            =   3915
         TabIndex        =   2
         Top             =   150
         Width           =   1350
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   10
         Left            =   7755
         TabIndex        =   4
         Top             =   135
         Width           =   1515
      End
      Begin VB.Label lblField 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Person"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   7770
         TabIndex        =   10
         Top             =   405
         Width           =   1095
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   3
      Left            =   12645
      TabIndex        =   23
      Top             =   2475
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
      Picture         =   "frmPOS_Register.frx":0000
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   12645
      TabIndex        =   20
      Top             =   1635
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Browse"
      AccessKey       =   "B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPOS_Register.frx":077A
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   12645
      TabIndex        =   19
      Top             =   1215
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&New"
      AccessKey       =   "N"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPOS_Register.frx":0EF4
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   12645
      TabIndex        =   21
      Top             =   2055
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Update"
      AccessKey       =   "U"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPOS_Register.frx":166E
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   4
      Left            =   12645
      TabIndex        =   22
      Top             =   2055
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Save"
      AccessKey       =   "S"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPOS_Register.frx":1DE8
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   5
      Left            =   12645
      TabIndex        =   24
      Top             =   2475
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmPOS_Register.frx":2562
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   6
      Left            =   12645
      TabIndex        =   26
      ToolTipText     =   "Void Transaction"
      Top             =   795
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Void"
      AccessKey       =   "V"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPOS_Register.frx":2CDC
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin VB.Label lblField 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   135
      TabIndex        =   25
      Top             =   1020
      Width           =   3765
   End
End
Attribute VB_Name = "frmPOS_Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset

Dim txtfieldGotfocus As Boolean
Dim pbnewitem As Boolean
Dim psSelected() As String

Dim pnUserRights As Integer
Dim psUserID As String
Dim psUserName As String
Dim Address As String
Dim Code As String
Dim Branch As String

Dim lsSQL As String
Dim pnindex As Integer
Dim pnCtr As Integer
Dim lrs As ADODB.Recordset
Dim lsSearch As String
Dim TranStat As String
Dim Time As String

Dim psClientID As String
Dim psUpdate_Load As Boolean
Dim psUpdate As Boolean

Property Let ClientID(ClientID As String)
   psClientID = ClientID
End Property

Property Let Update_Load(Update_Load As String)
   psUpdate_Load = Update_Load
End Property

Private Sub cmdButton_Click(Index As Integer)
Dim lnCtr As Integer

   Select Case Index
      Case 0 'New Record
            ClearFields
            EmptyGrid
      Case 1 'Update CP_SO_Master
            If txtfield(0).Text <> "" Or txtfield(2).Text <> "" Then
               If Code = oApp.BranchCode Then
                  InitButton xeModeReady
                  Select Case MSFlexGrid1.TextMatrix(1, 11)
                     Case 2, 3
                        psUpdate = True
                     Case 1, 4
                        psUpdate = False
                  End Select
                  For lnCtr = 0 To 8
                     Select Case lnCtr
                        Case 1, 3, 6, 8
                           If txtfield(lnCtr).Enabled = False Then txtfield(lnCtr).Enabled = True
                        Case 0
                           If txtfield(lnCtr).Enabled = True Then txtfield(lnCtr).Enabled = False
                     End Select
                  Next
                  txtfield(1).SetFocus
               Else
                  MsgBox "Update of Branch Transaction Not Permitted!!!" & vbCrLf & vbCrLf & _
                  "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
                  psUpdate = False
               End If
            Else
               MsgBox "No Active Transaction!!!", vbInformation, "Information"
               psUpdate = False
            End If
      Case 2 'Search
            SearchTrans
            psUpdate = False
      Case 3 'Close
            Unload Me
      Case 4 'Save
            If txtfield(1).Text <> "" Then
               If psUpdate_Load = True Then
                  Delete_Load
                  Unload frmLoadRetail_POS
                  Unload frmLoadWallet_POS
               Else
                  UpdateCP_SO_Master
               End If
            Else
               MsgBox "Invalid Transaction Date!!!", vbCritical, "Warning"
               txtfield(1).SetFocus
            End If
            psUpdate = False
     Case 5 'Cancel
            InitButton xeModeAddNew
            ClearFields
            EmptyGrid
      Case 6 'Cancel_Transaction
            If Trim(txtfield(8).Text) = "" Then
               MsgBox "Specify Reason for Cancelling Transaction!!!", vbCritical, "Warning"
               If txtfield(8).Enabled = False Then txtfield(8).Enabled = True
               txtfield(8).SetFocus
            Else
               If Code = oApp.BranchCode Then
                  Delete_Transaction
               Else
                  MsgBox "Update of Branch Transaction Not Permitted!!!" & vbCrLf & vbCrLf & _
                  "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
               End If
               psUpdate = False
               Unload frmLoadRetail_POS
               Unload frmLoadWallet_POS
            End If
      Case 7 'Payment
         Select Case TranStat
            Case 1   'Credit Card
               frmPOS_Credit_Register.Show
'            Case 2   'Cheque
'               txtfield(7).Text = "Cheque"
'            Case 3   'Installment
'               txtfield(7).Text = "Installment"
'            Case 4   'Cancelled
'               txtfield(7).Text = "Cancelled"
         End Select
   End Select

End Sub

Private Sub Form_Activate()

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      ClearFields
      bLoaded = True
      txtfield(2).SetFocus
      oDriver.HideButton 6
      Code = oApp.BranchCode
   End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = MSFlexGrid1.hWnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      Case 27
         Call Modified("CP_SO_Master", "sTransNox = '" & txtfield(0).Text & "' ")
   End Select
End Sub

Private Sub Form_Load()
Dim lsSQL As String
Dim lnCtr As Integer
   
   CenterChildForm mdiMain, Me
   bLoaded = False
   
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance

   InitButton xeModeAddNew

   InitGrid
   EmptyGrid
   txtfield(0).Tag = 0 'Do not Update msFlexGrid
   psUpdate = False
            
End Sub

Private Sub ClearFields()
Dim lnCtr As Integer

For pnCtr = 1 To 8
   txtfield(pnCtr) = ""
Next

For lnCtr = 0 To 8
   Select Case lnCtr
      Case 1, 3, 6, 8
         If txtfield(lnCtr).Enabled = True Then txtfield(lnCtr).Enabled = False
      Case 0, 2
         If txtfield(lnCtr).Enabled = False Then txtfield(lnCtr).Enabled = True
   End Select
Next

txtfield(0).Text = ""
txtfield(1).Tag = ""
txtfield(2).SetFocus
xrFrame3.Enabled = False
oDriver.HideButton 6
Code = oApp.BranchCode

psUpdate = False
psUpdate_Load = False
Unload frmLoadWallet_POS
Unload frmLoadRetail_POS

End Sub

Private Sub InitButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = xeModeReady, False, True)
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow
   cmdButton(6).Visible = lbShow
  
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   txtfield(1).Enabled = Not lbShow
   txtfield(3).Enabled = Not lbShow
   txtfield(6).Enabled = Not lbShow
   txtfield(8).Enabled = Not lbShow
End Sub

Private Sub InitGrid()

   With MSFlexGrid1
      .Rows = 2
      .Cols = 13
      .Font = "Arial"
      
      'column title
      .TextMatrix(0, 1) = "Bar Code"
      .TextMatrix(0, 2) = "Brand & Model"
      .TextMatrix(0, 3) = "Unit Price"
      .TextMatrix(0, 4) = "Qty"
      .TextMatrix(0, 5) = "%"
      .TextMatrix(0, 6) = "Amt"
      .TextMatrix(0, 7) = "Stock ID"
      .TextMatrix(0, 8) = "IMEI No. / Cell #"
      .TextMatrix(0, 9) = "Serial ID. / Ref. #"
      .TextMatrix(0, 10) = "Sub Total"
      .TextMatrix(0, 11) = "Category"
      .TextMatrix(0, 12) = "Load Pur Price"
      
      'column width
      .ColWidth(0) = 300
      .ColWidth(1) = 1550
      .ColWidth(2) = 3780
      .ColWidth(3) = 950
      .ColWidth(4) = 380
      .ColWidth(5) = 380
      .ColWidth(6) = 580
      .ColWidth(7) = 0
      .ColWidth(8) = 1600
      .ColWidth(9) = 1400
      .ColWidth(10) = 1200
      .ColWidth(11) = 0
      .ColWidth(12) = 0

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 6
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
      .ColAlignment(6) = 6
      .ColAlignment(7) = 6
      .ColAlignment(8) = 1
      .ColAlignment(9) = 1
      .Row = 1
   End With

End Sub

Private Sub EmptyGrid()
Dim lnCtr As Integer

With MSFlexGrid1
   .Rows = 2
   For lnCtr = 1 To .Cols - 1
      .TextMatrix(1, lnCtr) = ""
   Next
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
   Unload frmPOS_Credit_Register
   Unload frmLoadRetail_POS
   Unload frmLoadWallet_POS
End Sub

Private Sub MSFlexGrid1_Click()
Dim lnCtr As Integer
   
   If psUpdate = True Then
      With MSFlexGrid1
         Select Case MSFlexGrid1.TextMatrix(1, 11)
            Case 2 'Update LoadWallet
               frmLoadWallet_POS.oForm = "Register"
               frmLoadWallet_POS.oStock = MSFlexGrid1.TextMatrix(1, 7)
               frmLoadWallet_POS.oBarrcode = MSFlexGrid1.TextMatrix(1, 1)
               frmLoadRetail_POS.oBrand = MSFlexGrid1.TextMatrix(1, 2)
               frmLoadWallet_POS.txtfield(2).Text = txtfield(5).Text
               With frmLoadWallet_POS.GridEditor1
                  .Rows = MSFlexGrid1.Rows
                  For lnCtr = 1 To MSFlexGrid1.Rows
                     If lnCtr = .Rows Then Exit For
                     .TextMatrix(lnCtr, 0) = MSFlexGrid1.TextMatrix(lnCtr, 0)
                     .TextMatrix(lnCtr, 1) = MSFlexGrid1.TextMatrix(lnCtr, 8)
                     .TextMatrix(lnCtr, 2) = MSFlexGrid1.TextMatrix(lnCtr, 9)
                     .TextMatrix(lnCtr, 3) = Format(MSFlexGrid1.TextMatrix(lnCtr, 3), "#,##0.00")
                     .TextMatrix(lnCtr, 4) = MSFlexGrid1.TextMatrix(lnCtr, 5)
                     .TextMatrix(lnCtr, 5) = Format(MSFlexGrid1.TextMatrix(lnCtr, 10), "#,##0.00")
                  Next
               End With
               frmLoadWallet_POS.Show
               
            Case 3   'Update LoadRetail
               frmLoadRetail_POS.oForm = "Register"
               frmLoadRetail_POS.oStock = MSFlexGrid1.TextMatrix(1, 7)
               frmLoadRetail_POS.oBarrcode = MSFlexGrid1.TextMatrix(1, 1)
               frmLoadRetail_POS.oBrand = MSFlexGrid1.TextMatrix(1, 2)
               frmLoadRetail_POS.txtfield(2).Text = txtfield(5).Text
               With frmLoadRetail_POS.GridEditor1
                  .Rows = MSFlexGrid1.Rows
                  For lnCtr = 1 To MSFlexGrid1.Rows
                     If lnCtr = .Rows Then Exit For
                     .TextMatrix(lnCtr, 0) = MSFlexGrid1.TextMatrix(lnCtr, 0)
                     .TextMatrix(lnCtr, 1) = MSFlexGrid1.TextMatrix(lnCtr, 8)
                     .TextMatrix(lnCtr, 2) = MSFlexGrid1.TextMatrix(lnCtr, 9)
                     .TextMatrix(lnCtr, 3) = Format(MSFlexGrid1.TextMatrix(lnCtr, 3), "#,##0.00")
                     .TextMatrix(lnCtr, 4) = Format(MSFlexGrid1.TextMatrix(lnCtr, 12), "#,##0.00")
                  Next
               End With
               frmLoadRetail_POS.Show
         End Select
      End With
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfieldGotfocus = True
   pnindex = Index
   txtfield(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   If Index = 1 Then
      If Not IsDate(txtfield(Index).Text) Then
         txtfield(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Else
         txtfield(Index).Text = Format(txtfield(Index).Text, "MMMM DD, YYYY")
      End If
   End If
   If txtfield(3).Text = "" Then txtfield(3).Tag = ""
End Sub

Private Sub ShowMaster()
   'Show Master
   lsSQL = "SELECT" _
            & " a.sTransNox, " _
            & " a.sSalesInv, " _
            & " a.nTranTotl, " _
            & " a.nAmtPaidx, " _
            & " a.sCashierx, " _
            & " a.dTransact, " _
            & " b.sLastName + ' , ' + b.sFrstName + ' ' + b.sMiddName as xFullName," _
            & " a.cTranStat, " _
            & " c.sLastName + ' , ' + c.sFrstName + ' ' + c.sMiddName as xSalesPer," _
            & " b.sClientID, " _
            & " a.sRemarksx, " _
            & " d.sBranchCd, " _
            & " d.sBranchNm, " _
            & " d.sAddressx + ' ' + e.sTownName xAddressx " _
         & " FROM CP_SO_Master a " _
            & " LEFT JOIN Client_Master b " _
               & " ON a.sClientID = b.sClientID " _
            & " LEFT JOIN Sales_Person c " _
               & " ON a.sCashierx = c.sEmployID " _
            & " LEFT JOIN Branch d " _
               & " ON left(a.sTransNox,2) = d.sBranchCd " _
            & " LEFT JOIN TownCity e " _
               & " ON d.sTownIDxx = e.sTownIDxx " _
         & " WHERE left(a.sTransNox,2) = '" & Code & "' "
   Set lrs = New ADODB.Recordset

End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

   If KeyCode = 13 Or KeyCode = vbKeyF3 Then
      Select Case Index
         Case 2   'Search Transaction
            SearchTrans
            ShowGrid
         Case 3      'Search Client
            SearchClient False
         Case 6
            SearchSales False
      End Select
      If txtfield(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
   End If

End Sub

Private Sub SearchTrans()

   ShowMaster
   Select Case pnindex
   Case 0
         lsSQL = lsSQL & " AND a.sTransNox like '%" & txtfield(0).Text & "' "
   Case 2
         lsSQL = lsSQL & " AND a.sSalesInv like '%" & txtfield(2).Text & "%' "
   Case Else
         lsSQL = lsSQL
   End Select
      
   lsSQL = lsSQL & " ORDER BY a.dTransact"
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   ClearFields

      Select Case lrs.RecordCount
         Case 1
            Code = lrs("sBranchCd")
            txtfield(0).Text = lrs("sTransNox")
            txtfield(1).Text = Format(lrs("dTransact"), "MMMM dd, yyyy")
            txtfield(1).Tag = Format(lrs("dTransact"), "MMMM dd, yyyy")
            If Not IsNull(lrs("sSalesInv")) Then txtfield(2).Text = lrs("sSalesInv")
            If Not IsNull(lrs("xFullName")) Then txtfield(3).Text = lrs("xFullName")
            txtfield(3).Tag = IIf(IsNull(lrs("sClientID")), "", lrs("sClientID"))
            txtfield(4).Text = Format(lrs("nAmtPaidx"), "#,##0.00")
            txtfield(5).Text = Format(lrs("nTranTotl"), "#,##0.00")
            txtfield(6).Text = IIf(IsNull(lrs("xSalesPer")), "", lrs("xSalesPer"))
            txtfield(6).Tag = lrs("sCashierx")
            txtfield(8).Text = lrs("sRemarksx")
            TranStat = lrs("cTranStat")
            oDriver.ShowButton 6
         Case Is > 1
            lsSearch = KwikBrowse(oApp, lrs, _
                           "sSalesInv»dTransact»xFullName»nTranTotl»nAmtPaidx", _
                           "Invoice»Date»Customer Name»Tran Total»Amount Paid", _
                           "@»MMMM dd, yyyy»@»#,##0.00»#,##0.00")
            If lsSearch <> "" Then
               psSelected = Split(lsSearch, "»")
               Code = psSelected(11)
               txtfield(0).Text = psSelected(0)
               txtfield(1).Text = Format(psSelected(5), "MMMM dd, yyyy")
               txtfield(1).Tag = Format(psSelected(5), "MMMM dd, yyyy")
               If Not IsNull(psSelected(1)) Then txtfield(2).Text = psSelected(1)
               If Not IsNull(psSelected(6)) Then txtfield(3).Text = psSelected(6)
               txtfield(3).Tag = psSelected(9)
               txtfield(4).Text = Format(psSelected(3), "#,##0.00")
               txtfield(5).Text = Format(psSelected(2), "#,##0.00")
               txtfield(6).Text = psSelected(8)
               txtfield(6).Tag = psSelected(4)
               txtfield(8).Text = psSelected(10)
               TranStat = psSelected(7)
               oDriver.ShowButton 6
            End If
         Case 0
            ClearFields
            EmptyGrid
            MsgBox "No Record Found!!!", vbInformation, "Notice"
            Exit Sub
      End Select
      
      TransactionType
      ShowGrid

End Sub
Private Sub TransactionType()
   Select Case TranStat
      Case 0
         txtfield(7).Text = "Cash"
      Case 1
         txtfield(7).Text = "Credit Card"
      Case 2
         txtfield(7).Text = "Cheque"
      Case 3
         txtfield(7).Text = "Installment"
      Case 4
         txtfield(7).Text = "Cancelled"
      Case 5
         txtfield(7).Text = "Replaced"
   End Select
End Sub

Private Sub SearchClient(ByVal SearchValue As Boolean)
   
   Set lrs = New ADODB.Recordset
   lsSQL = "SELECT" _
               & " sClientID, " _
               & " sLastName + ', ' + sFrstName + ' ' + sMiddName FullName," _
               & " sAddressx " _
            & " FROM Client_Master " _

   If SearchValue Then
      lsSQL = lsSQL & " WHERE sLastName + ', ' + sFrstName + ' ' + sMiddName = '" & txtfield(3).Text & "'"
   Else
      lsSQL = lsSQL & " WHERE sLastName + ', ' + sFrstName + ' ' + sMiddName LIKE '" & txtfield(3).Text & "%' "
   End If
   lsSQL = lsSQL & " ORDER BY sLastName + ', ' + sFrstName + ' ' + sMiddName"
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.RecordCount = 1 Then
      txtfield(3).Tag = lrs("sClientID")
      txtfield(3).Text = lrs("FullName")
      
   ElseIf lrs.RecordCount > 1 Then
        lsSearch = KwikBrowse(oApp, lrs, _
                          "sClientID»" _
                        & "FullName»" _
                     & "sAddressx", _
                          "Client ID»" _
                        & "Name»" _
                        & "Address")

        If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            txtfield(3).Tag = psSelected(0)
            txtfield(3).Text = psSelected(1)
        End If
   Else
      frmCustomer.Client = txtfield(3).Text
      frmCustomer.oForm = "Register"
      frmCustomer.Show 1
   End If
   Set lrs = Nothing

End Sub

Private Sub SearchSales(ByVal SearchValue As Boolean)
   
   Set lrs = New ADODB.Recordset
   lsSQL = "SELECT" _
               & " sEmployID, " _
               & " sFrstName + ' ' + sLastName xFullName " _
            & " FROM Sales_Person " _
            & " WHERE cRecdStat = 1 "
   If SearchValue Then
      lsSQL = lsSQL & " AND sFrstName + ' ' + sLastName = '" & txtfield(6).Text & "'"
   Else
      lsSQL = lsSQL & " AND sFrstName + ' ' + sLastName LIKE '" & txtfield(6).Text & "%' "
   End If
                  
   lsSQL = lsSQL & " ORDER BY xFullname"
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.RecordCount = 1 Then
      txtfield(6).Tag = lrs("sEmployID")
      txtfield(6).Text = lrs("xFullName")
      
   ElseIf lrs.RecordCount > 1 Then
        lsSearch = KwikBrowse(oApp, lrs, _
                          "sEmployID»xFullName", _
                          "Emp. ID»Sales Person")

        If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            txtfield(6).Tag = psSelected(0)
            txtfield(6).Text = psSelected(1)
        End If
   Else
      frmSales_Person.Show
   End If
   Set lrs = Nothing

End Sub


Private Sub ShowGrid()
Dim lsSQL As String
Dim lnCtr As Integer

   'Show Detail
   lsSQL = "SELECT" _
         & " Distinct " _
         & " a.nEntryNox, " _
         & " a.nQuantity, " _
         & " a.nUnitPrce, " _
         & " a.nDiscount, " _
         & " a.nDiscAmnt, " _
         & " a.nSubTotal, " _
         & " a.nPurPrice, " _
         & " b.sSerialID, " _
         & " c.sIMEINoxx, " _
         & " d.sStockIDx, " _
         & " d.sBarrCode, " _
         & " d.sDescript, " _
         & " d.cWdSerial, " _
         & " e.sPhoneNum, " _
         & " e.sReferNox, " _
         & " f.sBrandNme, " _
         & " g.sModelNme, " _
         & " h.sColorNme, " _
         & " d.cCellLoad, " _
         & " d.cWalletxx  " _

   lsSQL = lsSQL _
         & " FROM CP_SO_Detail a " _
            & " LEFT JOIN CP_SO_Serial b " _
               & " ON a.sTransNox = b.sTransNox " _
               & " AND a.nEntryNox = b.nEntryNox " _
            & " LEFT JOIN CP_Serial_Master c " _
               & " ON b.sSerialID = c.sSerialID " _
            & " LEFT JOIN CP_Inventory d " _
               & " ON a.sStockIDx = d.sStockIDx " _
            & " LEFT join ELoad_Ledger e " _
               & " ON a.sTransnox = e.ssourceno " _
                  & " AND a.nEntryNox = e.sTransNox " _
            & " LEFT JOIN Brand f " _
               & " ON d.sBrandIdx = f.sBrandIDx " _
            & " LEFT JOIN Model g " _
               & " ON d.sModelIDx = g.sModelIDx " _
            & " LEFT JOIN Color h " _
               & " ON d.sColorIDx = h.sColorIdx " _
         & " WHERE a.sTransNox = '" & txtfield(0).Text & "' " _
         & " ORDER by a.nEntryNox "

   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   With MSFlexGrid1
      If oRS.RecordCount <> 0 Then
         .Rows = oRS.RecordCount + 1
         For lnCtr = 0 To oRS.RecordCount - 1
            .TextMatrix(lnCtr + 1, 0) = oRS("nEntryNox")
            .TextMatrix(lnCtr + 1, 1) = oRS("sBarrCode")
            .TextMatrix(lnCtr + 1, 2) = Trim(IIf(IsNull(oRS("sBrandNme")), "", oRS("sBrandNme")) _
                                    & " " & IIf(IsNull(oRS("sModelNme")), "", oRS("sModelNme")) _
                                    & " " & IIf(IsNull(oRS("sDescript")), "", oRS("sDescript")) _
                                    & " " & IIf(IsNull(oRS("sColorNme")), "", oRS("sColorNme")))
            .TextMatrix(lnCtr + 1, 3) = Format(oRS("nUnitPrce"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 4) = oRS("nQuantity")
            .TextMatrix(lnCtr + 1, 5) = oRS("nDiscount")
            .TextMatrix(lnCtr + 1, 6) = oRS("nDiscAmnt")
            .TextMatrix(lnCtr + 1, 7) = oRS("sStockIDx")
            .TextMatrix(lnCtr + 1, 10) = Format(oRS("nSubTotal"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 12) = Format(oRS("nPurPrice"), "#,##0.00")
            Select Case oRS("cWdSerial")
            Case 1
               .TextMatrix(lnCtr + 1, 8) = IIf(IsNull(oRS("sIMEINoxx")), "", oRS("sIMEINoxx"))
               .TextMatrix(lnCtr + 1, 9) = IIf(IsNull(oRS("sSerialID")), "", oRS("sSerialID"))
               .TextMatrix(lnCtr + 1, 11) = 1

            Case 0
               If oRS("cWalletxx") = 1 Or oRS("cCellLoad") = 1 Then
                  .TextMatrix(lnCtr + 1, 8) = IIf(IsNull(oRS("sPhoneNum")), "", oRS("sPhoneNum"))
                  .TextMatrix(lnCtr + 1, 9) = IIf(IsNull(oRS("sReferNox")), "", oRS("sReferNox"))
                  If oRS("cWalletxx") = 1 Then
                     .TextMatrix(lnCtr + 1, 11) = 2
                  ElseIf oRS("cCellLoad") = 1 Then
                     .TextMatrix(lnCtr + 1, 11) = 3
                  End If
               Else
                  .TextMatrix(lnCtr + 1, 8) = ""
                  .TextMatrix(lnCtr + 1, 9) = ""
                  .TextMatrix(lnCtr + 1, 11) = 4
               End If
            End Select
            oRS.MoveNext
         Next
      Else
         .Rows = 2
      End If
      If .Rows > 12 Then
         .ColWidth(2) = 3500
      Else
         .ColWidth(2) = 3780
      End If
      
   End With
   Set oRS = Nothing
   txtfield(2).SetFocus

End Sub

Private Function Delete_Load() As Boolean
Dim lnrow As Long
Dim lnCtr As Integer
Dim lsApproval As Integer
Dim ctr As Integer
Dim QOH As Double
Dim lrs As New ADODB.Recordset
Dim oRS As New ADODB.Recordset
Dim Entry As Integer

Delete_Load = True
oApp.Connection.BeginTrans
On Error Goto errProc

   'Roll Back QOH in CP_Inventory_Master
   lrs.Open "SELECT sStockIDx, " _
               & " sTransNox, " _
               & " nPurPrice " _
            & "From CP_SO_Detail " _
            & "WHERE sTransNox ='" & txtfield(0).Text & "' " _
            , oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

   If lrs.RecordCount <> 0 Then
      Do While Not lrs.EOF
         'Roll Back QOH in CP_Inventory_Master
          lsSQL = "UPDATE CP_Inventory_Master SET" _
               & " nQtyOnHnd = nQtyOnHnd + '" & CDbl(lrs("nPurPrice")) & "'," _
               & " dModified = getdate() " _
         & " WHERE sStockIDx = '" & lrs("sStockIDx") & "' " _
               & " And sBranchCd = '" & oApp.BranchCode & "' "
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
         lrs.MoveNext
      Loop
   End If
   Set lrs = Nothing

   'Select nEntryNox of Current Transaction
   lsSQL = "SELECT sSourceNo, " _
               & " nEntryNox  " _
         & "From ELoad_Ledger " _
         & "WHERE sSourceNo ='" & txtfield(0).Text & "' " _
            & " AND sStockIDx = '" & MSFlexGrid1.TextMatrix(1, 7) & "'"
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   Entry = oRS("nEntryNox")
   Set oRS = Nothing
      
   'Select QOH of Prev Transaction
   lsSQL = "SELECT sSourceNo, " _
               & " nEntryNox, " _
               & " nQtyOnHnd  " _
         & " From ELoad_Ledger " _
         & " WHERE sStockIDx = '" & MSFlexGrid1.TextMatrix(1, 7) & "'" _
            & " AND nEntryNox = '" & Entry - 1 & "'" _
            & " AND sBranchcd = '" & oApp.BranchCode & "'" _
         & " ORDER BY sTransNox Desc "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   QOH = oRS("nQtyOnHnd")
   Set oRS = Nothing
      
   'Delete CP_SO_Detail Transaction
   lsSQL = "DELETE CP_SO_Detail " _
            & " WHERE sTransNox = '" & txtfield(0).Text & "'"
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
   If lnrow <> 0 Then oApp.RegisDelete lsSQL

   'Delete Eload_Ledger Transaction
   lsSQL = "DELETE ELoad_Ledger " _
            & " WHERE sSourceNo = '" & txtfield(0).Text & "'" _
               & " AND sStockIDx = '" & MSFlexGrid1.TextMatrix(1, 7) & "'" _
               & " AND sBranchCd = '" & oApp.BranchCode & "'"
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
   If lnrow <> 0 Then oApp.RegisDelete lsSQL

   With MSFlexGrid1
      For lnCtr = 1 To .Rows - 1

         'Insert New CP_SO_Detail
         lsSQL = "INSERT INTO CP_SO_Detail " _
            & "( sTransNox, " _
            & "  nEntryNox, " _
            & "  sStockIDx, " _
            & "  nQuantity, " _
            & "  nPurPrice, " _
            & "  nUnitPrce, " _
            & "  nDiscount, " _
            & "  nDiscAmnt, " _
            & "  nSubTotal, " _
            & "  dModified) " _
               & "VALUES " _
                  & "('" & txtfield(0).Text & "', " _
                  & "'" & .TextMatrix(lnCtr, 0) & "', " _
                  & "'" & .TextMatrix(lnCtr, 7) & "', " _
                  & "'" & CLng(.TextMatrix(lnCtr, 4)) & "', " _
                  & "'" & CDbl(.TextMatrix(lnCtr, 12)) & "', " _
                  & "'" & CDbl(.TextMatrix(lnCtr, 3)) & "', " _
                  & "'" & CLng(.TextMatrix(lnCtr, 5)) & "', " _
                  & "'" & CDbl(.TextMatrix(lnCtr, 6)) & "', " _
                  & "'" & CDbl(.TextMatrix(lnCtr, 10)) & "', " _
                  & " getdate())"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText

         'Insert New ELOad_Ledger
         lsSQL = "INSERT INTO ELoad_Ledger " _
                     & "( sStockIDx, " _
                     & "  sBranchCd, " _
                     & "  sLocation, " _
                     & "  dTransact, " _
                     & "  sReferNox, " _
                     & "  sPhoneNum, " _
                     & "  sSourceCd, " _
                     & "  sSourceNo, " _
                     & "  sTransNox, " _
                     & "  nQtyInxxx, " _
                     & "  nQtyOutxx, " _
                     & "  nEntryNox, " _
                     & "  nQtyOnHnd, " _
                     & "  sModified, " _
                     & "  dModified) "
         lsSQL = lsSQL _
                  & "VALUES " _
                     & "('" & .TextMatrix(lnCtr, 7) & "' ," _
                     & "'" & oApp.BranchCode & "', " _
                     & "'" & oApp.BranchCode & "', " _
                     & "'" & CDate(txtfield(1).Text) & "', " _
                     & "'" & .TextMatrix(lnCtr, 9) & "', " _
                     & "'" & .TextMatrix(lnCtr, 8) & "', " _
                     & " 'CPSl', " _
                     & "'" & txtfield(0).Text & "', " _
                     & "'" & .TextMatrix(lnCtr, 0) & "', " _
                     & "'0', " _
                     & "'" & CDbl(.TextMatrix(lnCtr, 12)) & "', " _
                     & "'" & Entry & "', " _
                     & "'" & CDbl(QOH) - CDbl(.TextMatrix(lnCtr, 12)) & "', " _
                     & "'" & Encrypt(oApp.UserID) & "'," _
                     & " getdate())"
         oApp.Connection.Execute lsSQL, lnrow, adCmdText

         QOH = CDbl(QOH) - CDbl(.TextMatrix(lnCtr, 12))

      Next
   End With
   
   'Select QOH of Current Transaction
   lsSQL = "SELECT sStockIDx, " _
               & " nEntryNox, " _
               & " nQtyOnHnd " _
         & "From ELoad_Ledger " _
         & "WHERE sSourceNo ='" & txtfield(0).Text & "' " _
            & " AND sStockIDx = '" & MSFlexGrid1.TextMatrix(1, 7) & "'" _
         & " ORDER BY sTransNox Desc "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   QOH = oRS("nQtyOnHnd")
   
   lrs.Open "SELECT sStockIDx, " _
            & " sTransNox, " _
            & " nQtyOnHnd, " _
            & " sBranchCd, " _
            & " nEntryNox, " _
            & " nQtyInxxx, " _
            & " nQtyOutxx  " _
         & " From ELoad_Ledger " _
         & " WHERE sStockIDx ='" & oRS("sStockidx") & "'" _
            & " AND sBranchCd = '" & oApp.BranchCode & "'" _
            & " AND nEntryNox > '" & oRS("nEntryNox") & "'" _
         & " ORDER BY nEntryNox, sTransNox " _
         , oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

   'Update ELoad_Ledger, QOH
   Do While Not lrs.EOF
      lsSQL = "UPDATE ELoad_Ledger SET " _
               & " dModified = getdate(), " _
               & " nQtyOnHnd = '" & CDbl(QOH) & "' + nQtyinxxx - nQtyOutxx "
      lsSQL = lsSQL & " WHERE sStockIDx = '" & lrs("sStockIDx") & "'" _
                     & " AND sBranchCd = '" & lrs("sBranchCd") & "'" _
                     & " AND nEntryNox = '" & lrs("nEntryNox") & "'" _
                     & " AND sTransNox = '" & lrs("sTransNox") & "'"
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
      
      QOH = CDbl(QOH) + CDbl(lrs("nQtyInxxx")) - CDbl(lrs("nQtyOutxx"))
      lrs.MoveNext
   Loop
   Set lrs = Nothing
   Set oRS = Nothing

   'Update SO_Master
   lsSQL = "UPDATE CP_SO_Master SET " _
               & " nTranTotl = '" & CDbl(txtfield(5).Text) & "'," _
               & " nAmtPaidx = '" & CDbl(txtfield(5).Text) & "'," _
               & " dTransact = '" & CDate(txtfield(1).Text) & "'," _
               & " sRemarksx = '" & txtfield(8).Text & "'," _
               & " sModified = '" & Encrypt(oApp.UserID) & "', " _
               & " dModified = getdate() " _
         & " WHERE sTransNox = '" & txtfield(0).Text & "' "
   oApp.Connection.Execute lsSQL, lnrow, adCmdText

   'Update QOH, CP_Inventory_Master
   lsSQL = "UPDATE CP_Inventory_Master SET" _
         & " nQtyOnHnd = '" & CDbl(QOH) & "', " _
         & " sModified = '" & Encrypt(oApp.UserID) & "', " _
         & " dModified = getdate() " _
   & " WHERE sStockIDx = '" & MSFlexGrid1.TextMatrix(1, 7) & "' " _
         & " And sBranchCd = '" & oApp.BranchCode & "' "
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
   
   If lnrow <= 0 Then
      MsgBox "Unable to Update Transaction!!!", vbCritical, "Warning"
      GoTo endProc
   End If

   MsgBox "Record Successfully Updated!!!", vbInformation, "Information"
   InitButton xeModeAddNew
   oDriver.HideButton 6

   Unload frmLoadRetail_POS
   Unload frmLoadWallet_POS

endProc:
   oApp.Connection.CommitTrans
   Exit Function
errProc:
   oApp.Connection.RollbackTrans
   Delete_Load = False
   MsgBox Err.Description, vbCritical, "Warning"

End Function

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(1).Text = Format(txtfield(1).Text, "MMMM dd, yyyy")
   txtfield(Index).BackColor = &HFFFFFF
End Sub

Private Function UpdateCP_SO_Master() As Boolean
Dim lnrow As Long
Dim lnCtr As Integer
Dim lrs As New ADODB.Recordset

UpdateCP_SO_Master = True
On Error GoTo errProc

   Time = Format(Now, "hh:nn:ss AM/PM")

   If txtfield(1).Tag <> txtfield(1).Text Then
      lsSQL = "UPDATE CP_SO_Cheque SET" _
                  & " dTransact = '" & CDate(txtfield(1).Text) & " " & Time & "'," _
                  & " dModified = getdate() " _
            & " WHERE sTransNOx = '" & txtfield(0).Text & "' "
      oApp.Connection.Execute lsSQL, lnrow, adCmdText

      lsSQL = "UPDATE CP_SO_Credit SET" _
                  & " dTransact = '" & CDate(txtfield(1).Text) & " " & Time & "'," _
                  & " dModified = getdate() " _
            & " WHERE sTransNOx = '" & txtfield(0).Text & "' "
      oApp.Connection.Execute lsSQL, lnrow, adCmdText

      lsSQL = "UPDATE CP_SO_Installment SET" _
                  & " dTransact = '" & CDate(txtfield(1).Text) & " " & Time & "'," _
                  & " dModified = getdate() " _
            & " WHERE sTransNOx = '" & txtfield(0).Text & "' "
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
   End If

   If txtfield(3).Text = "" Then txtfield(3).Tag = ""

   With MSFlexGrid1
      For lnCtr = 1 To .Rows - 1
         Select Case .TextMatrix(.Row, 11)
            Case 2, 3  'ELoad
               If txtfield(1).Tag <> txtfield(1).Text Then
                  lsSQL = "UPDATE ELoad_Ledger SET" _
                              & " dTransact = '" & CDate(txtfield(1).Text) & " " & Time & "'," _
                              & " dModified = getdate() " _
                        & " WHERE sStockIDx = '" & .TextMatrix(lnCtr, 7) & "'" _
                              & " AND sSourceNo = '" & txtfield(0).Text & "'" _
                              & " AND sSourceCd = 'CPSl' " _
                              & " AND sBranchCd = '" & oApp.BranchCode & "'"
                  oApp.Connection.Execute lsSQL, lnrow, adCmdText
               End If
           Case 1, 4
               If txtfield(1).Tag <> txtfield(1).Text Then
                  lsSQL = "UPDATE CP_Inventory_Ledger SET" _
                              & " dTransact = '" & CDate(txtfield(1).Text) & " " & Time & "'," _
                              & " dModified = getdate() " _
                        & " WHERE sStockIDx = '" & .TextMatrix(lnCtr, 7) & "'" _
                              & " AND sSourceNo = '" & txtfield(0).Text & "'" _
                              & " AND sSourceCd = 'CPSl' " _
                              & " AND sBranchCd = '" & oApp.BranchCode & "'"
                  oApp.Connection.Execute lsSQL, lnrow, adCmdText

                  lsSQL = "UPDATE CP_Serial_Ledger SET" _
                              & " dTransact = '" & CDate(txtfield(1).Text) & " " & Time & "'," _
                              & " dModified = getdate() " _
                        & " WHERE sSerialID = '" & .TextMatrix(lnCtr, 9) & "' " _
                              & " AND sSourceNo = '" & txtfield(0).Text & "' " _
                              & " AND sSourceCd = 'CPSl' "
                  oApp.Connection.Execute lsSQL, lnrow, adCmdText
               End If

               lsSQL = "UPDATE CP_Serial_Master SET" _
                           & " sClientID = '" & txtfield(3).Tag & "'," _
                           & " sModified = '" & Encrypt(oApp.UserID) & "'," _
                           & " dModified = getdate() " _
                     & " WHERE sSerialID = '" & .TextMatrix(lnCtr, 9) & "'" _
                           & " AND sClientID <> '" & txtfield(3).Tag & "'"
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
         End Select

         If txtfield(1).Tag <> txtfield(1).Text Then
            lsSQL = "UPDATE CP_SO_Master SET " _
                        & " dTransact = '" & CDate(txtfield(1).Text) & "'," _
                        & " sRemarksx = '" & Trim(txtfield(8).Text) & "'," _
                        & " sSalesInv = '" & Trim(txtfield(2).Text) & "'," _
                        & " sClientID = '" & Trim(txtfield(3).Tag) & "'," _
                        & " sCashierx = '" & Trim(txtfield(6).Tag) & "'," _
                        & " sModified = '" & Encrypt(oApp.UserID) & "'," _
                        & " dModified = getdate() " _
                  & " WHERE sTransNox = '" & txtfield(0).Text & "' "
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
         Else
            lsSQL = "UPDATE CP_SO_Master SET " _
                        & " dTransact = '" & CDate(txtfield(1).Text) & "'," _
                        & " sRemarksx = '" & Trim(txtfield(8).Text) & "'," _
                        & " sSalesInv = '" & Trim(txtfield(2).Text) & "'," _
                        & " sClientID = '" & Trim(txtfield(3).Tag) & "'," _
                        & " sCashierx = '" & Trim(txtfield(6).Tag) & "'," _
                        & " sModified = '" & Encrypt(oApp.UserID) & "'," _
                        & " dModified = getdate() " _
                  & " WHERE (sRemarksx <> '" & Trim(txtfield(8)) & "'" _
                        & " OR sSalesInv <> '" & Trim(txtfield(2)) & "'" _
                        & " OR sClientID <> '" & Trim(txtfield(3).Tag & "'" _
                        & " OR sCashierx <> '" & txtfield(6).Tag) & "'" _
                        & " )AND sTransNox = '" & txtfield(0).Text & "'"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
         End If
      Next

'      If lnrow <= 0 Then
'         MsgBox "Unable to Update Record!!!" & vbCrLf & vbCrLf & _
'         "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
'         UpdateCP_SO_Master = False
'         GoTo endProc
'      End If

   End With
   MsgBox "Record Successfully Updated!!!", vbInformation, "Information"
   InitButton xeModeAddNew
   oDriver.HideButton 6

endProc:
   Exit Function
errProc:
   UpdateCP_SO_Master = False
   MsgBox Err.Description, vbCritical, "Warning"

End Function

Private Function Delete_Transaction() As Boolean
Dim lnrow As Long
Dim lnCtr As Integer
Dim lsApproval As Integer
Dim ctr As Integer
Dim QOH As Integer
Dim lrs As New ADODB.Recordset
Dim oRS As New ADODB.Recordset

Delete_Transaction = True
oApp.Connection.BeginTrans
On Error Goto errProc

   'Allow Cancel if 1 day after dtransact
   'Only Managers are Allowed to Cancel Transaction
   If DateDiff("d", txtfield(1).Text, Date) > 1 Then
      lsApproval = MsgBox("Cancel Not Permitted!!!" & vbCrLf & _
            "Sales Report Already Generated!!!" & vbCrLf & vbCrLf & _
            "Notify ROSALYN LAZO DESCALLAR for Assistance" & vbCrLf & vbCrLf & _
            "Seek for Approval!", vbQuestion + vbYesNo, "Notice")
      If lsApproval <> vbYes Then Exit Function
      If Not GetApproval(oApp, pnUserRights, psUserID, psUserName) Then Exit Function
         If pnUserRights < xeManager Then
            MsgBox "Approving User is Not Authorized!!!", vbCritical, "Warning"
            Exit Function
         End If
   Else
      If oApp.UserLevel = xeEncoder Then
         lsApproval = MsgBox("Seek for Approval?", vbQuestion + vbYesNo, "Confirm")
         If lsApproval = vbYes Then
            If Not GetApproval(oApp, pnUserRights, psUserID, psUserName) Then Exit Function
            If pnUserRights < xeManager Then
               MsgBox "Approving User is Not Authorized!!!", vbCritical, "Warning"
               Exit Function
            End If
         End If
      End If
   End If

   lsApproval = MsgBox("Are you sure you want to Cancel this Transaction?", vbQuestion + vbYesNo, "Confirm")
   If lsApproval <> vbYes Then Exit Function

   With MSFlexGrid1
      If lrs.State = adStateOpen Then lrs.Close
      lrs.Open "SELECT sSerialID, " _
                  & " sTransNox  " _
               & "From CP_SO_Serial " _
               & "WHERE sTransNox ='" & txtfield(0).Text & "' " _
                  & " ORDER BY sSerialID " _
               , oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
      If lrs.RecordCount <> 0 Then
         Do While Not lrs.EOF
            'Delete CP_Serial_Ledger
            lsSQL = "DELETE CP_Serial_Ledger " _
                        & " WHERE sSourceNo = '" & lrs("sTransNox") & "'" _
                           & " AND sSourceCd = 'CPSl' " _
                           & " AND sBranchCd = '" & oApp.BranchCode & "'" _
                           & " AND sSerialID = '" & lrs("sSerialID") & "'"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
            If lnrow <> 0 Then
               oApp.RegisDelete lsSQL

               'Update CP_Serial_Ledger nEntryNox
               Call Recalc_Serial("'" & lrs("sSerialID") & "'")

               'Update CP_Serial_Master
               lsSQL = "UPDATE CP_Serial_Master SET" _
                     & " cSoldStat = '0', " _
                     & " cLocation = '1', " _
                     & " sClientID = '', " _
                     & " dModified = getdate() " _
                     & " WHERE sSerialID = '" & lrs("sSerialID") & "' "
                     oApp.Connection.Execute lsSQL, lnrow, adCmdText
            End If
            lrs.MoveNext
         Loop
         Set lrs = Nothing
      End If
   End With

   'Update CP_SO_Master
   lsSQL = "UPDATE CP_SO_Master SET " _
               & " sRemarksx = '" & txtfield(8).Text & "'," _
               & " cTranStat = '4', " _
               & " dCancelxx = '" & oApp.ServerDate & "'," _
               & " sModified = '" & Encrypt(oApp.UserID) & "'," _
               & " dModified = getdate() " _
         & " WHERE sTransNox = '" & txtfield(0).Text & "' " _
               & " AND cTranStat <> '4' "
   oApp.Connection.Execute lsSQL, lnrow, adCmdText

   If lnrow = 0 Then
      MsgBox "Unable to Delete CP_Inventory!!!", vbCritical, "Warning"
      Delete_Transaction = True
      GoTo endProc
   End If

   If MSFlexGrid1.TextMatrix(1, 11) = 2 Or MSFlexGrid1.TextMatrix(1, 11) = 3 Then
      If lrs.State = adStateOpen Then lrs.Close
      lrs.Open "SELECT sStockIDx, " _
                  & " sTransNox, " _
                  & " nPurPrice  " _
               & "From CP_SO_Detail " _
               & "WHERE sTransNox ='" & txtfield(0).Text & "' " _
               , oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

      If lrs.RecordCount <> 0 Then
         If oRS.State = adStateOpen Then lrs.Close
            oRS.Open "SELECT nEntryNox, sTransNox " _
                     & "From ELoad_Ledger " _
                     & "WHERE sSourceNo ='" & txtfield(0).Text & "'" _
                        & "AND sSourcecd = 'CPSl' " _
                        & "AND sBranchcd = '" & oApp.BranchCode & "'" _
                     & " ORDER BY nEntryNox Desc " _
                     , oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

         Do While Not lrs.EOF
            'Delete ELoad Ledger
            lsSQL = "DELETE ELoad_Ledger " _
                     & " WHERE sSourceNo = '" & lrs("sTransNOx") & "'" _
                        & " AND sStockIDx = '" & lrs("sStockIDx") & "'" _
                        & " AND sSourceCd = 'CPSl' " _
                        & " AND sBranchCd = '" & oApp.BranchCode & "'"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
            If lnrow <> 0 Then
               oApp.RegisDelete lsSQL

               'Update ELoad_Ledger nEntryNox, nQtyOnHnd
               Call Recalc_Load("'" & oApp.BranchCode & "'", "'" & lrs("sStockIDx") & "'", _
                     "'" & CLng(oRS("nEntryNox")) & "'")

               'Roll Back QOH in CP_Inventory_Master
                lsSQL = "UPDATE CP_Inventory_Master SET" _
                     & " nQtyOnHnd = nQtyOnHnd + '" & CDbl(lrs("nPurPrice")) & "'," _
                     & " dModified = getdate() " _
               & " WHERE sStockIDx = '" & lrs("sStockIDx") & "' " _
                     & " And sBranchCd = '" & oApp.BranchCode & "' "
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
            End If
            lrs.MoveNext
         Loop
      End If
      Set lrs = Nothing
      Set oRS = Nothing
   Else
      'Roll Back Quantity
      Call RollBack_Qty("CP_SO_Detail", "'" & txtfield(0) & "'")

      'Roll Back EntryNo
      Call Recalc_Ledger("CP_SO_Detail", "'" & txtfield(0) & "'", "'CPSl'", "'" & oApp.BranchCode & "'")
   End If

   MsgBox "Transaction Successfully Cancelled!!!", vbInformation, "Information"
   InitButton xeModeAddNew
   oDriver.HideButton 6

endProc:
   oApp.Connection.CommitTrans
   Exit Function
errProc:
   oApp.Connection.RollbackTrans
   Delete_Transaction = False
   MsgBox Err.Description, vbCritical, "Warning"

End Function

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  March 18, 2008  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'Include cWdSerial in Table CP_Inventory
'Category .textmatrix(pnctr,11)
'1   w/ serial
'2   Load Wallet
'3   Load Retail
'4   Accessories




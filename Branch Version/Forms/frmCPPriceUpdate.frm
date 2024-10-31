VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPPriceUpdate 
   BorderStyle     =   0  'None
   Caption         =   "CP Price Update"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   16095
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2340
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   4128
      BorderStyle     =   1
      Begin VB.CheckBox Check1 
         Caption         =   "Filter by Model"
         Height          =   195
         Left            =   3465
         TabIndex        =   4
         Tag             =   "wt0;fb0"
         Top             =   825
         Width           =   1335
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1500
         TabIndex        =   3
         Top             =   495
         Width           =   3255
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   2850
         MaxLength       =   128
         TabIndex        =   8
         Top             =   1530
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1830
         Width           =   1905
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1500
         TabIndex        =   1
         Top             =   195
         Width           =   3255
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   2850
         MaxLength       =   6
         TabIndex        =   6
         Top             =   1230
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP MODEL"
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
         Index           =   0
         Left            =   165
         TabIndex        =   2
         Top             =   540
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Effectivity Date"
         Height          =   195
         Index           =   3
         Left            =   1500
         TabIndex        =   9
         Top             =   1875
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Increase Amount"
         Height          =   195
         Index           =   8
         Left            =   1500
         TabIndex        =   7
         Top             =   1575
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP BRAND"
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
         Index           =   1
         Left            =   165
         TabIndex        =   0
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Increase Rate"
         Height          =   195
         Index           =   2
         Left            =   1500
         TabIndex        =   5
         Top             =   1275
         Width           =   1005
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   27
      Top             =   4980
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
      Picture         =   "frmCPPriceUpdate.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   28
      Top             =   1905
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCPPriceUpdate.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   29
      Top             =   2520
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Searc&h"
      AccessKey       =   "h"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPPriceUpdate.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   30
      Top             =   3750
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Del. Row"
      AccessKey       =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPPriceUpdate.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   31
      Top             =   4365
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&MarkUp"
      AccessKey       =   "M"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPPriceUpdate.frx":1DE8
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4185
      Index           =   1
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   2895
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   7382
      BorderStyle     =   1
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   2850
         TabIndex        =   26
         Top             =   1605
         Width           =   1905
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   1545
         TabIndex        =   24
         Top             =   3720
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1545
         TabIndex        =   22
         Top             =   3420
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1545
         TabIndex        =   20
         Top             =   3120
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   2850
         TabIndex        =   18
         Top             =   1305
         Width           =   1905
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   780
         Width           =   1905
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Price"
         Height          =   195
         Index           =   12
         Left            =   1800
         TabIndex        =   25
         Top             =   1650
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price 4"
         Height          =   195
         Index           =   11
         Left            =   195
         TabIndex        =   23
         Top             =   3765
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price 3"
         Height          =   195
         Index           =   10
         Left            =   195
         TabIndex        =   21
         Top             =   3465
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price 2"
         Height          =   195
         Index           =   9
         Left            =   195
         TabIndex        =   19
         Top             =   3165
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price"
         Height          =   195
         Index           =   7
         Left            =   1635
         TabIndex        =   17
         Top             =   1350
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   15
         Top             =   825
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Desc."
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   13
         Top             =   525
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   11
         Top             =   225
         Width           =   420
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6525
      Left            =   6585
      TabIndex        =   32
      Top             =   540
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   11509
      _Version        =   393216
      Appearance      =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   33
      Top             =   3135
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&ADD"
      AccessKey       =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPPriceUpdate.frx":2562
   End
End
Attribute VB_Name = "frmCPPriceUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmMCCashPriceUpdate"

Private WithEvents oTrans As ggcPriceUpdate.clsCPPriceUpdate
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnCtr As Integer, pnIndex As Integer
Dim pbGridFocus As Boolean, pbSave As Boolean, pbLoaded As Boolean

Private Sub Check1_Click()
   With Check1
      If .Value = xeYes Then
         If oTrans.FilterModel(txtField(0)) Then
            InitForm
            LoadDetail
         End If
         
      Else
         If oTrans.RemoveFilter Then
            InitForm
            LoadDetail
         End If
         
         txtField(0) = ""
      End If
      
      txtField(0).Locked = .Value
      txtField(0).SetFocus
   End With
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   txtField_LostFocus pnIndex
   With MSFlexGrid1
      Select Case Index
      Case 0 'Save
         If .Rows > 2 Then
            pnCtr = 0
            Do While pnCtr < .Rows
               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
                  .Row = pnCtr
                  If oTrans.deleteDetail(.Row - 1) Then
                     InitForm
                     LoadDetail
                  End If
               Else
                  pnCtr = pnCtr + 1
               End If
            Loop

            .ColWidth(1) = 3350
            If .Rows > 27 Then .ColWidth(1) = 3200
         End If

         If isEntryOk Then
            If oTrans.SaveTransaction = True Then
               MsgBox "Transaction Saved Successfully!!!", vbInformation, "Notice"
               Call InitForm
               Call ClearFields

               pbSave = True
               txtField(1).SetFocus
            Else
               MsgBox "Unable to Save Transaction!!!", vbCritical, "Warning"
            End If
         End If
      Case 1 'Search
         If pbGridFocus Then
            If txtOthers(1).hwnd Then
               If oTrans.searchDetail(.Row - 1, pnIndex, txtOthers(1)) Then
                  .Row = .Rows - 1
                  Call MSFlexGrid1_Click
               End If
            End If
         Else
            oTrans.SearchMaster pnIndex
            txtField(pnIndex).SetFocus
         End If
         .Refresh
      Case 2 'Delete Row
         If .TextMatrix(.Rows - 1, 1) = "" Then Exit Sub
         If MsgBox("This Model will be Set to Inactive!!! Do you want to continue?", vbQuestion & vbYesNo, "Confirm") = vbYes Then
            If .Rows > 2 Then
               If oTrans.deleteDetail(.Row - 1) Then
                  InitForm
                  LoadDetail
               End If
            End If
         End If
      Case 3 'Markup
         If .Rows > 2 Then
            If .TextMatrix(1, 1) = Empty Then Exit Sub

            Call oTrans.MarkUpPrice
            LoadDetail
         End If
      Case 4 'Close
         If pbSave = False And oTrans.Master("sBrandNme") <> Empty Then
            lnRep = MsgBox("Entry is in Update Mode!!!" & vbCrLf & _
                           "Do you want to Continue Closing this Entry!!!", vbYesNo + vbQuestion, "Confirm")

            If lnRep = vbYes Then
               Unload Me
            Else
               txtField(pnIndex).SetFocus
            End If
         Else
            Unload Me
         End If
      Case 5 'add
         With MSFlexGrid1
            If .TextMatrix(.Rows - 1, 1) = "" Then Exit Sub
               
            oTrans.addDetail
      
            LoadDetail
            
            If .Rows > 27 Then
            End If
            .Row = .Rows - 1
            MSFlexGrid1_Click
         End With
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   pbLoaded = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = MSFlexGrid1.hwnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New ggcPriceUpdate.clsCPPriceUpdate
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

   InitForm
   ClearFields
   initButton xeModeAddNew

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   pbLoaded = False
End Sub

'Private Sub GridEditor1_AddingRow(Cancel As Boolean)
'   With GridEditor1
'      If .TextMatrix(.Row, 1) = "" Then
'         Cancel = True
'      End If
'      If Not Cancel Then oTrans.addDetail
'
'      If .Rows > 27 Then .ColWidth(1) = 3450
'   End With
'End Sub

Private Sub MSFlexGrid1_Click()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      If pbLoaded Then If .Col <> 0 Then txtOthers(.Col).SetFocus
      
      For lnCtr = 1 To 8
         txtOthers(lnCtr) = .TextMatrix(.Row, lnCtr)
      Next
      
      .Col = 0
      .ColSel = .Cols - 1
   End With
   
   pbGridFocus = True
End Sub

Private Sub oTrans_DetailRetrieved()
   With MSFlexGrid1
      For pnCtr = 1 To 8
         Select Case pnCtr
         Case 4, 5, 6, 7, 8
            .TextMatrix(.Row, pnCtr) = Format(oTrans.Detail(.Row - 1, pnCtr), "#,##0.00")
         Case Else
            .TextMatrix(.Row, pnCtr) = IFNull(oTrans.Detail(.Row - 1, pnCtr))
         End Select
      Next
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   With txtField(Index)
      If .Text <> oTrans.Master(Index) Then

         .Text = oTrans.Master(Index)
         Call LoadDetail
      End If
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 1 Then .Text = Format(.Text, "MM/DD/YY")
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pbGridFocus = False
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         If Index = 1 Then
            If KeyCode = vbKeyF3 Then
               oTrans.SearchMaster Index, .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oTrans.SearchMaster Index, .Text
               If .Text <> "" Then SetNextFocus
            End If
         End If
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)

   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow

   xrFrame1(0).Enabled = lbShow

   If Not lbShow Then cmdButton(4).SetFocus
End Sub

Private Sub InitForm()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Clear
   
      .Cols = 9
      .Rows = 2
      .Font = "MS Sans Serif"

      'Column Title
      .TextMatrix(0, 1) = "Model Name"
      .TextMatrix(0, 2) = "Model Code"
      .TextMatrix(0, 3) = "Brand"
      .TextMatrix(0, 4) = "SRP"
      .TextMatrix(0, 5) = "SRP 2"
      .TextMatrix(0, 6) = "SRP 3"
      .TextMatrix(0, 7) = "SRP 4"
      .TextMatrix(0, 8) = "Last Price"

      .Row = 0

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      .Row = 1

      'Column Width
      .ColWidth(0) = 300
      .ColWidth(1) = 3450
      .ColWidth(2) = 1600
      .ColWidth(3) = 1600
      .ColWidth(4) = 1200
      .ColWidth(5) = 0
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      .ColWidth(8) = 1200

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      
      MSFlexGrid1_Click
   End With
End Sub

Private Sub ClearFields()
   txtField(1).Text = ""
   txtField(2).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   txtField(3).Text = "0.00"
   txtField(4).Text = "0.00"
   
   txtOthers(1) = ""
   txtOthers(2) = ""
   txtOthers(3) = ""
   txtOthers(4) = "0.00"
   txtOthers(5) = "0.00"
   txtOthers(6) = "0.00"
   txtOthers(7) = "0.00"
   txtOthers(8) = "0.00"
   
   InitForm
End Sub

Private Function isEntryOk() As Boolean
   If txtField(2).Text = "" Then
      MsgBox "Invalid Destination Detected!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      GoTo EntryNotOK
   End If

   With MSFlexGrid1
      If .TextMatrix(1, 1) = "" Then
         MsgBox "Detail is required!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         txtOthers(1).SetFocus
         GoTo EntryNotOK
      End If
   End With

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
End Function

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 1
         oTrans.Master(Index) = .Text
      Case 2
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
         oTrans.Master(Index) = .Text
      Case 3, 4
         If Not IsNumeric(.Text) Then .Text = 0#
         .Text = Format(.Text, "#,##0.00")
         oTrans.Master(Index) = CDbl(.Text)
      End Select
   End With
End Sub

Private Sub LoadDetail()
   Dim lnRow As Integer
   Dim lnCol As Integer
   Dim lasFormat(8) As String
   
   lasFormat(1) = "@"
   lasFormat(2) = "@"
   lasFormat(3) = "@"
   lasFormat(4) = "#,##0.00"
   lasFormat(5) = "#,##0.00"
   lasFormat(6) = "#,##0.00"
   lasFormat(7) = "#,##0.00"
   lasFormat(8) = "#,##0.00"

   With MSFlexGrid1
      .Rows = oTrans.ItemCount + 1
      
      .ColWidth(1) = 3450
      If .Rows > 27 Then .ColWidth(1) = 3200
      For lnRow = 0 To oTrans.ItemCount - 1
         For lnCol = 1 To .Cols - 1
            .TextMatrix(lnRow + 1, lnCol) = Format(IFNull(oTrans.Detail(lnRow, lnCol)), lasFormat(lnCol))
         Next
      Next
      
      .Row = 1
   End With
   Call MSFlexGrid1_Click
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With oApp
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
      If bEnd Then
         .xShowError
         End
      Else
         With Err
            .Raise .Number, .Source, .Description
         End With
      End If
   End With
End Sub

Private Sub txtOthers_GotFocus(Index As Integer)
   Call HighlightOn(txtOthers(Index))
   
   With txtOthers(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
   
   pbGridFocus = True
End Sub

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
      With MSFlexGrid1
         If Index = 1 Then
            If oTrans.searchDetail(.Row - 1, Index, txtOthers(Index)) Then txtOthers(4).SetFocus
            .Refresh
         End If
      End With
      KeyCode = 0
   End If
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   Call HighlightOff(txtOthers(Index))
End Sub

Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
   With MSFlexGrid1
      Select Case Index
      Case 4, 5, 6, 7, 8
         If .TextMatrix(.Row, 1) = "" Then
            txtOthers(Index) = ""
            Exit Sub
         End If
      
         If Not IsNumeric(txtOthers(Index)) Then
            Exit Sub
         Else
            txtOthers(Index) = Format(txtOthers(Index), "#,##0.00")
         End If
      
         .TextMatrix(.Row, Index) = txtOthers(Index)
         oTrans.Detail(.Row - 1, Index) = txtOthers(Index)
      End Select
   End With
End Sub

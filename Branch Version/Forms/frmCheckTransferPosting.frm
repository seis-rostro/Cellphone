VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCheckTransferPosting 
   BorderStyle     =   0  'None
   Caption         =   "Check Transfer Posting"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5085
      Index           =   1
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   1155
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   8969
      BackColor       =   12632256
      BorderStyle     =   1
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4860
         Left            =   5850
         TabIndex        =   21
         Top             =   90
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   8573
         _Version        =   393216
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   1425
         Left            =   90
         Tag             =   "wt0;fb0"
         Top             =   3525
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   2514
         BackColor       =   12632256
         Enabled         =   0   'False
         ClipControls    =   0   'False
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   1335
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   510
            Width           =   4230
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   1335
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   915
            Width           =   2340
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   1335
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   105
            Width           =   2340
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   165
            Width           =   930
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   705
            TabIndex        =   16
            Top             =   555
            Width           =   465
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   495
            TabIndex        =   18
            Top             =   975
            Width           =   675
         End
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   3420
         Left            =   90
         Tag             =   "wt0;fb0"
         Top             =   90
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   6033
         BackColor       =   12632256
         Enabled         =   0   'False
         ClipControls    =   0   'False
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   4
            Left            =   3060
            TabIndex        =   13
            Tag             =   "ht0;ft0"
            Top             =   2745
            Width           =   2490
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1155
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   135
            Width           =   2340
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1155
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   1425
            Width           =   4395
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1155
            TabIndex        =   7
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   1020
            Width           =   2340
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Index           =   5
            Left            =   1155
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   1830
            Width           =   4395
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL AMOUNT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   8
            Left            =   420
            TabIndex        =   12
            Top             =   2835
            Width           =   2535
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   375
            Left            =   1320
            Tag             =   "et0;ht2"
            Top             =   240
            Width           =   2265
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date "
            Height          =   195
            Index           =   0
            Left            =   660
            TabIndex        =   6
            Top             =   1110
            Width           =   390
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Trans. No."
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
            Left            =   135
            TabIndex        =   4
            Top             =   225
            Width           =   915
         End
         Begin VB.Label lblField 
            BackStyle       =   0  'Transparent
            Caption         =   "Destination"
            Height          =   195
            Index           =   5
            Left            =   255
            TabIndex        =   8
            Top             =   1515
            Width           =   795
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            Height          =   195
            Index           =   1
            Left            =   420
            TabIndex        =   10
            Top             =   1830
            Width           =   630
         End
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   13125
      TabIndex        =   22
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCheckTransferPosting.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   13125
      TabIndex        =   23
      Top             =   1770
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
      Picture         =   "frmCheckTransferPosting.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   585
      Index           =   0
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   1032
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   82
         Left            =   10320
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   105
         Width           =   2340
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   81
         Left            =   4665
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   4395
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   80
         Left            =   1080
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   90
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Recvd"
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
         Index           =   2
         Left            =   9180
         TabIndex        =   25
         Top             =   195
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
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
         Left            =   3600
         TabIndex        =   2
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trans No."
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
         Index           =   9
         Left            =   120
         TabIndex        =   0
         Top             =   180
         Width           =   855
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   13125
      TabIndex        =   26
      Top             =   1155
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Post"
      AccessKey       =   "P"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCheckTransferPosting.frx":0EF4
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Index           =   4
      Left            =   6765
      TabIndex        =   20
      Top             =   1950
      Width           =   570
   End
End
Attribute VB_Name = "frmCheckTransferPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmCheckTransfer"

Private WithEvents oTrans As clsCheckTransfer
Attribute oTrans.VB_VarHelpID = -1

Private oSkin As clsFormSkin

Dim pnIndex As Integer
Dim pnRow As Integer
Dim pnActiveRow As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim pnCtr As Integer
   Dim lnRep As Integer
   
   With MSFlexGrid1
      Select Case Index
      Case 3 'browse
         If oTrans.SearchAcceptance("", False) Then
            Call LoadMaster
            Call LoadDetail
         Else
            Call InitForm
            Call InitGrid
         End If
      Case 5 'cancel
         If oTrans.PostTransaction Then
            MsgBox "Transaction was posted successfuly.", vbInformation, pxeMODULENAME
            Call InitForm
            Call InitGrid
         End If
      Case 6 'Close
         Unload Me
      End Select
   End With
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   '''''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   '''''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsCheckTransfer
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction
   oTrans.Location = 10
   oTrans.SourceCd = "CkDv"

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   InitGrid
   InitForm
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub InitForm()
   Dim loTxt As TextBox

   For Each loTxt In txtField
      loTxt = ""
   Next

   txtOthers(4) = ""
   txtOthers(6) = "0.00"
   txtOthers(8) = ""
End Sub

Private Sub LoadMaster()
   With oTrans
      txtField(0).Text = Format(.Master("sTransNox"), "@@@@@@-@@@@@@")
      txtField(1).Text = strLongDate(.Master("dTransact"))
      txtField(2).Text = .Master("sDestinat")
      txtField(5).Text = .Master("sRemarksx")
      txtField(4).Text = Format(.Master("nTranTotl"), "#,##0.00")
      
      txtField(80) = Replace(txtField(0), "-", "")
      txtField(81) = txtField(2)
      
      txtField(80).SetFocus
   End With
End Sub

Private Sub MSFlexGrid1_Click()
   Dim lnCtr As Integer

   With oTrans
      pnActiveRow = MSFlexGrid1.Row
      pnRow = pnActiveRow
      
      Call showdetail
   End With
End Sub

Private Sub showdetail()
   With MSFlexGrid1
      If .Row = 0 Then
         txtOthers(4) = ""
         txtOthers(6) = "0.00"
         txtOthers(8) = ""
      Else
         txtOthers(4) = .TextMatrix(.Row, 2)
         txtOthers(6) = .TextMatrix(.Row, 3)
         txtOthers(8) = .TextMatrix(.Row, 1)
      End If
      
      pnActiveRow = .Row - 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
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

Private Sub InitGrid()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Clear

      .Cols = 4
      .Rows = 2
      .Font = "MS Sans Serif"
      .RowHeight(0) = 350

      'Column Title
      .TextMatrix(0, 1) = "Bank"
      .TextMatrix(0, 2) = "Check No."
      .TextMatrix(0, 3) = "Amount"

      .ColWidth(0) = 330
      .ColWidth(1) = 2800
      .ColWidth(2) = 1800
      .ColWidth(3) = 1800
      
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 6
      
      .Row = 1
      
      .TextMatrix(1, 0) = 1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = "0.00"
      
      pnActiveRow = .Row - 1
      
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Rows = oTrans.ItemCount + 1
      
      For lnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, "xBankName")
         .TextMatrix(lnCtr + 1, 2) = oTrans.Detail(lnCtr, "sCheckNox")
         .TextMatrix(lnCtr + 1, 3) = Format(oTrans.Detail(lnCtr, "nAmountxx"), "#,##0.00")
         
         .Row = .Rows - 1
         MSFlexGrid1_Click
      Next
   End With
End Sub

Public Function PrintTrans() As Boolean
   Dim lrs As New ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String

   lsOldProc = "printTrans"
   '''On Error GoTo errProc
   
   PrintTrans = False
   
   If oTrans.Master("cTranStat") = xeStateCancelled Then Exit Function
   
   Set lrs = New ADODB.Recordset
   
   lrs.Fields.Append "nField01", adInteger, 10
   lrs.Fields.Append "sField01", adVarChar, 200
   lrs.Fields.Append "sField02", adVarChar, 200
   lrs.Open
      
   With MSFlexGrid1
      For lnCtr = 0 To oTrans.ItemCount - 1
         lrs.AddNew
         lrs("sField01").Value = "CH # " & oTrans.Detail(lnCtr, "sCheckNox")
         lrs("sField02").Value = oTrans.Detail(lnCtr, "xPrtclrDs")
      Next
   End With
   
   ' assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\TransmittalForm.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   oReport.Sections("RH").ReportObjects("txtRefNo").SetText Right(oTrans.Master("sTransNox"), 10)
   oReport.Sections("RH").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("PH").ReportObjects("txtTo").SetText txtField(2).Text
   oReport.Sections("PH").ReportObjects("txtFrom").SetText oApp.BranchName
   oReport.Sections("RF").ReportObjects("txtRemarks").SetText txtField(5).Text
   oReport.Sections("PF").ReportObjects("txtPrepared").SetText oApp.UserName
   
   oReport.PrintOutEx False, 1
   lrs.Close
   PrintTrans = True

endProc:
   If oTrans.Master("cTranStat") = xeStateOpen Then oTrans.CloseTransaction
   Set oReport = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub oTrans_MasterRetreived(ByVal Index As Integer)
   Select Case Index
   Case 9
      txtField(82) = strLongDate(oTrans.Master("dReceived"))
   End Select
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
      If pnIndex = 80 Then
         If oTrans.SearchTransaction(txtField(pnIndex), True, True) Then
            Call LoadMaster
            Call LoadDetail
         Else
            Call InitForm
            Call InitGrid
         End If
      Else
         If oTrans.SearchTransaction(txtField(pnIndex), False, False) Then
            Call LoadMaster
            Call LoadDetail
         Else
            Call InitForm
            Call InitGrid
         End If
      End If
   ElseIf KeyCode = vbKeyReturn Then
      If pnIndex = 80 Then
         If oTrans.SearchTransaction(txtField(pnIndex), True, True) Then
            Call LoadMaster
            Call LoadDetail
         Else
            Call InitForm
            Call InitGrid
         End If
      Else
         If oTrans.SearchTransaction(txtField(pnIndex), False, True) Then
            Call LoadMaster
            Call LoadDetail
         Else
            Call InitForm
            Call InitGrid
         End If
      End If
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = &HC0FFFF
      
      If .Text <> "" Then
         Select Case Index
         Case 80
            .Text = Replace(.Text, "-", "")
         Case 82
            .Text = strShortDate(.Text)
         End Select
      End If
      
      .SelStart = 0
      .SelLength = Len(.Text)
      pnIndex = Index
   End With
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = &H80000005
      
      If Index = 80 Then
         
      End If
      
      If .Text <> "" Then
         Select Case Index
         Case 80
            If Len(.Text) = 12 Then
               .Text = Format(.Text, "@@@@@@-@@@@@@")
            End If
         Case 82
            .Text = strLongDate(.Text)
         End Select
      End If
      
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      Select Case Index
      Case 82
         oTrans.Master("dReceived") = .Text
      End Select
   End With
End Sub

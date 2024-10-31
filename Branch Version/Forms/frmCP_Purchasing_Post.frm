VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCP_Purchasing_Post 
   BorderStyle     =   0  'None
   Caption         =   " Purchase Order"
   ClientHeight    =   7710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   10575
      TabIndex        =   3
      Top             =   630
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
      Picture         =   "frmCP_Purchasing_Post.frx":0000
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2955
      Index           =   0
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   1080
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   5212
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   24
         Left            =   1320
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   735
         Width           =   3030
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   16
         Left            =   1350
         TabIndex        =   29
         Top             =   1725
         Width           =   5295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   13
         Left            =   8265
         TabIndex        =   27
         Top             =   2055
         Width           =   1800
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   8265
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1725
         Width           =   1800
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1335
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   75
         Width           =   2265
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   8250
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1065
         Width           =   1815
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1350
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1065
         Width           =   5295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1350
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1395
         Width           =   5295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Index           =   5
         Left            =   1350
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2385
         Width           =   5310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   8265
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1395
         Width           =   1800
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1350
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2055
         Width           =   5310
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   31
         Top             =   780
         Width           =   660
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   30
         Top             =   1755
         Width           =   420
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO #"
         Height          =   195
         Index           =   9
         Left            =   6990
         TabIndex        =   28
         Top             =   2055
         Width           =   375
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
         Height          =   195
         Index           =   8
         Left            =   6990
         TabIndex        =   25
         Top             =   1725
         Width           =   360
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1350
         Tag             =   "et0;ht2"
         Top             =   165
         Width           =   2265
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
         Height          =   195
         Index           =   1
         Left            =   6990
         TabIndex        =   19
         Top             =   1065
         Width           =   1230
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1170
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1485
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   5
         Left            =   105
         TabIndex        =   16
         Top             =   2445
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Delivery"
         Height          =   195
         Index           =   6
         Left            =   6990
         TabIndex        =   15
         Top             =   1395
         Width           =   960
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivered To"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   2085
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "unknown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7575
         TabIndex        =   13
         Tag             =   "eb0;et0"
         Top             =   165
         Width           =   2385
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   7515
         Top             =   105
         Width           =   2520
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   7545
         Top             =   135
         Width           =   2460
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   10575
      TabIndex        =   4
      Top             =   1905
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Print"
      AccessKey       =   "Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_Purchasing_Post.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   10575
      TabIndex        =   5
      Top             =   2535
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
      Picture         =   "frmCP_Purchasing_Post.frx":0EF4
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   540
      Index           =   1
      Left            =   135
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   953
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   75
         Width           =   2145
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   4995
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   75
         Width           =   4965
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
         Height          =   285
         Index           =   9
         Left            =   75
         TabIndex        =   22
         Top             =   105
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   3615
         TabIndex        =   21
         Top             =   105
         Width           =   1410
      End
   End
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   3630
      Left            =   135
      TabIndex        =   23
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   4035
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   6403
      AllowBigSelection=   -1  'True
      AutoAdd         =   -1  'True
      AutoNumber      =   -1  'True
      BACKCOLOR       =   -2147483643
      BACKCOLORBKG    =   8421504
      BACKCOLORFIXED  =   -2147483633
      BACKCOLORSEL    =   -2147483635
      BORDERSTYLE     =   1
      COLS            =   2
      FILLSTYLE       =   0
      FIXEDCOLS       =   1
      FIXEDROWS       =   1
      FOCUSRECT       =   1
      EDITORBACKCOLOR =   -2147483643
      EDITORFORECOLOR =   -2147483640
      FORECOLOR       =   -2147483640
      FORECOLORFIXED  =   -2147483630
      FORECOLORSEL    =   -2147483634
      FORMATSTRING    =   ""
      Object.HEIGHT          =   3630
      GRIDCOLOR       =   12632256
      GRIDCOLORFIXED  =   0
      BeginProperty GRIDFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GRIDLINES       =   1
      GRIDLINESFIXED  =   2
      GRIDLINEWIDTH   =   1
      MOUSEICON       =   "frmCP_Purchasing_Post.frx":166E
      MOUSEPOINTER    =   0
      REDRAW          =   -1  'True
      RIGHTTOLEFT     =   0   'False
      ROWS            =   2
      SCROLLBARS      =   3
      SCROLLTRACK     =   0   'False
      SELECTIONMODE   =   0
      Object.TOOLTIPTEXT     =   ""
      WORDWRAP        =   0   'False
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   10
      Left            =   10560
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1260
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
      Picture         =   "frmCP_Purchasing_Post.frx":168A
   End
End
Attribute VB_Name = "frmCP_Purchasing_Post"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Purchasing"

Private WithEvents oTrans As clsCPPurchasing
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnCtr As Integer
Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pbSave As Boolean
Dim pbEditMode As Boolean

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
   With GridEditor1
      .Rows = 2
      .Cols = 7
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "BarrCode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Brand"
      .TextMatrix(0, 4) = "Model"
      .TextMatrix(0, 5) = "Quantity"
      .TextMatrix(0, 6) = "Unit Prc"
      
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      'column width
      .ColWidth(0) = 400
      .ColWidth(1) = 1700
      .ColWidth(2) = 3000
      .ColWidth(3) = 1400
      .ColWidth(4) = 1500
      .ColWidth(5) = 850
      .ColWidth(6) = 1200
      
      .ColFormat(5) = 0#
      .ColFormat(6) = "0.00"
            
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 6
      
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      
      .EditorBackColor = oApp.getColor("HT1")

      .Row = 1
      .Col = 1
   End With
End Sub

Public Function PrintTrans() As Boolean
   Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lrs As ADODB.Recordset
   Dim loRS As Recordset
   Dim lsAddress As String
   Dim lsSQL As String
   
   lsOldProc = "InitReport"

   Set lrs = New ADODB.Recordset
   With lrs
      .Fields.Append "nField01", adInteger, 3
      .Fields.Append "sField01", adVarChar, 64
      .Fields.Append "sField02", adVarChar, 64, adFldIsNullable
      .Fields.Append "sField03", adVarChar, 64, adFldIsNullable
      .Fields.Append "lField01", adCurrency
      .Open
      
      For lnCtr = 0 To oTrans.ItemCount - 1
         lrs.AddNew
         .Fields("nField01").Value = oTrans.Detail(lnCtr, "nQuantity")
         .Fields("sField01").Value = IFNull(oTrans.Detail(lnCtr, "sModelNme"), "")
         .Fields("sField02").Value = oTrans.Detail(lnCtr, "sColorNme")
         .Fields("sField03").Value = oTrans.Detail(lnCtr, "sModelCde")
         .Fields("lField01").Value = oTrans.Detail(lnCtr, "nUnitPrce")
      Next
   End With
   
   lsSQL = "SELECT CONCAT(a.sAddressx, ', ', b.sTownName, ' ', c.sProvName)" & _
                     " FROM Branch a, TownCity b, Province c" & _
                     " WHERE a.sBranchNm = " & strParm(txtField(3)) & _
                        " AND a.sTownIDxx = b.sTownIDxx AND b.sProvIDxx = c.sProvIDxx"
   Set loRS = New Recordset
   
   With loRS
      .Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If Not .EOF Then
         lsAddress = loRS(0)
      Else
         lsAddress = ""
      End If
   End With
   
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\MP-Purchase4Supplier.rpt")
    
   With oReport
      'assign important info to the report
      .DiscardSavedData
      .FieldMappingType = crAutoFieldMapping
      .Database.SetDataSource lrs
      
    .Sections("PH").ReportObjects("txtSupplier").SetText txtField(2)
    .Sections("PH").ReportObjects("txtDeliverTo").SetText txtField(4)
    .Sections("PH").ReportObjects("txtTerm").SetText txtField(9)
    .Sections("PH").ReportObjects("txtDateTransact").SetText txtField(1)
    .Sections("PH").ReportObjects("txtInvoiceTo").SetText oTrans.Master("sCompnyNm")
    .Sections("PH").ReportObjects("txtPONo").SetText oTrans.Master("sTransNox")
    .Sections("RF").ReportObjects("txtRemarks").SetText txtField(5)
    .Sections("PF").ReportObjects("txtRequested").SetText "JULIE MARTINEZ" 'oApp.UserName
 
      oReport.PrintOutEx False, 1
   End With

   If oTrans.Master("cTranStat") = xeStateOpen Then
      PrintTrans = oTrans.CloseTransaction
   Else
      PrintTrans = True
   End If
End Function

Private Sub LoadMaster()

   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "MMMM DD, YYYY")
      Case 2
         txtField(pnCtr).Text = oTrans.Master(2)
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 4
         txtField(pnCtr).Text = IFNull(oTrans.Master(8), "")
      Case 6
         txtField(pnCtr).Text = Format(oTrans.Master(11), "MMMM DD, YYYY")
      Case 7
         txtField(pnCtr).Text = Format(oTrans.Master(0), "@@@@-@@@@@@")
      Case 8
         txtField(pnCtr).Text = oTrans.Master(2)
      Case 9
        txtField(pnCtr).Text = IFNull(oTrans.Master(12), "")
    Case 10, 11, 12
      Case Else
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next
   txtField(13).Text = IFNull(oTrans.Master("sTransNox"), "")
   txtField(16).Text = IFNull(oTrans.Master("sBrandNme"), "")
   txtField(24).Text = IFNull(oTrans.Master("sCompnyNm"), "")
   
   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
   pbSave = True
End Sub

Private Sub LoadDetail()
Dim lnCtr As Integer

   With GridEditor1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)

      For pnCtr = 0 To oTrans.ItemCount - 1
         For lnCtr = 1 To .Cols - 1
            If lnCtr = 5 Then
               .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, 6)
            ElseIf lnCtr = 6 Then
               .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, 7)
            Else
               .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, lnCtr)
            End If
         Next
      Next
   End With
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsRep As String

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   With GridEditor1
      Select Case Index
      Case 4 ' Browse
         If oTrans.SearchTransaction() Then
            LoadMaster
            LoadDetail
         End If
      Case 7 ' code for close
         Unload Me
      Case 8 'Print
         If pbSave Then
            lsRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
            If lsRep = vbYes Then
               If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
            Else
               MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
            End If
         End If

         If GenerateReport Then
            Label2.Caption = TransStat(oTrans.Master("cTranStat"))
            
            If MsgBox("Do you want to eMail Purchase Order?", vbQuestion + vbYesNo, "") = vbYes Then
                If oTrans.Master("cTranStat") = xeStatePosted Or oTrans.Master("cTranStat") = xeStateCancelled Then
                    MsgBox "Unable to Upload purchase order.", vbCritical, "Warning"
                        GoTo endProc:
                Else
                    If Not EmailPO Then
                       MsgBox "Unable to Upload purchase order.", vbCritical, "Warning"
                    End If
                    oTrans.PostTransaction (oTrans.Master("sTransNox"))
                End If
            End If
         Else
            MsgBox "Unable to Close Transaction.", vbCritical, "Warning"
         End If
      Case 10
         If txtField(0) = "" Then Exit Sub
   
         If oTrans.PostTransaction(oTrans.Master("sTransNox")) Then
            Label2.Caption = TransStat(oTrans.Master("cTranStat"))
         
            MsgBox "Transaction Posted Successfuly.", vbInformation, "Success"

            If Not EmailPO Then
               If MsgBox("Unable to Upload purchase order. Do you want to print rather?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                  Call GenerateReport
               End If
            End If
         End If
      End Select
   End With
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   GridEditor1.Refresh

   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   End Sub
   Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hwnd Then Exit Sub
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

   Set oTrans = New clsCPPurchasing
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   InitGrid
   clearFields

   pbEditMode = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 6) = 0 Then
         Cancel = True
      End If
      If Not Cancel Then oTrans.addDetail

   End With
End Sub
Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   With GridEditor1
      oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
   End With
End Sub

Private Sub GridEditor1_GotFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("HT1")
   End With
   pbGridFocus = True
End Sub
Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "GridEditor1_KeyDown"
   ''On Error GoTo errProc

   With GridEditor1
      If KeyCode = vbKeyF3 Then
         If oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then .Col = 6
         KeyCode = 0
      End If
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub GridEditor1_LostFocus()
   With GridEditor1
      .EditorBackColor = oApp.getColor("EB")
      oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
   End With
End Sub


Private Sub clearFields()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 0
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
      Case 1, 6
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 2, 3, 4, 5
         txtField(pnCtr).Text = ""
      End Select
   Next
   
   txtField(7).Text = ""
   txtField(13).Text = ""
   txtField(24).Text = ""

   With GridEditor1
      .Rows = 2

      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
   End With

   pbSave = False
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer, ByVal Value As Variant)
      With GridEditor1
      If Index = 5 Then
         .TextMatrix(.Row, Index) = oTrans.Detail(.Row, 6)
      Else
         .TextMatrix(.Row, Index) = Value
      End If
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
'   txtField(Index).Text = Value
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      Select Case Index
      Case 1
         .Text = Format(.Text, "MM/DD/YY")
      End Select

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
         If KeyCode = vbKeyF3 Then
            oTrans.SearchTransaction .Text
            LoadMaster
            LoadDetail
            If .Text <> "" Then SetNextFocus
         Else
            If .Text <> "" Then oTrans.SearchTransaction .Text, False
               LoadMaster
               LoadDetail
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
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
      Case 5
         .Text = Format(.Text, ">")
      Case 10
         If .Text = "" Then
            clearFields
            Exit Sub
         End If

         If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then
            If oTrans.SearchTransaction(.Text, IIf(Index = 9, True, False)) Then
               LoadMaster
               LoadDetail
            Else
               clearFields
               .SetFocus
            End If
         End If
      End Select

      If Index < 9 Then oTrans.Master(Index) = .Text
   End With
End Sub

'Private Function EmailPO(Optional ByVal fsPath As String, Optional ByVal fsFile As String) As Boolean
'   Dim lsSQL As String
'   Dim lors As Recordset
'   Dim loRS2 As Recordset
'   Dim lnRet As Integer
'
'   'kalyptus - 2017.05.13 11:53am
'   lsSQL = "SELECT IFNULL(a.sEmailAdd, '') sEmailAdd, IFNULL(d.sEmailAdd, '') xEmailAdd" & _
'          " FROM Branch a" & _
'               " LEFT JOIN Branch_Others b ON a.sBranchCd = b.sBranchCD" & _
'               " LEFT JOIN Branch_Area c ON b.sAreaCode = c.sAreaCode" & _
'               " LEFT JOIN Client_Master d ON c.sAreaMngr = d.sClientID" & _
'          " WHERE a.sBranchCD = " & strParm(oTrans.Master("sBranchCD"))
'   Set loRS2 = oApp.Connection.Execute(lsSQL, , adCmdText)
'
'   lsSQL = "SELECT" & _
'                  "  sTransNox" & _
'                  ", dTransact" & _
'                  ", sMailFrom" & _
'                  ", sMailToxx" & _
'                  ", sMailCCxx" & _
'                  ", sMailBCCx" & _
'                  ", sSubjectx" & _
'                  ", sMailBody" & _
'                  ", sAttached" & _
'                  ", sSourceCD" & _
'                  ", sSourceNo" & _
'                  ", cStatusxx" & _
'                  ", dPostedxx" & _
'                  ", sModified" & _
'                  ", dModified" & _
'         " FROM Send_Mail_Master" & _
'         " WHERE 0=1"
'   Set lors = New Recordset
'   Debug.Print lsSQL
'   lors.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
'   Set lors.ActiveConnection = Nothing
'
'   lors.AddNew
'
'   lors("dTransact") = oApp.ServerDate
'   lors("sMailFrom") = oApp.getConfiguration("MPPrProcM1")
'
'   'Send the mail to the supplier
'   If InStr(1, oTrans.Master("sSupplier"), "C0W110000001", vbTextCompare) Then
'      lors("sMailToxx") = oApp.getConfiguration("SamsungMl")
''   ElseIf InStr(1, oTrans.Master("sSupplrNm"), "Suzuki", vbTextCompare) Then
''      loRS("sMailToxx") = oApp.getConfiguration("SuzkiMCMl")
''   ElseIf InStr(1, oTrans.Master("sSupplrNm"), "Yamaha", vbTextCompare) Then
''      loRS("sMailToxx") = oApp.getConfiguration("YamhaMCMl")
''   ElseIf InStr(1, oTrans.Master("sSupplrNm"), "Kawasaki", vbTextCompare) Then
''      loRS("sMailToxx") = oApp.getConfiguration("KwskiMCMl")
'   End If
'
'   lors("sMailCCxx") = oApp.getConfiguration("MailMPAGM")   'AGM will always received a CC of the PO
'   'kalyptus - 2017.05.13 11:58am
'   'Add email of branch and area manager to the email of AGM
'   If loRS2("sEmailAdd") <> "" Then
'      lors("sMailCCxx") = lors("sMailCCxx") & ";" & loRS2("sEmailAdd")
'   End If
'
'   If loRS2("xEmailAdd") <> "" Then
'      lors("sMailCCxx") = lors("sMailCCxx") & ";" & loRS2("xEmailAdd")
'   End If
'   lors("sMailBCCx") = oApp.getConfiguration("MSMPDptMl")
'
'   lors("sSubjectx") = "PURCHASE ORDER - " & oTrans.Master("sTransNox")
'   lors("sMailBody") = "Please see attached file(s) and kindly acknowledge upon receipt thru this email addresses:" & vbCrLf & _
'                        "sirabanal@guanzongroup.com.ph" & vbCrLf & _
'                        "sirabanal@guanzongroup.com.ph"
'   lors("sAttached") = oTrans.Master("sTransNox") & ".pdf"
'   lors("sSourceCD") = "MCPO"
'   lors("sSourceNo") = oTrans.Master("sTransNox")
'
'   lors("cStatusxx") = "0"
'   lors("sTransNox") = GetNextCode("Send_Mail_Master", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)
'
'   lsSQL = ADO2SQL(lors, "Send_Mail_Master", , oApp.UserID, oApp.ServerDate)
'   Debug.Print lsSQL
'   oApp.BeginTrans
'
'   oApp.Connection.Execute lsSQL, , adCmdText
'
'   'UPLOAD THE FILE TO THE SERVER
'   lsSQL = oApp.getConfiguration("MailUplApp") & " " & _
'            lors("sTransNox") & " " & _
'            oApp.AppPath & "/Temp/Upload/ " & _
'            oTrans.Master("sTransNox") & ".pdf"
'   Debug.Print lsSQL
'   lnRet = ExecCmd(lsSQL)
'
'   If lnRet = 0 Then
'      lsSQL = "UPDATE Send_Mail_Master" & _
'             " SET cStatusxx = '1'" & _
'             " WHERE sTransNox = " & strParm(lors("sTransNox"))
'      oApp.Connection.Execute lsSQL, , adCmdText
'      oApp.CommitTrans
'      MsgBox "Creating of email queue was successfully! PO will be emailed within a few minutes..."
'      EmailPO = True
'   Else
'      oApp.RollbackTrans
'      MsgBox "Unable to create email queue! Please try again later...."
'      EmailPO = False
'   End If
'End Function

Private Function EmailPO(Optional ByVal fsPath As String, Optional ByVal fsFile As String) As Boolean
   
   Dim lsSQL As String
   Dim loRS As Recordset
   Dim lnRet As Integer
   Dim loRS2 As Recordset
   Dim loData As Recordset
   Dim lsOldProc As String
   
   'kalyptus - 2017.05.13 11:41am
   lsSQL = "SELECT IFNULL(a.sEmailAdd, '') sEmailAdd, IFNULL(d.sEmailAdd, '') xEmailAdd" & _
          " FROM Branch a" & _
               " LEFT JOIN Branch_Others b ON a.sBranchCd = b.sBranchCD" & _
               " LEFT JOIN Branch_Area c ON b.sAreaCode = c.sAreaCode" & _
               " LEFT JOIN Client_Master d ON c.sAreaMngr = d.sClientID" & _
          " WHERE a.sBranchCD = " & strParm(oTrans.Master("sBranchCD"))
   Set loRS2 = oApp.Connection.Execute(lsSQL, , adCmdText)
   
   lsSQL = "SELECT" & _
                  "  sTransNox" & _
                  ", dTransact" & _
                  ", sMailFrom" & _
                  ", sMailToxx" & _
                  ", sMailCCxx" & _
                  ", sMailBCCx" & _
                  ", sSubjectx" & _
                  ", sMailBody" & _
                  ", sAttached" & _
                  ", sSourceCD" & _
                  ", sSourceNo" & _
                  ", cStatusxx" & _
                  ", dPostedxx" & _
                  ", sModified" & _
                  ", dModified" & _
         " FROM Send_Mail_Master" & _
         " WHERE 0=1"
   Set loRS = New Recordset
   loRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   Set loRS.ActiveConnection = Nothing
   
   loRS.AddNew
   
   loRS("dTransact") = oApp.ServerDate
   loRS("sMailFrom") = oApp.getConfiguration("MSCCProcMl")
    Debug.Print loRS("sMailFrom")
   'Send the mail to the supplier
   Select Case oTrans.Master("sSupplier")
   Case "C0W210000162" 'Iridium Technologies Inc.
        loRS("sMailToxx") = oApp.getConfiguration("IridiumMl")
   Case "C0W218000005" 'DMEG
        loRS("sMailToxx") = oApp.getConfiguration("DMEGM1")
   Case "C0W217000004" 'cognetics
        loRS("sMailToxx") = oApp.getConfiguration("CognetcM1")
   Case "C0W115000003" 'Twireless
      Select Case oTrans.Master("sBrandIdx")
      Case "C0W1133" 'vivo
        loRS("sMailToxx") = oApp.getConfiguration("VIVOM1")
      Case "C0W1115" 'huawei
        loRS("sMailToxx") = oApp.getConfiguration("HUAWEIM1")
      Case "C0W1098" 'Oppo
        loRS("sMailToxx") = oApp.getConfiguration("TOPPOM1")
      Case "C0W2256" 'realme
        loRS("sMailToxx") = oApp.getConfiguration("REALMEM1")
      Case "C0W1151" 'Tecno
        loRS("sMailToxx") = oApp.getConfiguration("TECNOM1")
      Case Else
         MsgBox "Please set the Brand to be able to send the PO to specific Supplier!!"
         GoTo endProc
      End Select
   Case "C0W214000024", "C0W114000008" 'oppo
        loRS("sMailToxx") = oApp.getConfiguration("OppoM1")
   Case "C0W121000002" 'EMERALD WIRELESS TECH INC.,
      If oTrans.Master("sBrandIdx") = "C0W1098" Then
         loRS("sMailToxx") = oApp.getConfiguration("EOPPOM1")
      Else
          MsgBox "Auto Email for this supplier is not yet available!!!"
        GoTo endProc
      End If
   Case "M02915002137" 'One Orange Telecommunications Trading
      loRS("sMailToxx") = oApp.getConfiguration("ONERGEM1")
   Case "C0W121000004" 'Xeeme Inc.
      loRS("sMailToxx") = oApp.getConfiguration("XeemeM1")
   Case "C0W121000009" 'Menandro Tagle
      loRS("sMailToxx") = oApp.getConfiguration("MenandM1")
   Case Else
    MsgBox "Auto Email for this supplier is not yet available!!!"
        GoTo endProc
   End Select
   
   loRS("sMailCCxx") = oApp.getConfiguration("MailMPAGM")   'AGM will always received a CC of the PO
   If loRS2("xEmailAdd") <> "" Then
      loRS("sMailCCxx") = "sirabanal@guanzongroup.com.ph" 'lors("sMailCCxx") & ";" & loRS2("xEmailAdd")
   End If
   
   loRS("sMailBCCx") = oApp.getConfiguration("MSMPDptMl")
   loRS("sSubjectx") = "PURCHASE ORDER - " & oTrans.Master("sTransNox")
   loRS("sMailBody") = "Please see attached file..."
   loRS("sAttached") = oTrans.Master("sTransNox") & ".pdf"
   loRS("sSourceCD") = "CPPO"
   loRS("sSourceNo") = oTrans.Master("sTransNox")
      
   loRS("cStatusxx") = "0"
   loRS("sTransNox") = GetNextCode("Send_Mail_Master", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)
   
   lsSQL = ADO2SQL(loRS, "Send_Mail_Master", , oApp.UserID, oApp.ServerDate)
   Debug.Print lsSQL
   oApp.BeginTrans
   
   oApp.Connection.Execute lsSQL, , adCmdText

   'UPLOAD THE FILE TO THE SERVER
   lsSQL = oApp.getConfiguration("MailUplApp") & " " & _
            loRS("sTransNox") & " " & _
            oApp.AppPath & "/Temp/Upload" & " " & _
            oTrans.Master("sTransNox") & ".pdf"
   Debug.Print lsSQL
   'lnRet = ExecCmd(lsSQL)
   Debug.Print "D:/GGC_Java_Systems/ftp_upload.bat " & lsSQL
   lnRet = RMJExecute("D:/GGC_Java_Systems/ftp_upload.bat " & lsSQL)
   
   If lnRet = 0 Then
      lsSQL = "UPDATE Send_Mail_Master" & _
             " SET cStatusxx = '1'" & _
             " WHERE sTransNox = " & strParm(loRS("sTransNox"))
      oApp.Connection.Execute lsSQL, , adCmdText
      
      'update also the CP PO master
      lsSQL = "UPDATE CP_PO_Master" & _
             " SET cEmailSnt = '1'" & _
             " WHERE sTransNox = " & strParm(loRS("sTransNox"))
      oApp.Connection.Execute lsSQL, , adCmdText
      
      oApp.CommitTrans
      MsgBox "Creating of email queue was successfully! PO will be emailed within a few minutes..."
      EmailPO = True
   Else
      oApp.RollbackTrans
      MsgBox "Unable to create email queue! Please try again later...."
      EmailPO = False
   End If
   
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"

End Function


Private Function GenerateReport(Optional ByVal lbOpenFile As Boolean = True) As Boolean
    Dim lsOldProc As String
   Dim lnCtr As Integer
   Dim lrs As ADODB.Recordset
   Dim loRS As Recordset
   Dim lsAddress As String
   Dim lsSQL As String
   Dim lsReferNox As String
   
   lsOldProc = "InitReport"

   Set lrs = New ADODB.Recordset
   With lrs
      .Fields.Append "nField01", adInteger, 3
      .Fields.Append "sField01", adVarChar, 64
      .Fields.Append "sField02", adVarChar, 64, adFldIsNullable
      .Fields.Append "sField03", adVarChar, 64, adFldIsNullable
      .Fields.Append "lField01", adCurrency
      .Open
      
      For lnCtr = 0 To oTrans.ItemCount - 1
         lrs.AddNew
         .Fields("nField01").Value = oTrans.Detail(lnCtr, "nQuantity")
         .Fields("sField01").Value = IFNull(oTrans.Detail(lnCtr, "sModelNme"), "")
         .Fields("sField02").Value = oTrans.Detail(lnCtr, "sColorNme")
         .Fields("sField03").Value = oTrans.Detail(lnCtr, "sModelCde")
         .Fields("lField01").Value = oTrans.Detail(lnCtr, "nUnitPrce")
      Next
   End With
   
   lsSQL = "SELECT CONCAT(a.sAddressx, ', ', b.sTownName, ' ', c.sProvName)" & _
                     " FROM Branch a, TownCity b, Province c" & _
                     " WHERE a.sBranchNm = " & strParm(txtField(3)) & _
                        " AND a.sTownIDxx = b.sTownIDxx AND b.sProvIDxx = c.sProvIDxx"
   Set loRS = New Recordset
   
   With loRS
      .Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If Not .EOF Then
         lsAddress = loRS(0)
      Else
         lsAddress = ""
      End If
   End With
    
    Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\MP-Purchase4Supplier.rpt")
    
   With oReport
      'assign important info to the report
      .DiscardSavedData
      .FieldMappingType = crAutoFieldMapping
      .Database.SetDataSource lrs
      
    .Sections("PH").ReportObjects("txtSupplier").SetText txtField(2)
    .Sections("PH").ReportObjects("txtDeliverTo").SetText txtField(4)
    .Sections("PH").ReportObjects("txtTerm").SetText txtField(9)
    .Sections("PH").ReportObjects("txtDateTransact").SetText txtField(1)
    .Sections("PH").ReportObjects("txtInvoiceTo").SetText oTrans.Master("sCompnyNm")
    .Sections("PH").ReportObjects("txtPONo").SetText oTrans.Master("sTransNox")
    .Sections("RF").ReportObjects("txtRemarks").SetText txtField(5)
    .Sections("PF").ReportObjects("txtRequested").SetText "JULIE MARTINEZ" 'oApp.UserName
      
      With .ExportOptions
         .DestinationType = crEDTDiskFile
         .DiskFileName = oApp.AppPath & "/Temp/Upload/" & oTrans.Master("sTransNox") & ".pdf"
         .FormatType = crEFTPortableDocFormat
         .PDFExportAllPages = True
      End With
   
      .Export False
      
      'open file for printing
      If lbOpenFile Then ShellExecute 0&, "OPEN", .ExportOptions.DiskFileName, "", "", 1
   End With
   
   If oTrans.Master("cTranStat") = xeStateOpen Then
      GenerateReport = oTrans.CloseTransaction
   Else
      GenerateReport = True
   End If
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function


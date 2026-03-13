VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport_Cash 
   BorderStyle     =   0  'None
   Caption         =   "Sales Report"
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2850
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   2205
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   3889
      ClipControls    =   0   'False
      Begin VB.OptionButton optTran 
         Caption         =   "Cheque Trans."
         Height          =   195
         Index           =   0
         Left            =   1725
         TabIndex        =   5
         Tag             =   "et0;fb0"
         Top             =   705
         Width           =   1560
      End
      Begin VB.OptionButton optTran 
         Caption         =   "Credit Card Trans."
         Height          =   195
         Index           =   1
         Left            =   1725
         TabIndex        =   4
         Tag             =   "et0;fb0"
         Top             =   450
         Width           =   1620
      End
      Begin VB.OptionButton optTran 
         Caption         =   "Installment"
         Height          =   195
         Index           =   2
         Left            =   3510
         TabIndex        =   6
         Tag             =   "et0;fb0"
         Top             =   465
         Width           =   1215
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   3000
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1530
         Width           =   1700
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   765
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1530
         Width           =   1700
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Detailed"
         Height          =   210
         Index           =   0
         Left            =   195
         TabIndex        =   1
         Tag             =   "et0;fb0"
         Top             =   450
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Summarized"
         Height          =   225
         Index           =   1
         Left            =   195
         TabIndex        =   2
         Tag             =   "et0;fb0"
         Top             =   705
         Width           =   1170
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
         Left            =   165
         TabIndex        =   7
         Tag             =   "et0;fb0"
         Top             =   1170
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Payment Type"
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
         Left            =   1635
         TabIndex        =   3
         Tag             =   "et0;fb0"
         Top             =   120
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Presentation"
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
         Left            =   150
         TabIndex        =   0
         Tag             =   "et0;fb0"
         Top             =   105
         Width           =   1170
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Index           =   1
         Left            =   75
         Top             =   1230
         Width           =   4755
      End
      Begin VB.Shape Shape1 
         Height          =   885
         Index           =   0
         Left            =   60
         Top             =   195
         Width           =   1470
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Index           =   1
         Left            =   2655
         TabIndex        =   10
         Top             =   1545
         Width           =   270
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Index           =   5
         Left            =   285
         TabIndex        =   8
         Top             =   1530
         Width           =   525
      End
      Begin VB.Shape Shape1 
         Height          =   885
         Index           =   2
         Left            =   1560
         Top             =   195
         Width           =   3270
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   5265
      TabIndex        =   13
      Top             =   1620
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "Pre&view"
      AccessKey       =   "v"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmReport_Cash.frx":0000
      CaptionAlign    =   0
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin MSComCtl2.Animation Progress 
      Height          =   720
      Left            =   5565
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "wt0;wb0"
      Top             =   540
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1270
      _Version        =   393216
      BackColor       =   1842204
      FullWidth       =   61
      FullHeight      =   48
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   5265
      TabIndex        =   14
      Top             =   2040
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
      Picture         =   "frmReport_Cash.frx":1112
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmReport_Cash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Dim psSelected() As String

Dim lrsReport As ADODB.Recordset
Dim lrs As ADODB.Recordset
Dim lrsLabor As ADODB.Recordset
Dim lrsExpense As ADODB.Recordset

Dim lsSQL As String
Dim Address As String
Dim Code As String
Dim Branch As String

Private Sub cmdButton_Click(Index As Integer)
Dim lnctr As Integer
   Select Case Index
      Case 0 'OK
         Progress.Open App.Path & "\images\FINDFILE.AVI"
         Progress.Play
         ReportPreview
      Case 1 'Cancel
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   If bLoaded = False Then
      bLoaded = True
   End If
End Sub

Private Sub Form_Deactivate()
   Progress.Stop
   Progress.Close
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
      
   txtfield(1).Text = Format(oApp.ServerDate, "MMMM dd, yyyy")
   txtfield(2).Text = Format(oApp.ServerDate, "MMMM dd, yyyy")
   
   Check1(0).Value = 1
   Check1(1).Value = 0
   optTran(0).Value = True
         
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   txtfield(Index).BackColor = &HE1FEFF
   oDriver.ColumnIndex = Index
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

Private Sub Check1_Click(Index As Integer)
   If Check1(0).Value = 1 Then
      Check1(1).Value = 0
   ElseIf Check1(0).Value = 0 Then
      Check1(1).Value = 1
   End If
End Sub

Private Function ReportPreview() As Boolean
Dim Index As Integer

ReportPreview = True
On Error GoTo errProc

   Branch = oApp.BranchCode
   getBranch Code, Branch, Address
   
   If optTran(0).Value = True Then
      Cheque_Trans
   ElseIf optTran(1).Value = True Then
      Card_Trans
   ElseIf optTran(2).Value = True Then
      Installment_Trans
   End If
   
endProc:
   Progress.Stop
   Progress.Close
   Exit Function
errProc:
   ReportPreview = False
End Function

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = vbKeyF3 Then
   If Index = 0 Then
      Branch = txtfield(Index)
      getBranch Code, Branch, Address
      txtfield(Index) = Branch
   End If
End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   If Check1(0).Value = 1 Then
      If Index = 1 Then
         txtfield(2).Text = txtfield(1).Text
      End If
   End If
   txtfield(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1, 2
         If Not IsDate(txtfield(Index).Text) Then
            txtfield(Index).Text = Format(oApp.ServerDate, "MMMM dd, yyyy")
         Else
            txtfield(Index).Text = Format(txtfield(Index).Text, "MMMM dd,yyyy")
         End If
   End Select
End Sub

Private Sub Cheque_Trans()
   Dim lnctr As Integer
   Dim lrsReport As ADODB.Recordset
   Dim lrs As ADODB.Recordset

   Set lrs = New ADODB.Recordset
   Set lrsReport = New ADODB.Recordset

   lrs.Fields.Append "dField01", adDBDate
   lrs.Fields.Append "sField01", adVarChar, 150
   lrs.Fields.Append "sField02", adVarChar, 150
   lrs.Fields.Append "sField03", adVarChar, 150
   lrs.Fields.Append "sField04", adVarChar, 150
   lrs.Fields.Append "sField05", adVarChar, 150
   lrs.Fields.Append "sField06", adVarChar, 150
   lrs.Fields.Append "lField01", adCurrency, 10
   lrs.Fields.Append "lField02", adCurrency, 10
   lrs.Fields.Append "lField03", adCurrency, 10
   lrs.Open
                        
   'Cheque Transaction
   lsSQL = " SELECT" _
            & " a.sTransNox, " _
            & " a.dTransact, " _
            & " a.nTranTotl, " _
            & " a.nCashAmnt, " _
            & " a.nCheqAmnt, " _
            & " a.sAccntNum, " _
            & " a.sSalesInv, " _
            & " b.sLastName + ', ' + b.sFrstName + ' ' + b.sMiddName as xFullName, " _
            & " c.sBankName, " _
            & " c.sAddressx + ' ' + d.sTownName xAddressx, " _
            & " e.sRemarksx  " _
         & " FROM CP_SO_Cheque a " _
            & " LEFT JOIN Client_Master b " _
               & " ON a.sClientID = b.sClientID " _
            & " LEFT JOIN Banks c " _
               & " ON a.sBankIDxx = c.sBankIDxx " _
            & " LEFT JOIN TownCity d " _
               & " ON c.sTownIDxx = d.sTownIDxx " _
            & " LEFT JOIN CP_SO_Master e " _
               & " ON a.sTransNox = e.sTransNox " _
         & " WHERE a.dTransact between '" & (txtfield(1).Text) & "' " _
            & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
            & " AND Left(a.sTransNox,2) = '" & oApp.BranchCode & "'" _
         & " ORDER BY a.dTransact "

      If lrsReport.State = adStateOpen Then lrsReport.Close
      lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If lrsReport.EOF Then
         Progress.Stop
         Progress.Close
         MsgBox "No Record Found!!!" & vbCrLf & _
                  "Please Verify your Entry then Try Again!!!", vbCritical, "Information"
         Exit Sub
      End If

      For lnctr = 0 To lrsReport.RecordCount - 1
         lrs.AddNew
         lrs("dField01").Value = Format(lrsReport("dTransact"), "MMM dd, yyyy")
         lrs("sField01").Value = lrsReport("sRemarksx")
         lrs("sField02").Value = lrsReport("sSalesInv")
         Select Case Check1(0).Value
            Case 1   'Detailed
               lrs("sField03").Value = lrsReport("xFullName")
               lrs("sField04").Value = lrsReport("sAccntNum")
               lrs("sField05").Value = IIf(IsNull(lrsReport("sBankName")), "", lrsReport("sBankName"))
               lrs("sField06").Value = lrsReport("xAddressx")
               lrs("lField01").Value = Format(lrsReport("nTranTotl"), "#,##0.00")
               lrs("lField02").Value = Format(lrsReport("nCheqAmnt"), "#,##0.00")
               lrs("lField03").Value = Format(lrsReport("nCashAmnt"), "#,##0.00")
            Case 0   'Summarized
               lrs("sField03").Value = lrsReport("sAccntNum")
               lrs("sField04").Value = IIf(IsNull(lrsReport("sBankName")), "", lrsReport("sBankName"))
               lrs("lField01").Value = Format(lrsReport("nCheqAmnt"), "#,##0.00")
         End Select
         lrsReport.MoveNext
      Next

      If Check1(0).Value = 1 Then
         Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_DetailedChequeTran.rpt")
      Else
         Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_SummarizedChequeTran.rpt")
      End If
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText "Cheque Transaction Summary"
         .Sections("PH").ReportObjects("txtReportDate").SetText _
                        txtfield(1).Text & " - " & txtfield(2).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      End With

      Set lrs = Nothing
      Set lrsReport = Nothing

      frmViewer.CRViewer91.ReportSource = oReport
      frmViewer.CRViewer91.ViewReport
      frmViewer.Show

End Sub

Private Sub Installment_Trans()
   Dim lnctr As Integer
   Dim lrsReport As ADODB.Recordset
   Dim lrs As ADODB.Recordset

   Set lrs = New ADODB.Recordset
   Set lrsReport = New ADODB.Recordset

   lrs.Fields.Append "dField01", adDBDate
   lrs.Fields.Append "sField01", adVarChar, 150
   lrs.Fields.Append "sField02", adVarChar, 150
   lrs.Fields.Append "sField03", adVarChar, 150
   lrs.Fields.Append "lField01", adCurrency, 10
   lrs.Fields.Append "lField02", adCurrency, 10
   lrs.Fields.Append "lField03", adCurrency, 10
   lrs.Fields.Append "lField04", adCurrency, 10
   lrs.Fields.Append "nField01", adInteger, 5
   lrs.Open
                        
      'Installment Transaction
      lsSQL = " SELECT" _
               & " a.sTransNox, " _
               & " a.dTransact, " _
               & " a.nTranTotl, " _
               & " a.nDownPaym, " _
               & " a.nBalancex, " _
               & " a.nPaymTerm, " _
               & " a.nMonthlyP, " _
               & " a.sSalesInv, " _
               & " b.sLastName + ', ' + b.sFrstName + ' ' + b.sMiddName as xFullName, " _
               & " c.sRemarksx  " _
            & " FROM CP_SO_Installment a " _
               & " LEFT JOIN Client_Master b " _
                  & " ON a.sClientID = b.sClientID " _
               & " LEFT JOIN CP_SO_Master c " _
                  & " ON a.sTransNOx = c.sTransNox " _
            & " WHERE Left(a.sTransNox,2) = '" & oApp.BranchCode & "' " _
               & " AND a.dTransact between '" & (txtfield(1).Text) & "' " _
               & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
            & " ORDER BY a.dTransact " _

      If lrsReport.State = adStateOpen Then lrsReport.Close
      lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

      If lrsReport.EOF Then
         Progress.Stop
         Progress.Close
         MsgBox "No Record Found!!!" & vbCrLf & _
                  "Please Verify your Entry then Try Again!!!", vbCritical, "Information"
         Exit Sub
      End If

      For lnctr = 0 To lrsReport.RecordCount - 1
         lrs.AddNew
            lrs("dField01").Value = lrsReport("dTransact")
            lrs("sField01").Value = lrsReport("sRemarksx")
            lrs("sField02").Value = lrsReport("sSalesInv")
            lrs("sField03").Value = lrsReport("xFullName")
            lrs("lField01").Value = Format(lrsReport("nTranTotl"), "#,##0.00")
            lrs("lField02").Value = Format(lrsReport("nDownPaym"), "#,##0.00")
            lrs("lField03").Value = Format(lrsReport("nBalancex"), "#,##0.00")
            lrs("lField04").Value = Format(lrsReport("nMonthlyP"), "#,##0.00")
            lrs("nField01").Value = lrsReport("nPaymTerm")
         
         DoEvents
         lrsReport.MoveNext
      Next

      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_SummarizedInstallment.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText "Installment Transaction Summary"
         .Sections("PH").ReportObjects("txtReportDate").SetText _
                        txtfield(1).Text & " - " & txtfield(2).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      End With

      Set lrs = Nothing
      Set lrsReport = Nothing

      frmViewer.CRViewer91.ReportSource = oReport
      frmViewer.CRViewer91.ViewReport
      frmViewer.Show

End Sub

Private Sub Card_Trans()
   Dim lnctr As Integer
   Dim lrsReport As ADODB.Recordset
   Dim lrs As ADODB.Recordset

   Set lrs = New ADODB.Recordset
   Set lrsReport = New ADODB.Recordset

   lrs.Fields.Append "dField01", adDBDate
   lrs.Fields.Append "sField01", adVarChar, 150
   lrs.Fields.Append "sField02", adVarChar, 150
   lrs.Fields.Append "sField03", adVarChar, 150
   lrs.Fields.Append "sField04", adVarChar, 150
   lrs.Fields.Append "sField05", adVarChar, 150
   lrs.Fields.Append "sField06", adVarChar, 150
   lrs.Fields.Append "lField01", adCurrency, 10
   lrs.Fields.Append "lField02", adCurrency, 10
   lrs.Fields.Append "lField03", adCurrency, 10
   lrs.Fields.Append "lField04", adCurrency, 10
   lrs.Open
                   
      'Credit Card Transaction
      lsSQL = " SELECT" _
               & " a.sTransNox, " _
               & " a.dTransact, " _
               & " a.nTranTotl, " _
               & " a.nCashAmnt, " _
               & " a.nCardAmnt, " _
               & " a.sAcctNmbr, " _
               & " a.nPercentx, " _
               & " a.sSalesInv, " _
               & " b.sLastName + ', ' + b.sFrstName + ' ' + b.sMiddName as xFullName, " _
               & " c.sCreditNm, " _
               & " d.sRemarksx, " _
               & " d.cTranStat, " _
               & " a.nCashTotl  " _
               
      lsSQL = lsSQL _
            & " FROM CP_SO_Credit a " _
               & " LEFT JOIN Client_Master b " _
                  & " ON a.sClientID = b.sClientID " _
               & " LEFT JOIN Credit_Card c " _
                  & " ON a.sCreditID = c.sCreditID " _
               & " LEFT JOIN CP_SO_Master d " _
                  & " ON a.sTransNOx = d.sTransNox " _
            & " WHERE Left(a.sTransNox,2) = '" & oApp.BranchCode & "' " _
               & " AND a.dTransact between '" & (txtfield(1).Text) & "' " _
               & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
               & " AND d.cTranStat = 1 " _
            & " ORDER BY a.dTransact " _

      If lrsReport.State = adStateOpen Then lrsReport.Close
      lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

      If lrsReport.EOF Then
         Progress.Stop
         Progress.Close
         MsgBox "No Record Found!!!" & vbCrLf & _
                  "Please Verify your Entry then Try Again!!!", vbCritical, "Information"
         Exit Sub
      End If

      For lnctr = 0 To lrsReport.RecordCount - 1
         lrs.AddNew
         lrs("dField01").Value = lrsReport("dTransact")
         lrs("sField01").Value = IIf(IsNull(lrsReport("sRemarksx")), "", lrsReport("sRemarksx"))
         lrs("sField02").Value = IIf(IsNull(lrsReport("sSalesInv")), "", lrsReport("sSalesInv"))
         lrs("sField04").Value = IIf(IsNull(lrsReport("sCreditNm")), "", lrsReport("sCreditNm"))
         Select Case Check1(0).Value
            Case 1   'Detailed
               lrs("sField03").Value = IIf(IsNull(lrsReport("xFullName")), "", lrsReport("xFullName"))
               lrs("sField05").Value = IIf(IsNull(lrsReport("sAcctNmbr")), "", lrsReport("sAcctNmbr"))
               lrs("lField01").Value = Format(lrsReport("nTranTotl"), "#,##0.00")
               lrs("lField02").Value = Format(lrsReport("nCardAmnt"), "#,##0.00")
               lrs("lField03").Value = Format(lrsReport("nCashAmnt"), "#,##0.00")
               lrs("lField04").Value = Format(lrsReport("nCashTotl"), "#,##0.00")
            Case 0 'Summarized
               lrs("sField03").Value = IIf(IsNull(lrsReport("xFullName")), "", lrsReport("xFullName"))
               lrs("lField01").Value = Format(lrsReport("nCardAmnt"), "#,##0.00")
         End Select
            
         lrsReport.MoveNext
      Next

      If Check1(0).Value = 1 Then
         Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_DetailedCardTran.rpt")
      Else
         Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_SummarizedCardTran.rpt")
      End If
         
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText "Credit Card Transaction Summary"
         .Sections("PH").ReportObjects("txtReportDate").SetText _
                        txtfield(1).Text & " - " & txtfield(2).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      End With

      Set lrs = Nothing
      Set lrsReport = Nothing

      frmViewer.CRViewer91.ReportSource = oReport
      frmViewer.CRViewer91.ViewReport
      frmViewer.Show

End Sub





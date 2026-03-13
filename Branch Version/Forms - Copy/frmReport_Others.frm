VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport_Others 
   BorderStyle     =   0  'None
   Caption         =   "Reports"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2400
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1770
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   3122
      ClipControls    =   0   'False
      Begin VB.CheckBox Check1 
         Caption         =   "Summarized"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "et0;fb0"
         Top             =   660
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Detailed"
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Tag             =   "et0;fb0"
         Top             =   405
         Width           =   1065
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   765
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1305
         Width           =   1700
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   3000
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1305
         Width           =   1700
      End
      Begin VB.OptionButton optTran 
         Caption         =   "Expense"
         Height          =   195
         Index           =   1
         Left            =   3180
         TabIndex        =   5
         Tag             =   "et0;fb0"
         Top             =   540
         Width           =   1605
      End
      Begin VB.OptionButton optTran 
         Caption         =   "Job Order"
         Height          =   195
         Index           =   0
         Left            =   1785
         TabIndex        =   4
         Tag             =   "et0;fb0"
         Top             =   525
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Transaction"
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
         Left            =   1665
         TabIndex        =   3
         Tag             =   "et0;fb0"
         Top             =   105
         Width           =   1065
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
         TabIndex        =   6
         Tag             =   "et0;fb0"
         Top             =   1050
         Width           =   510
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
         Height          =   780
         Index           =   2
         Left            =   1560
         Top             =   195
         Width           =   3270
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   7
         Top             =   1305
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Index           =   1
         Left            =   2655
         TabIndex        =   9
         Top             =   1320
         Width           =   270
      End
      Begin VB.Shape Shape1 
         Height          =   780
         Index           =   0
         Left            =   60
         Top             =   195
         Width           =   1470
      End
      Begin VB.Shape Shape1 
         Height          =   540
         Index           =   1
         Left            =   75
         Top             =   1110
         Width           =   4755
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   5265
      TabIndex        =   12
      Top             =   1485
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
      Picture         =   "frmReport_Others.frx":0000
      CaptionAlign    =   0
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin MSComCtl2.Animation Progress 
      Height          =   720
      Left            =   5565
      TabIndex        =   11
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
      TabIndex        =   13
      Top             =   1905
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
      Picture         =   "frmReport_Others.frx":1112
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmReport_Others"
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
Dim lnCtr As Integer
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
      
   txtField(1).Text = Format(oApp.ServerDate, "MMMM dd, yyyy")
   txtField(2).Text = Format(oApp.ServerDate, "MMMM dd, yyyy")
   
   Check1(0).Value = 1
   Check1(1).Value = 0
   optTran(0).Value = True
         
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   txtField(Index).BackColor = &HE1FEFF
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
On Error Goto errProc

   Branch = oApp.BranchCode
   getBranch Code, Branch, Address

   If optTran(1).Value = True Then
      If Check1(0).Value = 1 Then 'Detailed
         Detailed_Expense
      Else
         Summarized_Expense
      End If
   ElseIf optTran(0).Value = True Then
      If Check1(0).Value = 1 Then 'Detailed
         Detailed_JO
      Else
         Summarized_JO
      End If
   End If
      
endProc:
   Exit Function
   Progress.Stop
   Progress.Close
errProc:
   ReportPreview = False
End Function


Private Sub txtField_LostFocus(Index As Integer)
   If Check1(0).Value = 1 Then
      If Index = 1 Then
         txtField(2).Text = txtField(1).Text
      End If
   End If
   txtField(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)

Select Case Index
   Case 1, 2
      If Not IsDate(txtField(Index).Text) Then
         txtField(Index).Text = Format(oApp.ServerDate, "MMMM dd, yyyy")
      Else
         txtField(Index).Text = Format(txtField(Index).Text, "MMMM dd,yyyy")
      End If
   End Select
End Sub

Private Sub Detailed_Expense()
   Dim lnCtr As Integer
   Dim lrsReport As ADODB.Recordset
   Dim lrs As ADODB.Recordset

   Set lrs = New ADODB.Recordset
   Set lrsReport = New ADODB.Recordset

   lrs.Fields.Append "sField01", adVarChar, 150
   lrs.Fields.Append "dField01", adDBDate
   lrs.Fields.Append "lField01", adCurrency, 10
   lrs.Open
            
      'Expense
      lsSQL = " SELECT" _
               & " a.sTransNox, " _
               & " a.sDescript, " _
               & " a.nAmountxx, " _
               & " b.dTrandate, " _
               & " b.nTotalExp  " _
            & " FROM CP_Expense_Detail a " _
               & " LEFT JOIN CP_Expense_Master b " _
                  & " ON a.sTransNox = b.sTransNox  " _
            & " WHERE Left(a.sTransNox,2) = '" & oApp.BranchCode & "' " _
               & " AND b.dTrandate between '" & (txtField(1).Text) & "' " _
               & " AND '" & (txtField(2).Text & " 23:59:59") & "'" _
            & " ORDER BY b.dTranDate " _
   
      If lrsReport.State = adStateOpen Then lrsReport.Close
      lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

      If lrsReport.EOF Then
         Progress.Stop
         Progress.Close
         MsgBox "No Record Found!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         Exit Sub
      End If

      For lnCtr = 0 To lrsReport.RecordCount - 1
         lrs.AddNew
         lrs("sField01").Value = lrsReport("sDescript")
         lrs("dField01").Value = lrsReport("dTranDate")
         lrs("lField01").Value = Format(lrsReport("nAmountxx"), "#,##0.00")
         
         lrsReport.MoveNext
      Next

      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_DetailedExpense.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText "Detailed Expense Summary"
         .Sections("PH").ReportObjects("txtReportDate").SetText _
                        txtField(1).Text & " - " & txtField(2).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      End With

      Set lrs = Nothing
      Set lrsReport = Nothing

      frmViewer.CRViewer91.ReportSource = oReport
      frmViewer.CRViewer91.ViewReport
      frmViewer.Show

End Sub

Private Sub Detailed_JO()
   Dim lnCtr As Integer
   Dim lrsReport As ADODB.Recordset
   Dim lrs As ADODB.Recordset

   Set lrs = New ADODB.Recordset
   Set lrsReport = New ADODB.Recordset

   lrs.Fields.Append "dField01", adDBDate
   lrs.Fields.Append "sField02", adVarChar, 150
   lrs.Fields.Append "sField03", adVarChar, 150
   lrs.Fields.Append "sField04", adVarChar, 150
   lrs.Fields.Append "sField05", adVarChar, 150
   lrs.Fields.Append "sField06", adVarChar, 150
   lrs.Fields.Append "sField07", adVarChar, 150
   lrs.Fields.Append "sField08", adVarChar, 150
   lrs.Fields.Append "lField01", adCurrency, 10
   lrs.Fields.Append "lField02", adCurrency, 10
   lrs.Fields.Append "lField03", adCurrency, 10
   lrs.Fields.Append "lField04", adCurrency, 10
   lrs.Open
                        
   lsSQL = "SELECT" _
                  & " a.dTransact, " _
                  & " a.sJobOrdNo, " _
                  & " d.sLastName + ', ' + d.sFrstName + ' ' + d.sMiddName as xFullName, " _
                  & " a.sIMEINoxx, " _
                  & " b.sBrandNme+' '+ c.sModelNme as BrandModel, " _
                  & " a.sTransNox, " _
                  & " a.cTranstat, " _
                  & " a.cWarranty, " _
                  & " a.cCategory, " _
                  & " a.sCategory, " _
                  & " a.sComplent, " _
                  & " a.nTranTotl, " _
                  & " a.sBckJobNo, " _
                  & " a.nMiscChrg, " _
                  & " a.nLaborTot, " _
                  & " a.nPartsTot, " _
                  & " a.nAmtPaidx " _

   lsSQL = lsSQL _
               & " FROM CP_JobOrder_Master a " _
                  & " LEFT JOIN Brand b " _
                     & " ON a.sBrandIDx = b.sBrandIDx " _
                  & " LEFT JOIN Model c " _
                     & " ON a.sModelIDx = c.sModelIDx " _
                  & " LEFT JOIN Client_Master d " _
                     & " ON a.sClientID = d.sClientID " _
               & " WHERE a.dTransact between '" & (txtField(1).Text) & "' " _
                  & " AND '" & (txtField(2).Text & " 23:59:59") & "'" _
                  & " AND Left(a.sTransNox,2) = '" & oApp.BranchCode & "' " _
               & " ORDER BY a.dTransact, a.sJobOrdNo " _

      If lrsReport.State = adStateOpen Then lrsReport.Close
      lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

      If lrsReport.EOF Then
         Progress.Stop
         Progress.Close
         MsgBox "No Record Found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         Exit Sub
      End If

      For lnCtr = 0 To lrsReport.RecordCount - 1
         lrs.AddNew
         lrs("dField01").Value = lrsReport("dTransact")
         lrs("sField02").Value = lrsReport("sJobOrdNo")
         lrs("sField03").Value = IIf(IsNull(lrsReport("xFullName")), "Warranty", lrsReport("xFullName"))
         lrs("sField04").Value = lrsReport("sIMEINoxx")
         lrs("sField05").Value = lrsReport("BrandModel")
         lrs("sField08").Value = lrsReport("sComplent")
         lrs("lField01").Value = Format(lrsReport("nLaborTot"), "#,##0.00")
         lrs("lField02").Value = Format(lrsReport("nPartsTot"), "#,##0.00")
         lrs("lField03").Value = Format(lrsReport("nMiscChrg"), "#,##0.00")
         lrs("lField04").Value = Format(lrsReport("nTranTotl"), "#,##0.00")
         
         Select Case lrsReport("cWarranty")
            Case 1
               lrs("sField06").Value = "Void Warranty"
            Case 2
               lrs("sField06").Value = "Limited Warranty"
            Case 3
               lrs("sField06").Value = "Back Job" & " " & lrsReport("sBckJobNo")
         End Select
         
         Select Case lrsReport("cTranStat")
            Case 0
               lrs("sField07").Value = "Pending"
            Case 1
               lrs("sField07").Value = "Claimed"
         End Select
         
         lrsReport.MoveNext
      Next

      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_DetailedJobOrder.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText "Detailed Job Order Summary"
         .Sections("PH").ReportObjects("txtReportDate").SetText _
                        txtField(1).Text & " - " & txtField(2).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      End With

      Set lrs = Nothing
      Set lrsReport = Nothing

      frmViewer.CRViewer91.ReportSource = oReport
      frmViewer.CRViewer91.ViewReport
      frmViewer.Show

End Sub

Private Sub Summarized_Expense()
   Dim lnCtr As Integer
   Dim lrsReport As ADODB.Recordset
   Dim lrs As ADODB.Recordset

   Set lrs = New ADODB.Recordset
   Set lrsReport = New ADODB.Recordset

   lrs.Fields.Append "dField01", adDBDate
   lrs.Fields.Append "lField01", adCurrency, 10
   lrs.Open
            
      'Expense
      lsSQL = " SELECT" _
               & " sTransNox, " _
               & " dTrandate, " _
               & " nTotalExp  " _
            & " FROM CP_Expense_Master " _
            & " WHERE Left(sTransNox,2) = '" & oApp.BranchCode & "' " _
               & " AND dTrandate between '" & (txtField(1).Text) & "' " _
               & " AND '" & (txtField(2).Text & " 23:59:59") & "'" _
            & " ORDER BY dTranDate " _

      If lrsReport.State = adStateOpen Then lrsReport.Close
      lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

      If lrsReport.EOF Then
         Progress.Stop
         Progress.Close
         MsgBox "No Record Found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         Exit Sub
      End If

      For lnCtr = 0 To lrsReport.RecordCount - 1
         lrs.AddNew
         lrs("dField01").Value = lrsReport("dTranDate")
         lrs("lField01").Value = Format(lrsReport("nTotalExp"), "#,##0.00")
         
         lrsReport.MoveNext
      Next

      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_SummarizedExpenseSummary.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText "Expense Summary"
         .Sections("PH").ReportObjects("txtReportDate").SetText _
                        txtField(1).Text & " - " & txtField(2).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      End With

      Set lrs = Nothing
      Set lrsReport = Nothing

      frmViewer.CRViewer91.ReportSource = oReport
      frmViewer.CRViewer91.ViewReport
      frmViewer.Show

End Sub

Private Sub Summarized_JO()
   Dim lnCtr As Integer
   Dim lrsReport As ADODB.Recordset
   Dim lrs As ADODB.Recordset

   Set lrs = New ADODB.Recordset
   Set lrsReport = New ADODB.Recordset

   lrs.Fields.Append "dField01", adDBDate
   lrs.Fields.Append "sField01", adVarChar, 20
   lrs.Fields.Append "sField02", adVarChar, 150
   lrs.Fields.Append "sField03", adVarChar, 50
   lrs.Fields.Append "sField04", adVarChar, 20
   lrs.Fields.Append "sField05", adVarChar, 150
   lrs.Fields.Append "sField06", adVarChar, 150
   lrs.Fields.Append "lField01", adCurrency, 10
   lrs.Fields.Append "lField02", adCurrency, 10
   lrs.Fields.Append "lField03", adCurrency, 10
   lrs.Open
                        
   lsSQL = "SELECT" _
                  & " a.dPaymentx, " _
                  & " a.sJobOrdNo, " _
                  & " b.sLastName, " _
                  & " b.sFrstName, " _
                  & " b.sMiddName, " _
                  & " c.sBrandNme, " _
                  & " d.sModelNme, " _
                  & " a.sIMEINoxx, " _
                  & " a.sComplent, " _
                  & " a.nLaborTot, " _
                  & " a.nPartsTot, " _
                  & " a.cCategory  " _

   lsSQL = lsSQL _
               & " FROM CP_JobOrder_Master a " _
                  & " LEFT JOIN Client_Master b " _
                     & " ON a.sClientID = b.sClientID " _
                  & " LEFT JOIN Brand c " _
                     & " ON a.sBrandIdx = c.sBrandIDx " _
                  & " LEFT JOIN Model d " _
                     & " ON a.sModelIDx = d.sModelIDx " _
               & " WHERE a.dPaymentx between '" & (txtField(1).Text) & "' " _
                  & " AND '" & (txtField(2).Text & " 23:59:59") & "'" _
                  & " AND Left(a.sTransNox,2) = '" & oApp.BranchCode & "' " _
                  & " AND a.cCategory <> 5 " _
               & " ORDER BY a.dTransact, a.sJobOrdNo " _

   If lrsReport.State = adStateOpen Then lrsReport.Close
   lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If lrsReport.EOF Then
      Progress.Stop
      Progress.Close
      MsgBox "No Record Found!!!" & vbCrLf & _
          "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      Exit Sub
   End If


      For lnCtr = 0 To lrsReport.RecordCount - 1
         lrs.AddNew
         lrs("dField01").Value = lrsReport("dPaymentx")
         lrs("sField01").Value = lrsReport("sJobOrdNo")
         lrs("sField02").Value = Trim(IIf(IsNull(lrsReport("sLastName")), "", lrsReport("sLastName")) & ", " & _
                                 IIf(IsNull(lrsReport("sFrstName")), "", lrsReport("sFrstName")))
         lrs("sField03").Value = Trim(IIf(IsNull(lrsReport("sBrandNme")), "", lrsReport("sBrandNme")) & " " & _
                                 IIf(IsNull(lrsReport("sModelNme")), "", lrsReport("sModelNme")))
         lrs("sField04").Value = lrsReport("sIMEINoxx")
         lrs("sField05").Value = lrsReport("sComplent")
         lrs("lField01").Value = Format(lrsReport("nLaborTot"), "#,##0.00")
         lrs("lField02").Value = Format(lrsReport("nPartsTot"), "#,##0.00")
         lrs("lField03").Value = Format(lrsReport("nLaborTot") - lrsReport("nPartsTot"), "#,##0.00")
         
         lrsReport.MoveNext
      Next

      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_SummarizedJobOrder.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText "Job Order Summary"
         .Sections("PH").ReportObjects("txtReportDate").SetText _
                        txtField(1).Text & " - " & txtField(2).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      End With

      Set lrs = Nothing
      Set lrsReport = Nothing

      frmViewer.CRViewer91.ReportSource = oReport
      frmViewer.CRViewer91.ViewReport
      frmViewer.Show

End Sub




VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport_Weekly 
   BorderStyle     =   0  'None
   Caption         =   "Sales Report"
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1770
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   3122
      ClipControls    =   0   'False
      Begin VB.OptionButton optTran 
         Caption         =   "Summarized Sales Report"
         Height          =   450
         Index           =   2
         Left            =   2910
         TabIndex        =   3
         Tag             =   "et0;fb0"
         Top             =   390
         Width           =   1860
      End
      Begin VB.OptionButton optTran 
         Caption         =   "Weekly Inventory Sales Report"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Tag             =   "et0;fb0"
         Top             =   390
         Width           =   2655
      End
      Begin VB.OptionButton optTran 
         Caption         =   "Staff Sales Report"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Tag             =   "et0;fb0"
         Top             =   660
         Width           =   1605
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   3000
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1305
         Width           =   1700
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   765
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1305
         Width           =   1700
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Report Name"
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
         Left            =   135
         TabIndex        =   0
         Tag             =   "et0;fb0"
         Top             =   105
         Width           =   1320
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
         TabIndex        =   4
         Tag             =   "et0;fb0"
         Top             =   1050
         Width           =   510
      End
      Begin VB.Shape Shape1 
         Height          =   540
         Index           =   1
         Left            =   75
         Top             =   1110
         Width           =   4755
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Index           =   1
         Left            =   2655
         TabIndex        =   7
         Top             =   1320
         Width           =   270
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   5
         Top             =   1305
         Width           =   525
      End
      Begin VB.Shape Shape1 
         Height          =   780
         Index           =   2
         Left            =   60
         Top             =   195
         Width           =   4770
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   5280
      TabIndex        =   9
      Top             =   1515
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
      Picture         =   "frmReport_Weekly.frx":0000
      CaptionAlign    =   0
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin MSComCtl2.Animation Progress 
      Height          =   720
      Left            =   5580
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "wt0;wb0"
      Top             =   465
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
      Left            =   5280
      TabIndex        =   10
      Top             =   1935
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
      Picture         =   "frmReport_Weekly.frx":1112
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmReport_Weekly"
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
      Code = oApp.BranchCode
      optTran(0).SetFocus
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

Private Function ReportPreview() As Boolean
Dim Index As Integer

ReportPreview = True
On Error Goto errProc

   Branch = oApp.BranchCode
   getBranch Code, Branch, Address

   If optTran(0).Value = True Then
      Weekly_Inventory
   ElseIf optTran(1).Value = True Then
      Staff_Sales
   ElseIf optTran(2).Value = True Then
      Monthly_TranSummary
   End If
      
endProc:
   Exit Function
   Progress.Stop
   Progress.Close
errProc:
   ReportPreview = False
End Function

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &HFFFFFF
   If Index = 1 Then
      txtfield(2).Text = Format(DateAdd("d", 5, CDate(txtfield(1).Text)), "MMMM dd,yyyy")
   End If
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

Private Sub Weekly_Inventory()
Dim lnCtr As Integer
Dim lrsReport As ADODB.Recordset
Dim lrs As ADODB.Recordset
Dim nSRP As Double
Dim nQTY As Integer
Dim nAmount As Double
Dim sItem As String
Dim sDesc As String
Dim dDate As Date


   Set lrs = New ADODB.Recordset
   Set lrsReport = New ADODB.Recordset

   lrs.Fields.Append "dField01", adDBDate
   lrs.Fields.Append "sField01", adVarChar, 250
   lrs.Fields.Append "sField03", adVarChar, 250
   lrs.Fields.Append "lField01", adCurrency, 10
   lrs.Fields.Append "lField02", adCurrency, 10
   lrs.Fields.Append "nField01", adInteger, 5
   lrs.Open
            
      'SO
      lsSQL = "SELECT" _
                  & " a.nQuantity, " _
                  & " b.dTransact, " _
                  & " h.sCategNme, " _
                  & " d.sBrandNme, " _
                  & " e.sModelnme, " _
                  & " c.sDescript, " _
                  & " g.sColorNme, " _
                  & " a.nUnitPrce, " _
                  & " a.nSubTotal  " _
               & " FROM CP_SO_Detail a " _
                  & " LEFT JOIN CP_SO_Master b" _
                     & " ON a.sTransNox = b.sTransNox" _
                  & " LEFT JOIN CP_Inventory c " _
                     & " ON a.sStockIDx = c.sStockIDx " _
                  & " LEFT JOIN Brand d " _
                     & " ON c.sBrandIDx = d.sBrandIDx " _
                  & " LEFT JOIN MOdel e " _
                     & " ON c.sModelIDx = e.sModelIDx " _
                  & " LEFT JOIN Color f " _
                     & " ON c.sColorIDx = f.sColorIDx " _
                  & " LEFT JOIN Color g " _
                     & " ON c.sColorIDx = g.sColorIDx " _
                  & " LEFT JOIN Category h " _
                     & " ON c.sCategIDx = h.sCategIDx "
                     
      lsSQL = lsSQL _
               & " WHERE Left(b.sTransNox,2) = '" & Code & "' " _
                  & " AND b.dTransact between '" & (txtfield(1).Text) & "' " _
                  & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
                  & " AND h.cIncentiv = 1 " _
                  & " AND b.cTranStat <> 4 " _
               & " ORDER BY b.dTransact, d.sBrandNme, e.sModelNme, c.sDescript, g.sColorNme "
   
      If lrsReport.State = adStateOpen Then lrsReport.Close
      lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

      If lrsReport.State = adStateOpen Then lrsReport.Close
      lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If lrsReport.EOF Then
         Progress.Stop
         Progress.Close
         MsgBox "No Record Found!!!" & vbCrLf & _
                  "Please Verify your Entry then Try Again!!!", vbCritical, "Information"
         Exit Sub
      End If

      nSRP = 0#
      nAmount = 0#
      sItem = ""
      sDesc = ""
      dDate = Now()

      For lnCtr = 0 To lrsReport.RecordCount - 1
         sDesc = Trim(IIf(IsNull(lrsReport("sBrandNme")), "", lrsReport("sBrandNme")) + _
               " " + IIf(IsNull(lrsReport("sModelNme")), "", lrsReport("sModelNme")) + _
               " " + IIf(IsNull(lrsReport("sDescript")), "", lrsReport("sDescript")) + _
               " " + IIf(IsNull(lrsReport("sColorNme")), "", lrsReport("sColorNme")))

         If Trim(sDesc) = sItem And lrsReport("nUnitPrce") = nSRP And _
            lrsReport("nSubTotal") = nAmount And Format(lrsReport("dTransact"), "MM/dd/yy") = Format(CDate(dDate), "MM/dd/yy") Then
            lrs("nField01").Value = nQTY + 1
            nQTY = lrs("nField01").Value
         Else
            lrs.AddNew
            nQTY = 0
            lrs("dField01").Value = Format(lrsReport("dTransact"), "MMMM dd, yyyy")
            lrs("sField01").Value = Trim(IIf(IsNull(lrsReport("sBrandNme")), "", lrsReport("sBrandNme")) + _
                        " " + IIf(IsNull(lrsReport("sModelNme")), "", lrsReport("sModelNme")) + _
                        " " + IIf(IsNull(lrsReport("sDescript")), "", lrsReport("sDescript")) + _
                        " " + IIf(IsNull(lrsReport("sColorNme")), "", lrsReport("sColorNme")))
            lrs("sField03").Value = Trim(lrsReport("sCategNme"))
            lrs("lField01").Value = Format(lrsReport("nUnitPrce"), "#,##0.00")
            lrs("lField02").Value = Format(lrsReport("nSubTotal"), "#,##0.00")
            lrs("nField01").Value = nQTY + 1
            nQTY = lrs("nField01").Value
            nSRP = lrsReport("nUnitPrce")
            nAmount = lrsReport("nSubTotal")
            sItem = Trim(IIf(IsNull(lrsReport("sBrandNme")), "", lrsReport("sBrandNme")) + _
                        " " + IIf(IsNull(lrsReport("sModelNme")), "", lrsReport("sModelNme")) + _
                        " " + IIf(IsNull(lrsReport("sDescript")), "", lrsReport("sDescript")) + _
                        " " + IIf(IsNull(lrsReport("sColorNme")), "", lrsReport("sColorNme")))
            dDate = Format(lrsReport("dTransact"), "MM/dd/yy")
         End If
         lrs("lField02").Value = Format(lrs("nField01") * lrsReport("nSubTotal"), "#,##0.00")
         lrsReport.MoveNext
      Next

      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\Weekly_Inventory_Sales_Report.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("PH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText "Weekly Inventory Sales Report"
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

Private Sub Staff_Sales()
Dim lnCtr As Integer
Dim lrsReport As ADODB.Recordset
Dim lrs As ADODB.Recordset
Dim nQTY As Integer
Dim sItem As String
Dim sDesc As String
Dim sName As String
Dim dDate As Date


   Set lrs = New ADODB.Recordset
   Set lrsReport = New ADODB.Recordset

   lrs.Fields.Append "dField01", adDBDate
   lrs.Fields.Append "sField01", adVarChar, 250
   lrs.Fields.Append "sField02", adVarChar, 250
   lrs.Fields.Append "sField03", adVarChar, 250
   lrs.Fields.Append "nField01", adInteger, 5
   lrs.Open
            
      'SO
      lsSQL = "SELECT" _
                  & " b.sSalesInv, " _
                  & " a.nQuantity, " _
                  & " b.dTransact, " _
                  & " d.sBrandNme, " _
                  & " e.sModelnme, " _
                  & " c.sDescript, " _
                  & " g.sColorNme, " _
                  & " f.sLastName + ', ' + f.sFrstName as sSaleName, " _
                  & " h.sCategNme " _
               & " FROM CP_SO_Detail a " _
                  & " LEFT JOIN CP_SO_Master b" _
                     & " ON a.sTransNox = b.sTransNox" _
                  & " LEFT JOIN CP_Inventory c " _
                     & " ON a.sStockIDx = c.sStockIDx " _
                  & " LEFT JOIN Brand d " _
                     & " ON c.sBrandIDx = d.sBrandIDx " _
                  & " LEFT JOIN Model e " _
                     & " ON c.sModelIDx = e.sModelIDx " _
                  & " LEFT JOIN Sales_Person f " _
                     & " ON b.sCashierx = f.sEmployID " _
                  & " LEFT JOIN Color g " _
                     & " ON c.sColorIDx = g.sColorIDx " _
                  & " LEFT JOIN Category h " _
                     & " ON c.sCategIDx = h.sCategIDx "
                     
      lsSQL = lsSQL _
               & " WHERE Left(b.sTransNox,2) = '" & Code & "' " _
                  & " AND b.dTransact between '" & (txtfield(1).Text) & "' " _
                  & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
                  & " AND h.cIncentiv = 1 " _
                  & " AND b.cTranStat <> 4 " _
               & " ORDER BY sSaleName, b.dTransact, d.sBrandNme, e.sModelNme, c.sDescript, g.sColorNme  "
      
      If lrsReport.State = adStateOpen Then lrsReport.Close
      lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      
      If lrsReport.State = adStateOpen Then lrsReport.Close
      lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If lrsReport.EOF Then
         Progress.Stop
         Progress.Close
         MsgBox "No Record Found!!!" & vbCrLf & _
                  "Please Verify your Entry then Try Again!!!", vbCritical, "Information"
         Exit Sub
      End If

      sItem = ""
      sName = ""
      sDesc = ""
      dDate = Now()

      For lnCtr = 0 To lrsReport.RecordCount - 1
         sDesc = Trim(IIf(IsNull(lrsReport("sBrandNme")), "", lrsReport("sBrandNme")) + _
                        " " + IIf(IsNull(lrsReport("sModelNme")), "", lrsReport("sModelNme")) + _
                        " " + IIf(IsNull(lrsReport("sDescript")), "", lrsReport("sDescript")) + _
                        " " + IIf(IsNull(lrsReport("sColorNme")), "", lrsReport("sColorNme")))
         If Trim(sDesc) = sItem And Trim(lrsReport("sSaleName")) = sName _
            And Format(lrsReport("dTransact"), "MM/dd/yy") = Format(CDate(dDate), "MM/dd/yy") Then
            lrs("nField01").Value = nQTY + 1
            nQTY = lrs("nField01").Value
         Else
            lrs.AddNew
            nQTY = 0
            If IsNull(lrsReport("sSaleName")) Then
               MsgBox "Check Invoice No." & " " & lrsReport("sSalesInv") & vbCrLf & vbCrLf & _
               " Update Transaction Sales Person ", vbInformation, "Notice"
               frmPOS_Register.txtfield(2).Text = lrsReport("sSalesInv")
               frmPOS_Register.Show
               Exit Sub
            End If
            lrs("dField01").Value = Format(lrsReport("dTransact"), "MMMM dd, yyyy")
            lrs("sField01").Value = Trim(IIf(IsNull(lrsReport("sBrandNme")), "", lrsReport("sBrandNme")) + _
                        " " + IIf(IsNull(lrsReport("sModelNme")), "", lrsReport("sModelNme")) + _
                        " " + IIf(IsNull(lrsReport("sDescript")), "", lrsReport("sDescript")) + _
                        " " + IIf(IsNull(lrsReport("sColorNme")), "", lrsReport("sColorNme")))
            lrs("sField03").Value = Trim(lrsReport("sCategNme"))
            lrs("sField02").Value = IIf(IsNull(Trim(lrsReport("sSaleName"))), "", Trim(lrsReport("sSaleName")))
            lrs("nField01").Value = nQTY + 1
            nQTY = lrs("nField01").Value
            sItem = Trim(IIf(IsNull(lrsReport("sBrandNme")), "", lrsReport("sBrandNme")) + _
                        " " + IIf(IsNull(lrsReport("sModelNme")), "", lrsReport("sModelNme")) + _
                        " " + IIf(IsNull(lrsReport("sDescript")), "", lrsReport("sDescript")) + _
                        " " + IIf(IsNull(lrsReport("sColorNme")), "", lrsReport("sColorNme")))
            sName = IIf(IsNull(Trim(lrsReport("sSaleName"))), "", Trim(lrsReport("sSaleName")))
            dDate = Format(lrsReport("dTransact"), "MM/dd/yy")
         End If
         lrsReport.MoveNext
      Next

      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\Staff_Sales_Report.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs
      
      With oReport
         .Sections("PH").ReportObjects("txtReportName").SetText "Staff Sales Report"
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

Private Sub Monthly_TranSummary()
   Dim lnCtr As Integer
   Dim lrsReport As ADODB.Recordset
   Dim lrs As ADODB.Recordset
   Dim lsTransNox As String
   Dim lnCashTotl As Double

   Set lrs = New ADODB.Recordset
   Set lrsReport = New ADODB.Recordset

   lrs.Fields.Append "sField01", adVarChar, 150
   lrs.Fields.Append "nField01", adInteger, 5         'Date
   lrs.Fields.Append "nField02", adInteger, 5         'Units Sold
   lrs.Fields.Append "lField01", adCurrency, 10
   lrs.Fields.Append "lField02", adCurrency, 10
   lrs.Fields.Append "lField03", adCurrency, 10
   lrs.Fields.Append "lField04", adCurrency, 10
   lrs.Fields.Append "lField05", adCurrency, 10
   lrs.Fields.Append "lField06", adCurrency, 10
   lrs.Fields.Append "lField07", adCurrency, 10
   lrs.Fields.Append "lField08", adCurrency, 10
   lrs.Fields.Append "lField09", adCurrency, 10
   lrs.Fields.Append "lField10", adCurrency, 10
   lrs.Fields.Append "lField11", adCurrency, 10
   lrs.Fields.Append "lField12", adCurrency, 10
   lrs.Open

         'SO
         lsSQL = "SELECT" _
                     & " a.sTransNox as xTransNox, " _
                     & " a.nEntryNox, " _
                     & " a.nQuantity, " _
                     & " a.nSubTotal as xxPricexx," _
                     & " b.dTransact as xTranDate, " _
                     & " c.sCategIDx, " _
                     & " b.nAmtPaidx as xAmtPaidx, " _
                     & " 'Sales' xTranType " _

         lsSQL = lsSQL _
                  & " FROM CP_SO_Detail a " _
                     & " LEFT JOIN CP_SO_Master b" _
                        & " ON a.sTransNox = b.sTransNox" _
                     & " LEFT JOIN CP_Inventory c " _
                        & " ON a.sStockIDx = c.sStockIDx " _
                  & " WHERE Left(a.sTransNox,2) = '" & Code & "' " _
                     & " AND b.dTransact between '" & (txtfield(1).Text) & "' " _
                     & " AND '" & (txtfield(2).Text & " 23:59:59") & "'"

         'JO
         lsSQL = lsSQL _
                  & " UNION " _
                  & " SELECT" _
                     & " a.sTransNox as xTransNox, " _
                     & " 1 nEntryNox, " _
                     & " 1 nQuantity, " _
                     & " a.nTranTotl as xxPricexx," _
                     & " a.dPaymentx as xTranDate, " _
                     & "'a' sCategIDx, " _
                     & " a.nAmtPaidx as xAmtPaidx, " _
                     & " 'Repair' xTranType " _

         lsSQL = lsSQL _
                  & " FROM CP_JobOrder_Master a " _
                  & " WHERE Left(a.sTransNox,2) = '" & Code & "' " _
                     & " AND a.dPaymentx between '" & (txtfield(1).Text) & "' " _
                     & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
                     & " AND cTranStat = 1 " _

         'Expense
         lsSQL = lsSQL _
                  & " UNION " _
                  & " SELECT" _
                     & " a.sTransNox as xTransNox, " _
                     & " 1 nEntryNox, " _
                     & " 1 nQuantity, " _
                     & " a.nTotalExp as xxPricexx, " _
                     & " a.dTrandate as xTranDate, " _
                     & "'a' sCategIDx, " _
                     & " 0 xAmtPaidx, " _
                     & " 'Expense' xTranType " _

         lsSQL = lsSQL _
                  & " FROM CP_Expense_Master a " _
                  & " WHERE Left(a.sTransNox,2) = '" & Code & "' " _
                     & " AND a.dTrandate between '" & (txtfield(1).Text) & "' " _
                     & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _

         If lrsReport.State = adStateOpen Then lrsReport.Close
         lrsReport.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText


   lsTransNox = ""
   lnCashTotl = 0#
   For lnCtr = 0 To lrsReport.RecordCount - 1
      lrs.AddNew
      lrs("nField01").Value = Format(lrsReport("xTranDate"), "dd")
         
      Select Case lrsReport("xTranType")
         Case "Sales"
            Select Case lrsReport("sCategIDx")
               Case "01001" 'Cellphone
                  lrs("nField02").Value = IIf(IsNull(lrsReport("nQuantity")), 0, lrsReport("nQuantity"))
                  lrs("lField01").Value = IIf(Not IsNull(lrsReport("xxPricexx")), _
                                          Format(lrsReport("xxPricexx"), "#,##0.00"), 0)
               Case "01002" 'Mic
                  lrs("lField09").Value = IIf(Not IsNull(lrsReport("xxPricexx")), _
                                          Format(lrsReport("xxPricexx"), "#,##0.00"), 0)
               Case "01004" 'MP3
                  lrs("lField08").Value = IIf(Not IsNull(lrsReport("xxPricexx")), _
                                          Format(lrsReport("xxPricexx"), "#,##0.00"), 0)
               Case "01005", "01011" 'Accesories
                  lrs("lField02").Value = IIf(Not IsNull(lrsReport("xxPricexx")), _
                                          Format(lrsReport("xxPricexx"), "#,##0.00"), 0)
               Case "01003", "01006" 'CellCard, Sim Card
                  lrs("lField03").Value = IIf(Not IsNull(lrsReport("xxPricexx")), _
                                          Format(lrsReport("xxPricexx"), "#,##0.00"), 0)
               Case "01007", "01009" 'Load Retail
                  lrs("lField04").Value = IIf(Not IsNull(lrsReport("xxPricexx")), _
                                          Format(lrsReport("xxPricexx"), "#,##0.00"), 0)
               Case "01008", "01010" 'Load Wallet
                  lrs("lField05").Value = IIf(Not IsNull(lrsReport("xxPricexx")), _
                                          Format(lrsReport("xxPricexx"), "#,##0.00"), 0)
            End Select
         Case "Repair"
               lrs("lField06").Value = IIf(lrsReport("xxPricexx") <> "", _
                                       Format(lrsReport("xxPricexx"), "#,##0.00"), 0)
         Case "Expense"
               lrs("lField07").Value = IIf(Not IsNull(lrsReport("xxPricexx")), _
                                       Format(lrsReport("xxPricexx"), "#,##0.00"), 0)
      End Select
      If IsNull(lrs("lfield07").Value) Or lrs("lfield07").Value = "" Then lrs("lfield07") = 0
      lrs("lField10").Value = Format(CDbl(lrs("lfield01").Value + lrs("lfield02").Value + _
                              lrs("lfield03").Value + lrs("lfield04").Value + _
                              lrs("lfield05").Value + lrs("lfield06").Value + _
                              lrs("lfield08").Value + lrs("lfield09").Value), "#,##0.00")
      If lsTransNox <> lrsReport("xTransNox") & " " & lrsReport("xTranType") & " " & _
                        Format(lrsReport("xTranDate"), "dd") Then
         lnCashTotl = lrsReport("xAmtPaidx")
         lsTransNox = lrsReport("xTransNox") & " " & lrsReport("xTranType") & " " & _
                        Format(lrsReport("xTranDate"), "dd")
      Else
         lnCashTotl = 0
      End If
      lrs("lfield12").Value = lnCashTotl
      lrsReport.MoveNext
   Next

      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_MonthlyTransaction.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText "Sales Transaction Summary Report"
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




VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport_Sales 
   BorderStyle     =   0  'None
   Caption         =   "Sales Report"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1560
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1335
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   2752
      ClipControls    =   0   'False
      Begin VB.OptionButton Option1 
         Caption         =   "Summarized Report"
         Height          =   195
         Index           =   1
         Left            =   2685
         TabIndex        =   4
         Tag             =   "et0;fb0"
         Top             =   390
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Detailed Report"
         Height          =   195
         Index           =   0
         Left            =   855
         TabIndex        =   3
         Tag             =   "et0;fb0"
         Top             =   390
         Width           =   1725
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   900
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1095
         Width           =   1700
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   3000
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1095
         Width           =   1700
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   60
         TabIndex        =   14
         Top             =   2175
         Visible         =   0   'False
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
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
         TabIndex        =   5
         Tag             =   "et0;fb0"
         Top             =   810
         Width           =   390
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
         TabIndex        =   2
         Tag             =   "et0;fb0"
         Top             =   75
         Width           =   1170
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
         TabIndex        =   15
         Tag             =   "et0;fb0"
         Top             =   810
         Width           =   630
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Index           =   5
         Left            =   345
         TabIndex        =   6
         Top             =   1095
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Index           =   1
         Left            =   2685
         TabIndex        =   8
         Top             =   1110
         Width           =   270
      End
      Begin VB.Shape Shape1 
         Height          =   540
         Index           =   0
         Left            =   60
         Top             =   180
         Width           =   4770
      End
      Begin VB.Shape Shape1 
         Height          =   540
         Index           =   1
         Left            =   75
         Top             =   900
         Width           =   4755
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   5265
      TabIndex        =   11
      Top             =   1200
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
      Picture         =   "frmReport_Sales.frx":0000
      CaptionAlign    =   0
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   5265
      TabIndex        =   12
      Top             =   1620
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmReport_Sales.frx":1112
      CaptionAlign    =   0
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   705
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   1244
      Begin xrControl.xrFrame xrFrame2 
         Height          =   480
         Index           =   0
         Left            =   75
         Tag             =   "wt0;wb0"
         Top             =   90
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   847
         Begin VB.TextBox txtfield 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   0
            Left            =   990
            MaxLength       =   50
            TabIndex        =   1
            Top             =   105
            Width           =   3630
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
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
            Index           =   6
            Left            =   120
            TabIndex        =   0
            Tag             =   "ebo"
            Top             =   120
            Width           =   885
         End
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   5265
      TabIndex        =   13
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
      Picture         =   "frmReport_Sales.frx":188C
      CaptionAlign    =   0
   End
   Begin MSComCtl2.Animation Progress 
      Height          =   720
      Left            =   5550
      TabIndex        =   10
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
End
Attribute VB_Name = "frmReport_Sales"
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

Dim lrs As New ADODB.Recordset

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
      Case 1
         SearchCategory False
      Case 2 'Cancel
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   If bLoaded = False Then
      bLoaded = True
      Option1(0).Value = True
      txtfield(2).Locked = True
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
            
End Sub

Private Sub SearchCategory(ByVal SearchValue As Boolean)
   Dim lsSearch As String
   Dim lsSQL As String
   Dim oRS As ADODB.Recordset

   Set oRS = New ADODB.Recordset
   
   'Category
   lsSQL = "SELECT" _
               & " sCategory, " _
               & " sCatDescx  " _
         & " FROM Category_Master " _
         & " WHERE cRecdStat = " & strParm(xeRecStateActive) _

   If SearchValue Then
      lsSQL = lsSQL & " AND sCatDescx = '" & txtfield(0).Text & "'"
   Else
      lsSQL = lsSQL & " AND sCatDescx LIKE '" & txtfield(0).Text & "%' "
   End If
            
   lsSQL = lsSQL & " ORDER BY sCatDescx"
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If oRS.RecordCount = 1 Then
      txtfield(0).Text = oRS(1)
      txtfield(0).Tag = oRS(0)
   ElseIf oRS.RecordCount > 1 Then
        lsSearch = KwikBrowse(oApp, oRS, _
                          "sCatDescx", _
                          "Category")
                        
        If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            txtfield(0).Text = psSelected(1)
            txtfield(0).Tag = psSelected(0)
        End If
   Else
      txtfield(0).Text = ""
      txtfield(0).Tag = ""
      txtfield(0).SetFocus
   End If
   Set oRS = Nothing
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub Option1_LostFocus(Index As Integer)
   Select Case Index
      Case 0
         If Option1(Index).Value = True Then txtfield(2).Locked = True
      Case 1
         If Option1(Index).Value = True Then txtfield(2).Locked = False
   End Select
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

   Select Case Option1(0).Value
      Case 0 'Summarized
         Select Case txtfield(0).Tag
         Case ""
            DCPR_Summary
         Case Else
            Sales_Summary
         End Select
      Case 1 'Detailed
         Select Case txtfield(0).Tag
            Case ""
               DCPR_Details
            Case Else
               Sales_Details
         End Select
   End Select
            
endProc:
   Progress.Stop
   Progress.Close
   Exit Function
errProc:
   ReportPreview = False
End Function

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Or KeyCode = vbKeyF3 Then
      Select Case Index
      Case 0
         SearchCategory False
      End Select
      If txtfield(Index).Text = "" Then SetNextFocus
      KeyCode = 0
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &HFFFFFF
   Select Case Index
      Case 0
         If txtfield(Index).Text = "" Then txtfield(Index).Tag = ""
      Case 1
         If Option1(0).Value = True Then
            txtfield(2).Text = txtfield(1).Text
         End If
   End Select
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

Private Sub DCPR_Details()
Dim lnCtrDetail As Integer
Dim lorsDetail As ADODB.Recordset
Dim lrsDetail As ADODB.Recordset
   
   Set lorsDetail = New ADODB.Recordset
   Set lrsDetail = New ADODB.Recordset
     
   lorsDetail.Fields.Append "sField01", adVarChar, 20  'SI No.
   lorsDetail.Fields.Append "sField02", adVarChar, 150 'Particulars
   lorsDetail.Fields.Append "sField03", adVarChar, 50  'Category Description
   lorsDetail.Fields.Append "sField04", adVarChar, 20  'Category Master
   lorsDetail.Fields.Append "sField05", adVarChar, 50  'Transaction Type
   lorsDetail.Fields.Append "sField06", adVarChar, 50  'Transaction No.
   lorsDetail.Fields.Append "nField01", adInteger, 5   'Quantity
   lorsDetail.Fields.Append "lField01", adCurrency, 10 'Sub Total
'   lorsDetail.Fields.Append "lField02", adCurrency, 10 'Amount Paid
   lorsDetail.Open
   
'SO
   lsSQL = "SELECT" _
               & " a.sTransNox as xTransNox, " _
               & " a.nEntryNox, " _
               & " a.sStockIDx, " _
               & " a.nQuantity, " _
               & " a.nSubTotal, " _
               & " 1 xxTotalxx, " _
               & " b.nAmtPaidx,  " _
               & " b.dTransact as xTransact, " _
               & " b.sSalesInv as xReferens, " _
               & " b.nTranTotl as xTranTotl, " _
               & " g.sCategory, " _
               & " g.sCategNme, " _
               & " g.cJoinValx, " _
               & " i.sCatDescx, " _
               & " d.sBrandNme, " _
               & " e.sModelNme, " _
               & " c.sDescript, " _
               & " f.sColorNme, " _
               & " b.sRemarksx as xRemarksx, " _
               & " c.cCellCard as xCellCard, " _

   lsSQL = lsSQL _
            & " 'Sales' xTranType " _
            & " FROM CP_SO_Detail a" _
               & " LEFT JOIN ELoad_Ledger h " _
                  & " ON a.sTransNox = h.sSourceNo " _
                  & " AND a.sStockIDx = h.sStockIDx " _
                  & " AND a.nEntryNox = h.sTransNox " _
               & " LEFT JOIN CP_SO_Master b" _
                  & " ON a.sTransNox = b.sTransNox" _
               & " LEFT JOIN CP_Inventory c " _
                  & " ON a.sStockIDx = c.sStockIDx" _
               & " LEFT JOIN Brand d " _
                  & " ON c.sBrandIDx = d.sBrandIDx" _
               & " LEFT JOIN Model e " _
                  & " ON c.sModelIdx = e.sModelIDx " _
               & " LEFT JOIN Color f " _
                  & " ON c.sColorIDx = f.sColorIDx " _
               & " LEFT JOIN Category g " _
                  & " ON c.sCategIDx = g.sCategIDx " _
               & " LEFT JOIN Category_Master i " _
                  & " ON g.sCategory = i.sCategory " _
            & " WHERE Left(a.sTransNox,2) = '" & oApp.BranchCode & "' " _
               & " AND b.dTransact between '" & (txtfield(1).Text) & "' " _
               & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
               & " AND b.cTranStat <> 4 "

   'JO
   lsSQL = lsSQL _
         & " UNION " _
         & "SELECT" _
               & " sTransNox as xTransNox, " _
               & " 1 nEntryNox, " _
               & "' ' sStockIDx," _
               & " 1 nQuantity, " _
               & " 1 nSubTotal, " _
               & " nTranTotl as xxTotalxx,   " _
               & " nAmtPaidx,   " _
               & " dPaymentx as xTransact, " _
               & " sJobOrdNo as xReferens," _
               & " nTranTotl as xTrantotl, " _
               & "' ' sCategory," _
               & "' ' sCategNme," _
               
   lsSQL = lsSQL _
               & "' ' cJoinValx," _
               & "' ' sCatDescx, " _
               & "' ' sBrandNme, " _
               & "' ' sModelNme, " _
               & "' ' sDescript, " _
               & "' ' sColorNme, " _
               & " sComplent as xRemarksx, " _
               & "' ' xCellCard, " _
               & "'Job Order' xTranType " _

   lsSQL = lsSQL _
         & " FROM CP_JobOrder_Master " _
         & " WHERE Left(sTransNox,2) = '" & oApp.BranchCode & "' " _
               & " AND dPaymentx between '" & (txtfield(1).Text) & "' " _
               & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
               & " AND cTranStat = 1 " _
               & " ORDER BY xTransNox "
   
   If lrsDetail.State = adStateOpen Then lrsDetail.Close
   lrsDetail.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrsDetail.EOF Then
      Progress.Stop
      Progress.Close
      MsgBox "No Record Found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Information"
      Exit Sub
   End If
      
   
For lnCtrDetail = 0 To lrsDetail.RecordCount - 1   'Total Cash
      Select Case lrsDetail("xTranType")
         Case "Sales"
            If Right(Left(lrsDetail("xReferens"), 4), 2) = "SI" Then
               If lrsDetail("cJoinValx") = 0 Then
                  If lorsDetail.RecordCount <> 0 Then lorsDetail.MoveFirst
                  lorsDetail.Find "sField04 = " & strParm(lrsDetail("sCategory")), 0, adSearchForward
                  'New Category
                  If lorsDetail.EOF = True Then
                     lorsDetail.AddNew
                     lorsDetail("sField01").Value = ""
                     lorsDetail("sField02").Value = ""
                     lorsDetail("sField03").Value = lrsDetail("sCatDescx")
                     lorsDetail("sField04").Value = lrsDetail("sCategory")
                     lorsDetail("sField06").Value = lrsDetail("xTransNox")
                     lorsDetail("nField01").Value = lrsDetail("nQuantity")
                     lorsDetail("lField01").Value = Format(lrsDetail("nSubTotal"), "#,##0.00")
                  Else 'Existing Category
                     'Find if SI Existing
                     lorsDetail.Find "sField06 = " & strParm(lrsDetail("xTransNox")), 0, adSearchForward
                     'New SI
                     If lorsDetail.EOF = True Then
                        lorsDetail.AddNew
                        lorsDetail("sField01").Value = ""
                        lorsDetail("sField02").Value = ""
                        lorsDetail("sField03").Value = lrsDetail("sCatDescx")
                        lorsDetail("sField04").Value = lrsDetail("sCategory")
                        lorsDetail("sField06").Value = lrsDetail("xTransNox")
                        lorsDetail("nField01").Value = lrsDetail("nQuantity")
                        lorsDetail("lField01").Value = Format(lrsDetail("nSubTotal"), "#,##0.00")
                     Else
                        'Old SI and New Category
                        If lorsDetail("sField04").Value = lrsDetail("sCategory") Then
                           lorsDetail("nField01").Value = Format(lorsDetail("nField01").Value + lrsDetail("nQuantity"), "#,##0")
                           lorsDetail("lField01").Value = Format(lorsDetail("lField01").Value + lrsDetail("nSubTotal"), "#,##0.00")
                        Else
                           'Old SI and Old Category
                           lorsDetail.Find "sField04 = " & strParm(lrsDetail("sCategory")), 0, adSearchForward
                           If lorsDetail.EOF = True Then
                              lorsDetail.AddNew
                              lorsDetail("sField01").Value = ""
                              lorsDetail("sField02").Value = ""
                              lorsDetail("sField03").Value = lrsDetail("sCatDescx")
                              lorsDetail("sField04").Value = lrsDetail("sCategory")
                              lorsDetail("sField06").Value = lrsDetail("xTransNox")
                              lorsDetail("nField01").Value = lrsDetail("nQuantity")
                              lorsDetail("lField01").Value = Format(lrsDetail("nSubTotal"), "#,##0.00")
                           Else
                              lorsDetail("nField01").Value = Format(lorsDetail("nField01").Value + lrsDetail("nQuantity"), "#,##0")
                              lorsDetail("lField01").Value = Format(lorsDetail("lField01").Value + lrsDetail("nSubTotal"), "#,##0.00")
                           End If
                        End If
                     End If
                  End If
               Else
                  lorsDetail.AddNew
               End If
            Else
               lorsDetail.AddNew
               lorsDetail("sField01").Value = Trim(lrsDetail("xReferens"))
               lorsDetail("sField02").Value = Trim(IIf(IsNull(lrsDetail("sBrandNme")), "", lrsDetail("sBrandNme")) + " " + _
                                       IIf(IsNull(lrsDetail("sModelNme")), "", lrsDetail("sModelNme")) + " " + _
                                       IIf(IsNull(lrsDetail("sDescript")), "", lrsDetail("sDescript")))
               lorsDetail("sField03").Value = lrsDetail("sCatDescx")
               lorsDetail("sField04").Value = lrsDetail("sCategory")
               lorsDetail("nField01").Value = lrsDetail("nQuantity")
               lorsDetail("lField01").Value = Format(lrsDetail("nSubTotal"), "#,##0.00")
'               lorsDetail("lField02").Value = IIf(lrsDetail("nSubTotal") <> 0, Format(lrsDetail("nAmtPaidx"), "#,##0.00"), 0)
            End If
            lorsDetail("sField05").Value = lrsDetail("xTranType")
         Case "Job Order"
            lorsDetail.AddNew
            lorsDetail("sField01").Value = Trim(lrsDetail("xReferens"))
            lorsDetail("sField02").Value = Trim(lrsDetail("xRemarksx"))
            lorsDetail("sField03").Value = "Job Order"
            lorsDetail("lField01").Value = Format(lrsDetail("nAmtPaidx"), "#,##0.00")
      End Select
      lorsDetail("sField05").Value = lrsDetail("xTranType")
      lrsDetail.MoveNext
   Next
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\DCPR.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lorsDetail
         
      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText "DCPR Details"
         .Sections("PH").ReportObjects("txtReportDate").SetText _
                        txtfield(1).Text & " - " & txtfield(2).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      End With
      
      Set lorsDetail = Nothing
      Set lrsDetail = Nothing
      
   frmViewer.CRViewer91.ReportSource = oReport
   frmViewer.CRViewer91.ViewReport
   frmViewer.Show
      
End Sub

Private Sub DCPR_Summary()
Dim lnCtrDetail As Integer
Dim lorsDetail As ADODB.Recordset
Dim lrsDetail As ADODB.Recordset

Dim lnCtr As Integer

Dim lrsTotalCash As New ADODB.Recordset
Dim lnTotlCash As Double
Dim lrsExpense As New ADODB.Recordset
Dim lnExpenses As Double
Dim lrsJobOrder As ADODB.Recordset
Dim lnJobOrder As Double
   
   Set lorsDetail = New ADODB.Recordset
   Set lrsDetail = New ADODB.Recordset
     
   lorsDetail.Fields.Append "sField01", adVarChar, 20  'Category
   lorsDetail.Fields.Append "sField02", adVarChar, 20  'Category ID
   lorsDetail.Fields.Append "sField03", adVarChar, 20  'sTransNox
   lorsDetail.Fields.Append "lField01", adCurrency, 10 'Sub Total
   lorsDetail.Open
   
   'SO
   lsSQL = "SELECT" _
            & " Distinct " _
            & " a.sTransNox, " _
            & " a.nEntryNox, " _
            & " a.nSubTotal, " _
            & " b.nAmtPaidx,  " _
            & " b.dTransact, " _
            & " b.nTranTotl, " _
            & " g.sCategory, " _
            & " g.sCategNme, " _
            & " i.sCatDescx, " _
            & " c.sStockIdx  " _

   lsSQL = lsSQL _
            & " FROM CP_SO_Detail a" _
               & " LEFT JOIN CP_SO_Master b" _
                  & " ON a.sTransNox = b.sTransNox" _
               & " LEFT JOIN CP_Inventory c " _
                  & " ON a.sStockIDx = c.sStockIDx" _
               & " LEFT JOIN Category g " _
                  & " ON c.sCategIDx = g.sCategIDx " _
               & " LEFT JOIN Category_Master i " _
                  & " ON g.sCategory = i.sCategory " _
            & " WHERE Left(a.sTransNox,2) = '" & oApp.BranchCode & "' " _
               & " AND b.dTransact between '" & (txtfield(1).Text) & "' " _
               & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
               & " AND b.cTranStat <> 4 " _
            & "Order by g.sCategNme "
   
   If lrsDetail.State = adStateOpen Then lrsDetail.Close
   lrsDetail.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrsDetail.EOF Then
      Progress.Stop
      Progress.Close
      MsgBox "No Record Found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Information"
      Exit Sub
   End If
      

   If lrsDetail.RecordCount = 0 Then Exit Sub
   For lnCtrDetail = 0 To lrsDetail.RecordCount - 1
      If lorsDetail.RecordCount <> 0 Then lorsDetail.MoveFirst
      lorsDetail.Find "sField02 = " & strParm(lrsDetail("sCategory")), 0, adSearchForward
      If lorsDetail.EOF Then
         lorsDetail.AddNew
         lorsDetail("sField01").Value = lrsDetail("sCategNme")
         lorsDetail("sField02").Value = lrsDetail("sCategory")
         lorsDetail("sField03").Value = lrsDetail("sTransNox")
         lorsDetail("lField01").Value = Format(lrsDetail("nSubTotal"), "#,##0.00")
      Else
         lorsDetail("lField01").Value = Format(lorsDetail("lField01").Value + lrsDetail("nSubTotal"), "#,##0.00")
      End If
      lrsDetail.MoveNext
   Next
   
   'Total Cash
   lsSQL = "SELECT" _
            & " nAmtPaidx, " _
            & " dTransact  " _
         & " FROM CP_SO_Master " _
         & " WHERE Left(sTransNox,2) = '" & oApp.BranchCode & "' " _
            & " AND dTransact between '" & (txtfield(1).Text) & "' " _
            & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
            & " AND cTranStat <> 4 " _
         & " ORDER BY sTransNox "
   Set lrsTotalCash = New ADODB.Recordset
   lrsTotalCash.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   lnTotlCash = 0#
   For lnCtr = 0 To lrsTotalCash.RecordCount - 1  'Total Cash Collection
      lnTotlCash = lnTotlCash + lrsTotalCash("nAmtPaidx")
      lrsTotalCash.MoveNext
   Next
   
   'Job Order
   lsSQL = "SELECT" _
            & " nAmtPaidx, " _
            & " dPaymentx  " _
         & " FROM CP_JobOrder_Master " _
         & " WHERE Left(sTransNox,2) = '" & oApp.BranchCode & "' " _
            & " AND dPaymentx between '" & (txtfield(1).Text) & "' " _
            & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
            & " AND cTranStat = 1 " _
         & " ORDER BY sTransNox "
   Set lrsJobOrder = New ADODB.Recordset
   lrsJobOrder.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   lnJobOrder = 0#
   For lnCtr = 0 To lrsJobOrder.RecordCount - 1  'Total Job Order
      lnJobOrder = lnJobOrder + lrsJobOrder("nAmtPaidx")
      lrsJobOrder.MoveNext
   Next

   'Expenses
   lsSQL = "SELECT " _
               & " sTransNox, " _
               & " nTotalExp, " _
               & " dTranDate  " _
         & " FROM CP_Expense_Master " _
         & " WHERE dTranDate between '" & (txtfield(1).Text) & "' " _
               & " AND '" & (txtfield(1).Text & " 23:59:59") & "'" _
               & " AND sBranchCd = '" & oApp.BranchCode & "'"
   Set lrsExpense = New ADODB.Recordset
   lrsExpense.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   lnExpenses = 0#
   For lnCtr = 0 To lrsExpense.RecordCount - 1  'Total Expense
      lnExpenses = lnExpenses + lrsExpense("nTotalExp")
      lrsExpense.MoveNext
   Next
   
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\DCPR_Summary.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lorsDetail
         
      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText "DCPR Summary"
         .Sections("PH").ReportObjects("txtReportDate").SetText _
                        txtfield(1).Text & " - " & txtfield(2).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      
         .Sections("RF").ReportObjects("txtTotlCash").SetText Format(lnTotlCash + lnJobOrder, "#,##0.00")
         .Sections("RF").ReportObjects("txtExpenses").SetText Format(lnExpenses, "#,##0.00")
         .Sections("RF").ReportObjects("txtJobOrder").SetText Format(lnJobOrder, "#,##0.00")
         .Sections("RF").ReportObjects("txtGrndTotl").SetText Format _
                                       (lnTotlCash + lnJobOrder - lnExpenses, "#,##0.00")
      End With
      
      Set lorsDetail = Nothing
      Set lrsDetail = Nothing
      Set lrsExpense = Nothing
      Set lrsJobOrder = Nothing
      
   frmViewer.CRViewer91.ReportSource = oReport
   frmViewer.CRViewer91.ViewReport
   frmViewer.Show
      
End Sub

Private Sub Sales_Details()
Dim lnCtr As Integer
Dim lrsDetail As ADODB.Recordset
Dim sCategory As String
   
   Set lrs = New ADODB.Recordset
   Set lrsDetail = New ADODB.Recordset

   'SO
   lsSQL = "SELECT" _
               & " a.nEntryNox, " _
               & " a.sStockIDx, " _
               & " a.nQuantity, " _
               & " a.nUnitPrce, " _
               & " a.nDiscAmnt, " _
               & " c.sCategIDx, " _
               & " d.sBrandNme, " _
               & " e.sModelNme, " _
               & " c.sDescript, " _
               & " f.sColorNme, " _
               & " b.dTransact, " _
               & " a.nSubTotal, " _
               & " b.nAmtPaidx, " _
               & " g.sFrstName + ' ' + left(g.sLastName,1) as xFullName, " _
               & " i.sIMEINoxx, " _
               & " c.cCellPhon, " _
               & " c.cCellCard, " _
               & " c.cCellLoad, " _
               & " c.cWalletxx, " _
               & " c.cMicrofon, " _
               & " c.cWdSerial  "
               
   lsSQL = lsSQL _
            & " FROM CP_SO_Detail a" _
               & " LEFT JOIN CP_SO_Master b" _
                  & " ON a.sTransNox = b.sTransNox" _
               & " LEFT JOIN CP_Inventory c " _
                  & " ON a.sStockIDx = c.sStockIDx" _
               & " LEFT JOIN Category j " _
                  & " ON c.sCategIDx = j.sCategIDx" _
               & " LEFT JOIN Category_Master k " _
                  & " ON j.sCategory = k.sCategory" _
               & " LEFT JOIN Brand d " _
                  & " ON c.sBrandIDx = d.sBrandIDx" _
               & " LEFT JOIN Model e " _
                  & " ON c.sModelIdx = e.sModelIDx " _
               & " LEFT JOIN Color f " _
                  & " ON c.sColorIDx = f.sColorIDx " _
               & " LEFT JOIN Sales_Person g " _
                  & " ON b.sCashierx = g.sEmployID " _
               & " LEFT JOIN CP_SO_Serial h " _
                  & " ON a.sTransNox = h.sTransNox " _
                  & "AND a.nEntryNox = h.nEntryNox " _
               & " LEFT JOIN CP_Serial_Master i " _
                  & " ON h.sSerialID = i.sSerialID "
                  
   lsSQL = lsSQL _
            & " WHERE Left(a.sTransNox,2) = '" & oApp.BranchCode & "' " _
               & " AND b.dTransact between '" & (txtfield(1).Text) & "' " _
               & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
               & " AND j.sCategory = '" & txtfield(0).Tag & "'" _
               & " AND b.cTranStat <> 4 "

   If lrsDetail.State = adStateOpen Then lrsDetail.Close
   lrsDetail.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If lrsDetail.EOF Then
      Progress.Stop
      Progress.Close
      MsgBox "No Record Found!!!" & vbCrLf & _
          "Please Verify your Entry then Try Again!!!", vbCritical, "Information"
      Exit Sub
   End If
   
   sCategory = ""
   If lrsDetail("cCellLoad") = 0 And lrsDetail("cWalletxx") = 0 _
      And lrsDetail("cCellCard") = 0 Then
      
      'Cellphone,Microphone,MP3/MP4/Ipod,Accessories,Digital Camera
      lrs.Fields.Append "dField01", adDBDate          'dTransact
      lrs.Fields.Append "sField01", adVarChar, 150    'sBrandNme
      lrs.Fields.Append "sField02", adVarChar, 150    'sModelNme, sDescript, sColorNme
      lrs.Fields.Append "sField03", adVarChar, 50     'sIMEINoxx
      lrs.Fields.Append "sField05", adVarChar, 150    'xFullName
      lrs.Fields.Append "nField01", adInteger, 5      'nQuantity
      lrs.Fields.Append "lField01", adCurrency, 10    'nUnitPrce
      lrs.Fields.Append "lField02", adCurrency, 10    'nDiscAmnt
      lrs.Fields.Append "lField03", adCurrency, 10    'nSubTotal
      lrs.Open
      
      For lnCtr = 0 To lrsDetail.RecordCount - 1
         lrs.AddNew
         lrs("dField01").Value = lrsDetail("dTransact")
         lrs("sField01").Value = IIf(IsNull(lrsDetail("sBrandNme")), "", lrsDetail("sBrandNme"))
         lrs("sField03").Value = IIf(IsNull(lrsDetail("sIMEINoxx")), "", lrsDetail("sIMEINoxx"))
         lrs("sField05").Value = IIf(IsNull(lrsDetail("xFullName")), "", lrsDetail("xFullName"))
         lrs("nField01").Value = lrsDetail("nQuantity")
         lrs("lField01").Value = lrsDetail("nUnitPrce")
         If lrsDetail("nUnitPrce") = lrsDetail("nDiscAmnt") Then
            lrs("sField02").Value = Trim(IIf(IsNull(lrsDetail("sModelNme")), "", lrsDetail("sModelNme")) + " " + _
                        IIf(IsNull(lrsDetail("sDescript")), "", lrsDetail("sDescript")) + " " + _
                        IIf(IsNull(lrsDetail("sColorNme")), "", lrsDetail("sColorNme")) + "F R E E")
         Else
            lrs("sField02").Value = Trim(IIf(IsNull(lrsDetail("sModelNme")), "", lrsDetail("sModelNme")) + " " + _
                        IIf(IsNull(lrsDetail("sDescript")), "", lrsDetail("sDescript")) + " " + _
                        IIf(IsNull(lrsDetail("sColorNme")), "", lrsDetail("sColorNme")))
         End If
         lrs("lField02").Value = lrsDetail("nDiscAmnt")
         lrs("lField03").Value = lrsDetail("nAmtPaidx")
         lrsDetail.MoveNext
      Next

      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Sales_DetailedCP.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText txtfield(0).Text & " " & "Detailed Sales Report"
         .Sections("PH").ReportObjects("txtReportDate").SetText txtfield(1).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      End With
      Set lrs = Nothing
      Set lrsDetail = Nothing
      
   ElseIf lrsDetail("cCellLoad") = 1 Or lrsDetail("cWalletxx") = 1 Then
      'Load Retail, Load Wallet
      lrs.Fields.Append "sField01", adVarChar, 150   'sPhoneNum
      lrs.Fields.Append "sField02", adVarChar, 150   'sReferNox
      lrs.Fields.Append "sField03", adVarChar, 150   'sDescript
      lrs.Fields.Append "lField01", adCurrency, 10   'nUnitPrce
      lrs.Fields.Append "lField02", adCurrency, 10   'nPurPrice/nDiscAmnt
      lrs.Fields.Append "lField03", adCurrency, 10   'nQtyOnHnd
      lrs.Open

      lsSQL = "SELECT" _
               & " DISTINCT " _
               & " b.sSourceNo, " _
               & " a.nEntryNox, " _
               & " b.sTransNox, " _
               & " a.nUnitPrce, " _
               & " a.nPurPrice, " _
               & " c.dTransact, " _
               & " b.sPhoneNum, " _
               & " b.sReferNox, " _
               & " b.nQtyOnHnd, " _
               & " d.sDescript, " _
               & " f.sCategory, " _
               & " d.cCellLoad, " _
               & " d.cWalletxx  " _

      lsSQL = lsSQL & " FROM CP_SO_Detail a " _
               & " LEFT JOIN ELoad_Ledger b " _
                  & " ON a.sTransNox = b.sSourceNo " _
                  & " AND a.nEntryNox = b.sTransNox " _
               & " Left JOIN CP_SO_Master c " _
                  & " ON a.sTransNox = c.sTransNox " _
               & " Left JOIN CP_Inventory d " _
                  & " ON a.sStockIDx = d.sStockIDx " _
               & " LEFT JOIN Category e " _
                  & " ON d.sCategIDx = e.sCategIDx" _
               & " LEFT JOIN Category_Master f " _
                  & " ON e.sCategory = f.sCategory" _
            & " WHERE Left(a.sTransNox,2) = '" & oApp.BranchCode & "' " _
                  & " AND c.dTransact between '" & (txtfield(1).Text) & "' " _
                  & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
                  & " AND (d.cCellLoad = '1' or d.cWalletxx = '1') " _
                  & " AND f.sCategory = '" & txtfield(0).Tag & "'" _
                  & " AND b.sSOurceCd = 'CPSl'" _
                  & " AND c.cTranStat <> 4 " _
            & " ORDER by c.dTransact, a.nEntryNox, b.sTransNox "
   
      If lrsDetail.State = adStateOpen Then lrsDetail.Close
      lrsDetail.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
      If lrsDetail.EOF Then
         Progress.Stop
         Progress.Close
         MsgBox "No Record Found!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Information"
         Exit Sub
      End If
      MsgBox lrsDetail("sSourceno")
      For lnCtr = 0 To lrsDetail.RecordCount - 1
         lrs.AddNew
         lrs("sField01").Value = IIf(IsNull(lrsDetail("sPhoneNum")), "", lrsDetail("sPhoneNum"))
         lrs("sField02").Value = IIf(IsNull(lrsDetail("sReferNox")), "", lrsDetail("sReferNox"))
         lrs("sField03").Value = IIf(IsNull(lrsDetail("sDescript")), "", lrsDetail("sDescript"))
         lrs("lField01").Value = Format(lrsDetail("nUnitPrce"), "#,##0.00")
         If lrsDetail("cCellLoad") = 1 Then
            lrs("lField02").Value = Format(lrsDetail("nPurPrice"), "#,##0.00")
            lrs("lField03").Value = IIf(IsNull(lrsDetail("nQtyOnHnd")), 0, Format(lrsDetail("nQtyOnHnd"), "#,##0.00"))
            sCategory = "Retail"
         ElseIf lrsDetail("cWalletxx") = 1 Then
            sCategory = "Wallet"
         End If
         lrsDetail.MoveNext
      Next
      
      Select Case sCategory
         Case "Retail"
            Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Sales_DetailedLoad.rpt")
            oReport.DiscardSavedData
            oReport.FieldMappingType = crAutoFieldMapping
            oReport.Database.SetDataSource lrs
            With oReport
               .Sections("RH").ReportObjects("txtBranch").SetText Branch
               .Sections("PH").ReportObjects("txtReportName").SetText "Load Retail Sales Detailed Report"
               .Sections("PH").ReportObjects("txtReportDate").SetText txtfield(1).Text
               .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
            End With
         Case "Wallet"
            Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Sales_DetailedWallet.rpt")
            oReport.DiscardSavedData
            oReport.FieldMappingType = crAutoFieldMapping
            oReport.Database.SetDataSource lrs
            With oReport
               .Sections("RH").ReportObjects("txtBranch").SetText Branch
               .Sections("PH").ReportObjects("txtReportName").SetText "Load Wallet Detailed Sales Report"
               .Sections("PH").ReportObjects("txtReportDate").SetText txtfield(1).Text
               .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
            End With
      End Select
      Set lrs = Nothing
      Set lrsDetail = Nothing
      
   ElseIf lrsDetail("cCellCard") = 1 Then
      'Cell Card and Sim Card
      lrs.Fields.Append "sField01", adVarChar, 150    'sCardName
      lrs.Fields.Append "sField02", adVarChar, 20     'sTransNox
      lrs.Fields.Append "lField01", adCurrency, 10    'nUnitPrce
      lrs.Fields.Append "nField01", adInteger, 5      'nQuantity
      lrs.Fields.Append "nField02", adInteger, 5      'nQtyOnHnd
      lrs.Open
   
      lsSQL = "SELECT" _
                  & " a.sTransNox, " _
                  & " a.nEntryNox, " _
                  & " a.nSubTotal, " _
                  & " a.nQuantity, " _
                  & " d.sCardName, " _
                  & " b.dTransact, " _
                  & " e.nQtyOnHnd, " _
                  & " b.sSalesInv  "
                  
      lsSQL = lsSQL _
               & " FROM CP_SO_Detail a" _
                  & " LEFT JOIN CP_SO_Master b" _
                     & " ON a.sTransNox = b.sTransNox" _
                  & " LEFT JOIN CP_Inventory c " _
                     & " ON a.sStockIDx = c.sStockIDx" _
                  & " LEFT JOIN CP_Inventory_Ledger e " _
                     & " ON a.sStockIDx = e.sStockIDx " _
                     & "AND a.sTransNox = e.sSourceNo " _
                  & " LEFT JOIN Card d " _
                     & " ON c.sCardIDxx = d.sCardIDxx " _
                  & " LEFT JOIN Category f " _
                     & " ON c.sCategIDx = f.sCategIDx" _
                  & " LEFT JOIN Category_Master g " _
                     & " ON f.sCategory = g.sCategory" _
               & " WHERE Left(a.sTransNox,2) = '" & oApp.BranchCode & "' " _
                  & " AND b.dTransact between '" & (txtfield(1).Text) & "' " _
                  & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
                  & " AND f.sCategory = '" & txtfield(0).Tag & "' " _
                  & " AND e.sBranchCd = '" & oApp.BranchCode & "'" _
                  & " AND b.cTranStat <> 4 " _
               & " ORDER BY a.sTransNox, a.nEntryNox "
   
      If lrsDetail.State = adStateOpen Then lrsDetail.Close
      lrsDetail.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
      If lrsDetail.EOF Then
         Progress.Stop
         Progress.Close
         MsgBox "No Record Found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Information"
         Exit Sub
      End If
   
      For lnCtr = 0 To lrsDetail.RecordCount - 1
         lrs.AddNew
         lrs("sField01").Value = IIf(IsNull(lrsDetail("sCardName")), "", lrsDetail("sCardName"))
         lrs("sField02").Value = lrsDetail("sSalesInv")
         lrs("lField01").Value = lrsDetail("nSubTotal")
         lrs("nField01").Value = lrsDetail("nQuantity")
         lrs("nField02").Value = lrsDetail("nQtyOnHnd")
         lrsDetail.MoveNext
      Next

      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Sales_DetailedCard.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText txtfield(0).Text & " " & "Detailed Sales Report"
         .Sections("PH").ReportObjects("txtReportDate").SetText txtfield(1).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      End With
      Set lrs = Nothing
      Set lrsDetail = Nothing
      
   End If
   
   frmViewer.CRViewer91.ReportSource = oReport
   frmViewer.CRViewer91.ViewReport
   frmViewer.Show

End Sub

Private Sub Sales_Summary()
Dim lnCtr As Integer
Dim lrsSummary As ADODB.Recordset
Dim lrs As ADODB.Recordset
Dim sCategory As String

   Set lrs = New ADODB.Recordset
   Set lrsSummary = New ADODB.Recordset

   'SO
   lsSQL = "SELECT" _
               & " a.nEntryNox, " _
               & " a.sStockIDx, " _
               & " a.nQuantity, " _
               & " c.sCategIDx, " _
               & " d.sBrandNme, " _
               & " b.dTransact, " _
               & " c.cCellCard, " _
               & " c.cCellLoad, " _
               & " c.cWalletxx, " _
               & " j.sCategNme  " _
               
   lsSQL = lsSQL _
            & " FROM CP_SO_Detail a" _
               & " LEFT JOIN CP_SO_Master b" _
                  & " ON a.sTransNox = b.sTransNox" _
               & " LEFT JOIN CP_Inventory c " _
                  & " ON a.sStockIDx = c.sStockIDx" _
               & " LEFT JOIN Category j " _
                  & " ON c.sCategIDx = j.sCategIDx" _
               & " LEFT JOIN Category_Master k " _
                  & " ON j.sCategory = k.sCategory" _
               & " LEFT JOIN Brand d " _
                  & " ON c.sBrandIDx = d.sBrandIDx" _
                  
   lsSQL = lsSQL _
            & " WHERE Left(a.sTransNox,2) = '" & oApp.BranchCode & "' " _
               & " AND b.dTransact between '" & (txtfield(1).Text) & "' " _
               & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
               & " AND j.sCategory = '" & txtfield(0).Tag & "'" _
               & " AND b.cTranStat <> 4 "
   
   If lrsSummary.State = adStateOpen Then lrsSummary.Close
   lrsSummary.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If lrsSummary.EOF Then
      Progress.Stop
      Progress.Close
      MsgBox "No Record Found!!!" & vbCrLf & _
          "Please Verify your Entry then Try Again!!!", vbCritical, "Information"
      Exit Sub
   End If
   
   sCategory = ""
   If lrsSummary("cCellLoad") = 0 And lrsSummary("cWalletxx") = 0 _
      And lrsSummary("cCellCard") = 0 Then
      
      'Cellphone,Microphone,MP3/MP4/Ipod,Accessories,Digital Camera
      lrs.Fields.Append "sField01", adVarChar, 50     'sBrandNme
      lrs.Fields.Append "sField02", adVarChar, 50     'sCategNme
      lrs.Fields.Append "nField01", adInteger, 5      'nQuantity
      lrs.Open

      For lnCtr = 0 To lrsSummary.RecordCount - 1
         lrs.AddNew
         lrs("sField01").Value = IIf(IsNull(lrsSummary("sBrandNme")), "", lrsSummary("sBrandNme"))
         lrs("sField02").Value = IIf(IsNull(lrsSummary("sCategNme")), "", lrsSummary("sCategNme"))
         lrs("nField01").Value = IIf(IsNull(lrsSummary("nQuantity")), 0, lrsSummary("nQuantity"))
         lrsSummary.MoveNext
      Next

      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Sales_SummaryCP.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText txtfield(0).Text & " " & "Sales Summary Report"
         .Sections("PH").ReportObjects("txtReportDate").SetText _
                        txtfield(1).Text & " - " & txtfield(2).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      End With
      
      Set lrs = Nothing
      Set lrsSummary = Nothing
      
   ElseIf lrsSummary("cCellLoad") = 1 Or lrsSummary("cWalletxx") = 1 Then
      'Load Retail, Load Wallet
      lrs.Fields.Append "dField01", adDBDate          'dTransact
      lrs.Fields.Append "sField01", adVarChar, 50     'sDescript
      lrs.Fields.Append "lField01", adCurrency, 10    'nUnitPrce
      lrs.Fields.Append "lField02", adCurrency, 10    'nPurPrice/nDiscAmnt
      lrs.Open

      lsSQL = "SELECT" _
              & " Distinct " _
              & " c.dTransact, " _
              & " a.sTransNox, " _
              & " a.nEntryNox, " _
              & " a.nUnitPrce, " _
              & " a.nPurPrice, " _
              & " d.sDescript, " _
              & " f.sCategory, " _
              & " a.nDiscAmnt, " _
              & " d.cCellCard, " _
              & " d.cCellLoad, " _
              & " d.cWalletxx  "

      lsSQL = lsSQL & " FROM CP_SO_Detail a " _
               & " LEFT JOIN ELoad_Ledger b " _
                  & " ON a.sTransNox = b.sSourceNo " _
                  & " AND a.nEntryNox = b.sTransNox " _
               & " Left JOIN CP_SO_Master c " _
                  & " ON a.sTransNox = c.sTransNox " _
               & " Left JOIN CP_Inventory d " _
                  & " ON a.sStockIDx = d.sStockIDx " _
               & " LEFT JOIN Category e " _
                  & " ON d.sCategIDx = e.sCategIDx" _
               & " LEFT JOIN Category_Master f " _
                  & " ON e.sCategory = f.sCategory" _
            & " WHERE Left(a.sTransNox,2) = '" & oApp.BranchCode & "' " _
                  & " AND c.dTransact between '" & (txtfield(1).Text) & "' " _
                  & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
                  & " AND f.sCategory = '" & txtfield(0).Tag & "'" _
                  & " AND c.cTranStat <> 4 " _
            & " ORDER by c.dTransact, a.sTransNox, a.nEntryNox "

      If lrsSummary.State = adStateOpen Then lrsSummary.Close
      lrsSummary.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

      For lnCtr = 0 To lrsSummary.RecordCount - 1
         lrs.AddNew
         lrs("sField01").Value = lrsSummary("sDescript")
         lrs("dField01").Value = Format(lrsSummary("dTransact"), "MMMM dd,yyyy")
         lrs("lField01").Value = Format(lrsSummary("nUnitPrce"), "#,##0.00")
         If lrsSummary("cCellLoad") = 1 Then
            lrs("lField02").Value = Format(lrsSummary("nPurPrice"), "#,##0.00")
            sCategory = "Retail"
         ElseIf lrsSummary("cWalletxx") = 1 Then
            lrs("lField02").Value = Format(lrsSummary("nDiscAmnt"), "#,##0.00")
            sCategory = "Wallet"
         End If
         lrsSummary.MoveNext
      Next
      
      Select Case sCategory
         Case "Retail"
            Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Sales_SummaryLoad.rpt")
            oReport.DiscardSavedData
            oReport.FieldMappingType = crAutoFieldMapping
            oReport.Database.SetDataSource lrs
   
            With oReport
               .Sections("RH").ReportObjects("txtBranch").SetText Branch
               .Sections("PH").ReportObjects("txtReportName").SetText txtfield(0).Text & " " & "Sales Summary"
               .Sections("PH").ReportObjects("txtReportDate").SetText _
                              txtfield(1).Text & " - " & txtfield(2).Text
               .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
            End With
         Case "Wallet"
            Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Sales_SummaryWallet.rpt")
            oReport.DiscardSavedData
            oReport.FieldMappingType = crAutoFieldMapping
            oReport.Database.SetDataSource lrs
   
            With oReport
               .Sections("RH").ReportObjects("txtBranch").SetText Branch
               .Sections("PH").ReportObjects("txtReportName").SetText txtfield(0).Text & " " & "Sales Summary"
               .Sections("PH").ReportObjects("txtReportDate").SetText _
                              txtfield(1).Text & " - " & txtfield(2).Text
               .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
            End With
      End Select
   ElseIf lrsSummary("cCellCard") = 1 Then
      lrs.Fields.Append "dField01", adDBDate          'dTransact
      lrs.Fields.Append "nField01", adInteger, 5      'nQuantity
      lrs.Fields.Append "sField01", adVarChar, 150    'sDescript
      lrs.Fields.Append "lField01", adCurrency, 10    'nUnitPrce
      lrs.Open

      'SO
      lsSQL = "SELECT" _
                  & " a.sStockIDx, " _
                  & " a.nQuantity, " _
                  & " a.nSubTotal, " _
                  & " c.sCategIDx, " _
                  & " b.dTransact, " _
                  & " j.sCategNme, " _
                  & " d.sCardName  " _
                  
      lsSQL = lsSQL _
               & " FROM CP_SO_Detail a" _
                  & " LEFT JOIN CP_SO_Master b" _
                     & " ON a.sTransNox = b.sTransNox" _
                  & " LEFT JOIN CP_Inventory c " _
                     & " ON a.sStockIDx = c.sStockIDx" _
                  & " LEFT JOIN Card d " _
                     & " ON c.sCardIDxx = d.sCardIDxx" _
                  & " LEFT JOIN Category j " _
                     & " ON c.sCategIDx = j.sCategIDx" _
                  & " LEFT JOIN Category_Master k " _
                     & " ON j.sCategory = k.sCategory" _

      lsSQL = lsSQL _
               & " WHERE Left(a.sTransNox,2) = '" & oApp.BranchCode & "' " _
                  & " AND b.dTransact between '" & (txtfield(1).Text) & "' " _
                  & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
                  & " AND j.sCategory = '" & txtfield(0).Tag & "'" _
                  & " AND b.cTranStat <> 4 "
   
      If lrsSummary.State = adStateOpen Then lrsSummary.Close
      lrsSummary.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

      For lnCtr = 0 To lrsSummary.RecordCount - 1
         lrs.AddNew
         lrs("dField01").Value = Format(lrsSummary("dTransact"), "MMMM dd,yyyy")
         lrs("sField01").Value = IIf(IsNull(lrsSummary("sCardName")), "", lrsSummary("sCardName"))
         lrs("nField01").Value = IIf(IsNull(lrsSummary("nQuantity")), 0, _
                                 Format(lrsSummary("nQuantity"), "#,##0"))
         lrs("lField01").Value = Format(lrsSummary("nSubTotal"), "#,##0.00")
         lrsSummary.MoveNext
      Next

      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Sales_SummaryCard.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Branch
         .Sections("PH").ReportObjects("txtReportName").SetText "Card Sales Summary Report"
         .Sections("PH").ReportObjects("txtReportDate").SetText _
                        txtfield(1).Text & " - " & txtfield(2).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      End With
   End If
   
   Set lrs = Nothing
   Set lrsSummary = Nothing

   frmViewer.CRViewer91.ReportSource = oReport
   frmViewer.CRViewer91.ViewReport
   frmViewer.Show

End Sub



VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReport_Transfer 
   BorderStyle     =   0  'None
   Caption         =   "Branch Transfer Summary"
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   750
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   1323
      Begin xrControl.xrFrame xrFrame2 
         Height          =   540
         Left            =   75
         Tag             =   "wt0;wb0"
         Top             =   90
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   953
         Begin VB.TextBox txtfield 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   0
            Left            =   1005
            MaxLength       =   50
            TabIndex        =   1
            Top             =   135
            Width           =   3615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Branch"
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
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Tag             =   "ebo"
            Top             =   150
            Width           =   885
         End
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   375
      Index           =   0
      Left            =   5295
      TabIndex        =   8
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
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
      Picture         =   "frmReport_Transfer.frx":0000
      CaptionAlign    =   0
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   5295
      TabIndex        =   9
      Top             =   1590
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
      Picture         =   "frmReport_Transfer.frx":1112
      CaptionAlign    =   0
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1050
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1365
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   1852
      ClipControls    =   0   'False
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   3000
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   465
         Width           =   1700
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   765
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   465
         Width           =   1700
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
         Height          =   315
         Index           =   3
         Left            =   165
         TabIndex        =   2
         Tag             =   "et0;fb0"
         Top             =   60
         Width           =   510
      End
      Begin VB.Shape Shape1 
         Height          =   810
         Index           =   1
         Left            =   75
         Top             =   120
         Width           =   4755
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   3
         Top             =   465
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Index           =   1
         Left            =   2625
         TabIndex        =   5
         Top             =   465
         Width           =   525
      End
   End
   Begin MSComCtl2.Animation Progress 
      Height          =   720
      Left            =   5595
      TabIndex        =   7
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
      Index           =   2
      Left            =   5295
      TabIndex        =   10
      Top             =   2010
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
      Picture         =   "frmReport_Transfer.frx":188C
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmReport_Transfer"
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

Dim lsSQL As String
Dim Address As String
Dim Code As String
Dim Branch As String

Private Sub cmdButton_Click(Index As Integer)
Dim lnCtr As Integer
   Select Case Index
      Case 0 'OK
         If txtfield(0).Tag <> "" Then
            Progress.Open App.Path & "\images\FINDFILE.AVI"
            Progress.Play
            ReportPreview
         Else
            Exit Sub
         End If
      Case 1
         Branch = txtfield(Index)
         getBranch Code, Branch, Address
         txtfield(Index) = Branch
         txtfield(Index).Tag = Code
      Case 2 'Cancel
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   If bLoaded = False Then
      bLoaded = True
      txtfield(0).SetFocus
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
      
   txtfield(0) = ""
   txtfield(0).Tag = ""
   txtfield(1) = Format(oApp.ServerDate, "MMMM dd, yyyy")
   txtfield(2) = Format(oApp.ServerDate, "MMMM dd, yyyy")
   
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

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lsSQL As String
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 0 Then
         Branch = txtfield(Index)
         getBranch Code, Branch, Address
         txtfield(Index) = Branch
         txtfield(Index).Tag = Code
      End If
      If txtfield(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)

Select Case Index
   Case 1, 2
      If Not IsDate(txtfield(Index).Text) Then
         txtfield(Index).Text = Format(oApp.ServerDate, "MMMM dd, yyyy")
      Else
         txtfield(Index).Text = Format(txtfield(Index).Text, "MMMM dd, yyyy")
      End If
   End Select
End Sub

Private Sub ReportPreview()
   Dim pnCtr As Integer
   Dim lnCtr As Integer
   Dim lrs As ADODB.Recordset
   Dim Detail As ADODB.Recordset
   Dim oRS As ADODB.Recordset

   Set oRS = New ADODB.Recordset
   Set Detail = New ADODB.Recordset
   
   Set lrs = New ADODB.Recordset
   lrs.Fields.Append "dField01", adDBDate
   lrs.Fields.Append "sField01", adVarChar, 150
   lrs.Fields.Append "sField02", adVarChar, 240
   lrs.Fields.Append "sField03", adVarChar, 20
   lrs.Fields.Append "sField04", adVarChar, 150
   lrs.Fields.Append "sField05", adVarChar, 15
   lrs.Fields.Append "nField01", adInteger, 5
   lrs.Fields.Append "lField01", adCurrency, 10
   lrs.Open
      'No Serial
      lsSQL = "SELECT" _
               & " a.nEntryNox, " _
               & " a.nQuantity, " _
               & " a.nUnitPrce as xUnitPrce, " _
               & " b.sDescript, " _
               & " c.sBrandNme, " _
               & " d.sModelNme, " _
               & " e.sColorNme, " _
               & " f.sReferNox as xReferNox, " _
               & " f.dTransact as xTransact, " _
               & " g.sBranchNm, " _
               & "'a' sIMEINoxx," _
               & " f.sRemarksx, " _
               & " 'No Serial' xTranType "
      lsSQL = lsSQL _
         & " FROM CP_Transfer_Detail a " _
            & " LEFT JOIN CP_Transfer_Master f " _
               & " ON a.sTransNox = f.sTransNox " _
            & " LEFT JOIN CP_Inventory b " _
               & " ON a.sStockIDx = b.sStockIDx " _
            & " LEFT JOIN Brand c " _
               & " ON b.sBrandIDx = c.sBrandIDx " _
            & " LEFT JOIN Model d " _
               & " ON b.sModelIDx = d.sModelIDx " _
            & " LEFT JOIN Color e " _
               & " ON b.sColorIDx = e.sColorIDx " _
            & " LEFT JOIN Branch g " _
               & " ON f.sDestinat = g.sBranchCd "
      lsSQL = lsSQL _
         & " WHERE b.cCellLoad = 0 " _
               & " AND b.cWalletxx = 0 " _
               & " AND f.cTranStat = 1 " _
               & " AND f.dTransact between '" & (txtfield(1).Text) & "' " _
               & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
               & " AND f.sDestinat = '" & txtfield(0).Tag & "'" _
               & " AND f.sOriginxx = '" & oApp.BranchCode & "'"
                              
      'Load Transfer
      lsSQL = lsSQL _
               & " UNION " _
               & " SELECT" _
               & " a.nEntryNox, " _
               & " 1 nQuantity, " _
               & " a.nQtyOutxx as xUnitPrce, " _
               & " b.sDescript, " _
               & " 'a' sBrandNme, " _
               & " 'a' sModelNme, " _
               & " 'a' sColorNme, " _
               & " a.sReferNox as xReferNox, " _
               & " a.dTransact as xTransact, " _
               & " g.sBranchNm, " _
               & " 'a' sIMEINoxx, " _
               & " 'a' sRemarksx, " _
               & " 'Load Wallet' xTranType "
      lsSQL = lsSQL _
         & " FROM ELoad_Ledger a " _
            & " LEFT JOIN CP_Inventory b " _
               & " ON a.sStockIDx = b.sStockIDx " _
            & " LEFT JOIN Branch g " _
               & " ON a.sBranchCd = g.sBranchCd "
      lsSQL = lsSQL _
         & " WHERE a.sSourceCd = 'CPLv' " _
               & " AND a.dTransact between '" & (txtfield(1).Text) & "' " _
               & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
               & " AND a.sBranchCd = '" & txtfield(0).Tag & "'" _

      'w/ Serial
      lsSQL = lsSQL _
               & " UNION " _
               & " SELECT" _
               & " a.nEntryNox, " _
               & " 1 nQuantity, " _
               & " a.nUnitPrce as xUnitPrce, " _
               & " b.sDescript, " _
               & " c.sBrandNme, " _
               & " d.sModelNme, " _
               & " e.sColorNme, " _
               & " f.sReferNox as xReferNox, " _
               & " f.dTransact as xTransact, " _
               & " g.sBranchNm, " _
               & " h.sIMEINoxx, " _
               & " f.sRemarksx, " _
               & " 'w/ Serial' xTranType "
      lsSQL = lsSQL _
         & " FROM CP_Serial_Transfer_Detail a " _
            & " LEFT JOIN CP_Serial_Transfer_Master f " _
               & " ON a.sTransNox = f.sTransNox " _
            & " LEFT JOIN CP_Serial_Master h " _
               & " ON a.sSerialID = h.sSerialID " _
            & " LEFT JOIN CP_Inventory b " _
               & " ON h.sStockIDx = b.sStockIDx " _
            & " LEFT JOIN Brand c " _
               & " ON b.sBrandIDx = c.sBrandIDx " _
            & " LEFT JOIN Model d " _
               & " ON b.sModelIDx = d.sModelIDx " _
            & " LEFT JOIN Color e " _
               & " ON b.sColorIDx = e.sColorIDx " _
            & " LEFT JOIN Branch g " _
               & " ON f.sDestinat = g.sBranchCd "
      lsSQL = lsSQL _
         & " WHERE b.cCellLoad = 0 " _
               & " AND b.cWalletxx = 0 " _
               & " AND f.cTranStat = 1 " _
               & " AND f.dTransact between '" & (txtfield(1).Text) & "' " _
               & " AND '" & (txtfield(2).Text & " 23:59:59") & "'" _
               & " AND f.sDestinat = '" & txtfield(0).Tag & "'" _
               & " AND f.sOriginxx = '" & oApp.BranchCode & "'" _
         & " ORDER BY f.dTransact, a.nEntryNox "
      If oRS.State = adStateOpen Then oRS.Close
      oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If oRS.EOF Then
         Progress.Stop
         Progress.Close
         MsgBox "No Record Found!!!" & vbCrLf & _
                  "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         Exit Sub
      End If

      For lnCtr = 0 To oRS.RecordCount - 1
         lrs.AddNew
         lrs("dField01").Value = oRS("xTransact")
         lrs("sField01").Value = oRS("sBranchNm")
         lrs("sField03").Value = oRS("xReferNox")
         lrs("sfield05").Value = oRS("xTranType")
         Select Case oRS("xTranType")
            Case "No Serial"
               lrs("sField02").Value = oRS("sRemarksx")
               lrs("sField04").Value = Trim(IIf(IsNull(oRS("sBrandNme")), "", oRS("sBrandNme")) + " " + _
                                       IIf(IsNull(oRS("sModelNme")), "", oRS("sModelNme")) + " " + _
                                       IIf(IsNull(oRS("sDescript")), "", oRS("sDescript")) + " " + _
                                       IIf(IsNull(oRS("sColorNme")), "", oRS("sColorNme")))
               lrs("nField01").Value = oRS("nQuantity")
               lrs("lField01").Value = Format(oRS("xUnitPrce"), "#,##0.00")
            Case "w/ Serial"
               lrs("sField04").Value = Trim(IIf(IsNull(oRS("sBrandNme")), "", oRS("sBrandNme")) + " " + _
                                       IIf(IsNull(oRS("sModelNme")), "", oRS("sModelNme")) + " " + _
                                       IIf(IsNull(oRS("sDescript")), "", oRS("sDescript")) + " " + _
                                       IIf(IsNull(oRS("sColorNme")), "", oRS("sColorNme")) + " " + _
                                       IIf(IsNull(oRS("sIMEINoxx")), "", oRS("sIMEINoxx")))
               lrs("nField01").Value = 1
               lrs("lField01").Value = Format(oRS("xUnitPrce"), "#,##0.00")
               lrs("sField02").Value = oRS("sRemarksx")
            Case "Load Wallet"
               lrs("sField04").Value = Trim(IIf(IsNull(oRS("sDescript")), "", oRS("sDescript")))
               lrs("nField01").Value = 1
               lrs("lField01").Value = Format(oRS("xUnitPrce"), "#,##0.00")
         End Select
         
         oRS.MoveNext
      Next

      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Transfer_Summary.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs

      With oReport
         .Sections("RH").ReportObjects("txtBranch").SetText Address
         .Sections("PH").ReportObjects("txtReportName").SetText "Branch Transfer Summary"
         .Sections("PH").ReportObjects("txtReportDate").SetText _
                        txtfield(1).Text & " - " & txtfield(2).Text
         .Sections("PF").ReportObjects("txtrptUser").SetText oApp.UserName
      End With

      Set lrs = Nothing
      Set oRS = Nothing
      Set Detail = Nothing

      frmViewer.CRViewer91.ReportSource = oReport
      frmViewer.CRViewer91.ViewReport
      frmViewer.Show

End Sub






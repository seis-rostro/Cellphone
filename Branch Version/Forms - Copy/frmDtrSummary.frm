VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmDtrSummary 
   BorderStyle     =   0  'None
   Caption         =   "DTR Summary"
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3525
      Index           =   1
      Left            =   180
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   6218
      BackColor       =   12632256
      BorderStyle     =   1
      Begin xrControl.xrFrame xrFrame2 
         Height          =   3255
         Index           =   0
         Left            =   105
         Tag             =   "wt0;fb0"
         Top             =   105
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   5741
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Index           =   7
            Left            =   5490
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   2445
            Width           =   2325
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Index           =   4
            Left            =   1635
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   2565
            Width           =   2325
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Index           =   6
            Left            =   5505
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   2010
            Width           =   2325
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Index           =   5
            Left            =   5505
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   1575
            Width           =   2325
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Index           =   3
            Left            =   1650
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   2130
            Width           =   2325
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Index           =   2
            Left            =   1650
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   1695
            Width           =   2325
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
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
            Height          =   360
            Index           =   1
            Left            =   1635
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   1275
            Width           =   2325
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
            Index           =   0
            Left            =   1500
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   120
            Width           =   2340
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
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
            Index           =   8
            Left            =   5655
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   180
            Width           =   2340
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly Paym"
            Height          =   300
            Index           =   5
            Left            =   375
            TabIndex        =   20
            Top             =   2190
            Width           =   1230
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Deduction"
            ForeColor       =   &H000000FF&
            Height          =   450
            Index           =   0
            Left            =   -255
            TabIndex        =   19
            Top             =   2595
            Width           =   1785
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Check"
            Height          =   300
            Index           =   1
            Left            =   4155
            TabIndex        =   18
            Top             =   2505
            Width           =   1230
         End
         Begin VB.Shape Shape3 
            Height          =   300
            Index           =   0
            Left            =   5685
            Top             =   750
            Width           =   2250
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5700
            TabIndex        =   17
            Tag             =   "eb0;et0"
            Top             =   750
            Width           =   2220
         End
         Begin VB.Shape Shape4 
            Height          =   360
            Index           =   0
            Left            =   5655
            Top             =   720
            Width           =   2325
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL for Deposit:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   10
            Left            =   3975
            TabIndex        =   11
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Finance"
            Height          =   300
            Index           =   8
            Left            =   4230
            TabIndex        =   9
            Top             =   2070
            Width           =   1230
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " Credit Card"
            Height          =   300
            Index           =   7
            Left            =   4185
            TabIndex        =   8
            Top             =   1635
            Width           =   1230
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CP Sales"
            Height          =   195
            Index           =   2
            Left            =   810
            TabIndex        =   6
            Top             =   1335
            Width           =   645
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   405
            Left            =   1575
            Tag             =   "et0;ht2"
            Top             =   240
            Width           =   2325
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
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
            Left            =   120
            TabIndex        =   5
            Top             =   210
            Width           =   1245
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Others"
            Height          =   195
            Index           =   5
            Left            =   990
            TabIndex        =   4
            Top             =   1710
            Width           =   465
         End
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   9015
      TabIndex        =   3
      Top             =   1830
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
      Picture         =   "frmDtrSummary.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   9015
      TabIndex        =   2
      Top             =   1185
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
      Picture         =   "frmDtrSummary.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   9015
      TabIndex        =   1
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Confirm"
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
      Picture         =   "frmDtrSummary.frx":0EF4
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   300
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   3375
      Width           =   915
   End
End
Attribute VB_Name = "frmDtrSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmDtrSummary"

Private pnIndex As Integer
Private oSkin As clsFormSkin
Private p_oRepSource As Object
Dim pbFormLoad As Boolean
Dim loCopyFile As New FileSystemObject
Dim lsFileName As String
Dim lorec As Recordset
Dim lsDate As String

Private Sub cmdButton_Click(Index As Integer)
   Dim lnMsg As String
   Dim lnRep As Integer
   Dim lsSQL As String
   
   Select Case Index
      Case 0 ' POST
         If lorec("cPostedxx") = xeStateClosed Then
            If txtField(0).Text = "" And txtField(8).Text < 0# Then
               MsgBox "Unable to post DTR Summary!!!", vbCritical, "NOTICE"
               Exit Sub
            Else
               If oApp.UserLevel = xeManager Then
                  If isUnEncodedTransOK = True Then
                     If autogenrep = True Then
                        If PostDTR = True Then
                           MsgBox "DTR Posted Successfully!!!", vbOKOnly, "INFO"
                           
                           lsFileName = ""
         
                           lsFileName = "DTR" & oApp.BranchCode & Format(txtField(0).Text)
                           loCopyFile.MoveFile oApp.AppPath & "/Temp/" & lsFileName & ".pdf", oApp.AppPath & "/Temp/Upload/" & lsFileName & "-P" & ".pdf"
'                           Kill oApp.AppPath & "/Temp/" & lsFileName
                        End If
                     End If
                     LoadData (txtField(0).Text)
                  End If
               Else
                  MsgBox "You are not authorized to POST DTR Transaction!!!", vbCritical, "NOTICE"
               End If
            End If
         Else
            MsgBox "DTR Summary not yet confirm!!!", vbInformation
         End If

      Case 1
         Unload Me
      Case 2 'confirm
         If lorec("nDepositd") >= 0# And lorec("cPostedxx") = xeStateOpen Then
            lnRep = MsgBox("Are you sure you want to CONFIRM DTR Summary?", vbYesNo + vbQuestion, "Confirm")
            
            If lnRep = vbYes Then
               If autogenrep = True Then
'                     lsFileName = ""
'                     lsFileName = "DTR" & oApp.BranchCode & Format(txtField(0).Text)
'                     loCopyFile.MoveFile oApp.AppPath & "/Temp/" & lsFileName & ".pdf", oApp.AppPath & "/Temp/Upload/" & lsFileName & ".pdf"
               
                    lsSQL = "UPDATE DTR_Summary SET cPostedxx = '1' WHERE sBranchCd = " & strParm(oApp.BranchCode) & _
                        " AND sTranDate = " & strParm(lorec("sTranDate"))
                    oApp.Execute lsSQL, "DTR_Summary"
                    MsgBox "DTR Confirmed Successfully!!!", vbOKOnly, "INFO"
               End If
            End If
        Else
            MsgBox "Unable to confirm DTR Summary!!!", vbCritical, "NOTICE"
        End If
   End Select
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
'   '''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If Not pbFormLoad Then
      pbFormLoad = True
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   Dim lsSQL As String
      
   lsOldProc = "Form_Load"
'   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   clearFields
   LoadData ("")
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing

   pbFormLoad = False
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 0
   txtField(Index) = IFNull(Value, 0#)
   Case 1 To 11
      txtField(Index) = IFNull(Format(Value, "#,##0.00"), 0#)
   End Select
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

Private Sub clearFields()
   txtField(0).Text = ""
   txtField(8).Text = 0#
   
   txtField(1).Text = 0#
   txtField(2).Text = 0#
   txtField(3).Text = 0#
   txtField(4).Text = 0#
   txtField(5).Text = 0#
   txtField(6).Text = 0#
   txtField(7).Text = 0#
 
End Sub

Private Sub ComputeTotal()
   Dim lnTotal As Currency
   Dim lnCtr As Integer
   Dim lnAmtLess As Currency
   
   lnTotal = 0#
   lnAmtLess = 0#
   
   For lnCtr = 1 To 4
      lnTotal = lnTotal + CDbl(txtField(lnCtr).Text)
   Next
   
   lnAmtLess = CDbl(txtField(4)) + CDbl(txtField(5)) + CDbl(txtField(6) + CDbl(txtField(7)))
   
   txtField(8).Text = Format(lnTotal - lnAmtLess, "#,#00.00")
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 0 Then
      If KeyCode = vbKeyReturn Then
         LoadData (txtField(0).Text)
      ElseIf KeyCode = vbKeyF3 Then
         LoadData (txtField(0).Text)
      End If
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   If Index = 0 Then
   Else
'      Call ComputeTotal
   End If
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 1, 2, 3, 4, 5, 6, 7, 8
      If Not IsNumeric(txtField(Index)) Then
         txtField(Index) = 0#
      Else
         txtField(Index).Text = Format(txtField(Index), "#,##0.00")
      End If
   End Select
'   Call ComputeTotal
End Sub

Function PostDTR() As Boolean
   
   Dim lsSQL As String
   Dim lorec As Recordset
   lsSQL = "SELECT d.sAreaDesc, b.sBranchNm, a.sTranDate" & _
            " FROM DTR_Summary a" & _
            ", Branch b" & _
               " LEFT JOIN Branch_Others c ON b.sBranchCd = c.sBranchCd" & _
               " LEFT JOIN Branch_Area d ON c.sAreaCode = d.sAreaCode" & _
            " WHERE a.sBranchCd = b.sBranchCd" & _
            " AND a.sTranDate = " & strParm(lsDate) & _
            " AND a.sBranchCd = " & strParm(oApp.BranchCode)

   Set lorec = New Recordset
   lorec.Open lsSQL, oApp.Connection, , , adCmdText
   
   If Not lorec.EOF Then
      lsSQL = "UPDATE DTR_Summary SET cPostedxx = '2'  WHERE sBranchCd = " & strParm(oApp.BranchCode) & _
               " AND sTranDate = " & strParm(txtField(0).Text)
      Debug.Print lsSQL
      oApp.Execute lsSQL, "DTR_Summary"
      PostDTR = True
   Else
      PostDTR = False
   End If
   
endProc:
End Function

Private Function isUnEncodedTransOK() As Boolean
   Dim lsSQL As String
   Dim loRec1 As Recordset
   Dim loRec2 As Recordset
   
   lsSQL = "SELECT b.*" & _
            " FROM DTR_Summary a" & _
            ", DTR_Summary_Detail b" & _
            " WHERE a.sBranchcd = b.sBranchcd" & _
            " AND a.sTranDate = b.sTranDate" & _
            " AND a.sTranDate = " & strParm(txtField(0).Text) & _
            " AND a.sBranchCd = " & strParm(oApp.BranchCode) & _
            " AND b.cHasEntry = '0'"
            
   Set loRec1 = New Recordset
   loRec1.Open lsSQL, oApp.Connection, , , adCmdText

   If loRec1.EOF Then
      isUnEncodedTransOK = True
      
       'check here if there is previous date that not yet posted
         lsSQL = "SELECT * FROM DTR_Summary WHERE sBranchCd = " & strParm(oApp.BranchCode) & _
            " AND sTranDate < " & strParm(txtField(0).Text) & _
            " AND cPostedxx IN('0','1')"
      Set loRec2 = New Recordset
      loRec2.Open lsSQL, oApp.Connection, , , adCmdText
   
      If Not loRec2.EOF Then
         isUnEncodedTransOK = False
            MsgBox "There are previous unposted DTR Summary!" & vbCrLf & _
                   "Please post the transaction first..."
            GoTo endProc
      End If
      
   Else
      MsgBox "Pls encode all unencoded transactions" & vbCrLf & _
                              "before posting of DTR!!", vbCritical, "NOTICE"
                              
      isUnEncodedTransOK = False
   End If

endProc:
End Function

Function autogenrep() As Boolean
   Dim loReports As clsCPBranchRep
   Dim loRepViewer As frmRepViewer

   Set loReports = New clsCPBranchRep
   
   With loReports
      Set .AppDriver = oApp
      
      If ShowReport Then
         Set loRepViewer = New frmRepViewer
         Set loRepViewer.ReportSource = p_oRepSource.Source

      Else
        MsgBox "DTR Date for posting is not equal to DTR Date generated"
        autogenrep = False
        Exit Function
      End If
      autogenrep = True
   End With
   
End Function

Public Function ShowReport() As Boolean
      Set p_oRepSource = CreateObject(Trim("ggcCPBranchRep") & "." & Trim("clsDailyTransSummary"))
      
      Set p_oRepSource.AppDriver = oApp
      
      p_oRepSource.InitReport "CPDTR", "Daily Transaction Report"
      If p_oRepSource.ProcessReport = False Then Exit Function
      If txtField(0).Text <> Format(p_oRepSource.DateFr, "YYYYMMDD") Then
        p_oRepSource.CloseReport
        ShowReport = False
      Else
         p_oRepSource.CloseReport
        ShowReport = True
      End If
      
End Function

Private Sub LoadData(lsTranDate As String)
   Dim lsSQL As String
   Dim lnCtr As Integer
   
   Set lorec = New Recordset

   lsDate = ""
   lsSQL = "SELECT c.dUnEncode, a.cPostedxx, a.*" & _
            " FROM DTR_Summary a" & _
            ", Branch b" & _
               " LEFT JOIN Branch_Others c ON b.sBranchCd = c.sBranchCd" & _
               " LEFT JOIN Branch_Area d ON c.sAreaCode = d.sAreaCode" & _
            " WHERE a.sBranchCd = b.sBranchCd" & _
            " AND a.sTranDate >= DATE_FORMAT(c.dUnEncode, '%Y%m%d')" & _
            IIf(lsTranDate <> "", " AND cPostedxx IN('0','1','2')", " AND cPostedxx = '0'") & _
            " AND a.sBranchCd = " & strParm(oApp.BranchCode) & _
         IIf(lsTranDate = "", " ORDER BY sTranDate Asc LIMIT 1", " AND a.sTranDate = " & strParm(lsTranDate))
   Debug.Print lsSQL
   lorec.Open lsSQL, oApp.Connection, , , adCmdText
   
   If Not lorec.EOF Then
      txtField(0).Text = lorec("sTranDate")
      
      If lorec("cPostedxx") = xeStateClosed Then
         For lnCtr = 1 To 8
            txtField(lnCtr).Enabled = True
            txtField(lnCtr).Text = 0#
         Next
         txtField(0).Enabled = True
         
      End If
         txtField(1).Text = Format(lorec("nTotalSle"), "#,##0.00")
         txtField(2).Text = Format(lorec("nOthersxx"), "#,##0.00")
         txtField(5).Text = Format(lorec("nCrdtCard"), "#,##0.00")
         txtField(6).Text = Format(lorec("nFinAmtxx"), "#,##0.00")
         txtField(8).Text = Format(lorec("nDepositd"), "#,##0.00")
         Label2.Caption = TransStat(CInt(lorec("cPostedxx")))
         lsDate = lorec("sTranDate")
   Else
      MsgBox "No Record Found!!", vbOKOnly
      txtField(0).Text = ""
      txtField(8).Text = 0#
      Label2.Caption = "UNKNOWN"
      lsDate = ""
   End If
   Label1(0).ForeColor = &HFF&
End Sub


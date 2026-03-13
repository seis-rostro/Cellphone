VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmRefundableDep 
   BorderStyle     =   0  'None
   Caption         =   "Transaction Encoding"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   22380
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   8
      Top             =   1305
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmRefundableDep.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   9
      Top             =   660
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
      Picture         =   "frmRefundableDep.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5235
      Index           =   0
      Left            =   1605
      Tag             =   "wt0;fb0"
      Top             =   600
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   9234
      BorderStyle     =   1
      Begin VB.ComboBox cmbField 
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
         Index           =   0
         ItemData        =   "frmRefundableDep.frx":0EF4
         Left            =   1800
         List            =   "frmRefundableDep.frx":0F07
         TabIndex        =   24
         Top             =   2610
         Width           =   1935
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
         Index           =   6
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3435
         Width           =   1995
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
         Height          =   1260
         Index           =   7
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   3825
         Width           =   5745
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
         Height          =   360
         Index           =   4
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2190
         Width           =   5745
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
         Height          =   360
         Index           =   3
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1800
         Width           =   5745
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
         Height          =   360
         Index           =   1
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   645
         Width           =   2655
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
         Height          =   360
         Index           =   2
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1410
         Width           =   2655
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
         Index           =   5
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3045
         Width           =   1995
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
         Index           =   0
         Left            =   1800
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7560
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   4860
         Top             =   255
         Width           =   2715
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "UNKNOWN"
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
         Left            =   4920
         TabIndex        =   19
         Tag             =   "eb0;et0"
         Top             =   285
         Width           =   2540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amt Paid"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   3555
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   17
         Top             =   3945
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   16
         Top             =   2310
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No"
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1530
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   765
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   2730
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Check Amt"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   3165
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
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
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1920
         Tag             =   "et0;ht2"
         Top             =   315
         Width           =   2655
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5235
      Index           =   1
      Left            =   9345
      Tag             =   "wt0;fb0"
      Top             =   600
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   9234
      BorderStyle     =   1
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4695
         Left            =   0
         TabIndex        =   21
         Top             =   465
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   8281
         _Version        =   393216
      End
      Begin VB.Label Label3 
         Caption         =   "LIST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   735
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   22
      Top             =   1305
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
      Picture         =   "frmRefundableDep.frx":0F35
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   23
      Top             =   660
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmRefundableDep.frx":16AF
   End
End
Attribute VB_Name = "frmRefundableDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmRefundableDep"

Private WithEvents oTrans As clsRefundableDep
Attribute oTrans.VB_VarHelpID = -1

Private oSkin As clsFormSkin


Dim pbSave As Boolean
Dim psSelected() As String
Dim pnIndex As Integer
Dim pnCtr As Integer
Dim pnRow As Integer
Dim pbFormLoad As Boolean
Dim pbMasterGotFocus As Boolean
Dim pbGridGotFocus As Boolean
Dim pnActiveRow As Integer

Private Sub cmbField_Validate(Index As Integer, Cancel As Boolean)
    On Error GoTo errProc

    cmbField(Index).Text = TitleCase(cmbField(Index).Text)

    If Index = 0 Then ' Payment type
        Debug.Print cmbField(0).ListIndex
        oTrans.Master("cPaymForm") = cmbField(0).ListIndex
    End If

    Exit Sub

errProc:
    ShowError "cmbField_Validate(" & Index & ", " & Cancel & ")", True
End Sub


Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub
Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean


   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(4).Visible = Not lbShow
   'cmdButton(5).Visible = Not lbShow

   cmdButton(1).Visible = Not lbShow 'new
   cmdButton(2).Visible = lbShow 'save
   cmdButton(3).Visible = lbShow 'cancel
   cmdButton(4).Visible = Not lbShow 'close

'   xrFrame2.Enabled = lbShow
'   xrFrame3.Enabled = lbShow

   If Not lbShow Then cmdButton(1).SetFocus
End Sub
Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   txtField_LostFocus pnIndex
      Select Case Index
        'Case 0 'Browse
            'oTrans.SearchTransaction
            'LoadDetail
        Case 1 'New
            oTrans.InitTransaction
            oTrans.NewTransaction
            initButton xeModeAddNew
            InitGrid
            InitFields
            LoadDetail
            xrFrame1(0).Enabled = True
            xrFrame1(1).Enabled = True
            MSFlexGrid1.Row = 1
        Case 2 'save
            
            If IsEntryOkay Then
                Debug.Print ("sample " + txtField(5).Text)
                If oTrans.SaveTransaction Then
                   MsgBox "Transaction Saved Successfully!!!", vbInformation, "Notice"
                   initButton xeModeReady
                   ClearAll
                Else
                   MsgBox "Unable to Saved Successfully!!!", vbInformation, "Notice"
                End If
            End If
        
        Case 3 'Cancel
        Dim lsRep As Long
'        If Not oTrans.OpenTransaction(oTrans.Master("sTransNox")) Then Call InitFields
'         Call initButton(xeModeReady)
             lsRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'
         If lsRep = vbYes Then
            ClearAll
            xrFrame1(0).Enabled = False
            xrFrame1(1).Enabled = False
            initButton xeModeReady
         Else
            txtField(pnIndex).SetFocus
         End If
        Case 4 'close
               Unload Me
      End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub LoadMaster()
   With oTrans
      txtField(0).Text = .Master(0)
      txtField(1).Text = strLongDate(.Master(1))
   End With
End Sub

Private Sub LoadDetail()
    Dim rsDetail As ADODB.Recordset
    Dim lnRow As Integer
    Dim lnCtr As Integer

    ' Call the class function to load the data
    If Not oTrans.LoadDetail() Then
        MsgBox "No data available to load.", vbExclamation, "Load Detail"
        Exit Sub
    End If

    ' Retrieve the result set using the Recordset property
    Set rsDetail = oTrans.Recordset

    ' Bind the result set to the grid
    With MSFlexGrid1
        ' Set up the grid
        .Rows = 1 ' Clear existing rows, keeping only the header row
        If Not rsDetail.EOF Then
            rsDetail.MoveLast
            lnRow = rsDetail.RecordCount
            rsDetail.MoveFirst
            .Rows = lnRow + 1 ' Header + data rows
        Else
            .Rows = 2 ' Just header + empty row
        End If

        ' Adjust column width based on row count
        If .Rows > 16 Then
            .ColWidth(1) = 2950
        Else
            .ColWidth(1) = 3100
        End If

        ' Populate the grid
        lnCtr = 1 ' Start populating from the second row
        Do Until rsDetail.EOF
            .TextMatrix(lnCtr, 0) = lnCtr ' Row number
            .TextMatrix(lnCtr, 1) = IFNull(rsDetail("sBranchNm"), "")
            .TextMatrix(lnCtr, 2) = IFNull(rsDetail("sDescript"), "")
            .TextMatrix(lnCtr, 3) = IFNull(rsDetail("sPayeeNme"), "")
            .TextMatrix(lnCtr, 4) = Format(IFNull(rsDetail("dCheckDte"), Date), "MMMM dd, yyyy")
            .TextMatrix(lnCtr, 5) = Format(IFNull(rsDetail("nAmountxx"), "0.00"), "#,##0.00")
            .TextMatrix(lnCtr, 6) = IFNull(rsDetail("sRemarksx"), "") ' Remarks (hidden column)
            .TextMatrix(lnCtr, 7) = IFNull(rsDetail("sBranchCd"), "") ' BranchCd (hidden column)
            .TextMatrix(lnCtr, 8) = IFNull(rsDetail("sPrtclrID"), "") ' sPrtclrID (hidden column)
            .TextMatrix(lnCtr, 9) = IFNull(rsDetail("sSourceCd"), "") ' BranchCd (hidden column)
            .TextMatrix(lnCtr, 10) = IFNull(rsDetail("sTransNox"), "") ' sPrtclrID (hidden column)
            lnCtr = lnCtr + 1
            
            rsDetail.MoveNext
        Loop
    End With

    ' Clean up
    Set rsDetail = Nothing
End Sub


Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsRefundableDep
   Set oTrans.AppDriver = oApp
   
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
   oTrans.NewTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction
   

   InitGrid
   InitFields
   LoadDetail

   'cmbTranType.ListIndex = 1
'   ClearFields
'   initButton xeModeAddNew

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub
Private Sub clearFields()
   Dim loTxt As TextBox
   Dim lnCtr As Integer
   
   For lnCtr = 1 To 3
      txtField(lnCtr) = ""
   Next
   
   For Each loTxt In txtField
      loTxt.BackColor = oApp.getColor("EB")
   Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Clear
   
      .Cols = 6
      .Rows = 2
      .Font = "MS Sans Serif"
      .RowHeight(0) = 350

      ' Column Titles
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "Particular"
      .TextMatrix(0, 3) = "Payee"
      .TextMatrix(0, 4) = "Check Date"
      .TextMatrix(0, 5) = "Amount"
      .Row = 0

      ' Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 5 ' Center alignment
      Next

      .ColWidth(0) = 400
      .ColWidth(1) = 2850
      .ColWidth(2) = 2850
      .ColWidth(3) = 2850
      .ColWidth(4) = 1700
      .ColWidth(5) = 1460

      ' Column Specific Alignment (for Amount column)
      .ColAlignment(1) = 1 ' Left-aligned for Branch
      .ColAlignment(2) = 1 ' Left-aligned for Particular
      .ColAlignment(3) = 1 ' Left-aligned for Payee
      .ColAlignment(4) = 1 ' Left-aligned for Check Date
      .ColAlignment(5) = 7 ' Right-aligned for Amount (currency)

      ' Add an extra column for remarks
      .Cols = 11
      .ColWidth(6) = 0 ' Set the remarks column width to 0 to hide it
      .ColWidth(7) = 0 ' Set the branchcd column width to 0 to hide it
      .ColWidth(8) = 0 ' Set the particularid column width to 0 to hide it
      .ColWidth(9) = 0 ' Set the particularid column width to 0 to hide it
      .ColWidth(10) = 0 ' Set the particularid column width to 0 to hide it
      ' Select entire range for column selection (if needed)
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub



Private Sub InitFields()
   Dim loTxt As TextBox
   With oTrans
      For Each loTxt In txtField
         Select Case loTxt.Index
         Case 0
            loTxt = Format(.Master(loTxt.Index), "@@@@@@-@@@@@@")
         Case 1
            loTxt = Format(.Master(loTxt.Index), "MMMM DD, YYYY")
         
         Case Else
            loTxt = ""
         End Select
      Next
'      txtOthers(1) = ""
'      txtOthers(2) = "0"
'      txtOthers(3) = "0.00"
   End With

   With MSFlexGrid1
      .Rows = 2
      .TextMatrix(1, 0) = 1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = "0"

      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   pnRow = 1
   
   Call LoadDetail
   
End Sub

Private Sub MSFlexGrid1_Click()
Dim lnCtr As Integer
   With oTrans
      pnActiveRow = MSFlexGrid1.Row
      pnRow = pnActiveRow
      Call showdetail
   End With
   
   If xrFrame1(1).Enabled = True Then txtField(2).SetFocus
   
End Sub

Private Sub txtField_KeyPress(Index As Integer, keyascii As Integer)
     Select Case Index
        Case 2
            ' Limit input to 12 characters for txtField(2)
            If Len(txtField(Index).Text) >= 12 And keyascii <> 8 Then
                keyascii = 0 ' Discard additional input
            End If
        Case Else
            ' Optional: Handle unexpected indices, if necessary
    End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub
Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
    Dim lsOldProc As String

   lsOldProc = "txtField_Validate"
   ''On Error GoTo errProc

    With txtField(Index)
      .Text = TitleCase(.Text)
      With MSFlexGrid1
            Select Case Index
                Case 2 'reference
                    oTrans.Master(17) = txtField(2).Text
                Case 4 'check amt
                    oTrans.Master(4) = IIf(.TextMatrix(.Row, 5) = "", txtField(7).Text, .TextMatrix(.Row, 5))
                Case 7 'remarks
                    oTrans.Master(7) = IIf(.TextMatrix(.Row, 6) = "", txtField(7).Text, .TextMatrix(.Row, 6))
                    
            End Select
        End With
    End With
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub
Private Sub showdetail()
   cmbField(0).ListIndex = 0
   With MSFlexGrid1
      If .Row = 0 Then
         txtField(2) = ""
         txtField(3) = ""
         txtField(4) = ""
         txtField(5) = "0.00"
         txtField(6) = "0.00"
         txtField(7) = ""
         cmbField(0) = ""
      Else
         'txtField(2) = .TextMatrix(.Row, 1) 'refernce
         txtField(3) = .TextMatrix(.Row, 1) 'branch
         txtField(4) = .TextMatrix(.Row, 2) 'particular
         txtField(5) = .TextMatrix(.Row, 5) 'check amt
         txtField(6) = "0.00" 'amt paid
         txtField(7) = .TextMatrix(.Row, 6) 'remarks
         
          
      End If
      Debug.Print "Check amt == "; Replace(Trim(.TextMatrix(.Row, 5)), ",", "")
      oTrans.Master(6) = "0"
      oTrans.Master(4) = Replace(Trim(.TextMatrix(.Row, 5)), ",", "") 'check amt
      oTrans.Master(7) = .TextMatrix(.Row, 6) 'remarks
      oTrans.Master(8) = .TextMatrix(.Row, 7) 'branchcd
      oTrans.Master(9) = .TextMatrix(.Row, 8) 'particularid
      oTrans.Master(10) = .TextMatrix(.Row, 9) 'sSourceCd
      oTrans.Master(11) = .TextMatrix(.Row, 10) 'sSourceNo
      
      
   End With
End Sub

'Public Sub setMaster()
'
'    With MSFlexGrid1
'      oTrans.Master("sBranchCd") = .TextMatrix(.Row, 7)
'      oTrans.Master(17) = txtField(2).Text
'      oTrans.Master("sPrtclrID") = .TextMatrix(.Row, 8)
'
'   End With
'End Sub


Private Sub ClearAll()

    cmbField(0).ListIndex = 0
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 1

    txtField(2) = ""
    txtField(3) = ""
    txtField(4) = ""
    txtField(5) = "0.00"
    txtField(6) = "0.00"
    txtField(7) = ""
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 1
    txtField(2) = ""
    txtField(3) = ""
    txtField(4) = ""
    txtField(5) = "0.00"
    txtField(6) = "0.00"
    txtField(7) = ""
End Sub

Private Function IsEntryOkay() As Boolean
    Dim isValid As Boolean
    
        isValid = True
    If cmbField(0).ListIndex = -1 Then
        isValid = False
        MsgBox "Unable to save transactions!!!" & vbCrLf & _
                  "Pls check your payment type then try again!!" & "WARNING"
        cmbField(0).SetFocus
        GoTo endProc
    End If
    
    If Trim(txtField(2).Text) = "" Then
        isValid = False
        MsgBox "Unable to save transactions!!!" & vbCrLf & _
                  "Pls check your Reference No then try again!!" & "WARNING"
        txtField(2).SetFocus
        GoTo endProc
    End If
    
    If Trim(txtField(3).Text) = "" Then
        isValid = False
        MsgBox "Unable to save transactions!!!" & vbCrLf & _
                  "Pls check Branch then try again!!" & "WARNING"
        txtField(3).SetFocus
        GoTo endProc
    End If
    
'    If Trim(txtField(4).Text) = "" Then
'        isValid = False
'        MsgBox "Unable to save transactions!!!" & vbCrLf & _
'                  "Pls check Branch then try again!!" & "WARNING"
'        txtField(4).SetFocus
'        GoTo endProc
'    End If
    IsEntryOkay = True
    
endProc:
   Exit Function
End Function


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


VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{305FFFF2-59E4-4627-A34F-B5BA746770B4}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmReceipt 
   BorderStyle     =   0  'None
   Caption         =   "Receipt"
   ClientHeight    =   6225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5550
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   9790
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1215
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "frmReceipt.frx":0000
         Top             =   1470
         Width           =   4035
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   6780
         MaxLength       =   6
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   630
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1215
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmReceipt.frx":0008
         Top             =   2130
         Width           =   4035
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   5
         Left            =   6780
         MaxLength       =   25
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1470
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1215
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmReceipt.frx":0010
         Top             =   1800
         Width           =   4035
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   6780
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   960
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   6
         Left            =   6780
         MaxLength       =   25
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1980
         Width           =   1950
      End
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   1515
         Left            =   105
         TabIndex        =   26
         Top             =   2835
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   2672
         AllowBigSelection=   -1  'True
         AutoAdd         =   0   'False
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
         FORECOLOR       =   -2147483640
         FORECOLORFIXED  =   -2147483630
         FORECOLORSEL    =   -2147483634
         FORMATSTRING    =   ""
         Object.HEIGHT          =   1515
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
         MOUSEICON       =   "frmReceipt.frx":0018
         MOUSEPOINTER    =   0
         REDRAW          =   -1  'True
         RIGHTTOLEFT     =   0   'False
         ROWS            =   6
         SCROLLBARS      =   3
         SCROLLTRACK     =   0   'False
         SELECTIONMODE   =   0
         Object.TOOLTIPTEXT     =   ""
         WORDWRAP        =   0   'False
      End
      Begin VB.Label lblHead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tel No: 075-5221085"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   1440
         TabIndex        =   25
         Top             =   585
         Width           =   3570
      End
      Begin VB.Label lblPayment 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Text1"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   6180
         TabIndex        =   16
         Tag             =   "ht0;ft0"
         Top             =   4440
         Width           =   2550
      End
      Begin VB.Label lblChange 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Text1"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   6180
         TabIndex        =   18
         Tag             =   "et0;hb0"
         Top             =   4920
         Width           =   2550
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Check Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   5370
         TabIndex        =   12
         Top             =   2040
         Width           =   1440
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O.R.  No."
         Height          =   195
         Index           =   10
         Left            =   5970
         TabIndex        =   0
         Top             =   690
         Width           =   675
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   0
         Left            =   5970
         TabIndex        =   2
         Top             =   1020
         Width           =   345
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   8
         Top             =   2190
         Width           =   630
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   8
         Left            =   5370
         TabIndex        =   10
         Top             =   1500
         Width           =   1290
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check"
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
         Index           =   7
         Left            =   105
         TabIndex        =   14
         Top             =   2580
         Width           =   555
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "T O T A L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   4515
         TabIndex        =   15
         Top             =   4545
         Width           =   1575
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   5280
         TabIndex        =   17
         Top             =   5010
         Width           =   810
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   1215
         TabIndex        =   22
         Tag             =   "et0"
         Top             =   1800
         Width           =   4035
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1215
         TabIndex        =   21
         Tag             =   "et0"
         Top             =   1470
         Width           =   4035
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   600
         Picture         =   "frmReceipt.frx":0034
         Top             =   525
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "frmReceipt.frx":0CFE
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblHead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perez Blvd., Dagupan City"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   1440
         TabIndex        =   24
         Top             =   390
         Width           =   3570
      End
      Begin VB.Label lblHead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guanzon Merchandising Corporation"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   1395
         TabIndex        =   23
         Top             =   210
         Width           =   3570
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   990
         Index           =   1
         Left            =   180
         Top             =   150
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         Height          =   1020
         Index           =   0
         Left            =   165
         Top             =   135
         Width           =   1110
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   6
         Top             =   1845
         Width           =   570
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receive From"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   4
         Top             =   1500
         Width           =   990
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   9195
      TabIndex        =   20
      Top             =   1815
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
      Picture         =   "frmReceipt.frx":19C8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   9195
      TabIndex        =   19
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Ok"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmReceipt.frx":2142
   End
End
Attribute VB_Name = "frmReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oSkin As FormSkin

Private p_oAppDrivr As AppDriver
Private p_oMod As New MainModules
Private p_bCancelxx As Boolean
Private p_bEmptyORx As Boolean
Private p_nAmtPaidx As Double

Dim pnCtr As Integer
Dim pbFocus As Boolean

Property Set AppDriver(Value As AppDriver)
10       Set p_oAppDrivr = Value
End Property

Property Let AllowEmptyOR(ByVal Value As Boolean)
10       p_bEmptyORx = Value
End Property

Property Let AmountPaid(ByVal Value As Double)
10       p_nAmtPaidx = Value
End Property

Property Get Cancelled() As Boolean
10       Cancelled = p_bCancelxx
End Property

Function GetNextOR() As String
10       Dim lors As Recordset
20       Dim lsCtr As Long

30       Set lors = New Recordset
40       lors.Open "SELECT TOP 1 sORNoxxxx FROM Receipt_Master" & _
      " WHERE LEFT(sTransNox, 2) = " & p_oMod.strParm(p_oAppDrivr.BranchCode) & _
         " AND sORNoxxxx <> " & p_oMod.strParm("") & _
      " ORDER BY dTransact DESC, sORNoxxxx DESC", p_oAppDrivr.Connection, , , adCmdText

50       If lors.EOF Then
60          GetNextOR = Format(1, "000")
70       Else
80          GetNextOR = Format(CLng(lors(0)) + 1, String(Len(lors(0)), "0"))
90       End If

100      Set lors = Nothing
End Function

Private Sub cmdButton_Click(Index As Integer)
10       Select Case Index
   Case 0
      'check first the content of the or
20          If Not isEntryOK() Then
30             MsgBox "Official Receipt contain/s invalid entries!!!" & vbCrLf & _
            "Verify your entries then try again!!!", vbCritical + vbOKOnly, "Warning"
40             Exit Sub
50          End If
      
60          p_bCancelxx = False
70       Case 1
80          p_bCancelxx = True
90       End Select
100      Me.Hide
End Sub

Private Sub Form_Activate()
10       txtField(0).SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10       Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
20          If KeyCode <> vbKeyReturn And pbFocus Then Exit Sub
30          Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
40             p_oMod.SetNextFocus
50          Case vbKeyUp
60             p_oMod.SetPreviousFocus
70          End Select
80       End Select
End Sub

Private Sub Form_Load()
10       If p_oAppDrivr Is Nothing Then Exit Sub
20       If Not (p_oAppDrivr.MDIMain Is Nothing) Then p_oMod.CenterChildForm p_oAppDrivr.MDIMain, Me
   
30       Set oSkin = New FormSkin
40       Set oSkin.AppDriver = p_oAppDrivr
50       Set oSkin.Form = Me
60       oSkin.ApplySkin xeFormTransDetail
70       oSkin.DisableClose = True
   
80       InitGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
10       Set oSkin = Nothing
20       Set p_oMod = Nothing
End Sub

Private Sub InitGrid()
10       With GridEditor1
20          .Cols = 6
30          .Font = "MS Sans Serif"

      'column title
40          .TextMatrix(0, 1) = "Check No"
50          .TextMatrix(0, 2) = "Account No"
60          .TextMatrix(0, 3) = "Bank Name"
70          .TextMatrix(0, 4) = "Check Date"
80          .TextMatrix(0, 5) = "Amount"
90          .Row = 0

      'column alignment
100         For pnCtr = 0 To .Cols - 1
110            .Col = pnCtr
120            .CellFontBold = True
130            .CellAlignment = 1
140         Next

      'column width
150         .ColWidth(0) = 330
160         .ColWidth(1) = 1600
170         .ColWidth(2) = 1600
180         .ColWidth(3) = 2800
190         .ColWidth(4) = 1110
200         .ColWidth(5) = 1100
  
210         .ColFormat(1) = ">"
220         .ColFormat(2) = ">"
230         .ColFormat(5) = "#,##0.00"

240         .ColNumberOnly(5) = True
250         .ColLimit(1) = 15
260         .ColLimit(2) = 15
270         .ColLimit(3) = 30
280         .ColMaxValue(5) = 9999999.99
290         .ColDefault(4) = Format(Now, "MM/DD/YYYY")
300         .ColDefault(5) = "0.00"
    
310         .ColAlignment(1) = 1
320         .ColAlignment(2) = 1
330         .ColAlignment(3) = 1
340         .ColAlignment(4) = 6
350         .ColAlignment(5) = 6
    
360         .Row = 1
370         .Col = 1
380      End With
End Sub

Private Sub Form_Initialize()
10       p_bCancelxx = False
20       p_bEmptyORx = False
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
10       With GridEditor1
20          Select Case .Col
      Case 4
30             If .TextMatrix(.Row, .Col) <> "" Then
40                If IsDate(.TextMatrix(.Row, .Col)) Then
50                   .TextMatrix(.Row, .Col) = Format(.TextMatrix(.Row, .Col), "MM/DD/YYYY")
60                Else
70                   .TextMatrix(.Row, .Col) = Format(txtField(1).Text, "MM/DD/YYYY")
80                End If
90             End If
100         Case 5
110            If CDbl(.TextMatrix(.Row, .Col)) > p_nAmtPaidx Then .TextMatrix(.Row, .Col) = 0#
120            If TotalCheckPayment + CDbl(txtField(5).Text) > p_nAmtPaidx Then .TextMatrix(.Row, .Col) = 0#
         
130            txtField(6).Text = Format(TotalCheckPayment, "#,##0.00")
140            lblChange.Caption = Format(TotalChange, "#,##0.00")
150         End Select
160      End With
End Sub

Private Sub GridEditor1_GotFocus()
10       pbFocus = True
End Sub

Private Sub GridEditor1_LostFocus()
10       pbFocus = False
End Sub

Private Sub txtField_GotFocus(Index As Integer)
10       If txtField(Index).Text <> "" Then
20          txtField(Index).SelStart = 0
30          txtField(Index).SelLength = Len(txtField(Index).Text)
40       End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
10       Dim lvContent
 
   ' put all validation here
20       lvContent = txtField(Index)
30       Select Case Index
   Case 0
40          If p_bEmptyORx = False Then
50             If Not IsNumeric(lvContent) Then
60                txtField(Index).Text = GetNextOR
70                txtField(Index).SetFocus
80             End If
90          Else
100            If Not IsNumeric(lvContent) Then txtField(Index) = ""
110         End If
120      Case 5
130         If Not IsNumeric(lvContent) Then
140            txtField(Index).Text = 0#
150         Else
160            If CDbl(txtField(Index).Text) + CDbl(txtField(6)) > p_nAmtPaidx Then
170               txtField(Index).Text = 0#
180            End If
190         End If
200         txtField(Index).Text = Format(txtField(Index).Text, "#,##0.00")
210         lblChange = Format(TotalChange, "#,##0.00")
220      End Select
End Sub

Private Function isEntryOK() As Boolean
10       Dim lnCtr As Integer
20       Dim lnCash As Double
30       Dim lnCheck As Double
40       Dim lnTotal As Double
  
50       isEntryOK = False
   
   ' check first the OR No
60       If txtField(0) = Empty And p_bEmptyORx = False Then GoTo endProc
70       If txtField(2) = Empty Then GoTo endProc
   
80       lnCash = CDbl(txtField(5))
90       lnCheck = CDbl(txtField(6))
100      lnTotal = lnCash + lnCheck
   
110      If lnTotal <> p_nAmtPaidx Then GoTo endProc
   
'   If CDbl(lnTotal) < CDbl(lblPayment) Then GoTo endProc
'   If lnCheck > CDbl(txtField(5).Text) Then GoTo endProc
   
120      isEntryOK = True
   
endProc:
130      Exit Function
End Function

Private Function TotalCheckPayment() As Double
10       Dim lnCtr As Integer
20       Dim lnSum As Double

30       lnSum = 0
40       With GridEditor1
50          For lnCtr = 1 To .Rows - 1
60             lnSum = lnSum + CDbl(.TextMatrix(lnCtr, 5))
70          Next
80       End With
   
90       TotalCheckPayment = lnSum
End Function

Private Function TotalChange() As Double
10       TotalChange = CDbl(txtField(6)) + CDbl(txtField(5)) - CDbl(lblPayment)
End Function


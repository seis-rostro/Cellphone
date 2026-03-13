VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Sales_Return 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Sales Return"
   ClientHeight    =   9420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   140.807
   ScaleMode       =   0  'User
   ScaleWidth      =   100.902
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame3 
      Height          =   1275
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   4290
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   2249
      ClipControls    =   0   'False
      Begin VB.TextBox txtDetail 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   2
         Left            =   8175
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   525
         Width           =   1755
      End
      Begin VB.TextBox txtDetail 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   60
         TabIndex        =   20
         Top             =   525
         Width           =   8085
      End
      Begin VB.OptionButton optClientTp 
         Caption         =   "CHARGE INVOICE"
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
         Left            =   60
         TabIndex        =   15
         Tag             =   "wt0;fb0"
         Top             =   90
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton optClientTp 
         Caption         =   "CP SALES"
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
         Index           =   1
         Left            =   60
         TabIndex        =   16
         Tag             =   "wt0;fb0"
         Top             =   300
         Width           =   2115
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   4305
         TabIndex        =   18
         Top             =   60
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&IMEI / BARCODE"
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
         Index           =   15
         Left            =   60
         TabIndex        =   19
         Top             =   1020
         Width           =   8085
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&QTY"
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
         Index           =   7
         Left            =   8190
         TabIndex        =   21
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&REFERNCE NO"
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
         Index           =   1
         Left            =   2955
         TabIndex        =   17
         Top             =   105
         Width           =   1335
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3705
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   5610
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   6535
      BackColor       =   14737632
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3510
         Left            =   75
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   90
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   6191
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   75
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2385
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
      Picture         =   "frmCP_Sales_Return.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   540
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
      Picture         =   "frmCP_Sales_Return.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   31
      TabStop         =   0   'False
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
      Picture         =   "frmCP_Sales_Return.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1770
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Del. Row"
      AccessKey       =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_Sales_Return.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   75
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1155
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_Sales_Return.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   540
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
      Picture         =   "frmCP_Sales_Return.frx":2562
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3735
      Left            =   1590
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   6588
      BorderStyle     =   1
      Begin VB.CheckBox chkClientTp 
         Caption         =   "Company / Institution"
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
         Left            =   1365
         TabIndex        =   4
         Tag             =   "wt0;fb0"
         Top             =   1215
         Width           =   2415
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1365
         TabIndex        =   10
         Top             =   2370
         Width           =   4950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   915
         Index           =   5
         Left            =   1365
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2685
         Width           =   4950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   600
         Index           =   3
         Left            =   1365
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1740
         Width           =   4950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1365
         MaxLength       =   50
         TabIndex        =   3
         Top             =   735
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   1365
         TabIndex        =   1
         Top             =   255
         Width           =   2310
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   7830
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1275
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1365
         TabIndex        =   6
         Top             =   1425
         Width           =   4950
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*PIC"
         Height          =   195
         Index           =   5
         Left            =   705
         TabIndex        =   9
         Top             =   2415
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   270
         Index           =   12
         Left            =   -45
         TabIndex        =   11
         Top             =   2685
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   11
         Left            =   -45
         TabIndex        =   7
         Top             =   1740
         Width           =   1065
      End
      Begin VB.Label lblTrantotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "999,000.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   630
         Left            =   6585
         TabIndex        =   25
         Top             =   480
         Width           =   3240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   6585
         TabIndex        =   24
         Top             =   255
         Width           =   2070
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   10
         Left            =   150
         TabIndex        =   2
         Top             =   765
         Width           =   1065
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
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   1065
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1455
         Tag             =   "et0;ht2"
         Top             =   375
         Width           =   2325
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   195
         Index           =   14
         Left            =   6570
         TabIndex        =   13
         Top             =   1290
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Top             =   1470
         Width           =   660
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   75
      TabIndex        =   32
      Top             =   1155
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Print"
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
      Picture         =   "frmCP_Sales_Return.frx":2CDC
   End
End
Attribute VB_Name = "frmCP_Sales_Return"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Sales_Invoice"
Private Const pxeAPPNAME = "CP Sales Invoice"
Private WithEvents oTrans As clsCPSalesReturn
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim psTransNox As String
Dim pnIndex As Integer
Dim pnCtr As Integer

Dim pbLoaded As Boolean
Dim pbSave As Boolean
Dim bLoadRecord As Boolean
Dim pnRow As Integer

Private Sub chkClientTp_Click()
10    oTrans.Master("cClientTp") = chkClientTp.Value
End Sub

Private Sub cmdButton_Click(Index As Integer)
1     Dim lsOldProc As String
2     Dim lnRep As Long

3     lsOldProc = "cmdButton_Click"
10    ''On Error GoTo errProc

20    Select Case Index
      Case 0 'save
30       If Not isEntryOk Then Exit Sub

40       If oTrans.SaveTransaction Then
50          MsgBox "Transaction saved successfuly.", vbInformation, pxeAPPNAME

            're-open the previous trasaction made for printing.
60          If Not oTrans.OpenTransaction(psTransNox) Then
70             MsgBox "Unable to open transaction.", vbCritical, pxeAPPNAME
80          Else
90             bLoadRecord = True
               lnRep = MsgBox("Do you want to confirm transaction???", vbQuestion + vbYesNo)
               If lnRep = vbYes Then
                  If oTrans.CloseTransaction(oTrans.Master("sTransNox")) Then
                     Call InitForm
                     Call InitGrid
                  End If
               End If
100         End If

110         Call initButton
120         cmdButton(6).SetFocus
130      Else
140         MsgBox "Unable to save transaction.", vbCritical, pxeAPPNAME
150      End If
160   Case 1 'search
170      Select Case pnIndex
         Case 2
180         Call txtField_KeyDown(pnIndex, vbKeyF3, 0)
190      End Select
200   Case 2 'delrow
210      If oTrans.deleteDetail(pnRow) Then
220         Call refreshGrid
230      End If
235      txtDetail(1).SetFocus
240   Case 3 'cancel
250      If oTrans.InitTransaction Then
260         Call InitForm
270         Call InitGrid
280         cmdButton(4).SetFocus
290         bLoadRecord = False
300      End If
310   Case 4 'new
320      If oTrans.NewTransaction Then
340         Call InitEntry
            initButton
350      End If
360   Case 5 'close
370      Unload Me
380   Case 6 'print
390      If Not bLoadRecord Then
400         MsgBox "Unable to Print Transaction.", vbCritical, pxeAPPNAME
410         Exit Sub
420      End If

430      lnRep = MsgBox("Do you want to print this transaction?", vbQuestion & vbYesNo, pxeAPPNAME)
440      If lnRep = vbYes Then
450         If PrintTrans Then
460            If oTrans.CloseTransaction(psTransNox) Then MsgBox "Printing..."
470         End If

480         If MsgBox("Reprint?", vbQuestion & vbYesNo, pxeAPPNAME) = vbYes Then PrintTrans
490      End If
500   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Activate()
10    If Not pbLoaded Then pbLoaded = True

20    oApp.MenuName = Me.Tag
30    Me.ZOrder 0

40    bLoadRecord = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10    Select Case KeyCode
      Case vbKeyReturn, vbKeyUp, vbKeyDown
20       Select Case KeyCode
            Case vbKeyReturn, vbKeyDown
30             If GetFocus = txtDetail(1).hwnd Then Exit Sub
40             SetNextFocus
50          Case vbKeyUp
60             SetPreviousFocus
70       End Select
80    Case vbKeyF9
90       txtDetail(1).Enabled = True
100      txtDetail(1).SetFocus
110   Case vbKeyF10
120      txtDetail(2).Enabled = True
130      txtDetail(2).SetFocus
140   Case vbKeyF11
150      txtDetail(3).Enabled = True
160      txtDetail(3).SetFocus
170   Case vbKeyF12
180      txtField(13).Enabled = True
190      txtField(13).SetFocus
200   End Select
End Sub

Private Sub Form_Load()
10    Dim lsOldProc As String

20    lsOldProc = "Form_Load"
30    ''On Error GoTo errProc

40    CenterChildForm mdiMain, Me

50    Set oSkin = New clsFormSkin
60    Set oSkin.AppDriver = oApp
70    Set oSkin.Form = Me
80    oSkin.ApplySkin xeFormTransEqualLeft

90    Set oTrans = New clsCPSalesReturn
100   With oTrans
110      Set .AppDriver = oApp
120      .Branch = oApp.BranchCode
130      .InitTransaction
140      .QueryDetailTable = "CP_CO_Detail"
150      .QueryMasterTable = "CP_CO_Master"

160      If .NewTransaction Then
180         Call InitEntry
190         Call InitGrid

200         bLoadRecord = False
210      End If
220   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
10    Set oSkin = Nothing
20    Set oTrans = Nothing
30    pbLoaded = False
End Sub

Private Sub InitEntry()
10    Dim lsOldProc As String

20    lsOldProc = pxeMODULENAME & ".InitEntry"
30    ''On Error GoTo errProc

40    With oTrans
50       psTransNox = .Master("sTransNox")
60       txtField(0) = Format(.Master("sTransNox"), "@@@@@@-@@@@@@")
70       txtField(1) = Format(.Master("dTransact"), "MMMM DD, YYYY")
80       txtField(2) = ""
90       txtField(3) = ""
100      txtField(4) = IFNull(.Master("sCPerson1"), "")
110      txtField(5) = .Master("sRemarksx")
120      txtField(6) = Format(.Master("nABalance"), "#,##0.00")

130      txtDetail(0) = ""
140      txtDetail(1) = ""
150      txtDetail(2) = 0

160      lblTrantotal = Format(.Master("nTranTotl"), "#,##0.00")
170      chkClientTp.Value = .Master("cClientTp")

180      pnRow = 0
190   End With
      With MSFlexGrid1
         .TextMatrix(.Row, .Col) = ""
      End With
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub initButton()
10    With oTrans
20       cmdButton(0).Visible = .EditMode = xeModeAddNew
30       cmdButton(1).Visible = .EditMode = xeModeAddNew
40       cmdButton(2).Visible = .EditMode = xeModeAddNew
50       cmdButton(3).Visible = .EditMode = xeModeAddNew
60       cmdButton(4).Visible = .EditMode = xeModeReady
70       cmdButton(5).Visible = .EditMode = xeModeReady
80       cmdButton(6).Visible = .EditMode = xeModeReady

90       xrFrame1.Enabled = .EditMode = xeModeAddNew
100      xrFrame2.Enabled = .EditMode = xeModeAddNew
110      xrFrame3.Enabled = .EditMode = xeModeAddNew
120   End With
End Sub

Private Sub InitForm()
10    txtField(0) = ""
20    txtField(1) = ""
30    txtField(2) = ""
40    txtField(3) = ""
50    txtField(4) = ""
60    txtField(5) = ""
70    txtField(6) = Format(0#, "##0.00 %")

80    txtDetail(0) = ""
90    txtDetail(1) = ""
100   txtDetail(2) = ""

110   lblTrantotal = Format(0#, "#,##0.00")
120   chkClientTp.Value = 0

130   Call initButton

140   pnRow = 0
End Sub

Private Sub InitGrid()
10    Dim lnCtr As Integer

20    With MSFlexGrid1
30       .Clear
40       .Cols = 6
50       .Rows = 2

60       .TextMatrix(0, 0) = ""
70       .TextMatrix(0, 1) = "IMEI/Barcode"
80       .TextMatrix(0, 2) = "Description"
90       .TextMatrix(0, 3) = "Qty."
100      .TextMatrix(0, 4) = "Amount"
110      .TextMatrix(0, 5) = "Total Amount"

120      .Row = 0
130      'column alignment
140      For lnCtr = 0 To .Cols - 1
150         .Col = lnCtr
160         .CellFontBold = True
170         .CellAlignment = flexAlignCenterCenter
180      Next

190      .ColWidth(0) = "450"
200      .ColWidth(1) = "1600"
210      .ColWidth(2) = "4000"
220      .ColWidth(4) = "1190"
230      .ColWidth(5) = "1600"

240      .Row = 1
250      .Col = 0
260      .ColSel = .Cols - 1
270   End With
End Sub

Private Sub MSFlexGrid1_Click()
10   With MSFlexGrid1
20      .Col = 1
30      .ColSel = .Cols - 1
40   End With
End Sub

Private Sub MSFlexGrid1_RowColChange()
10    With MSFlexGrid1
20       .Col = 0
30       .ColSel = .Cols - 1

40       If .Row >= 1 Then
50          pnRow = .Row - 1
60          txtDetail(2) = .TextMatrix(.Row, 3)

70          If pbLoaded And xrFrame3.Enabled Then txtDetail(2).SetFocus
80       Else
90          pnRow = 0
100      End If
110   End With
End Sub

Private Sub optClientTp_Click(Index As Integer)
10    With oTrans
20       Select Case Index
         Case 0
40          .QueryDetailTable = "CP_CO_Detail"
50          .QueryMasterTable = "CP_CO_Master"
60       Case 1
70          .QueryDetailTable = "CP_SO_Detail"
80          .QueryMasterTable = "CP_SO_Master"
90       End Select
100   End With
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
10    With MSFlexGrid1
20       Select Case Index
         Case 1, 2
40          .TextMatrix(pnRow + 1, Index) = oTrans.Detail(pnRow, Index)
50       Case 7
60          .TextMatrix(pnRow + 1, 3) = oTrans.Detail(pnRow, Index)
70          Call refreshGrid
80       Case 8
90          .TextMatrix(pnRow + 1, 4) = Format(oTrans.Detail(pnRow, Index), "#,##0.00")
100         Call refreshGrid
110      Case 15
120         txtDetail(0) = oTrans.Master("sReferNox")
130      End Select
140   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
10    Select Case Index
      Case 0
20    Case 1
30       txtField(Index) = Format(oTrans.Master(Index), "MMMM DD, YYYY")
40    Case 6
50       txtField(Index) = Format(oTrans.Master(Index), "#,##0.00")
60    Case 7
70       lblTrantotal.Caption = Format(oTrans.Master(Index), "#,##0.00")
80    Case Else
90       txtField(Index) = oTrans.Master(Index)
100   End Select
End Sub

Private Sub txtDetail_GotFocus(Index As Integer)
10    With txtDetail(Index)
20       .SelStart = 0
30       .SelLength = Len(.Text)
40       .BackColor = oApp.getColor("HT1")
50    End With

60    pnIndex = Index
End Sub

Private Sub txtDetail_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
10    Select Case Index
      Case 1 'barcode/imei
20       Select Case KeyCode
         Case vbKeyReturn
30          Call prcSearch(Index)
40       Case vbKeyF3
50          Call prcSearch(Index, True)
60       End Select
70    End Select
End Sub

Private Sub prcSearch(ByVal lnIndex As Integer, Optional ByVal lbSearch As Boolean = False)
10    Dim lsValue As String
20    Dim lsBarrCode As String
30    Dim lsQty As String
40    Dim lnCtr As Integer
50    Dim lnQty As Integer
60    Dim lbDuplicate As Boolean
70    Dim lsOldProc As String

80    lsOldProc = pxeMODULENAME & ".prcSearch"
90    ''On Error GoTo errProc

100   With txtDetail(lnIndex)
110      'if no customer selected, don't allow for item search.
120      If Trim(oTrans.Master("sClientID")) = "" Then
130         MsgBox "No customer detected. Please input a customer name.", vbInformation, "Sales Return"

140         .Text = ""
150         txtField(2).SetFocus
160         Exit Sub
170      End If

180      lsValue = Trim(Left(.Text, 6))
190      lsBarrCode = .Text
200      lnQty = 0

210      If lsValue = "" And lbSearch Then
220         GoTo searchDetail
230      ElseIf lsValue = "" And Not lbSearch Then
240         Exit Sub
250      End If

260      For lnCtr = 1 To Len(lsValue)
270         If LCase(Left(Right(lsValue, lnCtr), 1)) = "x" Then
280            lsQty = Left(lsValue, Len(Trim(lsValue)) - lnCtr)
290            If IsNumeric(lsQty) Then
300               lnQty = lsQty
310               If Right(.Text, 1) = "x" Then
320                  lnQty = 1
330               Else
340                  lsBarrCode = Right(.Text, Len(.Text) - (Len(lsQty) + 1))
350               End If
360            Else
370               lnQty = 1
380               lsBarrCode = .Text
390            End If
400         End If
410      Next

420      With MSFlexGrid1
430         For lnCtr = 1 To .Rows - 1
440            If Trim(LCase(lsBarrCode)) = Trim(LCase(.TextMatrix(lnCtr, 1))) Then
450               .TextMatrix(lnCtr, 3) = CDbl(IIf(.TextMatrix(lnCtr, 3) = "", 0, .TextMatrix(lnCtr, 3))) + lnQty
460               oTrans.Detail(lnCtr - 1, "nQuantity") = CDbl(.TextMatrix(lnCtr, 3))
470               Call GrandTotal
480               lbDuplicate = True
490            End If
500         Next
510      End With

520      If Not lbDuplicate Then
530         If Trim(lsBarrCode) <> "" Then

searchDetail:
540            Call InsertDetail(lnQty, lsBarrCode)
550         End If
560      End If
570      .Text = ""

580      .SetFocus
590   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub txtDetail_LostFocus(Index As Integer)
10    With txtDetail(Index)
20       .BackColor = oApp.getColor("EB")
30    End With
End Sub

Private Sub txtDetail_Validate(Index As Integer, Cancel As Boolean)
10    With txtDetail(Index)
20       Select Case Index
         Case 0
40          oTrans.Master("sReferNox") = Trim(.Text)
80       Case 2
90          If Not IsNumeric(.Text) Then .Text = 0#
100         oTrans.Detail(pnRow, "nQuantity") = CInt(.Text)
110         .Text = ""
120      End Select
130   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
10    With txtField(Index)
20       Select Case Index
         Case 1
30         .Text = Format(.Text, "MM/DD/YYYY")
40       Case 9
50         .Text = Format(.Text, "##0.00")
60       End Select

70       .SelStart = 0
80       .SelLength = Len(.Text)
90       .BackColor = oApp.getColor("HT1")
100   End With

110   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
10    Dim lsOldProc As String

20    lsOldProc = "txtField_KeyDown"
30    ''On Error GoTo errProc

40    If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
50       With txtField(Index)
60          If KeyCode = vbKeyF3 Then
70             oTrans.SearchMaster Index, .Text
80             If .Text <> "" Then SetNextFocus
90          Else
100            If .Text <> "" Then oTrans.SearchMaster Index, .Text
110         End If
120      End With
130      KeyCode = 0
140   End If

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
10    With txtField(Index)
20       .BackColor = oApp.getColor("EB")
30    End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
10    With txtField(Index)
20       Select Case Index
         Case 1
40          If Not IsDate(.Text) Then .Text = oApp.ServerDate
50          .Text = Format(.Text, "MMMM DD, YYYY")

60          oTrans.Master(Index) = CDate(.Text)
70       Case 5
80          .Text = TitleCase(.Text)

90          oTrans.Master(Index) = .Text
100      Case Else
110         oTrans.Master(Index) = .Text
120      End Select
130   End With
End Sub

Private Sub GrandTotal()
10    Dim lsOldProc As String
20    Dim lnCtr As Integer
30    Dim lnTotal As Currency

40    lsOldProc = pxeMODULENAME & ".GrandTotal"
50    ''On Error GoTo errProc

60    With MSFlexGrid1
70       lnTotal = 0#
80       For lnCtr = 1 To .Rows - 1
90          lnTotal = lnTotal + CDbl(IIf(.TextMatrix(lnCtr, 5) = "", 0, .TextMatrix(lnCtr, 5)))
100      Next
110   End With

120   lblTrantotal.Caption = Format(lnTotal, "#,##0.00")
130   oTrans.Master("nTranTotl") = CDbl(lnTotal)

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub InsertDetail(ByVal Quantity As Integer, ByVal Value As String)
10    Dim lsOldProc As String

20    lsOldProc = pxeMODULENAME & ".InsertDetail"
30    ''On Error GoTo errProc

40    With MSFlexGrid1
50       If .Rows = 2 Then
60          If .TextMatrix(.Row, 1) <> "" Then
70             If oTrans.ItemCount <> .Row Then
80                oTrans.addDetail
90                oTrans.Detail(.Rows - 1, "xReferNox") = Value
100               If oTrans.Detail(.Rows - 1, "xReferNox") <> "" Then
110                  .Rows = .Rows + 1
120                  .Row = .Rows - 1
130                  .TextMatrix(.Row, 1) = Value
140                  .TextMatrix(.Row, 0) = .Row
150               Else
160                  oTrans.deleteDetail .Row
170                  Exit Sub
180               End If
190            Else
200               oTrans.addDetail
210               oTrans.Detail(.Row, "xReferNox") = Value
220               If oTrans.Detail(.Row, "xReferNox") <> "" Then
230                  .Rows = .Rows + 1
240                  .Row = .Rows - 1
250                  .TextMatrix(.Row, 1) = Value
260                  .TextMatrix(.Row, 0) = .Row
270               Else
280                  oTrans.deleteDetail .Row
290                  Exit Sub
300               End If
310            End If
320         Else
330            oTrans.Detail(.Row - 1, "xReferNox") = Value
340            If oTrans.Detail(.Row - 1, "xReferNox") <> "" Then .TextMatrix(.Row, 1) = Value
350            .TextMatrix(.Row, 0) = .Row
360         End If
370      Else
380         If oTrans.ItemCount <> .Row Then
390            oTrans.addDetail
400            oTrans.Detail(.Rows - 1, "xReferNox") = Value
410            If oTrans.Detail(.Rows - 1, "xReferNox") <> "" Then
420               .Rows = .Rows + 1
430               .Row = .Rows - 1
440               .TextMatrix(.Row, 1) = Value
450               .TextMatrix(.Row, 0) = .Row
460            Else
470               oTrans.deleteDetail .Rows
480               Exit Sub
490            End If
500         Else
510            oTrans.addDetail
520            oTrans.Detail(.Row, "xReferNox") = Value
530            If oTrans.Detail(.Row, "xReferNox") <> "" Then
540               .Rows = .Rows + 1
550               .Row = .Rows - 1
560               .TextMatrix(.Row, 1) = Value
570               .TextMatrix(.Row, 0) = .Row
580            Else
590               oTrans.deleteDetail .Row
600               Exit Sub
610            End If
620         End If
630      End If
640      Call refreshGrid

650      Call GrandTotal
660   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Function PrintTrans() As Boolean
10    Dim lsOldProc As String

20    lsOldProc = "PrintTrans"
30    ''On Error GoTo errProc

40    PrintTrans = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function


Private Sub refreshGrid()
10    Dim lnCtr As Integer
20    Dim lsOldProc As String

30    lsOldProc = pxeMODULENAME & ".refreshGrid"
40    ''On Error GoTo errProc

50    Call InitGrid

60    With MSFlexGrid1
70       .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
80       For lnCtr = 1 To .Rows - 1
90          .TextMatrix(lnCtr, 0) = lnCtr
100         .TextMatrix(lnCtr, 1) = oTrans.Detail(lnCtr - 1, "xReferNox")
110         .TextMatrix(lnCtr, 2) = oTrans.Detail(lnCtr - 1, "sDescript")
120         .TextMatrix(lnCtr, 3) = oTrans.Detail(lnCtr - 1, "nQuantity")
130         .TextMatrix(lnCtr, 4) = Format(oTrans.Detail(lnCtr - 1, "nUnitPrce"), "#,##0.00")
140         .TextMatrix(lnCtr, 5) = Format(CDbl(.TextMatrix(lnCtr, 3)) * CDbl(.TextMatrix(lnCtr, 4)), "#,##0.00")
150      Next

160      .Row = .Rows - 1
170      .ColSel = .Cols - 1

180      .ColWidth(2) = 4000
190      If .Rows > 21 Then .ColWidth(2) = 3750

200      pnRow = .Row - 1
210      Call GrandTotal
220   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Function isEntryOk() As Boolean
10    Dim lsOldProc As String
20    Dim lnCtr As Integer

30    lsOldProc = pxeMODULENAME & ".isEntryOK"
40    ''On Error GoTo errProc

50    For lnCtr = 0 To oTrans.ItemCount - 1
60       If oTrans.Detail(lnCtr, "nQuantity") = 0 Then
70          MsgBox "There is an item with no quantity." & vbCrLf & vbCrLf & _
                "Please verify your entry.", vbCritical, pxeAPPNAME

80          GoTo endProc
90       End If
100   Next

110   isEntryOk = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
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



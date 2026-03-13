VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Customer 
   BorderStyle     =   0  'None
   Caption         =   "MC Supplier Maintenance"
   ClientHeight    =   7395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7665
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5820
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   10266
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1485
         MaxLength       =   128
         TabIndex        =   20
         Top             =   3945
         Width           =   5880
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1485
         MaxLength       =   15
         TabIndex        =   18
         Top             =   3615
         Width           =   5880
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   5325
         MaxLength       =   25
         TabIndex        =   30
         Top             =   4275
         Width           =   2040
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   5325
         TabIndex        =   32
         Top             =   4605
         Width           =   2040
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
         Index           =   0
         Left            =   1485
         TabIndex        =   1
         Top             =   195
         Width           =   1950
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1485
         MaxLength       =   30
         TabIndex        =   16
         Top             =   3285
         Width           =   5880
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1485
         MaxLength       =   30
         TabIndex        =   22
         Top             =   4275
         Width           =   2295
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1485
         MaxLength       =   30
         TabIndex        =   24
         Top             =   4605
         Width           =   2295
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   5325
         TabIndex        =   34
         Top             =   4935
         Width           =   2040
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   5325
         TabIndex        =   36
         Top             =   5265
         Width           =   2040
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1485
         MaxLength       =   50
         TabIndex        =   3
         Top             =   765
         Width           =   5880
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   645
         Index           =   5
         Left            =   1485
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2295
         Width           =   5880
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   1485
         MaxLength       =   30
         TabIndex        =   26
         Top             =   4935
         Width           =   2295
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1485
         TabIndex        =   14
         Top             =   2955
         Width           =   5880
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   1485
         TabIndex        =   28
         Top             =   5265
         Width           =   2295
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   1185
         Left            =   1485
         Tag             =   "wt0;fb0"
         Top             =   1095
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   2090
         BackColor       =   12632256
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   6
            Top             =   90
            Width           =   4560
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   8
            Top             =   420
            Width           =   4560
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   4
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   10
            Top             =   750
            Width           =   4560
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   5
            Top             =   150
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Middle Name"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   9
            Top             =   810
            Width           =   930
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   19
         Top             =   4005
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   17
         Top             =   3675
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   3345
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term"
         Height          =   195
         Index           =   4
         Left            =   4005
         TabIndex        =   29
         Top             =   4335
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
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
         Left            =   90
         TabIndex        =   0
         Top             =   255
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. No."
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   21
         Top             =   4335
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax No."
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   23
         Top             =   4665
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Balance"
         Height          =   195
         Index           =   7
         Left            =   4005
         TabIndex        =   35
         Top             =   5310
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client Since"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   27
         Top             =   5325
         Width           =   840
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1575
         Tag             =   "et0;ht2"
         Top             =   300
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   2
         Top             =   825
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Address"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   11
         Top             =   2355
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Owner"
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   4
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "e-Mail Address"
         Height          =   195
         Index           =   16
         Left            =   90
         TabIndex        =   25
         Top             =   4995
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Town/City"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   13
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   195
         Index           =   18
         Left            =   4005
         TabIndex        =   31
         Top             =   4680
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit"
         Height          =   195
         Index           =   19
         Left            =   4005
         TabIndex        =   33
         Top             =   5010
         Width           =   840
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   6825
      TabIndex        =   44
      Top             =   6615
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Customer.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   6045
      TabIndex        =   43
      Top             =   6615
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Customer.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   3705
      TabIndex        =   38
      Top             =   6615
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Customer.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   2925
      TabIndex        =   37
      Top             =   6615
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Customer.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   6825
      TabIndex        =   45
      Top             =   6615
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Customer.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   7
      Left            =   4485
      TabIndex        =   40
      Top             =   6615
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCP_Customer.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   3705
      TabIndex        =   39
      Top             =   6615
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Update"
      AccessKey       =   "U"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_Customer.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   4485
      TabIndex        =   41
      Top             =   6615
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Delete"
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
      Picture         =   "frmCP_Customer.frx":3456
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   5265
      TabIndex        =   42
      Top             =   6615
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Ledger"
      AccessKey       =   "L"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_Customer.frx":3BD0
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmCP_Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmSP_Customer"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oForm As frmLedger
Private oSkin As clsFormSkin
Private bLoaded As Boolean
Private oRS As ADODB.Recordset

Dim psSupplier As String
Dim pnIndex As Integer
Dim pnCtr As Integer
Dim pbLoading As Boolean
Dim pbLoadOthers As Boolean

Dim pbtxtOthers As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   
   If pbtxtOthers Then
      txtOthers_LostFocus pnIndex
   Else
      txtField_LostFocus pnIndex
   End If
   
   Select Case Index
   Case 0
      pbLoading = True
      oDriver.RecordCancelUpdate
   Case 1
      oDriver.BrowseRecord
   Case 2
      oDriver.RecordSave
   Case 3
      pbLoading = False
      oDriver.RecordUpdate
   Case 4
      oDriver.RecordNew
   Case 5
      Unload Me
   Case 6
      oDriver.RecordDelete
   Case 7
      If pbtxtOthers Then
         If pnIndex = 5 Then SearchTerm Empty
      Else
         If Index = 1 Then
            SearchCompany Empty
         Else
            oDriver.RecordSearch
         End If
      End If
   Case 8
'      If Not pbNewSupplier Then
'         oForm.ClientID = oDriver.FieldValue(0)
'         oForm.Caption = "MC Supplier Ledger"
'         oForm.Show 1
'      Else
'         MsgBox "Unable to Load Supplier Ledger!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      End If
   End Select
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      oDriver.RecordNew
      oDriver.DisableTextbox 0
      bLoaded = True
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
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oForm = New frmLedger
   Set oRS = New ADODB.Recordset
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
   oDriver.RecQuery = "SELECT" _
                           & "  sClientID" _
                           & ", sCompnyNm" _
                           & ", sLastName" _
                           & ", sFrstName" _
                           & ", sMiddName" _
                           & ", sAddressx" _
                           & ", sTownIDxx" _
                           & ", sEmailAdd" _
                           & ", cRecdStat" _
                           & ", sModified" _
                           & ", dModified" _
                        & " FROM Client_Master"
                     
   oDriver.BrowseQuery = "SELECT" _
                           & "  a.sClientID" _
                           & ", a.sCompnyNm" _
                           & ", b.nABalance" _
                        & " FROM Client_Master a" _
                           & ", CP_Customer b" _
                        & " WHERE b.cRecdStat = " & strParm(xeRecStateActive) _
                           & " AND b.sClientID = a.sClientID" _
                           & " AND b.sBranchCd = " & strParm(oApp.BranchCode) _
                        & " ORDER BY b.sClientID"
   
   oDriver.InitRecForm
   
   oDriver.BrowseFTitle(0) = "Code"
   oDriver.BrowseFTitle(1) = "Company"
   oDriver.BrowseFTitle(2) = "Account Balance"
   oDriver.BrowseFFormat(0) = IIf(Len(oApp.BranchCode) = 2, "@@@@-@@@@@@", "@@@@@@-@@@@@@")
   oDriver.BrowseFFormat(2) = "#,##0.00"
   
   oDriver.LookupQuery(6) = "SELECT" _
                              & "  a.sTownIDxx" _
                              & ", CONCAT(a.sTownName, ', ', b.sProvName, ' ', a.sZippCode) as Town" _
                           & " FROM TownCity a" _
                              & " LEFT JOIN Province b" _
                                 & " ON a.sProvIDxx = b.sProvIDxx" _
                           & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY a.sTownName,b.sProvName"
   
   oDriver.LookupReference(6) = "a.sTownIDxx" _
                              & "»CONCAT(a.sTownName, ', ', b.sProvName " _
                              & ", ' ', a.sZippCode)"
   oDriver.LookupColumn(6) = "Town"
   oDriver.LookupTitle(6) = "Town"
   
   oDriver.FieldFormat(0) = IIf(Len(oApp.BranchCode) = 2, "@@@@-@@@@@@", "@@@@@@-@@@@@@")
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
   
   oDriver.FieldStart = 1
   
   psSupplier = "SELECT" _
                     & "  a.sCPerson1" _
                     & ", a.sCPPosit1" _
                     & ", a.sTelNoxxx" _
                     & ", a.sFaxNoxxx" _
                     & ", a.sRemarksx" _
                     & ", a.sTermIDxx" _
                     & ", a.nDiscount" _
                     & ", a.nCredLimt" _
                     & ", a.nABalance" _
                     & ", a.dCltSince" _
                     & ", a.nLedgerNo" _
                     & ", a.cRecdStat" _
                     & ", a.sClientID" _
                     & ", a.sBranchCd" _
                     & ", b.sTermName" _
                  & " FROM CP_Customer a" _
                     & " LEFT JOIN Term b" _
                        & " ON a.sTermIDxx = b.sTermIDxx" _
                  & " ORDER BY a.sBranchCd"
   
   Set oRS = New ADODB.Recordset
   lsSQL = AddCondition(psSupplier, "0 = 1")
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockPessimistic, adCmdText
   Set oRS.ActiveConnection = Nothing
   
   For pnCtr = 0 To txtOthers.Count - 1
      If pnCtr < 5 Then txtOthers(pnCtr).MaxLength = oRS(pnCtr).DefinedSize
   Next
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oRS = Nothing
   Set oDriver = Nothing
   Set oSkin = Nothing
   Set oForm = Nothing
End Sub

Private Sub oDriver_Delete(Deleted As Boolean)
   Deleted = True
End Sub

Private Sub oDriver_DeleteComplete()
   For pnCtr = 0 To txtOthers.Count - 1
      txtOthers(pnCtr).Text = ""
   Next
End Sub

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_InitValue"
   ''On Error GoTo errProc

   If oDriver.SetValue(0, GetNextCode("Client_Master", "sClientID", True, oApp.Connection, True, oApp.BranchCode)) = False Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(8) = 1
   
   oRS.AddNew
   InitOthers
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "oDriver_LoadOtherData"
   ''On Error GoTo errProc
   
   Set oRS = New ADODB.Recordset
   lsSQL = AddCondition(psSupplier, "a.sClientID = " & strParm(oDriver.FieldValue(0)) _
                                    & " AND a.sBranchCd = " & strParm(oApp.BranchCode)) _
                                    & " AND a.cRecdStat = " & strParm(xeRecStateActive)
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set oRS.ActiveConnection = Nothing
         
   If oRS.EOF Then
      oRS.AddNew
      InitOthers
      GoTo endProc
   End If
         
   For pnCtr = 0 To txtOthers.Count - 1
      Select Case pnCtr
      Case 5
         txtOthers(pnCtr).Text = IIf(IsNull(oRS("sTermName")), "", oRS("sTermName"))
         txtOthers(pnCtr).Tag = txtOthers(pnCtr).Text
      Case 6
         txtOthers(pnCtr).Text = Format(IIf(IsNull(oRS(pnCtr)), 0, oRS(pnCtr)), "##0.00 %")
      Case 7, 8
         txtOthers(pnCtr).Text = Format(IIf(IsNull(oRS(pnCtr)), 0, oRS(pnCtr)), "##0.00 php.")
      Case 9
         txtOthers(pnCtr).Text = Format(IIf(IsNull(oRS(pnCtr)), "1999-01-01", oRS(pnCtr)), "MMMM DD, YYYY")
      Case Else
         txtOthers(pnCtr).Text = IIf(IsNull(oRS(pnCtr)), "", oRS(pnCtr))
      End Select
   Next
   
   txtOthers(8).Enabled = True
   If CDbl(Left(txtOthers(8).Text, Len(txtOthers(8).Text) - 5)) > 0 Then txtOthers(8).Enabled = False
   pbLoadOthers = True
   pbLoading = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_Save(Saved As Boolean)
   Saved = False
End Sub

Private Sub oDriver_WillDelete(Cancel As Boolean)
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lsOldProc As String
   Dim lrs As ADODB.Recordset
   
   lsOldProc = "oDriver_WillDelete"
   ''On Error GoTo errProc
   
   lsSQL = "DELETE FROM SP_Customer " _
            & "WHERE sClientID =" & strParm(oDriver.FieldValue(0))
   lnRow = oApp.Execute(lsSQL, "SP_Customer", oApp.BranchCode)
   If lnRow <= 0 Then
      MsgBox "Unable to Delete MC Supplier!!!", vbCritical, "Warning"
      Cancel = True
      GoTo endProc
   End If
   
   Set lrs = New ADODB.Recordset
   lrs.Open "SELECT * From Client_Ledger " _
            & "WHERE sClientID = " & strParm(oDriver.FieldValue(0)) & "" _
               & " AND sBranchCd = " & strParm(oApp.BranchCode) _
            , oApp.Connection, adOpenDynamic, adLockReadOnly, adCmdText
            
   If Not lrs.EOF Then
      lsSQL = "DELETE FROM Client_Ledger " _
               & "WHERE sClientID = " & strParm(oDriver.FieldValue(0)) _
               & " AND sBranchCd = " & strParm(oApp.BranchCode)
      lnRow = oApp.Execute(lsSQL, "Client_Ledger", oApp.BranchCode)
      If lnRow <= 0 Then
         MsgBox "Unable to Client Ledger!!!", vbCritical, "Warning"
         Cancel = True
         GoTo endProc
      End If
   End If

endProc:
   Set lrs = Nothing
   Exit Sub
errProc:
   Cancel = True
   ShowError lsOldProc & "( " & Cancel & " )"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_WillSave"
   ''On Error GoTo errProc
   
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Company Name detected!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   ElseIf CDbl(Left(txtOthers(7).Text, Len(txtOthers(7)) - 5)) > 99999999.99 Then
      MsgBox "Invalid Credit Limit!!!", vbCritical, "Warning"
      txtOthers(7).SetFocus
      Cancel = True
   ElseIf CDbl(Left(txtOthers(8).Text, Len(txtOthers(8)) - 5)) > 99999999.99 Then
      MsgBox "Invalid Account Balance!!!", vbCritical, "Warning"
      txtOthers(8).SetFocus
      Cancel = True
   Else
      Cancel = Not UpdateMCSupplier
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )"
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   If pbLoading Then Exit Sub
   
   oDriver.ColumnIndex = Index
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
      If .MultiLine Then .SelStart = Len(.Text)
   End With
   
   pbtxtOthers = False
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(pnIndex)
         Select Case Index
         Case 1
            If oDriver.EditMode = xeModeAddNew Then
               If Trim(.Text) <> "" Then
                  If .Tag <> .Text Then SearchCompany .Text
                  If .Text <> "" Then SetNextFocus
               End If
            End If
         Case 6
            If KeyCode = vbKeyF3 Then
               oDriver.RecordSearch .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oDriver.RecordSearch .Text
            End If
         End Select
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

Private Sub oDriver_DisableOtherControl()
   For pnCtr = 0 To txtOthers.Count - 1
      txtOthers(pnCtr).Enabled = False
   Next
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   
   For pnCtr = 0 To txtOthers.Count - 1
      txtOthers(pnCtr).Enabled = True
   Next
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   ''On Error GoTo errProc
   
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 1
         If oDriver.EditMode = xeModeAddNew Then
            If Trim(.Text) <> "" Then
               If .Tag <> .Text Then SearchCompany .Text
            End If
         End If
      End Select
      Cancel = Not oDriver.ValidateField(Index)
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub txtOthers_GotFocus(Index As Integer)
   If pbLoading Then Exit Sub
   
   With txtOthers(Index)
      Select Case Index
      Case 6
         .Text = Left(.Text, Len(.Text) - 2)
      Case 7, 8
         .Text = Left(.Text, Len(.Text) - 5)
      Case 9
         .Text = Format(.Text, "MM/DD/YY")
      End Select
   
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   
   pbtxtOthers = True
   pnIndex = Index
End Sub

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtOthers_KeyDown"
   ''On Error GoTo errProc
   
   If pnIndex = 5 Then
      If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
         With txtOthers(pnIndex)
            If KeyCode = vbKeyF3 Then
               SearchTerm .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then SearchTerm .Text
            End If
         End With
         KeyCode = 0
      End If
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

Private Sub SearchTerm(ByVal lsValue As String)
   Dim lsSearch As String
   Dim lrs As ADODB.Recordset
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim lsSelected() As String
   
   lsOldProc = "SearchTerm"
   ''On Error GoTo errProc
   
   Set lrs = New ADODB.Recordset
  
   With txtOthers(5)
      lsSQL = "SELECT" _
                  & "  sTermIDxx" _
                  & ", sTermName" _
                  & ", nTermDays" _
                  & ", nDiscDays" _
                  & ", nDiscount" _
               & " FROM Term" _
               & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                  & " AND sTermName LIKE " & strParm(lsValue & "%") _
               & " ORDER BY sTermName"
   
      If .Text = .Tag Then Exit Sub
      lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
      If lrs.RecordCount = 1 Then
         oRS("sTermIDxx") = lrs("sTermIDxx")
         .Text = lrs("sTermName")
      ElseIf lrs.RecordCount > 1 Then
         lsSearch = KwikBrowse(oApp, lrs _
                        , "sTermName" _
                        & "»nTermDays" _
                        & "»nDiscDays" _
                        & "»nDiscount" _
                        , "Term»" _
                        & "Term Days" _
                        & "»Discount Days" _
                        & "»Discount" _
                        , "»0.00 day/s" _
                        & "»0.00 day/s" _
                        & "»0.00")
      
         If lsSearch <> "" Then
            lsSelected = Split(lsSearch, "»")
            oRS("sTermIDxx") = lsSelected(0)
            .Text = lsSelected(1)
         End If
      Else
         oRS("sTermIDxx") = ""
         .Text = ""
      End If
      .Tag = .Text
   End With
   
endProc:
   Set lrs = Nothing
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & lsValue & " )"
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub SearchCompany(Optional lsCompany As Variant)
   Dim lrs As ADODB.Recordset
   Dim lsBrowse As String
   Dim lsSQL As String
   Dim lsOldProc As String
   Dim lrsSupplier As ADODB.Recordset
   Dim lsSelected() As String
   
   lsOldProc = "SearchCompany"
   ''On Error GoTo errProc
   
   lsSQL = "SELECT" _
               & "  a.sClientID" _
               & ", a.sCompnyNm" _
               & ", a.sLastName" _
               & ", a.sFrstName" _
               & ", a.sMiddName" _
               & ", a.sAddressx" _
               & ", b.sTownName" _
               & ", a.sEmailAdd" _
            & " FROM Client_Master a" _
               & " LEFT JOIN TownCity b" _
                  & " ON a.sTownIDxx = b.sTownIDxx" _
            & IIf(IsMissing(lsCompany), "", _
            " WHERE a.sCompnyNm LIKE " & strParm(lsCompany & "%")) _
            & " ORDER BY" _
               & " a.sCompnyNm"

   Set lrs = New ADODB.Recordset
  
   If lrs.State = adStateOpen Then lrs.Close
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If lrs.EOF Then GoTo endProc
         
   lsBrowse = KwikBrowse(oApp, lrs _
               , "sCompnyNm»sLastName»sFrstName" _
               & "»sMiddName»sAddressx»sTownName" _
               , "Company»LastName»FirstName»MiddleName»Address»Town")

   If lsBrowse <> "" Then
      lsSelected = Split(lsBrowse, "»")
      oDriver.LookupValue(0) = lsSelected(0)
      oDriver.LoadRecord
      
      Set lrsSupplier = New ADODB.Recordset
      lsSQL = AddCondition(psSupplier, "a.sClientID = " & strParm(lrs("sClientID")))
      lrsSupplier.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If lrsSupplier.EOF Then
         MsgBox "Supplier not yet added to your branch!!!" & vbCrLf & _
                  "Please Save the record to activate!!!", vbInformation, "Notice"
         oDriver.RecordUpdate
         InitOthers
         GoTo endProc
      End If
      
      If pbLoadOthers Then Exit Sub
      lrsSupplier.Find "sBranchCd = " & strParm(oApp.BranchCode), 0, adSearchForward
      If lrsSupplier.EOF Then
         lrsSupplier.MoveFirst
         MsgBox "Supplier not yet added to your branch!!!" & vbCrLf & _
               "Please Save the record to activate!!!", vbInformation, "Notice"
         GoTo loadSupplier
      Else
         If lrsSupplier("cRecdStat") = xeRecStateInactive Then
            MsgBox "MC Supplier Status is Deactivated!!!" & vbCrLf & _
                     "Please Save the record to activate!!!", vbInformation, "Notice"
            oDriver.RecordUpdate
            lrs("cRecdStat") = xeRecStateActive
         End If
         GoTo loadSupplier
      End If
   End If
   
loadSupplier:
   InitOthers
   For pnCtr = 0 To txtOthers.Count - 1
      If pnCtr < 5 Then
         lrs(pnCtr) = IIf(IsNull(lrsSupplier(pnCtr)), "", lrsSupplier(pnCtr))
         txtOthers(pnCtr).Text = oRS(pnCtr)
      End If
   Next
   oRS("sClientID") = lrsSupplier("sClientID")
   oRS("sBranchCd") = lrsSupplier("sBranchCd")
   oRS("sTermIDxx") = IIf(IsNull(lrsSupplier("sTermIDxx")), "", lrsSupplier("sTermIDxx"))
   GoTo endProc
endProc:
   Set lrs = Nothing
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & IFNull(lsCompany) & " )"
End Sub

Private Sub InitOthers()
   For pnCtr = 0 To txtOthers.Count - 1
      Select Case pnCtr
      Case 6
         oRS(pnCtr) = 0
         txtOthers(pnCtr).Text = Format(oRS(pnCtr), "##0.00 %")
      Case 7, 8
         oRS(pnCtr) = 0#
         txtOthers(pnCtr).Text = Format(oRS(pnCtr), "##0.00 php.")
      Case 9
         oRS(pnCtr) = oApp.ServerDate
         txtOthers(pnCtr).Text = Format(oRS(pnCtr), "MMMM DD, YYYY")
      Case Else
         oRS(pnCtr) = Empty
         txtOthers(pnCtr).Text = oRS(pnCtr)
      End Select
   Next

   oRS("nLedgerNo") = 0
   oRS("cRecdStat") = xeRecStateActive
   oRS("sClientID") = oDriver.FieldValue(0)
   oRS("sBranchCd") = oApp.BranchCode
   oRS("sTermIDxx") = ""
   
   txtOthers(8).Enabled = True
   pbLoadOthers = False
End Sub

Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   ''On Error GoTo errProc
   
   With txtOthers(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 5
         If Trim(.Text) <> "" Then
            If .Tag <> .Text Then SearchTerm .Text
         End If
         .Tag = .Text
      Case 6
         If Not IsNumeric(.Text) Then .Text = 0
         oRS(Index) = .Text
         .Text = Format(.Text, "##0.00 %")
      Case 7, 8
         If Not IsNumeric(.Text) Then .Text = 0
         oRS(Index) = .Text
         .Text = Format(.Text, "##0.00 php.")
      Case 9
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         oRS(Index) = .Text
         .Text = Format(.Text, "MMMM DD, YYYY")
      Case Else
         oRS(Index) = .Text
      End Select
      
      .Tag = .Text
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Function UpdateMCSupplier() As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lrs As ADODB.Recordset

   lsOldProc = "UpdateMCSupplier"
'   ''On Error GoTo errProc

   Set lrs = New ADODB.Recordset
   lrs.Open "SELECT *" _
               & " FROM CP_Customer" _
               & " WHERE sClientID = " & strParm(oDriver.FieldValue(0)) _
                  & " AND sBranchCd = " & strParm(oApp.BranchCode) _
   , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.EOF Then
      UpdateMCSupplier = SaveOthers(oRS, "CP_Customer", xeModeAddNew)
   Else
      UpdateMCSupplier = SaveOthers(oRS, "CP_Customer", xeModeUpdate, "sClientID»sBranchCd")
   End If

endProc:
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & UpdateMCSupplier & " )", True
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

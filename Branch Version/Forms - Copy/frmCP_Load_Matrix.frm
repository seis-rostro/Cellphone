VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Load_Matrix 
   BorderStyle     =   0  'None
   Caption         =   "CP Load Matrix Maintenance"
   ClientHeight    =   6525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   4095
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1380
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   7223
      BorderStyle     =   1
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   7
         Left            =   5145
         TabIndex        =   29
         Top             =   3510
         Width           =   2115
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1485
         TabIndex        =   23
         Top             =   3510
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1095
         TabIndex        =   5
         Top             =   240
         Width           =   2820
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   5145
         TabIndex        =   27
         Top             =   3210
         Width           =   2115
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1485
         TabIndex        =   21
         Top             =   3210
         Width           =   1995
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   5145
         TabIndex        =   25
         Top             =   2910
         Width           =   2115
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1485
         TabIndex        =   19
         Top             =   2910
         Width           =   1995
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   5145
         TabIndex        =   17
         Text            =   "0,000.00"
         Top             =   2325
         Width           =   2115
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   0
         Left            =   1485
         TabIndex        =   15
         Top             =   2325
         Width           =   1995
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1095
         TabIndex        =   13
         Top             =   1710
         Width           =   6180
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1095
         TabIndex        =   11
         Top             =   1410
         Width           =   6180
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1095
         TabIndex        =   9
         Top             =   1110
         Width           =   6180
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1095
         TabIndex        =   7
         Top             =   810
         Width           =   6180
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   195
         X2              =   7230
         Y1              =   2145
         Y2              =   2145
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reserve Order"
         Height          =   195
         Index           =   16
         Left            =   3945
         TabIndex        =   28
         Top             =   3570
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Level"
         Height          =   195
         Index           =   25
         Left            =   255
         TabIndex        =   20
         Top             =   3255
         Width           =   780
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back Order"
         Height          =   195
         Index           =   14
         Left            =   4170
         TabIndex        =   26
         Top             =   3255
         Width           =   810
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beg. Inv. Date"
         Height          =   195
         Index           =   13
         Left            =   255
         TabIndex        =   22
         Top             =   3570
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min. Level"
         Height          =   195
         Index           =   24
         Left            =   255
         TabIndex        =   18
         Top             =   2940
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ReOrder Level"
         Height          =   210
         Index           =   23
         Left            =   3930
         TabIndex        =   24
         Top             =   2940
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amt On Hand"
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
         Index           =   21
         Left            =   3630
         TabIndex        =   16
         Top             =   2415
         Width           =   1365
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beg. Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   255
         TabIndex        =   14
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Shape Shape2 
         Height          =   3795
         Left            =   105
         Top             =   120
         Width           =   7275
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   195
         Index           =   8
         Left            =   195
         TabIndex        =   12
         Top             =   1725
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   10
         Top             =   1410
         Width           =   435
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   315
         Index           =   4
         Left            =   195
         TabIndex        =   8
         Top             =   1125
         Width           =   420
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1200
         Tag             =   "et0;ht2"
         Top             =   360
         Width           =   2820
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barr Code"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   270
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   825
         Width           =   795
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   6840
      TabIndex        =   36
      Top             =   5730
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
      Picture         =   "frmCP_Load_Matrix.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   6060
      TabIndex        =   35
      Top             =   5730
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
      Picture         =   "frmCP_Load_Matrix.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   3720
      TabIndex        =   30
      Top             =   5730
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
      Picture         =   "frmCP_Load_Matrix.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   3720
      TabIndex        =   32
      Top             =   5730
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
      Picture         =   "frmCP_Load_Matrix.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   6840
      TabIndex        =   37
      Top             =   5730
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
      Picture         =   "frmCP_Load_Matrix.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   4500
      TabIndex        =   31
      Top             =   5730
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
      Picture         =   "frmCP_Load_Matrix.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   4500
      TabIndex        =   33
      Top             =   5730
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
      Picture         =   "frmCP_Load_Matrix.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   3720
      TabIndex        =   39
      Top             =   5730
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
      Picture         =   "frmCP_Load_Matrix.frx":3456
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   5280
      TabIndex        =   34
      Top             =   5730
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
      Picture         =   "frmCP_Load_Matrix.frx":3BD0
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   9
      Left            =   3720
      TabIndex        =   38
      Top             =   5730
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Replace"
      AccessKey       =   "R"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_Load_Matrix.frx":434A
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   810
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   1429
      Begin VB.TextBox txtOthers 
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
         Height          =   285
         Index           =   9
         Left            =   1095
         TabIndex        =   3
         Top             =   390
         Width           =   6285
      End
      Begin VB.TextBox txtOthers 
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
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   1095
         MaxLength       =   50
         TabIndex        =   1
         Top             =   90
         Width           =   6285
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Description"
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
         Left            =   75
         TabIndex        =   2
         Top             =   435
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Barr C&ode"
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
         Left            =   75
         TabIndex        =   0
         Top             =   135
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmCP_Load_Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Load_Matrix"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset

Dim pbtxtOthers As Boolean
Dim pnCtr As Integer, pnIndex As Integer

Dim psCPInventory As String
Dim pbEnblButtons As Boolean
Dim pbNewInvntory As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsSearch As String
   Dim lnRep As Integer
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc
   
   Select Case Index
   Case 0 'cancel
      oDriver.RecordCancelUpdate
      pbEnblButtons = False
   Case 1 'browse
      oDriver.BrowseRecord
   Case 2 'save
      oDriver.RecordSave
   Case 3 'update
      oDriver.RecordUpdate
   Case 4 'new
      oDriver.RecordNew
   Case 5 'close
      Unload Me
   Case 6 'delete
      oDriver.RecordDelete
   Case 7 'search
      If pbtxtOthers Then
         oDriver.RecordSearch
         txtField(pnIndex).SetFocus
      End If
   Case 8 'ledger
      If Not pbNewInvntory Then
        With frmCP_MatrixLedger
            .txtDateFrom = Format(DateAdd("m", -1, oApp.ServerDate), "MMMM DD, YYYY")
            .txtDateThru = Format(oApp.ServerDate, "MMMM DD, YYYY")
            .txtField(0) = txtField(0)
            .txtField(1) = txtField(1)
            .txtField(2) = txtField(2)
            .txtField(3) = txtField(3)
            
            .StockID = oDriver.FieldValue(5)
            .Show 1
         End With
      Else
         MsgBox "Unable to Load Inventory Ledger!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   'On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      oDriver.RecordCancelUpdate
      oDriver_InitValue
      bLoaded = True
      txtOthers(8).SetFocus
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me
   
   bLoaded = False
   Set oRS = New ADODB.Recordset
   
   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
   oDriver.RecQuery = "SELECT" _
                           & "  sBarrcode" _
                           & ", sDescript" _
                           & ", sBrandIDx" _
                           & ", sModelIDx" _
                           & ", sCategID1" _
                           & ", sStockIDx" _
                           & ", cRecdStat" _
                           & ", sModified" _
                           & ", dModified" _
                        & " FROM CP_Load_Matrix"

   oDriver.BrowseQuery = "SELECT" _
                              & "  a.sBarrcode" _
                              & ", a.sDescript" _
                              & ", b.sBrandNme" _
                              & ", c.sModelNme" _
                              & ", d.sCategrNm" _
                              & ", e.nAmtonHnd" _
                           & " FROM CP_Load_Matrix a" _
                              & " LEFT JOIN CP_Brand b" _
                                 & " ON a.sBrandIDx = b.sBrandIDx" _
                              & " LEFT JOIN CP_Model c" _
                                 & " ON a.sModelIDx = c.sModelIDx" _
                              & " LEFT JOIN Category d" _
                                 & " ON a.sCategID1 = d.sCategrID" _
                              & ", CP_Load_Matrix_Master e" _
                           & " WHERE a.sStockIDx = e.sStockIDx" _
                              & " AND e.sBranchCd = " & strParm(oApp.BranchCode)
   oDriver.InitRecForm
      
   oDriver.BrowseColumn(0) = "sBarrcode"
   oDriver.BrowseColumn(1) = "sDescript"
   oDriver.BrowseColumn(2) = "sBrandNme"
   oDriver.BrowseColumn(3) = "sModelNme"
   oDriver.BrowseColumn(4) = "nAmtonHnd"
   oDriver.BrowseColumn(5) = "sCategrNm"
   
   oDriver.BrowseFTitle(0) = "Barcode"
   oDriver.BrowseFTitle(1) = "Description"
   oDriver.BrowseFTitle(2) = "Brand"
   oDriver.BrowseFTitle(3) = "Model"
   oDriver.BrowseFTitle(4) = "AMT"
   oDriver.BrowseFTitle(5) = "Category"
   
   oDriver.BrowseFFormat(4) = "#,##0.00"
   
   oDriver.LookupQuery(2) = "SELECT" _
                              & "  sBrandIDx" _
                              & ", sBrandNme" _
                           & " FROM CP_Brand" _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY sBrandNme"
   oDriver.LookupReference(2) = "sBrandIDx»sBrandNme"
   oDriver.LookupColumn(2) = "sBrandNme"
   oDriver.LookupTitle(2) = "Brand Name"

   oDriver.LookupQuery(3) = "SELECT" _
                              & " sModelIDx" _
                              & ",sModelNme " _
                           & "FROM CP_Model " _
                           & "WHERE cRecdStat = 1 " _
                           & "ORDER BY sModelNme"
   oDriver.LookupReference(3) = "sModelIDx»sModelNme"
   oDriver.LookupColumn(3) = "sModelNme"
   oDriver.LookupTitle(3) = "Model Name"

   oDriver.LookupQuery(4) = "SELECT" _
                              & "  sCategrID" _
                              & ", sCategrNm" _
                           & " FROM Category" _
                           & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY sCategrNm"
   oDriver.LookupReference(4) = "sCategrID»sCategrNm"
   oDriver.LookupColumn(4) = "sCategrNm"
   oDriver.LookupTitle(4) = "Category Name"
                           
   oDriver.FieldStart = 0
   oDriver.FieldFormat(0) = ">"
   
   psCPInventory = "SELECT" _
                     & "  nBegAmtxx" _
                     & ", nAmtOnHnd" _
                     & ", nMinLevel" _
                     & ", nMaxLevel" _
                     & ", dBegInvxx" _
                     & ", nReorderx" _
                     & ", nBackOrdr" _
                     & ", nResvOrdr" _
                     & ", nFloatAmt" _
                     & ", nLedgerNo" _
                     & ", dLastTran" _
                     & ", cRecdStat" _
                     & ", sStockIDx" _
                     & ", sBranchCd" _
                     & ", sModified" _
                     & ", dModified" _
                  & " FROM CP_Load_Matrix_Master" _
                  & " ORDER BY sBranchCd"
   
   Set oRS = New ADODB.Recordset
   lsSQL = AddCondition(psCPInventory, "0 = 1")
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockPessimistic, adCmdText
   Set oRS.ActiveConnection = Nothing
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
   Set oRS = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   For pnCtr = 0 To txtOthers.Count - 1
      txtOthers(pnCtr).Enabled = False
   Next
   
   txtOthers(8).Enabled = True
   txtOthers(9).Enabled = True
   
   oDriver.hideButton 6
   oDriver.hideButton 9
End Sub

Private Sub oDriver_EnableOtherControl()
   For pnCtr = 0 To txtOthers.Count - 1
      Select Case pnCtr
      Case 8, 9
         txtOthers(pnCtr).Enabled = False
      Case Else
         txtOthers(pnCtr).Enabled = True
      End Select
   Next

   If oDriver.EditMode = xeModeUpdate Then
      txtOthers(0).Enabled = pbEnblButtons
      txtOthers(1).Enabled = pbEnblButtons
      txtOthers(4).Enabled = pbEnblButtons
   End If
End Sub

Private Sub InitOthers()
   For pnCtr = 0 To txtOthers.Count - 3
      Select Case pnCtr
      Case 0 To 3, 5, 6, 7
         oRS(pnCtr) = 0
         txtOthers(pnCtr).Text = Format(oRS(pnCtr), "#,##0.00")
      Case 4
         oRS(pnCtr) = oApp.ServerDate
         txtOthers(pnCtr).Text = Format(oRS(pnCtr), "MMM DD, YYYY")
      End Select
   Next

   oRS("cRecdStat") = xeRecStateActive
   oRS("sStockIDx") = oDriver.FieldValue(5)
   oRS("sBranchCd") = oApp.BranchCode
   oRS("nFloatAmt") = 0
   oRS("nLedgerNo") = 0
End Sub

Private Sub oDriver_InitValue()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_InitValue"
   'On Error GoTo errProc

   oDriver.FieldReference(0) = True
   oDriver.FieldValue(0) = NewBarrCode
   txtField(0).Text = oDriver.FieldValue(0)
   oDriver.FieldValue(5) = GetNextCode("CP_Load_Matrix", "sStockIDx", True, oApp.Connection, True, oApp.BranchCode)
   oDriver.FieldValue(6) = xeRecStateActive
   
   For pnCtr = 0 To txtOthers.Count - 1
      Select Case pnCtr
      Case 0 To 3, 5, 6, 7
         txtOthers(pnCtr).Text = 0
      Case 4
         txtOthers(pnCtr).Text = Format(oApp.ServerDate, "MMM DD, YYYY")
      End Select
      txtOthers(pnCtr).Tag = ""
   Next
   
   txtOthers(0).Locked = False
   txtOthers(1).Locked = False
   
   oRS.AddNew
   InitOthers
   pbEnblButtons = True
   pbNewInvntory = True
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lsOldProc As String
   Dim lsSQL As String
   
   lsOldProc = "oDriver_LoadOtherData"
   'On Error GoTo errProc
   
   Set oRS = New ADODB.Recordset
   lsSQL = AddCondition(psCPInventory, "sStockIDx = " & strParm(oDriver.FieldValue(5)) _
                                    & " AND sBranchCd = " & strParm(oApp.BranchCode)) _
                                    & " AND cRecdStat = " & strParm(xeRecStateActive)
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set oRS.ActiveConnection = Nothing
   
   If oRS.EOF Then
      oRS.AddNew
      InitOthers
   Else
      For pnCtr = 0 To txtOthers.Count - 1
         Select Case pnCtr
         Case 0 To 3, 5, 6, 7
            txtOthers(pnCtr).Text = Format(IIf(IsNull(oRS(pnCtr)), 0, oRS(pnCtr)), "#,##0.00")
         Case 4
            txtOthers(pnCtr).Text = IIf(IsNull(oRS(pnCtr)), "", Format(oRS(pnCtr), "MMM DD, YYYY"))
         End Select
      Next
      pbNewInvntory = False
   End If

   txtOthers(8).Text = oDriver.FieldValue(0)
   txtOthers(8).Tag = txtOthers(8).Text

   txtOthers(9).Text = oDriver.FieldValue(1)
   txtOthers(9).Tag = txtOthers(9).Text
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_Save(Saved As Boolean)
   Saved = False
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   Dim lsOldProc As String
   Dim lnCtr As Integer

   lsOldProc = "oDriver_WillSave"
   'On Error GoTo errProc

   If oDriver.FieldValue(0) = "" Then
      MsgBox "Invalid BarrCode detected!!!", vbCritical, "Warning"
      txtField(0).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Description detected!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(5) = "" Then
      MsgBox "Invalid Stock ID Detected!!!" & vbCrLf & _
               "Please contact GMC_SEG for assistant!!!", vbCritical, "Warning"
      Cancel = True
   Else
      Cancel = Not UpdateCPInventory
      If pbNewInvntory Then Cancel = Not SaveCPInventoryLedger
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )"
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   
   oDriver.ColumnIndex = Index
   pbtxtOthers = False
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   'On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         If KeyCode = vbKeyF3 Then
            oDriver.RecordSearch .Text
            If .Text <> "" Then SetNextFocus
         Else
            If .Text <> "" Then oDriver.RecordSearch .Text
         End If
      End With
      KeyCode = 0
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift & " )", True
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   'On Error GoTo errProc
   
   With txtField(Index)
      Select Case Index
      Case 0
         .Text = UCase(.Text)
      Case Else
         .Text = TitleCase(.Text)
      End Select
      Cancel = Not oDriver.ValidateField(Index)
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index _
                       & ", " & Cancel & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtOthers_GotFocus(Index As Integer)
   With txtOthers(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   
   pbtxtOthers = True
   pnIndex = Index
End Sub

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsSearch() As String
   Dim lnCtr As Integer
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "txtOthers_KeyDown"
   'On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtOthers(Index)
         Select Case Index
         Case 8, 9
            Call txtOthers_Validate(Index, False)
         End Select
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift & " )", True
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("EB")
   End With
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

Private Sub oDriver_Delete(Deleted As Boolean)
   Deleted = True
End Sub

Private Sub oDriver_DeleteComplete()
   For pnCtr = 0 To txtOthers.Count - 1
      txtOthers(pnCtr).Text = ""
   Next
End Sub

Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   Dim lnCtr As Integer
   
   lsOldProc = "txtOthers_Validate"
   'On Error GoTo errProc
   
   With txtOthers(Index)
      .Text = TitleCase(.Text)
      
      Select Case Index
      Case 0 To 3, 5, 6, 7
         If Not IsNumeric(.Text) Then .Text = 0
         .Text = Format(.Text, "#,##0.00")
         oRS(Index) = CDbl(.Text)
      Case 8, 9
         If Trim(.Text) = "" Then
            If oRS.EOF Then Exit Sub
            InitOthers
            txtOthers(IIf(Index = 8, 9, 8)).Text = ""
            txtOthers(IIf(Index = 8, 9, 8)).Tag = ""
            
            For lnCtr = 0 To txtField.Count - 1
               txtField(lnCtr).Text = ""
               txtField(lnCtr).Tag = txtField(lnCtr).Text
            Next
            Exit Sub
         End If
         
         If Trim(LCase(.Tag)) <> Trim(LCase(.Text)) Then SearchBarCode .Text, IIf(Index = 8, True, False)
         .Tag = .Text
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )", True
End Sub

Function NewBarrCode() As String
   Dim lrs As Recordset
   Dim lsSQL As String
   Dim lnCtr As Long
   Dim lsOldProc As String
   Dim lrsBranch As Recordset
   Dim lsCode As String
   
   lsOldProc = "NewBarrCode"
   'On Error GoTo errProc
   
   Set lrsBranch = New ADODB.Recordset
   lrsBranch.Open "SELECT" _
                     & "  a.sCompnyCd" _
                  & " FROM Company a" _
                     & ", Branch b" _
                  & " WHERE a.sCompnyID = b.sCompnyID" _
                     & " AND b.sBranchCd = " & strParm(oApp.BranchCode) _
                  , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
                   
   lsCode = "GTC"
   If Not lrsBranch.EOF Then lsCode = lrsBranch("sCompnyCd")

   lsSQL = "SELECT" & _
               " sBarrCode" & _
            " FROM CP_Load_Matrix" & _
            " WHERE sBarrCode LIKE " & strParm(Format(Date, "yy") & "-" & lsCode & "-%") & _
            " ORDER BY sBarrCode DESC" & _
            " LIMIT 1"
   Set lrs = New Recordset
   lrs.Open lsSQL, oApp.Connection, , , adCmdText
   
   If lrs.EOF Then
      lnCtr = 1
   Else
      If Left(lrs("sBarrCode"), 2) = Format(Date, "yy") Then
         lnCtr = CLng(Right(lrs("sBarrCode"), 6)) + 1
      Else
         lnCtr = 1
      End If
   End If
   NewBarrCode = Format(Date, "yy") & "-" & lsCode & "-" & Format(lnCtr, "000000")
   
   Set lrs = Nothing
   Set lrsBranch = Nothing

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub SearchBarCode(ByVal lsValue As String, ByVal lbByCode As Boolean)
   Dim lrsCellphone As ADODB.Recordset
   Dim lrsInventoryx As ADODB.Recordset
   Dim lsSelected() As String
   Dim lsBrowse As String
   Dim lsOldProc As String
   Dim lsStockIDx As String
   Dim lsSQL As String
   
   lsOldProc = "SearchSpareParts"
   'On Error GoTo errProc
      
   lsSQL = "SELECT" _
               & "  a.sStockIDx" _
               & ", a.sBarrcode" _
               & ", a.sDescript" _
               & ", b.sBrandNme" _
               & ", c.sModelNme" _
               & ", d.sCategrNm" _
            & " FROM CP_Load_Matrix a" _
               & " LEFT JOIN CP_Brand b" _
                  & " ON a.sBrandIDx = b.sBrandIDx" _
               & " LEFT JOIN CP_Model c" _
                  & " ON a.sModelIDx = c.sModelIDx" _
               & " LEFT JOIN Category d" _
                  & " ON a.sCategID1 = d.sCategrID" _
            & " ORDER BY a.sBarrCode" _
               & ", a.sDescript"
                  
   lsSQL = AddCondition(lsSQL, IIf(lbByCode, "a.sBarrcode LIKE " & strParm(lsValue & "%"), "a.sDescript LIKE " & strParm(lsValue & "%")))
   Set lrsCellphone = New ADODB.Recordset
   With lrsCellphone
      .Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If .EOF Then
         oDriver_InitValue
         For pnCtr = 0 To txtField.Count - 1
            txtField(pnCtr).Text = Empty
         Next
         GoTo endProc
      End If
      
      If .RecordCount = 1 Then
         lsStockIDx = .Fields("sStockIDx")
         GoTo LoadRecord
      Else
         With txtOthers(pnIndex)
            .BackColor = oApp.getColor("EB")
            lsBrowse = KwikBrowse(oApp, lrsCellphone _
                                    , "sBarrcode»sDescript»sBrandNme»sModelNme»sCategrNm" _
                                    , "BarrCode»Description»Brand»Model»Category" _
                                    , "@»@»@»@»@" _
                                    , "a.sBarrcode»a.sDescript»b.sBrandNme»c.sModelNme»e.sCategrNm")
                                    
            If lsBrowse <> "" Then
               lsSelected = Split(lsBrowse, "»")
               lsStockIDx = lsSelected(0)
               GoTo LoadRecord
            End If
            .BackColor = oApp.getColor("HT1")
            .SelStart = 0
            .SelLength = Len(.Text)
         End With
      End If
      GoTo endProc
   End With
      
LoadRecord:
   lsSQL = "SELECT" _
               & "  a.sStockIDx" _
               & ", a.sBranchCd" _
               & ", a.cRecdStat" _
               & ", b.sBarrcode" _
            & " FROM CP_Load_Matrix_Master a" _
               & ", CP_Load_Matrix b" _
            & " WHERE a.sStockIDx = b.sStockIDx" _
            & " ORDER BY a.sBranchCd DESC"
   lsSQL = AddCondition(lsSQL, "a.sStockIDx = " & strParm(lsStockIDx))
   Set lrsInventoryx = New ADODB.Recordset
   
   pbEnblButtons = True
   pbNewInvntory = False
   With lrsInventoryx
      .Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      If Not .EOF Then
         .Find "sBranchCd = " & strParm(oApp.BranchCode), 0, adSearchForward
         If Not lrsInventoryx.EOF Then
            If .Fields("cRecdStat") = xeRecStateActive Then
               oDriver.LookupValue(0) = .Fields("sBarrcode")
               oDriver.LoadRecord
            
               pbEnblButtons = False
            Else
               MsgBox "CP Inventory Status is Deactivated!!!" & vbCrLf & _
                        "Please Save the record to activate!!!", vbInformation, "Notice"
               
               oDriver.LookupValue(0) = .Fields("sBarrcode")
               oDriver.LoadRecord
               oDriver.RecordUpdate
               
               txtOthers(0).SetFocus
               oRS.Fields("nBegAmtxx") = 0
               oRS.Fields("nAmtOnHnd") = 0
               oRS.Fields("cRecdStat") = xeRecStateActive
               
               txtOthers(0).Text = "0.00"
               txtOthers(1).Text = "0.00"
            End If
         Else
            MsgBox "No Inventory found in your warehouse!!!" & vbCrLf & _
                     "Plese Save the record to create!!!", vbInformation, "Notice"
            
            .MoveFirst
            oDriver.LookupValue(0) = .Fields("sBarrcode")
            oDriver.LoadRecord
            oDriver.RecordUpdate
            txtOthers(0).SetFocus
            
            pbNewInvntory = True
         End If
      Else
         'no record at all
         pbNewInvntory = True
         If Not lrsCellphone.EOF Then
            MsgBox "No Inventory found in your warehouse!!!" & vbCrLf & _
                     "Plese Save the record to create!!!", vbInformation, "Notice"
            
            oDriver.LookupValue(0) = lrsCellphone("sBarrcode")
            oDriver.LoadRecord
            oDriver.RecordUpdate
            txtOthers(0).SetFocus
            
            pbNewInvntory = True
         End If
      End If
      .Close
   End With
endProc:
   Set lrsCellphone = Nothing
   Set lrsInventoryx = Nothing
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & oDriver.FieldValue(2) & "»" & oDriver.FieldValue(3) & " )", True
End Sub

Private Function UpdateCPInventory() As Boolean
   Dim lsOldProc As String
   Dim lrs As ADODB.Recordset
   Dim lsSQL As String
   Dim lnRow As Integer
   
   lsOldProc = "UpdateCPInventory"
   'On Error GoTo errProc

   lsSQL = AddCondition(psCPInventory, "sStockIDx = " & strParm(oDriver.FieldValue(5)) _
                  & " AND sBranchCd = " & strParm(oApp.BranchCode))

   Set lrs = New ADODB.Recordset
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.EOF Then
      lsSQL = ADO2SQL(oRS, "CP_Load_Matrix_Master", "", oApp.UserID, oApp.ServerDate, "")
   Else
      lsSQL = ADO2SQL(oRS, "CP_Load_Matrix_Master", "sStockIDx = " & strParm(oDriver.FieldValue(5)) & " AND sBranchCd = " & strParm(oApp.BranchCode), oApp.UserID, oApp.ServerDate, "")
   End If
   
   If lsSQL <> "" Then
      lnRow = oApp.Execute(lsSQL, "CP_Load_Matrix_Master", oApp.BranchCode, "")
      If lnRow <= 0 Then
         MsgBox "Unable to Save Inventory" & vbCrLf & _
                  lsSQL, vbCritical, "Warning"
         GoTo endProc
      End If
   End If
   
   UpdateCPInventory = True

endProc:
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & UpdateCPInventory & " )", True
End Function

Private Function SaveCPInventoryLedger() As Boolean
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lsOldProc As String

   lsOldProc = "SaveSPInventoryLedger"
   'On Error GoTo errProc

   lsSQL = "INSERT INTO CP_Load_Matrix_Ledger SET" _
               & "  sStockIDx = " & strParm(oDriver.FieldValue(5)) _
               & ", sBranchCd = " & strParm(oApp.BranchCode) _
               & ", sSourceCd = 'CPAd'" _
               & ", sSourceNo = '9900000001'" _
               & ", nAmtInxxx = " & CLng(txtOthers(3).Text) _
               & ", nAmtOutxx = '0'" _
               & ", nAmtOrder = '0'" _
               & ", nAmtIssue = '0'" _
               & ", nLedgerNo = '000001'" _
               & ", nAmtOnHnd = '0'" _
               & ", dTransact = " & dateParm(oApp.ServerDate) _
               & ", dModified = " & dateParm(oApp.ServerDate)
   
   lnRow = oApp.Execute(lsSQL, "CP_Load_Matrix_Ledger", oApp.BranchCode)
   If lnRow <= 0 Then
      MsgBox "Unable to Save CP_Inventory_Ledger!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   SaveCPInventoryLedger = True

endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
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

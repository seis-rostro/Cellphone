VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmInventory 
   BorderStyle     =   0  'None
   Caption         =   "Inventory"
   ClientHeight    =   5835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3780
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1050
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   6668
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.CheckBox chkField 
         Caption         =   "Cell Load"
         Height          =   315
         Index           =   2
         Left            =   2115
         TabIndex        =   51
         Tag             =   "wt0;fb0"
         Top             =   1380
         Width           =   1020
      End
      Begin VB.CheckBox chkField 
         Caption         =   "Card"
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   50
         Tag             =   "wt0;fb0"
         Top             =   1380
         Width           =   690
      End
      Begin VB.CheckBox chkField 
         Caption         =   "Cellphone"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   49
         Tag             =   "wt0;fb0"
         Top             =   1380
         Width           =   1035
      End
      Begin VB.CheckBox chkField 
         Caption         =   "Load Wallet"
         Height          =   315
         Index           =   3
         Left            =   3225
         TabIndex        =   48
         Tag             =   "wt0;fb0"
         Top             =   1380
         Width           =   1200
      End
      Begin VB.CheckBox chkField 
         Caption         =   "Microphone"
         Height          =   315
         Index           =   4
         Left            =   4515
         TabIndex        =   47
         Tag             =   "wt0;fb0"
         Top             =   1380
         Width           =   1170
      End
      Begin VB.CheckBox chkField 
         Caption         =   "W/ Serial"
         Height          =   315
         Index           =   5
         Left            =   4620
         TabIndex        =   46
         Tag             =   "wt0;fb0"
         Top             =   3270
         Width           =   1080
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1095
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   645
         Width           =   2820
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   7200
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1200
         Width           =   1515
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   1095
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1890
         Width           =   4605
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1095
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   975
         Width           =   4605
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1095
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   2220
         Width           =   2820
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1095
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2550
         Width           =   2820
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   1095
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2880
         Width           =   2820
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   1095
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   3210
         Width           =   2820
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   7230
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   315
         Width           =   1515
      End
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
         Height          =   315
         Index           =   7
         Left            =   10005
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   1530
         Width           =   1515
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   7200
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1530
         Width           =   1515
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   10005
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1200
         Width           =   1515
      End
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   12
         Left            =   10020
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   3060
         Width           =   1515
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   10020
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   315
         Width           =   1515
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   10
         Left            =   7230
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   3060
         Width           =   1515
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1095
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   195
         Width           =   2820
      End
      Begin VB.Shape Shape5 
         Height          =   1830
         Left            =   105
         Top             =   1800
         Width           =   5715
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   195
         Index           =   8
         Left            =   195
         TabIndex        =   6
         Top             =   690
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min. Level"
         Height          =   195
         Index           =   2
         Left            =   6075
         TabIndex        =   24
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   10
         Top             =   1950
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   8
         Top             =   1005
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Code"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   240
         Width           =   660
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1140
         Tag             =   "et0;ht2"
         Top             =   225
         Width           =   2820
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   12
         Top             =   2265
         Width           =   420
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   14
         Top             =   2595
         Width           =   435
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Made"
         Height          =   195
         Index           =   6
         Left            =   195
         TabIndex        =   16
         Top             =   2925
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   7
         Left            =   195
         TabIndex        =   18
         Top             =   3270
         Width           =   360
      End
      Begin VB.Shape Shape2 
         Height          =   1635
         Left            =   105
         Top             =   105
         Width           =   5715
      End
      Begin VB.Shape Shape3 
         Height          =   2040
         Left            =   5895
         Top             =   120
         Width           =   5715
      End
      Begin VB.Shape Shape4 
         Height          =   1440
         Left            =   5895
         Top             =   2190
         Width           =   5715
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pricing Info"
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
         Index           =   16
         Left            =   6060
         TabIndex        =   32
         Top             =   2415
         Width           =   990
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beg. Balance"
         Height          =   195
         Index           =   20
         Left            =   6075
         TabIndex        =   20
         Top             =   375
         Width           =   960
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty. On Hand"
         Height          =   195
         Index           =   21
         Left            =   8895
         TabIndex        =   30
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ReOrder Level"
         Height          =   195
         Index           =   23
         Left            =   6075
         TabIndex        =   28
         Top             =   1590
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Level"
         Height          =   195
         Index           =   25
         Left            =   8880
         TabIndex        =   26
         Top             =   1290
         Width           =   780
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   6090
         X2              =   11500
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price"
         Height          =   195
         Index           =   27
         Left            =   8895
         TabIndex        =   35
         Top             =   3135
         Width           =   870
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   6060
         X2              =   11500
         Y1              =   2670
         Y2              =   2670
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Pur. Date"
         Height          =   195
         Index           =   32
         Left            =   6075
         TabIndex        =   33
         Top             =   3105
         Width           =   1020
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beg. Inv. Date"
         Height          =   195
         Index           =   34
         Left            =   8880
         TabIndex        =   22
         Top             =   375
         Width           =   1035
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   480
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   847
      BorderStyle     =   1
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   1
         Top             =   75
         Width           =   2820
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   5910
         TabIndex        =   3
         Top             =   75
         Width           =   5700
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Code"
         Height          =   285
         Index           =   0
         Left            =   195
         TabIndex        =   0
         Top             =   120
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   285
         Index           =   9
         Left            =   4935
         TabIndex        =   2
         Top             =   120
         Width           =   795
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   8010
      TabIndex        =   38
      Top             =   5055
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
      Picture         =   "frmInventory.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   11070
      TabIndex        =   44
      Top             =   5055
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
      Picture         =   "frmInventory.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   10305
      TabIndex        =   43
      Top             =   5055
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
      Picture         =   "frmInventory.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   11070
      TabIndex        =   45
      Top             =   5055
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
      Picture         =   "frmInventory.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   8775
      TabIndex        =   40
      Top             =   5055
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
      Picture         =   "frmInventory.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   9540
      TabIndex        =   42
      Top             =   5055
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
      Picture         =   "frmInventory.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   8775
      TabIndex        =   41
      Top             =   5055
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
      Picture         =   "frmInventory.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   8010
      TabIndex        =   39
      Top             =   5055
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
      Picture         =   "frmInventory.frx":3456
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   7245
      TabIndex        =   37
      Top             =   5055
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
      Picture         =   "frmInventory.frx":3BD0
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset

Dim pbnewitem As Boolean
Dim psSelected() As String
Dim pnindex As Integer
Dim pbBoolean As Boolean
Dim psValue(2) As String
Dim lsPriceCode As String

Dim pnUserRights As Integer
Dim psUserID As String
Dim psUserName As String
Dim pbEnabled As Boolean

Dim txtfieldGotfocus As Boolean
Dim txtOthersGotfocus As Boolean
Dim pnCtr As Integer

Private Sub chkField_Click(Index As Integer)
   If chkField(0).Value = 1 Then chkField(5).Value = 1
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsSearch As String
   Dim lsApproval As Integer
   Dim orig As String
   Dim lsSQL As String
   Dim lsCondition As String
   
   Select Case Index
      Case 0 'cancel
         oDriver.RecordCancelUpdate
         txtothers(0).SetFocus
      Case 1 'browse
         orig = oDriver.BrowseQuery
         lsSQL = oDriver.BrowseQuery
         oDriver.BrowseQuery = lsSQL & " AND e.sBranchCd = '" & oApp.BranchCode & "'"
         oDriver.BrowseRecord
         oDriver.BrowseQuery = orig
      Case 2 'save
            oDriver.RecordSave
      Case 3 'Update
         oDriver.RecordUpdate
         If txtfield(0).Enabled = True Then txtfield(0).Enabled = False
         txtfield(12).SetFocus
      Case 4 'New
         oDriver.RecordNew
      Case 5 'close
         Unload Me
      Case 6 'delete
         If oApp.UserLevel <> xeEngineer Then
            lsApproval = MsgBox("Seek for Approval?", vbQuestion + vbYesNo, "Confirm")
            If lsApproval = vbYes Then
               If Not GetApproval(oApp, pnUserRights, psUserID, psUserName) Then Exit Sub
               If pnUserRights < xeEngineer Then
                  MsgBox "Approving User is Not Authorized!!!", vbCritical, "Warning"
                  Exit Sub
               Else
                  oDriver.RecordDelete
               End If
            End If
         Else
            oDriver.RecordDelete
         End If
      
      Case 7 'search
         If txtfieldGotfocus Then
            Select Case pnindex
               Case 2
                  If chkField(1).Value = 1 Then SearchCard False
               Case 3 To 5, 7 To 8
                  oDriver.RecordSearch txtfield(pnindex).Text
               Case 6
                  orig = oDriver.LookupQuery(6)
                  lsCondition = " a.sBrandIDx = '" & oDriver.FieldValue(5) & "'"
                  lsSQL = AddCondition(oDriver.LookupQuery(6), lsCondition)
                  oDriver.LookupQuery(6) = lsSQL
                  oDriver.RecordSearch txtfield(Index).Text
                  oDriver.LookupQuery(6) = orig
               If txtfield(pnindex).Text <> "" Then SetNextFocus
            End Select
         End If
      Case 8 'ledger
         If Not pbnewitem Then
            If chkField(2).Value = 1 Or chkField(3).Value = 1 Then
               frmLoad_Ledger.BarrCode = oDriver.FieldValue(0)
               frmLoad_Ledger.Show 1
            Else
               frmInventory_Ledger.BarrCode = oDriver.FieldValue(0)
               frmInventory_Ledger.Branch = oApp.BranchCode
               frmInventory_Ledger.Show 1
            End If
         Else
            MsgBox "Unable to Load Inventory Ledger!!!" & vbCrLf & _
                   "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
         End If
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   If bLoaded = False Then
      oDriver.RecordNew
      oDriver_InitValue
      bLoaded = True
   End If
End Sub

Private Sub Form_Load()
Dim lsSQL As String
Dim lnctr As Integer

CenterChildForm mdiMain, Me

bLoaded = False

Set oDriver = New FormDriver
Set oDriver.AppDriver = oApp
Set oDriver.MainForm = Me

Set oSkin = New FormSkin
Set oSkin.AppDriver = oApp
Set oSkin.Form = Me
oSkin.ApplySkin

      oDriver.RecQuery = "SELECT" _
                        & " sBarrcode ," _
                        & " sStockIDx ," _
                        & " sDescript ," _
                        & " sCategIDx ," _
                        & " sSupplier ," _
                        & " sBrandIDx ," _
                        & " sModelIDx ," _
                        & " sMadeIDxx ," _
                        & " sColorIDx ," _
                        & " nLastPrce ," _
                        & " dLastDate ," _
                        & " nPurPrice ," _
                        & " nSelPrice ," _
                        & " cCellPhon ," _
                        & " cCellCard ," _
                        & " cCellLoad ," _
                        & " cWalletxx ," _
                        & " cMicrofon ," _
                        & " cWdSerial ," _
                        
      oDriver.RecQuery = oDriver.RecQuery _
                        & " cRecdStat ," _
                        & " sCardIDxx ," _
                        & " sModified ," _
                        & " dModified ," _
                        & " vTimeStmp  " _
                     & " FROM CP_Inventory " _
      
      oDriver.BrowseQuery = "SELECT" _
                  & " a.sBarrcode, " _
                  & " b.sBrandNme, " _
                  & " c.sModelNme, " _
                  & " e.nQtyOnHnd, " _
                  & " a.sDescript, " _
                  & " d.sColorNme  " _
               & " FROM CP_Inventory a " _
                  & " LEFT JOIN Brand b " _
                     & " ON a.sBrandIDx = b.sBrandIDx " _
                  & " LEFT JOIN Model c " _
                     & " ON a.sModelIDx = c.sModelIDx " _
                  & " LEFT JOIN Color d " _
                     & " ON a.sColorIDx = d.sColorIDx " _
                  & " LEFT JOIN CP_Inventory_Master e " _
                     & " ON a.sStockIDx = e.sStockIDx " _
               & " WHERE a.cRecdStat = 1 " _

      oDriver.InitRecForm

      oDriver.BrowseColumn(0) = "sBarrcode"
      oDriver.BrowseColumn(1) = "sBrandNme"
      oDriver.BrowseColumn(2) = "sModelNme"
      oDriver.BrowseColumn(3) = "sDescript"
      oDriver.BrowseColumn(4) = "sColorNme"
      oDriver.BrowseColumn(5) = "nQtyonHnd"
      
   
      oDriver.BrowseFTitle(0) = "Bar Code"
      oDriver.BrowseFTitle(1) = "Brand"
      oDriver.BrowseFTitle(2) = "Model"
      oDriver.BrowseFTitle(3) = "Description"
      oDriver.BrowseFTitle(4) = "Color"
      oDriver.BrowseFTitle(5) = "QOH"
      
      oDriver.BrowseFFormat(5) = "#,##0"

      'Category
      oDriver.LookupQuery(3) = "SELECT" _
                        & " sCategIDx, " _
                        & " sCategNme  " _
                     & " FROM Category " _
                     & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY sCategNme"
         
      oDriver.LookupReference(3) = "sCategIDx»sCategNme"
      oDriver.LookupColumn(3) = "sCategNme"
      oDriver.LookupTitle(3) = "Category"
      
      'Supplier
      oDriver.LookupQuery(4) = "SELECT" _
                                  & " sSupplyID, " _
                                  & " sSupplyNm, " _
                                  & " sAddressx  " _
                          & " FROM Supplier a " _
                          & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                          & " ORDER BY a.sSupplyNm"
      
      oDriver.LookupReference(4) = "sSupplyID»sSupplyNm»sAddressx"
      oDriver.LookupColumn(4) = "sSupplyNm»sAddressx"
      oDriver.LookupTitle(4) = "Supplier Name»Address"

      'Brand
      oDriver.LookupQuery(5) = "SELECT" _
                        & " sBrandIDx, " _
                        & " sBrandNme " _
                     & " FROM Brand " _
                     & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY sBrandNme"
      
      oDriver.LookupReference(5) = "sBrandIDx»sBrandNme"
      oDriver.LookupColumn(5) = "sBrandNme"
      oDriver.LookupTitle(5) = "Brand Name"

      'Model
      oDriver.LookupQuery(6) = "SELECT" _
                        & " a.sModelIDx, " _
                        & " a.sModelNme, " _
                        & " b.sBrandNme  " _
                     & "FROM Model a LEFT JOIN " _
                        & " Brand b " _
                           & " ON a.sBrandIDx = b.sBrandIDx " _
                     & "WHERE a.cRecdStat = 1 " _
                     & "ORDER BY a.sModelNme "
      
      oDriver.LookupReference(6) = "sModelIDx»sModelNme»sBrandNme"
      oDriver.LookupColumn(6) = "sModelNme»sBrandNme"
      oDriver.LookupTitle(6) = "Model»Brand"

      'Country
      oDriver.LookupQuery(7) = "SELECT" _
                        & " sMadeIDxx, " _
                        & " sMadeName " _
                     & " FROM Made " _
                     & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY sMadeName "
      
      oDriver.LookupReference(7) = "sMadeIDxx»sMadeName"
      oDriver.LookupColumn(7) = "sMadeName"
      oDriver.LookupTitle(7) = "Country"
      
      'Color
      oDriver.LookupQuery(8) = "SELECT" _
                        & " sColorIDx, " _
                        & " sColorNme " _
                     & " FROM Color" _
                     & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY sColorNme"
      
      oDriver.LookupReference(8) = "sColorIDx»sColorNme"
      oDriver.LookupColumn(8) = "sColorNme"
      oDriver.LookupTitle(8) = "Color"
      
      oDriver.FieldStart = 0
      oDriver.FieldFormat(0) = ">"
      oDriver.FieldFormat(10) = "MMMM DD, YYYY"
      oDriver.FieldFormat(12) = "#,##0.00"
      
      lsPriceCode = "PATRONIZEX"

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
Dim lnctr As Integer
Dim ctr As Integer

   For pnCtr = 2 To 7
      txtothers(pnCtr).Enabled = False
   Next
   
   For lnctr = 2 To 8
      txtfield(lnctr).Enabled = False
   Next
   
   For ctr = 0 To 3
      chkField(ctr).Enabled = False
   Next

   txtothers(0).Enabled = True
   txtothers(1).Enabled = True
   oDriver.DisableTextbox 10
End Sub

Private Sub oDriver_EnableOtherControl()
Dim lnctr As Integer
Dim ctr As Integer

   For pnCtr = 2 To 7
      txtothers(pnCtr).Enabled = False
   Next
   
   For lnctr = 2 To 8
      txtfield(lnctr).Enabled = False
   Next
   
   For ctr = 0 To 3
      chkField(ctr).Enabled = False
   Next

   txtothers(0).Enabled = False
   txtothers(1).Enabled = False
   oDriver.DisableTextbox 10
End Sub

Private Sub oDriver_InitValue()
oDriver.FieldReference(1) = True
oDriver.FieldValue(1) = getNextCode("CP_Inventory", "sStockIDx", True, oApp.Connection, True, oApp.BranchCode)
  
txtfield(0).Tag = oDriver.FieldValue(1)
  
    For pnCtr = 0 To txtothers.Count - 1
      txtothers(pnCtr).Text = ""
    Next
    
    oDriver.FieldValue(0) = NewBarrCode
    txtfield(0).Text = oDriver.FieldValue(0)
    
    txtothers(2).Tag = 0
    txtothers(2).Text = 0
    txtothers(3).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
    txtothers(4).Text = 1
    txtothers(5).Text = 1
    txtothers(6).Text = 1
    txtothers(7).Text = 0
    txtfield(2).Tag = ""
    txtfield(6).Tag = ""
    txtfield(7).Tag = ""
    txtfield(10).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
    txtfield(12).Text = "0.00"
    
    ClearChkField
        
    txtothers(2).Locked = False
    txtothers(3).Locked = False
    txtfield(8).Locked = False
    txtfield(10).Locked = False
    txtfield(12).Locked = False
    
    pbnewitem = True
    pbEnabled = True
    
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lsSQL As String
   Dim lnctr As Integer
   
   If oRS.State = adStateOpen Then oRS.Close
   pbnewitem = False
   
   lsSQL = "SELECT" _
               & " a.sStockIDx, " _
               & " a.sBarrCode, " _
               & " b.nBegQtyxx, " _
               & " b.dBegInvxx, " _
               & " b.nMinLevel, " _
               & " b.nMaxLevel, " _
               & " b.nReorderx, " _
               & " b.nQtyOnHnd, " _
               & " a.cCellPhon, " _
               & " a.cCellCard, " _
               & " a.cCellLoad, " _
               & " a.cWalletxx, " _
               & " a.cMicrofon, " _
               & " a.cWdSerial, " _
               & " a.sCardIDxx  " _
         & " FROM CP_Inventory a " _
               & " LEFT JOIN CP_Inventory_Master b " _
               & " ON a.sStockIDx = b.sStockIDx " _
         & " Where a.sStockIDx = '" & oDriver.FieldValue(1) & "' " _
               & " AND b.sBranchCd = '" & oApp.BranchCode & "'"
   oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   
    If Not oRS.EOF Then
      For lnctr = 2 To 7
         Select Case lnctr
            Case 2, 4 To 6
               txtothers(lnctr).Text = IIf(IsNull(oRS(lnctr)), 0, Format(oRS(lnctr), "#,##0"))
            Case 3
               txtothers(lnctr).Text = Format(oRS("dBegInvxx"), "MMMM DD, YYYY")
            Case 7
               If oDriver.FieldValue(15) = 1 Or oDriver.FieldValue(16) = 1 Then
                  txtothers(lnctr).Text = IIf(IsNull(oRS(lnctr)), 0, Format(oRS(lnctr), "#,##0.00"))
               Else
                  txtothers(lnctr).Text = IIf(IsNull(oRS(lnctr)), 0, Format(oRS(lnctr), "#,##0"))
               End If
         End Select
      Next
      txtfield(2).Tag = oRS("sCardIDxx")
      chkField(0).Value = oRS("cCellPhon")
      chkField(1).Value = oRS("cCellCard")
      chkField(2).Value = oRS("cCellLoad")
      chkField(3).Value = oRS("cWalletxx")
      chkField(4).Value = oRS("cMicrofon")
      chkField(5).Value = oRS("cWdSerial")
   Else
      For lnctr = 0 To 7
         Select Case lnctr
         Case 0 To 1
            txtothers(lnctr).Text = ""
            txtothers(lnctr).Tag = ""
         Case 2, 4 To 7
            txtothers(lnctr).Text = 0
         Case 3
            txtothers(lnctr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
         End Select
      Next
      txtfield(2).Tag = ""
      ClearChkField
   End If
   
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open "SELECT * From CP_Inventory_Ledger " _
            & "WHERE sStockIDx = '" & oDriver.FieldValue(1) & "'" _
               & " AND sBranchCd = '" & oApp.BranchCode & "'" _
            , oApp.Connection, adOpenDynamic, adLockReadOnly, adCmdText
   
   If oRS.RecordCount > 1 Then
      If txtothers(2).Locked = False Then txtothers(2).Locked = True
      If txtothers(3).Locked = False Then txtothers(3).Locked = True
      If txtothers(7).Locked = False Then txtothers(7).Locked = True
   Else
      If txtothers(2).Locked = True Then txtothers(2).Locked = False
      If txtothers(3).Locked = True Then txtothers(3).Locked = False
      If txtothers(7).Locked = True Then txtothers(7).Locked = False
   End If

   txtothers(0).Text = oDriver.FieldValue(0)
   txtothers(1).Text = oDriver.FieldValue(2)
   txtothers(0).Tag = oDriver.FieldValue(0)
   txtothers(1).Tag = oDriver.FieldValue(2)
    
End Sub

Private Sub oDriver_Save(Saved As Boolean)
   Saved = False
End Sub

Private Sub oDriver_SaveComplete()
   If txtothers(0).Enabled Then txtothers(0).SetFocus
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
Dim lnctr As Integer

   If CDbl(txtfield(12).Text) = 0# And _
      chkField(2).Value = 0 And chkField(3).Value = 0 Then 'NOT Load
      MsgBox "Invalid Selling Price Detected!!!", vbCritical, "Warning"
      txtfield(12).SetFocus
      Cancel = True
   Else
      If pbnewitem Then
         'New Item Not Allowed in Branches
         MsgBox "Adding New Item Not Permitted!!!" & vbCrLf & vbCrLf & _
         "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
         Exit Sub
      Else
         'Update Item
         Cancel = Not UpdateCPInventoryMaster
            If Cancel Then Exit Sub
         If chkField(2).Value = 0 And chkField(3).Value = 0 Then
            'Category not Load
            Cancel = Not UpdateCPInventoryLedger
               If Cancel Then Exit Sub
         Else
            Cancel = Not UpdateLoadLedger
               If Cancel Then Exit Sub
         End If
      End If
      oDriver.FieldValue(13) = chkField(0).Value
      oDriver.FieldValue(14) = chkField(1).Value
      oDriver.FieldValue(15) = chkField(2).Value
      oDriver.FieldValue(16) = chkField(3).Value
      oDriver.FieldValue(17) = chkField(4).Value
      oDriver.FieldValue(18) = chkField(5).Value
      oDriver.FieldValue(19) = xeRecStateActive
      oDriver.FieldValue(20) = txtfield(2).Tag
   End If
       
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfield(Index).BackColor = &HE1FEFF
   txtfieldGotfocus = True
   txtOthersGotfocus = False
   pnindex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lsSQL As String
Dim lsCondition As String
Dim orig As String

   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 2 And (chkField(1).Value = 1 Or oDriver.FieldValue(3) = "01006") Then SearchCard False
      If Index > 2 And Index < 9 Then
         orig = oDriver.LookupQuery(6)
         If Index = 6 Then
            lsCondition = " a.sBrandIDx = '" & oDriver.FieldValue(5) & "'"
            lsSQL = AddCondition(oDriver.LookupQuery(6), lsCondition)
            oDriver.LookupQuery(6) = lsSQL
            oDriver.RecordSearch txtfield(Index).Text
            oDriver.LookupQuery(6) = orig
         Else
            oDriver.RecordSearch txtfield(Index).Text
         End If
      End If
      If txtfield(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
   End If
   
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   Select Case Index
      Case 12
         If Not IsNumeric(txtfield(Index).Text) Then txtfield(Index).Text = 0#
            txtfield(Index).Text = Format(txtfield(Index).Text, "#,##0.00")
      Case 10
         If Not IsDate(txtfield(Index).Text) Then
            txtfield(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
         Else
            txtfield(Index).Text = Format(txtfield(Index).Text, "MMMM DD, YYYY")
         End If
   End Select
   txtfield(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
       Case 0
           txtfield(Index).Text = Format(txtfield(Index).Text, ">")
       Case 12
           If Not IsNumeric(txtfield(Index).Text) Then txtfield(Index).Text = 0#
           txtfield(Index).Text = CDbl(txtfield(Index).Text)
       Case 10
          If Not IsDate(txtfield(Index).Text) Then
             txtfield(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
          Else
             txtfield(Index).Text = Format(txtfield(Index).Text, "MMMM DD, YYYY")
          End If
   End Select
   txtfield(Index).Text = TitleCase(txtfield(Index).Text)
   Cancel = Not oDriver.ValidateField(Index)
End Sub

Private Sub txtothers_GotFocus(Index As Integer)
   If Index = 3 Then txtothers(Index).Text = Format(txtothers(Index).Text, "MM/DD/YY")
   If txtothers(Index).Text <> "" Then
      txtothers(Index).SelStart = 0
      txtothers(Index).SelLength = Len(txtothers(Index).Text)
   End If
   txtfieldGotfocus = False
   txtothers(Index).BackColor = &HE1FEFF
   pnindex = Index
End Sub

Private Sub txtothers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lsSearch As String
Dim lnctr As Integer
Dim lsSQL As String
Dim orig As String
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      Select Case Index
      Case 0, 1
         If bLoaded Then
            If Trim(txtothers(Index).Text) <> "" Then
               If txtothers(Index).Text <> txtothers(Index).Tag Then
                  orig = oDriver.BrowseQuery
                  lsSQL = oDriver.BrowseQuery
                  lsSQL = lsSQL & " AND e.sBranchCd = " & strParm(oApp.BranchCode & "")
                  If Index = 0 Then
                     oDriver.BrowseQuery = lsSQL & " AND a.sBarrcode Like '%" & txtothers(Index).Text & "%'"
                     oDriver.BrowseRecord
                  ElseIf Index = 1 Then
                     oDriver.BrowseQuery = lsSQL & " AND a.sDescript Like '%" & txtothers(Index).Text & "%'"
                     oDriver.BrowseRecord
                  End If
                  oDriver.BrowseQuery = orig
                                    
                  If Trim(oDriver.FieldValue(0)) = "" Then
                     For lnctr = 0 To 7
                        Select Case lnctr
                           Case 0 To 1
                              txtothers(lnctr).Text = ""
                              txtothers(lnctr).Tag = ""
                           Case 2, 4 To 7
                              txtothers(lnctr).Text = 0
                           Case 3
                              txtothers(lnctr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
                        End Select
                     Next
                     ClearChkField
                  End If
               End If
            End If
            txtothers(Index).Tag = txtothers(Index).Text
         End If
      End Select
      
      If Trim(txtfield(0).Text) = "" Then
         txtothers(0).Text = ""
         txtothers(1).Text = ""
      End If
      txtothers(Index).SelStart = 0
      txtothers(Index).SelLength = Len(txtothers(Index).Text)
      KeyCode = 0
   End If
   
End Sub

Private Sub ClearChkField()
Dim lnctr As Integer
   For lnctr = 0 To 5
      chkField(lnctr).Value = 0
   Next
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
      Case 27
         Call Modified("CP_Inventory", "sStockIDx = '" & oDriver.FieldValue(1) & "' ")
   End Select
End Sub

'For CP Only
Private Function SaveCPInventoryMaster() As Boolean
Dim lsSQL As String
Dim lnrow As Long
   
SaveCPInventoryMaster = True
On Error GoTo errProc
   
   lsSQL = "INSERT INTO CP_Inventory_Master " _
            & "( sStockIDx, " _
            & "  sBranchCd, " _
            & "  nBegQtyxx, " _
            & "  nQtyOnHnd, " _
            & "  nReorderx, " _
            & "  nMinLevel, " _
            & "  nMaxLevel, " _
            & "  dBegInvxx, " _
            & "  cRecdStat, " _
            & "  sModified, " _
            & "  dModified) " _
                & "VALUES " _
                & "('" & oDriver.FieldValue(1) & "', " _
                & "'" & oApp.BranchCode & "', " _
                & "'" & CLng(txtothers(2).Text) & "', " _
                & "'" & CLng(txtothers(7).Text) & "', " _
                & "'" & CLng(txtothers(6).Text) & "', " _
                & "'" & CLng(txtothers(4).Text) & "', " _
                & "'" & CLng(txtothers(5).Text) & "', " _
                & "'" & txtothers(3).Text & "', " _
                & " '" & xeRecStateActive & "', " _
                & " '" & Encrypt(oApp.UserID) & "', " _
                & " getdate())"
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
      
   If lnrow <= 0 Then
      MsgBox "Unable to Save CP_Inventory_Master!!!", vbCritical, "Warning"
      SaveCPInventoryMaster = False
      GoTo endProc
   End If

endProc:
   Exit Function
errProc:
   SaveCPInventoryMaster = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'For CP Only
Private Function SaveCPInventoryLedger() As Boolean
Dim lsSQL As String
Dim lnrow As Long

SaveCPInventoryLedger = True
On Error GoTo errProc
   
   lsSQL = "INSERT INTO CP_Inventory_Ledger " _
            & "( sStockIDx, " _
            & "  sBranchCd, " _
            & "  sLocation, " _
            & "  sSourceCd, " _
            & "  sSourceNo, " _
            & "  nQtyInxxx, " _
            & "  nQtyOutxx, " _
            & "  nQtyOnHnd, " _
            & "  nEntryNox, " _
            & "  dTransact, " _
            & "  dModified) " _
                & "VALUES " _
                & "('" & oDriver.FieldValue(1) & "' ," _
                & "'" & oApp.BranchCode & "', " _
                & "'" & oApp.BranchCode & "', " _
                & " 'CPAd', " _
                & " '99000001', " _
                & "'" & CLng(txtothers(2).Text) & "', " _
                & " '0', " _
                & "'" & CLng(txtothers(7).Text) & "', " _
                & " '1', " _
                & "'3/15/2007'," _
                & " getdate())"
                                             
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
   
   If lnrow <= 0 Then
      MsgBox "Unable to Save CP_Inventory_Ledger!!!", vbCritical, "Warning"
      SaveCPInventoryLedger = False
      GoTo endProc
   End If

endProc:
   Exit Function
errProc:
   SaveCPInventoryLedger = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'For Cell Load Only
Private Function SaveLoadLedger() As Boolean
Dim lsSQL As String
Dim lnrow As Long

SaveLoadLedger = True
On Error GoTo errProc
   
   lsSQL = "INSERT INTO ELoad_Ledger " _
               & "( sStockIDx, " _
               & "  sBranchCd, " _
               & "  dTransact, " _
               & "  sReferNox, " _
               & "  sPhoneNum, " _
               & "  sSourceCd, " _
               & "  sSourceNo, " _
               & "  sTransNox, " _
               & "  nQtyInxxx, " _
               & "  nQtyOutxx, " _
               & "  nEntryNox, " _
               & "  nQtyOnHnd, " _
               & "  sModified, " _
               & "  dModified) "
            
   lsSQL = lsSQL _
            & "VALUES " _
               & "('" & oDriver.FieldValue(1) & "' ," _
               & "'" & oApp.BranchCode & "', " _
               & "'" & oApp.ServerDate & "', " _
               & "'99000001', " _
               & "'', " _
               & "'CPAd', " _
               & "'99000001', " _
               & "'1', " _
               & "'" & CDbl(txtothers(2).Text) & "', " _
               & "'0', " _
               & "'1', " _
               & "'" & CDbl(txtothers(7).Text) & "', " _
               & "'" & Encrypt(oApp.UserID) & "', " _
               & " getdate())"
                              
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
   
   If lnrow <= 0 Then
      MsgBox "Unable to Save ELoad_Ledger!!!", vbCritical, "Warning"
      SaveLoadLedger = False
      GoTo endProc
   End If

endProc:
   Exit Function
errProc:
   SaveLoadLedger = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'For Cell Load Only
Private Function UpdateLoadLedger() As Boolean
Dim lsSQL As String
Dim lnrow As Long

UpdateLoadLedger = True
On Error GoTo errProc

   lsSQL = "Update ELoad_Ledger SET" _
               & " dTransact = '" & oApp.ServerDate & "', " _
               & " nQtyInxxx = '" & CDbl(txtothers(7).Text) & "', " _
               & " nQtyOutxx = '0', " _
               & " nQtyOnHnd = '" & CDbl(txtothers(7).Text) & "', " _
               & " sModified = '" & Encrypt(oApp.UserID) & "', " _
               & " dModified = getdate() " _
            & " WHERE sStockIDx =  '" & oDriver.FieldValue(1) & "'" _
               & " And sBranchCd = '" & oApp.BranchCode & "' " _
               & " And nEntryNox = 1 "
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
   
   If lnrow <= 0 Then
      SaveLoadLedger
      MsgBox "Unable to Update ELoad_Ledger!!!", vbCritical, "Warning"
      UpdateLoadLedger = False
      GoTo endProc
   End If

endProc:
   Exit Function
errProc:
   UpdateLoadLedger = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'For CP Only
Private Function UpdateCPInventoryMaster() As Boolean
Dim lsSQL As String
Dim lnrow As Long

UpdateCPInventoryMaster = True
On Error GoTo errProc
   
   lsSQL = "UPDATE CP_Inventory_Master SET" _
            & " nBegQtyxx = '" & CLng(txtothers(2).Text) & "', " _
            & " nQtyOnHnd = '" & CLng(txtothers(7).Text) & "', " _
            & " nReorderx = '" & CLng(txtothers(6).Text) & "', " _
            & " nMinLevel = '" & CLng(txtothers(4).Text) & "', " _
            & " nMaxLevel = '" & CLng(txtothers(5).Text) & "', " _
            & " dBegInvxx = '" & txtothers(3).Text & "', " _
            & " sModified = '" & Encrypt(oApp.UserID) & "', " _
            & " dModified = getdate() " _
            & " WHERE sStockIDx =  '" & oDriver.FieldValue(1) & "'" _
            & " And sBranchCd = '" & oApp.BranchCode & "' "
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
   
   If lnrow <= 0 Then
      MsgBox "Unable to Update CP_Inventory_Master!!!", vbCritical, "Warning"
      UpdateCPInventoryMaster = False
      GoTo endProc
   End If

endProc:
   Exit Function
errProc:
   UpdateCPInventoryMaster = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'For CP Only
Private Function UpdateCPInventoryLedger() As Boolean
Dim lsSQL As String
Dim lnrow As Long
Dim oRS As New ADODB.Recordset

UpdateCPInventoryLedger = True
On Error GoTo errProc

   Set oRS = New ADODB.Recordset
   lsSQL = "Select * from CP_Inventory_Ledger " _
      & " WHERE sStockIDx =  '" & oDriver.FieldValue(1) & "'" _
               & " And sBranchCd = '" & oApp.BranchCode & "' " _
      & " ORDER BY nEntryNox "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

   If oRS.RecordCount = 1 Then
      If oRS("sSourceCd") = "CPAd" Then
         lsSQL = "UPDATE CP_Inventory_Ledger SET " _
                     & " nQtyInxxx = '" & CLng(txtothers(2).Text) & "', " _
                     & " nQtyOnHnd = '" & CLng(txtothers(7).Text) & "', " _
                     & " dModified= getdate() " _
                  & " WHERE sStockIDx =  '" & oDriver.FieldValue(1) & "'" _
                     & " And sBranchCd = '" & oApp.BranchCode & "' " _
                     & " And nEntryNox = 1 " _
                     & " And sSourcecd = 'CPAd' "
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
      
         If lnrow <= 0 Then
            MsgBox "Unable to Update CP_Inventory_Ledger!!!", vbCritical, "Warning"
            UpdateCPInventoryLedger = False
            GoTo endProc
         End If
      End If
   End If

endProc:
   Exit Function
errProc:
   UpdateCPInventoryLedger = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function
Private Sub oDriver_Delete(Deleted As Boolean)
   Deleted = True
End Sub

Private Sub oDriver_DeleteComplete()
   For pnCtr = 0 To txtothers.Count - 1
      txtothers(pnCtr).Text = ""
   Next
   ClearChkField
   MsgBox "Record Successfully Deleted!!!", vbInformation, "Information"
End Sub

Private Sub oDriver_WillDelete(Cancel As Boolean)
Dim lsSQL As String
Dim lnrow As Long

On Error GoTo errProc
   
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open "SELECT * From CP_Inventory_Ledger " _
            & "WHERE sStockIDx = '" & oDriver.FieldValue(1) & "'" _
               & " AND sBranchCd = '" & oApp.BranchCode & "'" _
            , oApp.Connection, adOpenDynamic, adLockReadOnly, adCmdText
            
   If oRS.RecordCount = 1 Then
      lsSQL = "DELETE CP_Inventory_Ledger " _
               & "WHERE sStockIDx = '" & oDriver.FieldValue(1) & "' " _
                  & " AND sSourceCd = 'CPAd'" _
                  & " AND sBranchCd = '" & oApp.BranchCode & "'"
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
      If lnrow <> 0 Then oApp.RegisDelete lsSQL
      If lnrow = 0 Then
         MsgBox "Unable to Save CP_Inventory_Ledger!!!", vbCritical, "Warning"
         Cancel = True
         GoTo endProc
      End If
   Else
      MsgBox "Unit Has Other Transactions!!!" & vbCrLf & _
      "" & vbCrLf & _
      "Delete Not Permitted!!!", vbInformation, "Notice"
      Cancel = True
      GoTo endProc
   End If

   
   lsSQL = "DELETE CP_Inventory " _
            & "WHERE sStockIDx = '" & oDriver.FieldValue(1) & "' "
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
   If lnrow <> 0 Then oApp.RegisDelete lsSQL
   
   If lnrow = 0 Then
      MsgBox "Unable to Delete CP_Inventory!!!", vbCritical, "Warning"
      Cancel = True
      GoTo endProc
   End If
   
   lsSQL = "DELETE CP_Inventory_Master " _
            & "WHERE sStockIDx = '" & oDriver.FieldValue(1) & "' " _
               & " AND sBranchCd = '" & oApp.BranchCode & "' "
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
   If lnrow <> 0 Then oApp.RegisDelete lsSQL
   If lnrow = 0 Then
      MsgBox "Unable to Delete CP_Inventory_Master!!!", vbCritical, "Warning"
      Cancel = True
      GoTo endProc
   End If
                              
endProc:
   Exit Sub
errProc:
   Cancel = True
   MsgBox Err.Description, vbCritical, "Warning"
End Sub

Private Sub txtothers_LostFocus(Index As Integer)
   txtothers(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtothers_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
      Case 2, 4, 5, 6, 7
          If Not IsNumeric(txtothers(Index).Text) Then txtothers(Index).Text = 0
             txtothers(Index).Text = Format(txtothers(Index).Text, "#,##0")
      Case 10, 11
          If Not IsNumeric(txtothers(Index).Text) Then txtothers(Index).Text = 0#
          txtothers(Index).Text = Format(txtothers(Index).Text, "#,##0.00")
   Case 3
      If Not IsDate(txtothers(Index).Text) Then
         txtothers(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Else
         txtothers(Index).Text = Format(txtothers(Index).Text, "MMMM DD, YYYY")
      End If
   End Select
   txtothers(Index).Text = TitleCase(txtothers(Index).Text)
End Sub

Function NewBarrCode() As String
   Dim lrs As Recordset
   Dim lsSQL As String
   Dim lnctr As Long
   
   lsSQL = "SELECT TOP 1" & _
            " sBarrCode" & _
            " FROM CP_Inventory " & _
            " WHERE sBarrCode LIKE " & strParm(Format(Date, "yy") & "-GMC-%") & _
            " ORDER BY sBarrCode DESC"
   
   Set lrs = New Recordset
   lrs.Open lsSQL, oApp.Connection, , , adCmdText
   
   If lrs.EOF Then
      lnctr = 1
   Else
      If Left(lrs("sBarrCode"), 2) = Format(Date, "yy") Then
         lnctr = CLng(Right(lrs("sBarrCode"), 6)) + 1
      Else
         lnctr = 1
      End If
   End If
   NewBarrCode = Format(Date, "yy") & "-GMC-" & Format(lnctr, "000000")
   
   Set lrs = Nothing
End Function

Private Sub SearchCard(ByVal SearchValue As Boolean)
   Dim lsSearch As String
   Dim lrs As ADODB.Recordset
   Dim lsSQL As String
   
   Set lrs = New ADODB.Recordset
   lsSQL = "SELECT" _
            & " sCardIDxx, " _
            & " sCardName  " _
         & " FROM Card" _
         & " WHERE cRecdStat = 1 " _

   If SearchValue Then
      lsSQL = lsSQL & " AND sCardName = '" & txtfield(2).Text & "'"
   Else
      lsSQL = lsSQL & " AND sCardName LIKE '" & txtfield(2).Text & "%' "
   End If
            
   lsSQL = lsSQL & " ORDER BY sCardName"
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.RecordCount = 1 Then
      txtfield(2).Text = lrs("sCardName")
      txtfield(2).Tag = lrs("sCardIDxx")
   ElseIf lrs.RecordCount > 1 Then
        lsSearch = KwikBrowse(oApp, lrs, _
                          "sCardIDxx»" _
                        & "sCardName", _
                          "Card ID»" _
                        & "Card Name")
                        
        If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            txtfield(2).Text = psSelected(1)
            txtfield(2).Tag = psSelected(0)
        End If
   Else
      txtfield(2).Text = ""
      txtfield(2).Tag = ""
   End If
   Set lrs = Nothing
End Sub




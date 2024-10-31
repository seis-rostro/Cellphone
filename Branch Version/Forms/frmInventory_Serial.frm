VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmInventory_Serial 
   BorderStyle     =   0  'None
   Caption         =   "Inventory"
   ClientHeight    =   5835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame2 
      Height          =   3780
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1050
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   6668
      BackColor       =   12632256
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
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   7230
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   2730
         Width           =   1515
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   10
         Left            =   7230
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   3060
         Width           =   1515
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   10005
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   315
         Width           =   1515
      End
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   11
         Left            =   10020
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   2730
         Width           =   1515
      End
      Begin VB.TextBox txtfield 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   12
         Left            =   10020
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   3060
         Width           =   1515
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   10005
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   1080
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
         Top             =   1410
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
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   1410
         Width           =   1515
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   7230
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   315
         Width           =   1515
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   1095
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   3210
         Width           =   2820
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   1095
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2880
         Width           =   2820
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1095
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2550
         Width           =   2820
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   5
         Left            =   1095
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2220
         Width           =   2820
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
         Index           =   4
         Left            =   1095
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1890
         Width           =   4605
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   7200
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1080
         Width           =   1515
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
      Begin VB.CheckBox chkField 
         Caption         =   "Card"
         Height          =   315
         Index           =   1
         Left            =   3165
         TabIndex        =   11
         Tag             =   "wt0;fb0"
         Top             =   1365
         Width           =   750
      End
      Begin VB.CheckBox chkField 
         Caption         =   "Cellphone"
         Height          =   315
         Index           =   0
         Left            =   1635
         TabIndex        =   10
         Tag             =   "wt0;fb0"
         Top             =   1380
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beg. Inv. Date"
         Height          =   195
         Index           =   34
         Left            =   8880
         TabIndex        =   24
         Top             =   375
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Pur. Date"
         Height          =   195
         Index           =   32
         Left            =   6075
         TabIndex        =   37
         Top             =   3105
         Width           =   1020
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Pur. Price"
         Height          =   195
         Index           =   31
         Left            =   6075
         TabIndex        =   35
         Top             =   2790
         Width           =   1035
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   6060
         X2              =   11500
         Y1              =   2430
         Y2              =   2430
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price"
         Height          =   195
         Index           =   27
         Left            =   8895
         TabIndex        =   41
         Top             =   3135
         Width           =   870
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Price"
         Height          =   300
         Index           =   26
         Left            =   8895
         TabIndex        =   39
         Top             =   2775
         Width           =   1080
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   6090
         X2              =   11500
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Level"
         Height          =   195
         Index           =   25
         Left            =   8880
         TabIndex        =   30
         Top             =   1170
         Width           =   780
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ReOrder Level"
         Height          =   195
         Index           =   23
         Left            =   6075
         TabIndex        =   28
         Top             =   1470
         Width           =   1035
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty. On Hand"
         Height          =   195
         Index           =   21
         Left            =   8895
         TabIndex        =   32
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beg. Balance"
         Height          =   195
         Index           =   20
         Left            =   6075
         TabIndex        =   22
         Top             =   375
         Width           =   960
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
         TabIndex        =   34
         Top             =   2175
         Width           =   990
      End
      Begin VB.Shape Shape4 
         Height          =   1680
         Left            =   5895
         Top             =   1950
         Width           =   5715
      End
      Begin VB.Shape Shape3 
         Height          =   1800
         Left            =   5895
         Top             =   105
         Width           =   5715
      End
      Begin VB.Shape Shape2 
         Height          =   1635
         Left            =   105
         Top             =   105
         Width           =   5715
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   7
         Left            =   195
         TabIndex        =   20
         Top             =   3270
         Width           =   360
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Made"
         Height          =   195
         Index           =   6
         Left            =   195
         TabIndex        =   18
         Top             =   2925
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   16
         Top             =   2595
         Width           =   435
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   14
         Top             =   2265
         Width           =   420
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
         Caption         =   "Bar Code"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   4
         Top             =   240
         Width           =   660
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
         Caption         =   "Supplier"
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   12
         Top             =   1950
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min. Level"
         Height          =   195
         Index           =   2
         Left            =   6075
         TabIndex        =   26
         Top             =   1140
         Width           =   735
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
      Begin VB.Shape Shape5 
         Height          =   1830
         Left            =   105
         Top             =   1800
         Width           =   5715
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
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   5910
         TabIndex        =   3
         Top             =   75
         Width           =   5700
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   1
         Top             =   75
         Width           =   2820
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
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   9555
      TabIndex        =   43
      Top             =   5040
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
      Picture         =   "frmInventory_Serial.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   11085
      TabIndex        =   45
      Top             =   5040
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
      Picture         =   "frmInventory_Serial.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   2
      Left            =   10320
      TabIndex        =   44
      Top             =   5040
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
      Picture         =   "frmInventory_Serial.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   11085
      TabIndex        =   46
      Top             =   5040
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
      Picture         =   "frmInventory_Serial.frx":166E
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmInventory_Serial"
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
Dim pnCtr As Integer
Dim pnindex As Integer
Dim pbBoolean As Boolean
Dim txtfieldGotfocus As Boolean
Dim lsPriceCode As String
Dim pbEnabled As Boolean

Private Sub cmdButton_Click(Index As Integer)
Dim lsSearch As String
Dim lsRep As Integer
Dim orig As String
Dim lsSQL As String
Dim lsCondition As String
   
   Select Case Index
   Case 0 'cancel
      Unload Me
      frmPO_Receiving.GridEditor1.SetFocus
   Case 1 'save
         With frmPO_Receiving.GridEditor1
            .TextMatrix(.Row, 1) = txtField(0).Text
            .TextMatrix(.Row, 2) = Trim(txtField(5).Text & " " & _
                                    txtField(6).Text & " " & _
                                    txtField(2).Text)
            .TextMatrix(.Row, 3) = Format(txtField(12).Text, "#,##0.00")
            .TextMatrix(.Row, 4) = oDriver.FieldValue(1)
            .TextMatrix(.Row, 5) = txtothers(2).Text
            .TextMatrix(.Row, 6) = oDriver.FieldValue(3)
         End With
         oDriver.RecordSave
   Case 2 'search
      If txtfieldGotfocus Then
         Select Case pnindex
            Case 2
               If chkField(1).Value = 1 Then SearchCard False
               
            Case 3 To 5, 7 To 9
               oDriver.RecordSearch txtField(pnindex).Text
            Case 6
               orig = oDriver.LookupQuery(6)
               lsCondition = " a.sBrandIDx = '" & oDriver.FieldValue(5) & "'"
               lsSQL = AddCondition(oDriver.LookupQuery(6), lsCondition)
               oDriver.LookupQuery(6) = lsSQL
               oDriver.RecordSearch txtField(Index).Text
               oDriver.LookupQuery(6) = orig
            If txtField(pnindex).Text <> "" Then SetNextFocus
         End Select
      Else
         txtothers(pnindex).SetFocus
      End If
   Case 3 'close
      Unload Me
      frmPO_Receiving.GridEditor1.SetFocus
   End Select
   
End Sub

Private Sub HidePrice()
Dim lnCtr As Integer

For lnCtr = 9 To 12
   Select Case lnCtr
   Case 9, 11, 12
      If oApp.UserLevel = xeEncoder Or oApp.UserLevel = xeSupervisor Then
         txtField(lnCtr).Text = Price2Code(oDriver.FieldValue(lnCtr), lsPriceCode)
      Else
         txtField(lnCtr).Text = Format(oDriver.FieldValue(lnCtr), "#,##0.00")
      End If
   End Select
Next
If txtothers(0).Enabled Then txtothers(0).SetFocus

End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      oDriver.RecordNew
      bLoaded = True
   End If
End Sub

Private Sub Form_Load()
Dim lsSQL As String
Dim lnCtr As Integer

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
                        & " cRecdStat ," _
                        & " sCardIDxx ," _
                        & " sModified ," _
                        & " dModified ," _
                        & " vTimeStmp  " _
                     & " FROM CP_Inventory " _

      oDriver.BrowseQuery = "SELECT" _
                        & " a.sBarrcode, " _
                        & " b.sBrandNme+' '+c.sModelNme as xDescript, " _
                        & " e.nQtyOnHnd, " _
                        & " a.sDescript  " _
                     & " FROM CP_Inventory a " _
                        & " LEFT JOIN Brand b " _
                           & " ON a.sBrandIDx = b.sBrandIDx " _
                        & " LEFT JOIN Model c " _
                           & " ON a.sModelIDx = c.sModelIDx " _
                        & " LEFT JOIN CP_Inventory_Master e " _
                           & " ON a.sStockIDx = e.sStockIDx " _
                     & " WHERE e.sBranchCd = '" & oApp.BranchCode & "' " _

      oDriver.InitRecForm

      oDriver.BrowseColumn(0) = "sBarrcode"
      oDriver.BrowseColumn(1) = "xDescript"
      oDriver.BrowseColumn(2) = "sDescript"
      oDriver.BrowseColumn(3) = "nQtyonHnd"
      
   
      oDriver.BrowseFTitle(0) = "Bar Code"
      oDriver.BrowseFTitle(1) = "Brand and Model"
      oDriver.BrowseFTitle(2) = "Description"
      oDriver.BrowseFTitle(3) = "QOH"

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
                     & " FROM Supplier " _
                     & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY sSupplyNm "
      
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
      oDriver.LookupColumn(7) = "sMadeIDxx»sMadeName"
      oDriver.LookupTitle(7) = "CountryID»Country"
      
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
      oDriver.FieldFormat(9) = "#,##0.00"
      oDriver.FieldFormat(10) = "MMMM DD, YYYY"
      oDriver.FieldFormat(11) = "#,##0.00"
      oDriver.FieldFormat(12) = "#,##0.00"
      
      lsPriceCode = "PATRONIZEX"

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   For pnCtr = 2 To 7
      txtothers(pnCtr).Enabled = False
   Next
   txtothers(0).Enabled = True
   txtothers(1).Enabled = True
   txtothers(7).Enabled = False
End Sub

Private Sub oDriver_EnableOtherControl()
   For pnCtr = 0 To 7
      Select Case pnCtr
      Case 2, 4 To 7
         txtothers(pnCtr).Enabled = True
      Case Else
         txtothers(pnCtr).Enabled = pbEnabled
      End Select
   Next
   txtothers(0).Enabled = False
   txtothers(1).Enabled = False
   txtothers(7).Enabled = False
End Sub

Private Sub oDriver_InitValue()

oDriver.FieldReference(1) = True
oDriver.FieldValue(1) = getNextCode("CP_Inventory", "sStockIDx", True, oApp.Connection, True, oApp.BranchCode)
    
   For pnCtr = 0 To txtothers.Count - 1
     txtothers(pnCtr).Text = ""
   Next
        
   txtothers(2).Tag = 0
   txtothers(2).Text = 0
   txtothers(3).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   txtothers(4).Text = 1
   txtothers(5).Text = 1
   txtothers(6).Text = 1
   txtothers(7).Text = 0
   txtField(2).Tag = ""
   txtField(6).Tag = ""
   txtField(7).Tag = ""
   txtField(9).Text = "0.00"
   txtField(10).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   txtField(11).Text = "0.00"
   txtField(12).Text = "0.00"
   
   ClearChkField
       
   txtothers(2).Locked = False
   txtothers(3).Locked = False
   txtothers(7).Enabled = False
   txtField(9).Locked = False
   txtField(10).Locked = False
   txtField(11).Locked = False
   txtField(12).Locked = False
   
   pbnewitem = True
   pbEnabled = True
    
    
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lsSQL As String
   Dim lnCtr As Integer
   
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
                & " a.sCardIDxx  " _
        & " FROM CP_Inventory a " _
                & " LEFT JOIN CP_Inventory_Master b " _
                & " ON a.sStockIDx = b.sStockIDx " _
        & " Where a.sStockIDx = '" & oDriver.FieldValue(1) & "' " _

   oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   
    If Not oRS.EOF Then
      For lnCtr = 2 To 7
         Select Case lnCtr
            Case 2, 4 To 7
               txtothers(lnCtr).Text = IIf(IsNull(oRS(lnCtr)), 0, Format(oRS(lnCtr), "#,##0"))
            Case 3
               txtothers(lnCtr).Text = Format(oRS("dBegInvxx"), "MMMM DD, YYYY")
         End Select
      Next
         txtField(2).Tag = oRS("sCardIDxx")
         chkField(0).Value = oRS("cCellPhon")
         chkField(1).Value = oRS("cCellCard")
   Else
      For lnCtr = 0 To 7
         Select Case lnCtr
         Case 0 To 1
            txtothers(lnCtr).Text = ""
            txtothers(lnCtr).Tag = ""
         Case 2, 4 To 7
            txtothers(lnCtr).Text = 0
         Case 3
            txtothers(lnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
         End Select
      Next
      txtField(2).Tag = ""
      ClearChkField
   End If
   
   txtothers(2).Locked = True
   txtothers(3).Locked = True
   
   txtothers(0).Text = oDriver.FieldValue(0)
   txtothers(1).Text = oDriver.FieldValue(2)
   txtothers(0).Tag = oDriver.FieldValue(0)
   txtothers(1).Tag = oDriver.FieldValue(2)
    
End Sub

Private Sub oDriver_Save(Saved As Boolean)
   Saved = False
End Sub

Private Sub oDriver_SaveComplete()
   Unload Me
   frmPO_Receiving.GridEditor1.SetFocus
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   If oDriver.FieldValue(0) = "" Then
      MsgBox "Invalid BarrCode Detected!!!", vbCritical, "Warning"
      txtField(0).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Stock ID Detected!!!", vbCritical, "Warning"
      txtField(1).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(2) = "" And chkField(1).Value = 1 Then
      MsgBox "Invalid Description Detected!!!", vbCritical, "Warning"
      txtField(2).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(3) = "" Then
      MsgBox "Invalid Category Detected!!!", vbCritical, "Warning"
      txtField(3).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(5) = "" And chkField(0).Value = 1 Then
      MsgBox "Invalid Brand Detected!!!", vbCritical, "Warning"
      txtField(5).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(6) = "" And chkField(0).Value = 1 Then
      MsgBox "Invalid Model Detected!!!", vbCritical, "Warning"
      txtField(6).SetFocus
      Cancel = True
   ElseIf CDbl(txtField(12).Text) = 0 Then
      MsgBox "Invalid Selling Price Detected!!!", vbCritical, "Warning"
      txtField(12).SetFocus
      Cancel = True
   Else
      If pbnewitem Then
         Cancel = Not Save_CP_Inventory_Master
            If Cancel Then Exit Sub
      End If
      oDriver.FieldValue(0) = txtField(0).Text
      oDriver.FieldValue(13) = chkField(0).Value
      oDriver.FieldValue(14) = chkField(1).Value
      oDriver.FieldValue(15) = 0
      oDriver.FieldValue(16) = 0
      oDriver.FieldValue(17) = xeRecStateActive
      oDriver.FieldValue(18) = txtField(2).Tag
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfieldGotfocus = True
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
         End If
         oDriver.RecordSearch txtField(Index).Text
         oDriver.LookupQuery(6) = orig
      End If
      If txtField(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
   End If
   
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   Select Case Index
      Case 9, 11, 12
         If Not IsNumeric(txtField(Index).Text) Then txtField(Index).Text = 0#
            txtField(Index).Text = Format(txtField(Index).Text, "#,##0.00")
      Case 10
         If Not IsDate(txtField(Index).Text) Then
            txtField(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
         Else
            txtField(Index).Text = Format(txtField(Index).Text, "MMMM DD, YYYY")
         End If
   End Select
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            txtField(Index).Text = Format(txtField(Index).Text, ">")
        Case 9, 11, 12
            If Not IsNumeric(txtField(Index).Text) Then txtField(Index).Text = 0#
                If Index = 12 Then
                    If CDbl(txtField(12).Text) <= CDbl(txtField(11).Text) Then
                        MsgBox "Invalid Selling Price!!!", vbCritical, "Warning"
                        txtField(12).SetFocus
                    End If
                End If
            txtField(Index).Text = CDbl(txtField(Index).Text)
        Case 10
           If Not IsDate(txtField(Index).Text) Then
              txtField(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
           Else
              txtField(Index).Text = Format(txtField(Index).Text, "MMMM DD, YYYY")
           End If
    End Select
Cancel = Not oDriver.ValidateField(Index)
End Sub

Private Sub txtothers_GotFocus(Index As Integer)
   If Index = 3 Then txtothers(Index).Text = Format(txtothers(Index).Text, "MM/DD/YY")
   If txtothers(Index).Text <> "" Then
      txtothers(Index).SelStart = 0
      txtothers(Index).SelLength = Len(txtothers(Index).Text)
   End If
   txtfieldGotfocus = False
   pnindex = Index
End Sub

Private Sub txtothers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lsSearch As String
Dim lnCtr As Integer
Dim lsSQL As String

   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 0 Or Index = 1 Then
         If bLoaded Then
            If Index = 0 Then
               lsSQL = oDriver.BrowseQuery
               oDriver.BrowseQuery = lsSQL & " AND a.sBarrcode Like " & strParm(txtothers(Index).Text & "%")
               oDriver.BrowseRecord
               oDriver.BrowseQuery = lsSQL
            Else
               lsSQL = oDriver.BrowseQuery
               oDriver.BrowseQuery = lsSQL & " AND a.sDescript Like " & strParm(txtothers(Index).Text & "%")
               oDriver.BrowseRecord
               oDriver.BrowseQuery = lsSQL
            End If
                  If Trim(oDriver.FieldValue(0)) = "" Then
                     For lnCtr = 0 To 11
                        Select Case lnCtr
                           Case 0 To 1
                              txtothers(lnCtr).Text = ""
                              txtothers(lnCtr).Tag = ""
                           Case 2, 4 To 7
                              txtothers(lnCtr).Text = 0
                           Case 3
                              txtothers(lnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
                        End Select
                     Next
                     ClearChkField
                  End If
         End If
   
         If Trim(txtField(0).Text) = "" Then
            txtothers(0).Text = ""
            txtothers(1).Text = ""
         End If
         txtothers(Index).Tag = txtothers(Index).Text
      End If
      txtothers(Index).SelStart = 0
      txtothers(Index).SelLength = Len(txtothers(Index).Text)
      KeyCode = 0
   End If
   
End Sub
Private Sub ClearChkField()
   For pnCtr = 0 To 1
      chkField(pnCtr).Value = 0
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
   End Select
End Sub

Private Function Save_CP_Inventory_Master() As Boolean
Dim lsSQL As String
Dim lnrow As Long
   
Save_CP_Inventory_Master = True
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
      Save_CP_Inventory_Master = False
      GoTo endProc
   End If

endProc:
   Exit Function
errProc:
   Save_CP_Inventory_Master = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Sub txtothers_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 2, 4, 5, 6, 7
            If Not IsNumeric(txtothers(Index).Text) Then txtothers(Index).Text = 0
            If CDbl(txtothers(Index).Text) > 32767 Then txtothers(Index).Text = 0
                txtothers(Index).Text = Format(txtothers(Index).Text, "#,##0")
        Case 8, 10, 11
            If Not IsNumeric(txtothers(Index).Text) Then txtothers(Index).Text = 0#
            txtothers(Index).Text = Format(txtothers(Index).Text, "#,##0.00")
   
   Case 3, 9
      If Not IsDate(txtothers(Index).Text) Then
         txtothers(Index).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Else
         txtothers(Index).Text = Format(txtothers(Index).Text, "MMMM DD, YYYY")
      End If
   End Select

End Sub

Private Sub ComputePricing()
   Dim lnPurPrice As Double

   If oApp.UserLevel = xeEncoder Or oApp.UserLevel = xeSupervisor Then
      lnPurPrice = CDbl(Code2Price(txtField(11).Text, lsPriceCode))
   Else
      lnPurPrice = CDbl(txtField(11).Text)
   End If

   If lnPurPrice >= CDbl(txtField(12).Text) Then
      txtField(12).Text = "0.00"
      Exit Sub
   End If

End Sub

Private Function Price2Code(ByVal nPrice, ByVal sCode) As String
   Dim lacDigitCode(9) As String * 1
   Dim lsPrice, lsConPrice As String
   Dim lnCtr As Long
   
   lsPrice = Trim(Str(nPrice))
   If Len(sCode) < 10 Then
      If Len(sCode) <> 9 Then
         Price2Code = "XXX"
         Exit Function
      End If
      sCode = sCode & "Z"
   End If
   
   lacDigitCode(0) = Mid(sCode, 10, 1)
   For lnCtr = 1 To UBound(lacDigitCode)
      lacDigitCode(lnCtr) = Mid(sCode, lnCtr, 1)
   Next
   
   lsConPrice = ""
   For lnCtr = 1 To Len(lsPrice)
      If Mid(lsPrice, lnCtr, 1) = "." Then
         lsConPrice = lsConPrice & "."
      Else
         lsConPrice = lsConPrice & lacDigitCode(Int(Mid(lsPrice, lnCtr, 1)))
      End If
   Next
   If InStr(1, lsConPrice, ".", vbTextCompare) = 0 Then
      lsConPrice = lsConPrice & "." & lacDigitCode(0) & lacDigitCode(0)
   End If
   Price2Code = lsConPrice
End Function

Private Function Code2Price(ByVal sPrice, ByVal sCode) As Double
   Dim lsConPrice As String
   Dim lnCtr As Integer
   Dim lnValue As Integer

   sPrice = Trim(sPrice)
   If Len(sCode) < 10 Then
      If Len(sCode) <> 9 Then
         Code2Price = 0#
         Exit Function
      End If
      sCode = sCode & "Z"
   End If

   lsConPrice = ""
   For lnCtr = 1 To Len(sPrice)
      If Mid(sPrice, lnCtr, 1) = "." Then
         lsConPrice = lsConPrice & "."
      Else
         lnValue = InStr(1, sCode, Mid(sPrice, lnCtr, 1), vbTextCompare)
         If lnValue = 10 Then lnValue = 0

         lsConPrice = lsConPrice & Trim(Str(lnValue))
      End If
   Next
   If InStr(1, lsConPrice, ".", vbTextCompare) = 0 Then
      lsConPrice = lsConPrice & ".00"
   End If
   Code2Price = CDbl(lsConPrice)
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
      lsSQL = lsSQL & " AND sCardName = '" & txtField(2).Text & "'"
   Else
      lsSQL = lsSQL & " AND sCardName LIKE '" & txtField(2).Text & "%' "
   End If
            
   lsSQL = lsSQL & " ORDER BY sCardName"
   
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.RecordCount = 1 Then
      txtField(2).Text = lrs("sCardName")
      txtField(2).Tag = lrs("sCardIDxx")
      
   ElseIf lrs.RecordCount > 1 Then
        lsSearch = KwikBrowse(oApp, lrs, _
                          "sCardIDxx»" _
                        & "sCardName", _
                          "Card ID»" _
                        & "Card Name")
                        
        If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            txtField(2).Text = psSelected(1)
            txtField(2).Tag = psSelected(0)
        End If
   Else
      txtField(2).Text = ""
      txtField(2).Tag = ""
   End If
   Set lrs = Nothing

End Sub



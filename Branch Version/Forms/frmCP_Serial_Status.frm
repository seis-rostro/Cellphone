VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Serial_Status 
   BorderStyle     =   0  'None
   Caption         =   "CP Serial Status"
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4275
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1380
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   7541
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   14
         Left            =   1050
         TabIndex        =   48
         Top             =   1470
         Width           =   2595
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   4530
         TabIndex        =   47
         Text            =   "Combo2"
         Top             =   855
         Width           =   1710
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   4545
         TabIndex        =   39
         Top             =   2400
         Width           =   1725
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1050
         TabIndex        =   21
         Top             =   2745
         Width           =   5205
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1050
         TabIndex        =   19
         Top             =   2415
         Width           =   2595
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   4545
         TabIndex        =   31
         Top             =   1200
         Width           =   1725
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1050
         TabIndex        =   17
         Top             =   2100
         Width           =   2595
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   510
         Index           =   5
         Left            =   1050
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   3345
         Width           =   5205
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   4545
         TabIndex        =   35
         Top             =   1800
         Width           =   1725
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   4545
         TabIndex        =   37
         Top             =   2100
         Width           =   1725
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   4545
         TabIndex        =   33
         Top             =   1500
         Width           =   1725
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmCP_Serial_Status.frx":0000
         Left            =   4530
         List            =   "frmCP_Serial_Status.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   495
         Width           =   1725
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1050
         TabIndex        =   15
         Top             =   1785
         Width           =   2595
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1050
         TabIndex        =   23
         Top             =   3045
         Width           =   5205
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1050
         TabIndex        =   27
         Top             =   3870
         Width           =   5205
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1050
         TabIndex        =   13
         Top             =   1155
         Width           =   2595
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1050
         TabIndex        =   11
         Top             =   840
         Width           =   2595
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1050
         TabIndex        =   9
         Top             =   525
         Width           =   2595
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1050
         TabIndex        =   7
         Top             =   60
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   195
         Index           =   15
         Left            =   135
         TabIndex        =   49
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Type"
         Height          =   195
         Index           =   14
         Left            =   3690
         TabIndex        =   46
         Top             =   900
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DA S.I. No."
         Height          =   195
         Index           =   6
         Left            =   3690
         TabIndex        =   32
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   20
         Top             =   2775
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   24
         Left            =   135
         TabIndex        =   18
         Top             =   2445
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   195
         Index           =   22
         Left            =   3690
         TabIndex        =   28
         Top             =   570
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SO S.I No."
         Height          =   195
         Index           =   19
         Left            =   3705
         TabIndex        =   36
         Top             =   2145
         Width           =   870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Purc."
         Height          =   195
         Index           =   18
         Left            =   3705
         TabIndex        =   38
         Top             =   2415
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   16
         Left            =   150
         TabIndex        =   24
         Top             =   3390
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Refer. Date"
         Height          =   195
         Index           =   13
         Left            =   3690
         TabIndex        =   34
         Top             =   1830
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Refer. No."
         Height          =   195
         Index           =   12
         Left            =   3705
         TabIndex        =   30
         Top             =   1290
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   11
         Left            =   150
         TabIndex        =   16
         Top             =   2115
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   26
         Top             =   3900
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   22
         Top             =   3075
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Barrcode"
         Height          =   195
         Index           =   10
         Left            =   150
         TabIndex        =   10
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   14
         Top             =   1785
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial No"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   555
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial ID"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   105
         Width           =   765
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   1155
         Tag             =   "et0;ht2"
         Top             =   165
         Width           =   1920
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   5745
      TabIndex        =   44
      Top             =   5910
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
      Picture         =   "frmCP_Serial_Status.frx":0004
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   4965
      TabIndex        =   43
      Top             =   5910
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
      Picture         =   "frmCP_Serial_Status.frx":077E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   2625
      TabIndex        =   40
      Top             =   5910
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
      Picture         =   "frmCP_Serial_Status.frx":0EF8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   5745
      TabIndex        =   45
      Top             =   5910
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
      Picture         =   "frmCP_Serial_Status.frx":1672
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   3405
      TabIndex        =   41
      Top             =   5910
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
      Picture         =   "frmCP_Serial_Status.frx":1DEC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   4185
      TabIndex        =   42
      Top             =   5910
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
      Picture         =   "frmCP_Serial_Status.frx":2566
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   810
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1429
      BackColor       =   12632256
      Begin VB.TextBox txtOther 
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
         Index           =   11
         Left            =   1050
         TabIndex        =   1
         Top             =   105
         Width           =   2790
      End
      Begin VB.TextBox txtOther 
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
         Index           =   12
         Left            =   1050
         TabIndex        =   5
         Top             =   405
         Width           =   5220
      End
      Begin VB.TextBox txtOther 
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
         Index           =   10
         Left            =   4380
         TabIndex        =   3
         Top             =   105
         Width           =   1890
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Custo&mer"
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
         Index           =   17
         Left            =   165
         TabIndex        =   4
         Top             =   450
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Serial No."
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
         Index           =   8
         Left            =   165
         TabIndex        =   0
         Top             =   135
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "S. ID"
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
         Left            =   3885
         TabIndex        =   2
         Top             =   150
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmCP_Serial_Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Option Explicit
Private Const pxeMODULENAME = "frmCPSerialStatus"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pbLoadRecord As Boolean
Dim pbOtherField As Boolean
Dim pnCtr As Integer, pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsSearch As String
   Dim lsSelected() As String
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   
   Select Case Index
   Case 0
      oDriver.RecordCancelUpdate
   Case 1
      SearchSerial "", True
   Case 2
      oDriver.RecordSave
   Case 3
      If Trim(oDriver.FieldValue(0)) <> "" Then oDriver.RecordUpdate
   Case 5
      Unload Me
   Case 6
      MsgBox "Unable to Delete Record" & vbCrLf & _
             "Deleting Record is prohibited!!!", vbCritical, "Warning"
   Case 7
      oDriver.RecordSearch
      txtField(pnIndex).SetFocus
   Case 8
      If pbLoadRecord Then
         With frmCPSerialLedger
            .txtDateFrom = Format(DateAdd("m", -1, oApp.ServerDate), "MMMM DD, YYYY")
            .txtDateThru = Format(oApp.ServerDate, "MMMM DD, YYYY")
            .txtField(0) = txtField(0)
            .txtField(1) = txtField(1)
            .txtField(2) = txtOther(1) & " " & txtOther(2)
            .txtField(3) = txtOther(3)
            
            .SerialID = oDriver.FieldValue(0)
            .Show 1
         End With
      Else
         MsgBox "Unable to Load Serial Ledger!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Combo1_GotFocus()
   With Combo1
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      SetNextFocus
      KeyCode = 0
   End If
End Sub

Private Sub Combo1_LostFocus()
   With Combo1
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded Then
      oDriver.RecordCancelUpdate
      bLoaded = False
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
      
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin

   oDriver.RecQuery = "SELECT" _
                        & "  sSerialID" _
                        & ", sSerialNo" _
                        & ", sSupplier" _
                        & ", sStockIDx" _
                        & ", sBranchCd" _
                        & ", cLocation" _
                        & ", cUnitType" _
                        & ", sModified" _
                        & ", dModified" _
                     & " FROM CP_Inventory_Serial"
                     
oDriver.BrowseQuery = "SELECT Distinct" _
                        & "  a.sSerialID" _
                        & ", a.sSerialNo" _
                        & ", c.sBrandNme" _
                        & ", d.sModelNme" _
                        & ", e.sColorNme" _
                        & ", f.sCompnyNm" _
                        & ", a.sBranchCd" _
                     & " FROM CP_Inventory_Serial a" _
                        & " LEFT JOIN Client_Master f" _
                           & " ON a.sSupplier = f.sClientID" _
                        & ", CP_Inventory b" _
                           & " LEFT JOIN CP_Brand c" _
                              & " ON b.sBrandIDx = c.sBrandIDx" _
                           & " LEFT JOIN CP_Model d" _
                              & " ON b.sModelIDx = d.sModelIDx" _
                           & " LEFT JOIN Color e" _
                              & " ON b.sColorIDx = e.sColorIDx" _
                     & " WHERE a.sStockIDx = b.sStockIDx" _
                     & " ORDER BY a.sSerialID"
   
'   oDriver.BrowseQuery = "SELECT Distinct" _
'                        & "  a.sSerialID" _
'                        & ", a.sSerialNo" _
'                        & ", c.sBrandNme" _
'                        & ", d.sModelNme" _
'                        & ", e.sColorNme" _
'                        & ", f.sCompnyNm" _
'                        & ", g.sBranchCd" _
'                     & " FROM CP_Inventory_Serial a" _
'                        & " LEFT JOIN Client_Master f" _
'                           & " ON a.sSupplier = f.sClientID" _
'                        & ", CP_Inventory b" _
'                           & " LEFT JOIN CP_Brand c" _
'                              & " ON b.sBrandIDx = c.sBrandIDx" _
'                           & " LEFT JOIN CP_Model d" _
'                              & " ON b.sModelIDx = d.sModelIDx" _
'                           & " LEFT JOIN Color e" _
'                              & " ON b.sColorIDx = e.sColorIDx" _
'                        & ", CP_Inventory_Master g" _
'                     & " WHERE a.sStockIDx = b.sStockIDx" _
'                        & " AND b.sStockIDx = g.sStockIDx" _
'                     & " ORDER BY a.sSerialID"

   oDriver.InitRecForm
   
   oDriver.BrowseFReference(0) = True
   oDriver.BrowseFTitle(0) = "Serial ID"
   oDriver.BrowseFTitle(1) = "Serial No"
   oDriver.BrowseFTitle(2) = "Brand Name"
   oDriver.BrowseFTitle(3) = "Model Name"
   oDriver.BrowseFTitle(4) = "Color Name"
   oDriver.BrowseFTitle(5) = "Supplier"
   
   oDriver.LookupQuery(2) = "SELECT" _
                              & "  a.sClientID" _
                              & ", a.sCompnyNm" _
                           & " FROM Client_Master a" _
                              & ", CP_Supplier b" _
                           & " WHERE a.sClientID = b.sClientID" _
                              & " AND b.sBranchCd = " & strParm(oApp.BranchCode) _
                              & " AND b.cRecdStat = " & strParm(xeRecStateActive) _
                           & " ORDER BY a.sCompnyNm"
   
   oDriver.LookupReference(2) = "a.sClientID»a.sCompnyNm"
   oDriver.LookupColumn(2) = "a.sCompnyNm"
   oDriver.LookupTitle(2) = "Supplier"

   oDriver.LookupQuery(3) = "SELECT" _
                              & "  a.sStockIDx" _
                              & ", a.sBarrCode" _
                              & ", b.sBrandNme" _
                              & ", c.sModelNme" _
                           & " FROM CP_Inventory a" _
                              & " LEFT JOIN CP_Brand b" _
                                 & " ON a.sBrandIDx = b.sBrandIDx" _
                              & " LEFT JOIN CP_Model c" _
                                 & " ON a.sModelIDx = c.sModelIDx" _
                           & " ORDER BY a.sBarrcode"
                        
   oDriver.LookupReference(3) = "a.sStockIDx»a.sBarrCode»b.sBrandNme»c.sModelNme"
   oDriver.LookupColumn(3) = "a.sBarrCode»b.sBrandNme»c.sModelNme"
   oDriver.LookupTitle(3) = "BarrCode»Brand Name»Model Name"
   
   oDriver.LookupQuery(4) = "SELECT" _
                              & "  sBranchCd" _
                              & ", sBranchNm" _
                           & " FROM Branch" _
                           & " ORDER BY sBranchNm"
   
   oDriver.LookupReference(4) = "sBranchCd»sBranchNm"
   oDriver.LookupColumn(4) = "sBranchNm"
   oDriver.LookupTitle(4) = "Branch Name"
   
   oDriver.FieldFormat(0) = "@@@@-@@@@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
   
   oDriver.FieldStart = 1

   Combo1.ListIndex = -1
   Combo1.List(0) = "Warehouse"
   Combo1.List(1) = "Branch"
   Combo1.List(2) = "Supplier"
   Combo1.List(3) = "Customer"
   Combo1.List(4) = "Unknown"
   Combo1.List(5) = "Service Center"
   Combo1.List(6) = "Service Unit"

   Combo2.ListIndex = -1
   Combo2.List(0) = "LDU"
   Combo2.List(1) = "Regular"
   Combo2.List(2) = "Free"
   Combo2.List(3) = "Live"
   Combo2.List(4) = "Service"
   Combo2.List(5) = "RDU"
   Combo2.List(6) = "Others"

   bLoaded = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   Combo1.Enabled = False
   Combo2.Enabled = False
   txtOther(10).Enabled = True
   txtOther(11).Enabled = True
   txtOther(12).Enabled = True
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   oDriver.DisableTextbox 2
   oDriver.DisableTextbox 3
   oDriver.DisableTextbox 4
   
   oDriver.showButton 1
   Combo1.Enabled = True
   Combo2.Enabled = True
   
   txtOther(10).Enabled = False
   txtOther(11).Enabled = False
   txtOther(12).Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub oDriver_InitValue()
   pbLoadRecord = False
   
   For pnCtr = 0 To txtOther.Count - 1
      txtOther(pnCtr).Text = ""
      txtOther(pnCtr).Tag = ""
   Next
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_LoadOtherData"
   ''On Error GoTo errProc

   Dim lrs As ADODB.Recordset
   
   For pnCtr = 4 To 9
      txtOther(pnCtr).Text = ""
   Next
   txtOther(13).Text = ""
   
   Combo1.ListIndex = IIf(IsNull(oDriver.FieldValue(5)), -1, IIf(Trim(oDriver.FieldValue(5)) = "", -1, oDriver.FieldValue(5)))
   Combo2.ListIndex = IIf(IsNull(oDriver.FieldValue(6)), -1, IIf(Trim(oDriver.FieldValue(6)) = "", -1, oDriver.FieldValue(6)))
   txtOther(0).Text = Format(oDriver.FieldValue(0), "@@@@-@@@@@@")
   txtOther(1).Text = oDriver.FieldValue(1)

   Set lrs = New ADODB.Recordset
   lrs.Open "SELECT" _
               & "  a.sReferNox" _
               & ", a.sSalesInv" _
               & ", a.dTransact" _
               & ", c.sCompnyNm" _
            & " FROM CP_PO_Receiving_Master a" _
               & " LEFT JOIN Client_Master c ON a.sSupplier = c.sClientID" _
               & ", CP_PO_Receiving_Serial b" _
            & " WHERE a.sTransNox = b.sTransNox" _
               & " AND a.cTranStat <> '3' " _
               & " AND b.sSerialID = " & strParm(oDriver.FieldValue(0)) _
            , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If Not lrs.EOF Then
      txtOther(6).Text = lrs("sReferNox")
      txtOther(7).Text = lrs("sSalesInv")
      txtOther(8).Text = Format(lrs("dTransact"), "MMM-DD-YYYY")
      txtField(2).Text = lrs("sCompnyNm")
   Else
      'she 2023-02-02 for branches checking of supplier in case hindi nakarating sa kanila ung PO receiving master.
      'kailangan nila supplier to check if imei is included to the promo
      Set lrs = New ADODB.Recordset
      lrs.Open "SELECT" _
               & " b.sCompnyNm" _
            & " FROM CP_Inventory_Serial a" _
               & " LEFT JOIN Client_Master b ON a.sSupplier = b.sClientID" _
            & " WHERE a.sSerialID = " & strParm(oDriver.FieldValue(0)) _
            , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
            
      If Not lrs.EOF Then
         txtField(2).Text = lrs("sCompnyNm")
      End If
   End If
   
   lrs.Close
   
   lrs.Open "SELECT" _
               & "  a.sSalesInv" _
               & ", a.dTransact" _
               & ", CONCAT(c.sLastName, ', ', c.sFrstName, ' ' , c.sMiddName) xFullName" _
               & ", CONCAT(c.sAddressx, ', ', d.sTownName, ', ', e.sProvName, ' ', d.sZippCode) xAddressx" _
            & " FROM CP_SO_Master a" _
               & ", Client_Master c" _
                  & " LEFT JOIN TownCity d" _
                     & " ON c.sTownIDxx = d.sTownIDxx" _
                  & " LEFT JOIN Province e" _
                     & " ON d.sProvIDxx = e.sProvIDxx" _
               & ", CP_SO_Detail b" _
            & " WHERE a.sTransNox = b.sTransNox" _
               & " AND a.sClientID = c.sClientID" _
               & " AND b.sSerialID = " & strParm(oDriver.FieldValue(0)) _
               & " AND a.cTranStat <> " & strParm(xeStateCancelled) _
            , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If Not lrs.EOF Then
      txtOther(4).Text = lrs("xFullName")
      txtOther(5).Text = IFNull(lrs("xAddressx"), "")
      txtOther(9).Text = lrs("sSalesInv")
      txtOther(13).Text = Format(lrs("dTransact"), "MMM-DD-YYYY")
   End If
   lrs.Close
   
   txtOther(10).Text = txtField(0).Text
   txtOther(11).Text = txtField(1).Text
   
   pbLoadRecord = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_WillSave"
   ''On Error GoTo errProc
   
   If Combo1.ListIndex <> -1 Then oDriver.FieldValue(5) = CStr(Combo1.ListIndex)
   If Combo2.ListIndex <> -1 Then oDriver.FieldValue(6) = CStr(Combo2.ListIndex)
endProc:
   Exit Sub
errProc:
   Cancel = True
   ShowError lsOldProc & "( " & Cancel & " )"

End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   pnIndex = Index
   pbOtherField = False
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc
   
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
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
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True
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
      If Trim(.Text) <> "" Then .Text = UCase(Left(.Text, 1)) & Right(.Text, Len(.Text) - 1)
      
      Select Case Index
      Case 1
         .Text = UCase(.Text)
         Cancel = Not oDriver.ValidateField(Index)
      End Select
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub

Private Sub txtOther_GotFocus(Index As Integer)
   With txtOther(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   pnIndex = Index
   pbOtherField = True
End Sub

Private Sub txtOther_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtOther_KeyDown"
   ''On Error GoTo errProc

   If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
      With txtOther(Index)
         If Trim(.Text) = "" Then
            oDriver_InitValue
            KeyCode = 0
            Exit Sub
         End If
         
         If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then
            Select Case Index
            Case 10
               oDriver.LookupValue(0) = .Text
               oDriver.LoadRecord
            Case 11
               SearchSerial .Text, False
            Case 12
               SearchCustomer .Text, False
            End Select
        End If
      
        .Tag = .Text
        .SelStart = 0
        .SelLength = Len(.Text)
             
        
        If KeyCode = vbKeyF3 Then
            If .Text <> "" Then SetNextFocus
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

Private Sub SearchSerial(lsSerialNo As String, lbSearch As Boolean)
   Dim lrs As ADODB.Recordset
   Dim lsSQL As String, lsOldProc As String
   Dim lsCondition As String, lsSearch As String, lsSelected() As String
   
   lsOldProc = "searchSerial"
   ''On Error GoTo errProc
            
   lsSQL = "SELECT" & _
               "  a.sStockIDx" & _
               ", a.sSerialNo" & _
               ", b.sBarrCode" & _
               ", b.sDescript" & _
               ", c.sBrandNme" & _
               ", d.sModelNme" & _
               ", e.sColorNme" & _
               ", a.sSerialID" & _
               ", f.sCategrNm" & _
            " FROM CP_Inventory_Serial a" & _
               ", CP_Inventory b" & _
                  " LEFT JOIN CP_Brand c" & _
                     " ON b.sBrandIDx = c.sBrandIDx" & _
                  " LEFT JOIN CP_Model d" & _
                     " ON b.sModelIDx = d.sModelIDx" & _
                  " LEFT JOIN Color e" & _
                     " ON b.sColorIdx = e.sColorIDx" & _
                  " LEFT JOIN Category f" & _
                     " ON b.sCategID1 = f.sCategrID" & _
            " WHERE a.sStockIDx = b.sStockIDx" & _
               " AND b.cHsSerial = " & strParm(xeYes) & _
            " ORDER BY a.sSerialNo"

   lsCondition = " sSerialNo Like " & strParm("%" & lsSerialNo)
   lsSQL = AddCondition(lsSQL, lsCondition)
   Debug.Print lsSQL
   Set lrs = New ADODB.Recordset
   lrs.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If lrs.RecordCount = 1 Then
      oDriver.LookupValue(0) = lrs("sSerialID")
      oDriver.LoadRecord
      
      txtOther(0).Text = lrs("sDescript")
      txtOther(1).Text = IFNull(lrs("sBrandNme"), "")
      txtOther(2).Text = IFNull(lrs("sModelNme"), "")
      txtOther(3).Text = IFNull(lrs("sColorNme"), "")
      txtOther(14).Text = IFNull(lrs("sCategrNm"), "")
   ElseIf lrs.RecordCount > 1 Then
      lsSearch = KwikSearch(oApp, lsSQL _
                           , "sSerialID»sSerialNo»sBrandNme»sModelNme»sColorNme" _
                           , "Code»SerialNo»Brand»Model»Color" _
                           , "@»@»@»@»@")
      If lsSearch <> "" Then
         lsSelected = Split(lsSearch, "»")
         oDriver.LookupValue(0) = lsSelected(7)
         oDriver.LoadRecord
            
         txtOther(0).Text = lsSelected(3)
         txtOther(1).Text = lsSelected(4)
         txtOther(2).Text = lsSelected(5)
         txtOther(3).Text = lsSelected(6)
         txtOther(14).Text = lsSelected(8)
      End If
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & lsSerialNo _
                       & ", " & lbSearch & " )"
End Sub

Private Sub SearchCustomer(lsCustName As String, lbSearch As Boolean)
   Dim lrs As ADODB.Recordset
   Dim lsSQL As String, lsCondition As String
   Dim lsSearch As String, lsSelected() As String, lsOldProc As String
   
   lsOldProc = "SearchCustomer"
   ''On Error GoTo errProc

   lsSelected = GetSplitedName(lsCustName)
   lsCondition = lsCondition & _
            " b.sLastName LIKE " & strParm(lsSelected(0) & "%") & _
               " AND (b.sFrstName LIKE " & strParm(lsSelected(1) & "%") & _
                  " OR b.sFrstName LIKE " & strParm(lsSelected(1) & lsSelected(2) & "%") & _
                  IIf(lsSelected(2) = Empty, " )", _
                     " OR b.sMiddName LIKE " & strParm(lsSelected(2) & "%") & ")")
   
   lsSQL = "SELECT" _
               & "  a.sSerialID" _
               & ", a.sSerialNo" _
               & ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) as xFullName" _
               & ", e.sBarrCode" _
               & ", e.sDescript" _
               & ", f.sBrandNme" _
               & ", g.sModelNme" _
               & ", h.sColorNme" _
            & " FROM CP_Inventory_Serial a" _
               & ", Client_Master b" _
               & ", CP_SO_Master c" _
               & ", CP_SO_Detail d" _
               & ", CP_Inventory e" _
                  & " LEFT JOIN CP_Brand f" _
                     & " ON e.sBrandIDx = f.sBrandIDx" _
                  & " LEFT JOIN CP_Model g" _
                     & " ON e.sModelIDx = g.sModelIDx" _
                  & " LEFT JOIN Color h" _
                     & " ON e.sColorIDx = h.sColorIDx" _
            & " WHERE a.sSerialID = d.sSerialID" _
               & " AND c.sTransNox = d.sTransNox" _
               & " AND c.sClientID = b.sClientID" _
               & " AND a.sStockIDx = e.sStockIDx" _
            & " ORDER BY xFullName"
   
   lsSQL = AddCondition(lsSQL, lsCondition)
   
   Set lrs = New ADODB.Recordset
   lrs.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If lrs.RecordCount = 1 Then
      oDriver.LookupValue(0) = lrs("sSerialID")
      oDriver.LoadRecord
      txtOther(0).Text = lrs("sDescript")
      txtOther(1).Text = IFNull(lrs("sBrandNme"), "")
      txtOther(2).Text = IFNull(lrs("sModelNme"), "")
      txtOther(3).Text = IFNull(lrs("sColorNme"), "")
   ElseIf lrs.RecordCount > 1 Then
      lsSearch = KwikBrowse(oApp, lrs _
                           , "sSerialID»sSerialNo»xFullName»sBrandNme" _
                           , "Code»SerialNo»Name»Brand" _
                           , "@»@»@»@" _
                           , "a.sSerialID»a.sSerialNo»CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)»f.sBrandNme")
      If lsSearch <> "" Then
         lsSelected = Split(lsSearch, "»")
         oDriver.LookupValue(0) = lsSelected(0)
         oDriver.LoadRecord
         
         txtOther(0).Text = lsSelected(4)
         txtOther(1).Text = lsSelected(5)
         txtOther(2).Text = lsSelected(6)
         txtOther(3).Text = lsSelected(7)
      End If
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & lsCustName _
                       & ", " & lbSearch & " )"
End Sub

Private Sub txtOther_LostFocus(Index As Integer)
   With txtOther(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtOther_Validate(Index As Integer, Cancel As Boolean)
   With txtOther(Index)
      If Trim(.Text) <> "" Then .Text = UCase(Left(.Text, 1)) & Right(.Text, Len(.Text) - 1)
   End With
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

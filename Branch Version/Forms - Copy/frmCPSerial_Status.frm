VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Serial_Status 
   BorderStyle     =   0  'None
   Caption         =   "IMEI Status"
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   735
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1296
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   1
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   1035
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   345
         Width           =   5550
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   4305
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   90
         Width           =   2280
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   1035
         TabIndex        =   1
         Top             =   90
         Width           =   2280
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial ID"
         Height          =   195
         Index           =   5
         Left            =   3525
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&IMEI No."
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
         Index           =   9
         Left            =   120
         TabIndex        =   0
         Top             =   105
         Width           =   930
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3375
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1335
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5953
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   1035
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1860
         Width           =   5550
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1035
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   150
         Width           =   2280
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   1035
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   2955
         Width           =   5535
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   570
         Index           =   5
         Left            =   1035
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "frmCPSerial_Status.frx":0000
         Top             =   2370
         Width           =   2820
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1035
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   2115
         Width           =   5550
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   4
         Left            =   4305
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1350
         Width           =   2280
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmCPSerial_Status.frx":0006
         Left            =   4755
         List            =   "frmCPSerial_Status.frx":0016
         TabIndex        =   27
         Text            =   "Combo1"
         Top             =   2580
         Width           =   1815
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1035
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1605
         Width           =   5550
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1035
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1350
         Width           =   2280
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1035
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1095
         Width           =   2280
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1035
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   840
         Width           =   2280
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1035
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   585
         Width           =   2280
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   1875
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial ID"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   28
         Top             =   2970
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   225
         Index           =   12
         Left            =   120
         TabIndex        =   24
         Top             =   2385
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   225
         Index           =   11
         Left            =   120
         TabIndex        =   22
         Top             =   2130
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   10
         Left            =   3525
         TabIndex        =   16
         Top             =   1365
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   1620
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1125
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Code"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   870
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   240
         Left            =   1080
         Tag             =   "et0;ht2"
         Top             =   180
         Width           =   2280
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI NO."
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   615
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   195
         Index           =   3
         Left            =   3990
         TabIndex        =   26
         Top             =   2640
         Width           =   690
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   3015
      TabIndex        =   30
      Top             =   4935
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
      Picture         =   "frmCPSerial_Status.frx":003F
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   3015
      TabIndex        =   31
      Top             =   4935
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
      Picture         =   "frmCPSerial_Status.frx":07B9
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   3015
      TabIndex        =   32
      Top             =   4935
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
      Picture         =   "frmCPSerial_Status.frx":0F33
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   3780
      TabIndex        =   33
      Top             =   4935
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
      Picture         =   "frmCPSerial_Status.frx":16AD
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   3780
      TabIndex        =   34
      Top             =   4935
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
      Picture         =   "frmCPSerial_Status.frx":1E27
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   5310
      TabIndex        =   36
      Top             =   4935
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
      Picture         =   "frmCPSerial_Status.frx":25A1
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   4545
      TabIndex        =   35
      Top             =   4935
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
      Picture         =   "frmCPSerial_Status.frx":2D1B
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   6075
      TabIndex        =   37
      Top             =   4935
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
      Picture         =   "frmCPSerial_Status.frx":3495
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   6075
      TabIndex        =   38
      Top             =   4935
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
      Picture         =   "frmCPSerial_Status.frx":3C0F
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmCP_Serial_Status"
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
Dim pnctr As Integer
Dim pnindex As Integer
Dim txtfieldGotfocus As Boolean
Dim txtOthersGotfocus As Boolean
Dim pnUserRights As Integer
Dim psUserID As String
Dim psUserName As String

Private Sub cmdButton_Click(Index As Integer)
   Dim lsSearch As String
   Dim lsDel As Integer

   Select Case Index
   Case 0   'cancel
      oDriver.RecordCancelUpdate
      If txtothers(0).Enabled Then txtothers(0).SetFocus
      Combo1.Enabled = False
   Case 1   'browse
       Search_Serial False
   Case 2   'save
      oDriver.RecordSave
'   Case 3   'update
'      If oApp.UserLevel = xeEncoder Or oApp.UserLevel = xeSupervisor Then
'         lsDel = MsgBox("Seek for Approval?", vbQuestion + vbYesNo, "Confirm")
'         If lsDel = vbYes Then
'            If Not GetApproval(oApp, pnUserRights, psUserID, psUserName) Then Exit Sub
'            If pnUserRights < xeManager Then
'               MsgBox "Approving User is Not Authorized!!!", vbCritical, "Warning"
'               Exit Sub
'            Else
'               xrFrame1.Enabled = True
'               Combo1.Enabled = True
'               txtothers(0).Enabled = True
'               oDriver.RecordUpdate
'            End If
'         End If
'      Else
'         xrFrame1.Enabled = True
'         Combo1.Enabled = True
'         txtothers(0).Enabled = True
'         oDriver.RecordUpdate
'      End If
   Case 4   'new
      oDriver.RecordNew
   Case 5   'close
      Unload Me
   Case 6   'delete
      MsgBox "Delete Not Permitted!!!" & vbCrLf & vbCrLf & _
      "Please Notify ROSALYN LAZO DESCALLAR" & vbCrLf & _
      "for Assistance!!!", vbCritical, "Warning"
'      oDriver.RecordDelete
   Case 7   'search
      If txtfieldGotfocus And pnindex = 3 Then
         oDriver.RecordSearch txtfield(pnindex).Text
         txtothers(5).Text = getAddress("Client_Master a", "a.sClientID = '" & oDriver.FieldValue(3) & "'")
      End If
      If txtOthersGotfocus And pnindex = 0 Then SearchBarCode False
   Case 8   'ledger
      If Not pbnewitem Then
         frmCP_Serial_Ledger.BarrCode = oDriver.FieldValue(0)
         frmCP_Serial_Ledger.Show 1
      Else
         MsgBox "Unable to Load IMEI Ledger!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      oDriver_InitValue
      oDriver.DisableTextbox 0
      ClearFields
      bLoaded = True
      Combo1.Enabled = False
   End If

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
   oSkin.ApplySkin

   oDriver.RecQuery = "SELECT" _
                & " sSerialID, " _
                & " sIMEINoxx, " _
                & " sStockIDx, " _
                & " sClientID, " _
                & " sBranchCd, " _
                & " cLocation, " _
                & " cSoldStat, " _
                & " sModified, " _
                & " dModified, " _
                & " vTimeStmp  " _
            & " FROM CP_Serial_Master "
                                                                          
   oDriver.BrowseQuery = "SELECT" _
                  & " a.sIMEINoxx, " _
                  & " b.sBarrCode, " _
                  & " b.sDescript, " _
                  & " c.sBrandNme, " _
                  & " d.sModelNme, " _
                  & " f.sColorNme, " _
                  & " g.sBranchNm  " _
            & " FROM CP_Serial_Master a " _
               & " LEFT JOIN CP_Inventory b " _
                  & " ON a.sStockIDx = b.sStockIDx " _
               & " LEFT JOIN Brand c " _
                  & " ON b.sBrandIDx = c.sBrandIDx " _
               & " LEFT JOIN Model d " _
                  & " ON b.sModelIDx = d.sModelIDx " _
               & " LEFT JOIN Color f " _
                  & " ON b.sColorIdx = f.sColorIDx " _
               & " LEFT JOIN Branch g " _
                  & " ON a.sBranchCd = g.sBranchCd " _

   oDriver.InitRecForm

   oDriver.BrowseColumn(0) = "sIMEINoxx"
   oDriver.BrowseColumn(1) = "sBarrCode"
   oDriver.BrowseColumn(2) = "sBrandNme"
   oDriver.BrowseColumn(3) = "sModelNme"
   oDriver.BrowseColumn(4) = "sDescript"
   oDriver.BrowseColumn(5) = "sColorNme"
   oDriver.BrowseColumn(6) = "sBranchNm"
   
   oDriver.BrowseFTitle(0) = "IMEI No."
   oDriver.BrowseFTitle(1) = "Bar Code"
   oDriver.BrowseFTitle(2) = "Brand"
   oDriver.BrowseFTitle(3) = "Model"
   oDriver.BrowseFTitle(4) = "Description"
   oDriver.BrowseFTitle(5) = "Color"
   oDriver.BrowseFTitle(6) = "Branch"
   
   'Customer
   oDriver.LookupQuery(3) = "SELECT" _
                  & " a.sClientID, " _
                  & " a.sLastName + ', ' + a.sFrstName + ' ' + a.sMiddName xFullName," _
                  & " a.sAddressx + ', ' + b.sTownName as xAddressx" _
               & " FROM Client_Master a " _
                  & " LEFT JOIN TownCity b " _
                     & " ON a.sTownIDxx = b.sTownIDxx " _
               & " ORDER BY slastname, sfrstname, smiddname "

   
   oDriver.LookupReference(3) = "sClientID»a.sLastName + ', ' +" _
                                 & " a.sFrstName + ' ' +" _
                                 & " a.sMiddName" _
                                 & " »a.sAddressx + ', ' +" _
                                 & " b.sTownName"
   oDriver.LookupColumn(3) = "xFullName»xAddressx"
   oDriver.LookupTitle(3) = "Customer Name»Address"
   
   'Branch
   oDriver.LookupQuery(4) = "SELECT" _
                     & " a.sBranchCd, " _
                     & " a.sBranchNm, " _
                     & " a.sAddressx + ' ' + b.sTownName xAddressx " _
               & " FROM Branch a " _
                  & "LEFT JOIN TownCity b " _
                     & " ON a.sTownIdxx = b.sTownIDxx " _
               & " WHERE a.cRecdStat = '" & xeRecStateActive & "'" _
               & " ORDER BY sBranchNm "

   oDriver.LookupReference(4) = "sBranchCd»sBranchNm»xAddressx"
   oDriver.LookupColumn(4) = "sBranchNm»xAddressx"
   oDriver.LookupTitle(4) = "Branch»Address"
   

       
   oDriver.FieldFormat(0) = "@@-@@@@@@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
   
   oDriver.FieldStart = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()
   oDriver.FieldReference(0) = False
   pbnewitem = True
   ClearFields
End Sub

Private Sub ClearFields()
   For pnctr = 0 To txtothers.Count - 1
     txtothers(pnctr).Text = ""
   Next
   Combo1.ListIndex = 2
End Sub

Private Sub SearchBarCode(ByVal SearchValue As Boolean)
Dim lsSQL As String
Dim lsSearch As String
Dim lnctr As Integer
   
   lsSQL = "SELECT" _
          & " a.sBarrcode, " _
          & " b.sBrandNme, " _
          & " c.sModelNme, " _
          & " a.sDescript, " _
          & " d.sColorNme, " _
          & " a.sStockIDx, " _
          & " a.cWdSerial  " _
      & " FROM CP_Inventory a " _
          & " LEFT JOIN Brand b " _
            & " ON a.sBrandIdx = b.sBrandIdx " _
          & " LEFT JOIN Model c " _
            & " ON a.sModelIdx = c.sModelIdx " _
          & " LEFT JOIN Color d " _
            & " ON a.sColorIDx = d.sColorIDx " _
      & " WHERE cWdSerial = 1 " _
         & " AND a.sBarrCode LIKE '" & txtothers(0).Text & "%' " _
      & " ORDER BY sBarrCode"
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   If oRS.RecordCount = 1 Then
      For lnctr = 0 To 4
         txtothers(lnctr).Text = IIf(IsNull(oRS(lnctr)), "", oRS(lnctr))
         txtothers(lnctr).Enabled = False
      Next
      oDriver.FieldValue(2) = oRS(5)
      
   ElseIf oRS.RecordCount > 1 Then
        lsSearch = KwikBrowse(oApp, oRS, _
               "sBarrcode»sBrandNme»sModelNme»sDescript»sColorNme", _
               "Bar Code»Brand»Model»Description»Color")
                        
        If lsSearch <> "" Then
            psSelected = Split(lsSearch, "»")
            For lnctr = 0 To 4
               txtothers(lnctr).Text = IIf(IsNull(psSelected(lnctr)), "", psSelected(lnctr))
               txtothers(lnctr).Enabled = False
            Next
            oDriver.FieldValue(2) = psSelected(5)
        End If
   Else
      For lnctr = 0 To 4
         txtothers(lnctr).Text = ""
      Next
      MsgBox "Record Not Existing!!!", vbInformation, "Information"
   End If
   
   For pnctr = 1 To txtothers.Count - 1
     txtothers(pnctr).Enabled = False
   Next
   txtothers(0).Enabled = True

   Set oRS = Nothing

End Sub

Private Sub oDriver_LoadOtherData()
Dim lsSQL As String
Dim lnctr As Integer

pbnewitem = False

   lsSQL = "SELECT" _
               & " a.sBarrCode, " _
               & " b.sBrandNme, " _
               & " c.sModelNme, " _
               & " a.sDescript, " _
               & " d.sColorNme  " _
         & " FROM CP_Inventory a " _
            & " LEFT JOIN Brand b " _
               & " ON a.sBrandIDx = b.sBrandIDx " _
            & " LEFT JOIN Model c " _
               & " ON a.sModelIDx = c.sModelIDx " _
            & " LEFT JOIN Color d " _
               & " ON a.sColorIDx = d.sColorIDx " _
         & " Where a.sStockIDx = '" & oDriver.FieldValue(2) & "' "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

   For lnctr = 0 To 4
      txtothers(lnctr).Text = IIf(IsNull(oRS(lnctr)), "", oRS(lnctr))
      txtothers(lnctr).Enabled = False
   Next
   Set oRS = Nothing

   lsSQL = "SELECT" _
               & " a.sIMEINoxx, " _
               & " b.sLastName + ', ' + b.sFrstName + ' ' + b.sMiddName FullName," _
               & " a.sSerialID, " _
               & " a.cLocation, " _
               & " a.sclientID, " _
               & " h.sSupplyNm  " _
         & " FROM CP_Serial_Master a " _
               & " LEFT JOIN Client_Master b " _
                  & " ON a.sClientID = b.sClientID " _
               & " LEFT JOIN PO_Receiving_Serial e " _
                  & " ON a.sSerialID = e.sSerialID " _
               & " LEFT JOIN PO_Receiving_Master g " _
                  & " ON e.sTransNox = g.sTransNox " _
               & " LEFT JOIN Supplier h " _
                  & " ON g.sClientID = h.sSupplyID " _
         & " WHERE a.sSerialID =  '" & oDriver.FieldValue(0) & "' "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

   For pnctr = 1 To 3
      txtothers(pnctr + 5).Text = IIf(IsNull(oRS(pnctr - 1)), "", oRS(pnctr - 1))
   Next
   txtothers(9).Text = IIf(IsNull(oRS("sSupplyNm")), "", oRS("sSupplyNm"))
   txtothers(5).Enabled = False
   txtothers(9).Enabled = False
   txtfield(3).Text = IIf(IsNull(oRS(1)), "", oRS(1))
   oDriver.FieldValue(3) = IIf(IsNull(oRS(4)), "", oRS(4))

   Select Case oRS(3)
      Case 0
         Combo1.ListIndex = 2
      Case 1
         Combo1.ListIndex = 0
      Case 2
         Combo1.ListIndex = 1
      Case 3
         Combo1.ListIndex = 3
   End Select
   txtothers(5).Text = getAddress("Client_Master a", "a.sClientID = '" & oDriver.FieldValue(3) & "'")
   Set oRS = Nothing
   
End Sub

Private Sub oDriver_Save(Saved As Boolean)
   Saved = False
End Sub

Private Sub oDriver_SaveComplete()
   ClearFields
   Combo1.Enabled = False
   txtothers(6).Enabled = True
   If txtothers(6).Enabled Then txtothers(6).SetFocus
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid IMEI No. Detected!!!", vbCritical, "Warning"
      txtfield(1).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(2) = "" Then
      MsgBox "Invalid Bar Code Detected!!!", vbCritical, "Warning"
      txtothers(0).SetFocus
      Cancel = True
   Else
      Select Case Combo1.ListIndex
         Case 0
            oDriver.FieldValue(5) = 1
         Case 1
            oDriver.FieldValue(5) = 2
         Case 2
            oDriver.FieldValue(5) = 0
         Case 3
            oDriver.FieldValue(5) = 3
      End Select
   End If
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfieldGotfocus = True
   pnindex = Index
   txtfield(Index).BackColor = &HE1FEFF
   txtfield(Index).Text = TitleCase(txtfield(Index).Text)
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Or KeyCode = vbKeyF3 Then
      If Index = 3 Then
         oDriver.RecordSearch txtfield(Index).Text
         txtothers(5).Text = getAddress("Client_Master a", "a.sClientID = '" & oDriver.FieldValue(3) & "'")
         If txtfield(Index).Text <> "" Then SetNextFocus
      End If
      KeyCode = 0
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &H80000005
   txtfieldGotfocus = False
End Sub

Private Sub txtothers_GotFocus(Index As Integer)
   txtOthersGotfocus = True
   If txtothers(Index).Text <> "" Then
      txtothers(Index).SelStart = 0
      txtothers(Index).SelLength = Len(txtothers(Index).Text)
   End If
   txtfieldGotfocus = False
   txtothers(Index).BackColor = &HE1FEFF
   pnindex = Index
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
         Call Modified("CP_Serial_Master", "sSerialID = '" & oDriver.FieldValue(0) & "' ")
   End Select
End Sub

Private Sub oDriver_Delete(Deleted As Boolean)
   Deleted = True
End Sub

Private Sub txtothers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If pnindex = 0 Then
         SearchBarCode False
      ElseIf Index = 6 Then
         Search_Serial False
         txtothers(Index).SelStart = 0
         txtothers(Index).SelLength = Len(txtothers(Index).Text)
         txtothers(Index).SetFocus
      End If
   End If
End Sub

Private Sub txtothers_LostFocus(Index As Integer)
   txtothers(Index).BackColor = &H80000005
   txtOthersGotfocus = False
End Sub

Private Sub Search_Serial(ByVal SearchValue As Boolean)
Dim orig As String
Dim lsSQL As String
Dim lsCondition As String
   orig = oDriver.BrowseQuery
   Select Case pnindex
   Case 6
         If SearchValue Then
            lsCondition = " a.sIMEINoxx = '" & txtothers(pnindex).Text & "'"
         Else
            lsCondition = " a.sIMEINoxx like '%" & txtothers(pnindex).Text & "%'"
         End If
         lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
         oDriver.BrowseQuery = lsSQL
   Case 7
         If SearchValue Then
            lsCondition = " f.sLastName + ', ' + f.sFrstName + ' ' + f.sMiddName = '" & txtfield(pnindex).Text & "'"
         Else
            lsCondition = " f.sLastName + ', ' + f.sFrstName + ' ' + f.sMiddName like '" & txtfield(pnindex).Text & "%'"
         End If
         lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
         oDriver.BrowseQuery = lsSQL
   Case 8
         If SearchValue Then
            lsCondition = " a.sSerialID = '" & txtothers(pnindex).Text & "'"
         Else
            lsCondition = " a.sSerialID like '%" & txtothers(pnindex).Text & "%'"
         End If
         lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
         oDriver.BrowseQuery = lsSQL
   End Select
   oDriver.BrowseRecord
   oDriver.BrowseQuery = orig
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)

If txtfield(1).Text <> "" Then oDriver.FieldValue(1) = txtfield(1).Text
Cancel = Not oDriver.ValidateField(Index)
End Sub

Private Sub txtothers_Validate(Index As Integer, Cancel As Boolean)
   If Index = 8 Then
      If txtothers(Index).Text <> "" Then
         txtothers(Index).Text = Format(txtothers(Index), "@@-@@@@@@@@")
      End If
   End If
   txtothers(Index).Text = TitleCase(txtothers(Index).Text)
End Sub



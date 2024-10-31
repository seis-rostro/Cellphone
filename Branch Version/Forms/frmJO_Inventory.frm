VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmJO_Inventory 
   BorderStyle     =   0  'None
   Caption         =   "Inventory"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   LinkTopic       =   "Form3"
   ScaleHeight     =   3000
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame2 
      Height          =   2370
      Left            =   30
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   4180
      BackColor       =   12632256
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1095
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   900
         Width           =   4650
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1095
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   195
         Width           =   2820
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   8
         Left            =   1095
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1920
         Width           =   2820
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   1095
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1665
         Width           =   2820
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   1095
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1410
         Width           =   2820
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   1095
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1155
         Width           =   2820
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   13
         Top             =   915
         Width           =   795
      End
      Begin VB.Shape Shape2 
         Height          =   510
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
         TabIndex        =   9
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Made"
         Height          =   195
         Index           =   6
         Left            =   195
         TabIndex        =   8
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   7
         Top             =   1410
         Width           =   435
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   6
         Top             =   1170
         Width           =   420
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   270
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
         TabIndex        =   5
         Top             =   195
         Width           =   660
      End
      Begin VB.Shape Shape5 
         Height          =   1440
         Left            =   105
         Top             =   810
         Width           =   5715
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   6285
      TabIndex        =   10
      Top             =   1380
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      SizeCW          =   1
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
      Picture         =   "frmJO_Inventory.frx":0000
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   6285
      TabIndex        =   11
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmJO_Inventory.frx":077A
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   6285
      TabIndex        =   12
      Top             =   960
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      SizeCW          =   2
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
      Picture         =   "frmJO_Inventory.frx":0EF4
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmJO_Inventory"
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
Dim psValue(2) As String
Dim txtfieldGotfocus As Boolean
Dim pbEnabled As Boolean


Private Sub cmdButton_Click(Index As Integer)
Dim orig As String
Dim lsSQL As String
Dim lsCondition As String
   
   Select Case Index
      Case 0 'save
         oDriver.RecordSave
      Case 1 'cancel
         Unload Me
      Case 2 'Search
         If pnindex > 4 And pnindex < 9 Then
            orig = oDriver.LookupQuery(6)
            If pnindex = 6 Then
               lsCondition = " a.sBrandIDx = '" & oDriver.FieldValue(5) & "'"
               lsSQL = AddCondition(oDriver.LookupQuery(6), lsCondition)
               oDriver.LookupQuery(6) = lsSQL
            End If
            oDriver.RecordSearch txtfield(Index).Text
            oDriver.LookupQuery(6) = orig
         End If

   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      oDriver.RecordNew
      oDriver_InitValue
      oDriver.DisableTextbox 0
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
oSkin.ApplySkin xeFormTransDetail

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

      oDriver.InitRecForm


      'Brand
      oDriver.LookupQuery(5) = "SELECT" _
                        & " sBrandIDx, " _
                        & " sBrandNme " _
                     & " FROM Brand " _
                     & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY sBrandNme"
      
      oDriver.LookupReference(5) = "sBrandIDx製BrandNme"
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
      
      oDriver.LookupReference(6) = "sModelIDx製ModelNme製BrandNme"
      oDriver.LookupColumn(6) = "sModelNme製BrandNme"
      oDriver.LookupTitle(6) = "Model翡rand"

      'Country
      oDriver.LookupQuery(7) = "SELECT" _
                        & " sMadeIDxx, " _
                        & " sMadeName " _
                     & " FROM Made " _
                     & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY sMadeName "
      
      oDriver.LookupReference(7) = "sMadeIDxx製MadeName"
      oDriver.LookupColumn(7) = "sMadeIDxx製MadeName"
      oDriver.LookupTitle(7) = "Country ID翟ountry"
      
      'Color
      oDriver.LookupQuery(8) = "SELECT" _
                        & " sColorIDx, " _
                        & " sColorNme " _
                     & " FROM Color" _
                     & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                     & " ORDER BY sColorNme"
      
      oDriver.LookupReference(8) = "sColorIDx製ColorNme"
      oDriver.LookupColumn(8) = "sColorNme"
      oDriver.LookupTitle(8) = "Color"


      oDriver.FieldStart = 0
      oDriver.FieldFormat(0) = ">"

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_InitValue()
Dim lnCtr As Integer

oDriver.FieldReference(1) = True
oDriver.FieldValue(1) = getNextCode("CP_Inventory", "sStockIDx", True, oApp.Connection, True, oApp.BranchCode)
  
txtfield(0).Tag = oDriver.FieldValue(1)
      
    oDriver.FieldValue(0) = NewBarrCode
    txtfield(0).Text = oDriver.FieldValue(0)
    
    For lnCtr = 3 To 19
      Select Case lnCtr
         Case 3
            oDriver.FieldValue(3) = "01001"
         Case 4, 19
            oDriver.FieldValue(lnCtr) = ""
         Case 10
            oDriver.FieldValue(lnCtr) = Date
         Case 9, 11, 12, 14 To 16
            oDriver.FieldValue(lnCtr) = 0
         Case 13, 17
            oDriver.FieldValue(lnCtr) = 1
      End Select
    Next
    
    pbnewitem = True
    pbEnabled = True
    
End Sub

Private Sub oDriver_Save(Saved As Boolean)
   Saved = False
End Sub

Private Sub oDriver_SaveComplete()
   Unload Me
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
Dim lnCtr As Integer

   If oDriver.FieldValue(0) = "" Then
      MsgBox "Invalid BarrCode Detected!!!", vbCritical, "Warning"
      txtfield(0).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(5) = "" Then
      MsgBox "Invalid Brand Detected!!!", vbCritical, "Warning"
      txtfield(5).SetFocus
      Cancel = True
   ElseIf oDriver.FieldValue(6) = "" Then
      MsgBox "Invalid Model Detected!!!", vbCritical, "Warning"
      txtfield(6).SetFocus
      Cancel = True
   End If
       
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfield(Index).BackColor = &HE1FEFF
   txtfieldGotfocus = True
   pnindex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim lsSQL As String
Dim lsCondition As String
Dim orig As String

   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index > 4 And Index < 9 Then
         orig = oDriver.LookupQuery(6)
         If Index = 6 Then
            lsCondition = " a.sBrandIDx = '" & oDriver.FieldValue(5) & "'"
            lsSQL = AddCondition(oDriver.LookupQuery(6), lsCondition)
            oDriver.LookupQuery(6) = lsSQL
         End If
         oDriver.RecordSearch txtfield(Index).Text
         oDriver.LookupQuery(6) = orig
      End If
      If txtfield(Index).Text <> "" Then SetNextFocus
      KeyCode = 0
   End If
   
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
       Case 0
           txtfield(Index).Text = Format(txtfield(Index).Text, ">")
   End Select
   txtfield(Index).Text = TitleCase(txtfield(Index).Text)
   Cancel = Not oDriver.ValidateField(Index)
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

Function NewBarrCode() As String
   Dim lrs As Recordset
   Dim lsSQL As String
   Dim lnCtr As Long
   
   lsSQL = "SELECT TOP 1" & _
            " sBarrCode" & _
            " FROM CP_Inventory " & _
            " WHERE sBarrCode LIKE " & strParm(Format(Date, "yy") & "-GMC-%") & _
            " ORDER BY sBarrCode DESC"
   
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
   NewBarrCode = Format(Date, "yy") & "-GMC-" & Format(lnCtr, "000000")
   
   Set lrs = Nothing
End Function


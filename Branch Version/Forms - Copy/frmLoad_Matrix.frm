VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLoad_Matrix 
   BorderStyle     =   0  'None
   Caption         =   "Load Matrix Maintenance"
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3810
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   3525
      TabIndex        =   18
      Top             =   3030
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
      Picture         =   "frmLoad_Matrix.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   2760
      TabIndex        =   17
      Top             =   3030
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
      Picture         =   "frmLoad_Matrix.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   1230
      TabIndex        =   13
      Top             =   3030
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
      Picture         =   "frmLoad_Matrix.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   1230
      TabIndex        =   14
      Top             =   3030
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
      Picture         =   "frmLoad_Matrix.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   465
      TabIndex        =   12
      Top             =   3030
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
      Picture         =   "frmLoad_Matrix.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   3525
      TabIndex        =   19
      Top             =   3030
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
      Picture         =   "frmLoad_Matrix.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   1995
      TabIndex        =   15
      Top             =   3030
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
      Picture         =   "frmLoad_Matrix.frx":2CDC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   1995
      TabIndex        =   16
      Top             =   3030
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
      Picture         =   "frmLoad_Matrix.frx":3456
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2250
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   3969
      BorderStyle     =   1
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1485
         TabIndex        =   5
         Top             =   900
         Width           =   2505
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1485
         TabIndex        =   3
         Top             =   645
         Width           =   2505
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1485
         TabIndex        =   11
         Top             =   1815
         Width           =   1425
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   1485
         TabIndex        =   9
         Top             =   1560
         Width           =   1425
      End
      Begin VB.TextBox txtfield 
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
         Height          =   240
         Index           =   0
         Left            =   1485
         TabIndex        =   1
         Top             =   165
         Width           =   900
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1485
         TabIndex        =   7
         Top             =   1305
         Width           =   2505
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   4
         Top             =   915
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Code"
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   2
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   10
         Top             =   1845
         Width           =   1770
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Price"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   8
         Top             =   1590
         Width           =   1770
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Matrix ID"
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
         Index           =   2
         Left            =   165
         TabIndex        =   0
         Top             =   195
         Width           =   990
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Matrix Name"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   1335
         Width           =   1770
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   240
         Left            =   1530
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmLoad_Matrix"
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
Private pnindex As Integer
Dim psSelected() As String



Private Sub cmdButton_Click(Index As Integer)
    Select Case Index
        Case 0 'Cancel
            oDriver.RecordCancelUpdate
            txtothers(0).Text = ""
            txtothers(1).Text = ""
        Case 1 'Browse
            oDriver.BrowseRecord
        Case 2 'Save
            oDriver.RecordSave
        Case 3 'Update
            oDriver.RecordUpdate
        Case 4 'New
            oDriver.RecordNew
            txtothers(0).SetFocus
        Case 5 'Close
            Unload Me
        Case 6 'Delete
            MsgBox "Delete Not Permitted!!!" & vbCrLf & vbCrLf & _
            "Please Notify ROSALYN LAZO DESCALLAR" & vbCrLf & _
            "for Assistance!!!", vbCritical, "Warning"
'            oDriver.RecordDelete
        Case 7 'Search
            If pnindex = 0 Then SearchBarCode False
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
   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
      
   oDriver.RecQuery = "SELECT " _
                        & " sMatrixID, " _
                        & " sMatrixNm, " _
                        & " nAmountxx, " _
                        & " nSelPrice, " _
                        & " sStockIDx,  " _
                        & " sModified, " _
                        & " dModified, " _
                        & " vTimeStmp " _
                     & " FROM ELoad_Matrix"
   
   oDriver.BrowseQuery = "SELECT " _
                        & " a.sMatrixID, " _
                        & " b.sDescript, " _
                        & " a.sMatrixNm, " _
                        & " a.nAmountxx, " _
                        & " a.nSelPrice " _
                     & " FROM ELoad_Matrix a " _
                        & " LEFT JOIN CP_Inventory b " _
                           & " ON a.sStockIDx = b.sStockIDx " _
                     & " ORDER BY sMatrixID "
                     
   
   oDriver.InitRecForm
   
   oDriver.BrowseFTitle(0) = "Matrix ID"
   oDriver.BrowseFTitle(1) = "Network"
   oDriver.BrowseFTitle(2) = "Description"
   oDriver.BrowseFTitle(3) = "Purchase Price"
   oDriver.BrowseFTitle(4) = "Selling Price"
   
   
   oDriver.BrowseFFormat(0) = "@@-@@@"
   
   oDriver.FieldFormat(0) = "@@-@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
   oDriver.FieldStart = 1
   
   oDriver.FieldValue(2) = "#,##0.00"
   oDriver.FieldValue(3) = "#,##0.00"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()
   If oDriver.SetValue(0, getNextCode("ELoad_Matrix", "sMatrixID", False, oApp.Connection, True, oApp.BranchCode)) = False Then Exit Sub
   oDriver.FieldReference(0) = True
   
   Clear_Others
   txtothers(0).SetFocus
   txtfield(2).Text = "0.00"
   txtfield(3).Text = "0.00"
   
End Sub

Private Sub oDriver_LoadOtherData()
    Dim lsSQL As String
    Dim lnCtr As Integer
    
    If oRS.State = adStateOpen Then oRS.Close
    
    lsSQL = "SELECT" _
            & " sStockIDx, " _
            & " sBarrCode, " _
            & " sDescript  " _
        & " FROM CP_Inventory " _
        & " Where sStockIDx = '" & oDriver.FieldValue(4) & "' " _
    
    
    oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
    
    
    If Not oRS.EOF Then
        txtothers(0).Text = oRS("sBarrCode")
        txtothers(1).Text = oRS("sDescript")
    Else
        txtothers(0).Text = ""
        txtothers(1).Text = ""
    End If
    
    txtothers(1).Locked = True

End Sub

Private Sub oDriver_SaveComplete()
   txtothers(0).SetFocus
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)

   If oDriver.FieldValue(1) = "" Then
      MsgBox "Invalid Netwrok detected!!!", vbCritical, "Warning"
      txtfield(1).SetFocus
      Cancel = True
   ElseIf CDbl(txtfield(2).Text) = 0# Then
      MsgBox "Invalid Amount detected!!!", vbCritical, "Warning"
      txtfield(2).SetFocus
      Cancel = True
   ElseIf CDbl(txtfield(3).Text) = 0# Then
      MsgBox "Invalid Selling Price detected!!!", vbCritical, "Warning"
      txtfield(3).SetFocus
      Cancel = True
   End If
   
   
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfield(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   Select Case Index
      Case 2, 3
         If Not IsNumeric(txtfield(Index).Text) Then txtfield(Index).Text = 0#
         txtfield(Index).Text = Format(txtfield(Index).Text, "#,##0.00")
   End Select
   txtfield(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 2, 3
         If Not IsNumeric(txtfield(Index).Text) Then txtfield(Index).Text = 0#
         If Index = 3 Then
            If CDbl(txtfield(3).Text) < CDbl(txtfield(2).Text) Then
               MsgBox "Invalid Selling Price!!!", vbCritical, "Warning"
               txtfield(3).SetFocus
            End If
         End If
         txtfield(Index).Text = Format(txtfield(Index).Text, "#,##0.00")
   End Select
   txtfield(Index).Text = TitleCase(txtfield(Index).Text)
   Cancel = Not oDriver.ValidateField(Index)
End Sub

Private Sub SearchBarCode(ByVal SearchValue As Boolean)
   Dim lsSearch As String
   Dim lrs As ADODB.Recordset
   Dim lsSQL As String
   
   Set lrs = New ADODB.Recordset
   
   lsSQL = "SELECT" _
            & " sStockIDx, " _
            & " sBarrcode, " _
            & " sDescript  " _
         & " FROM CP_Inventory " _
         & " WHERE cRecdStat = 1 " _
            & " AND cCellLoad = 1 " _
               
   If SearchValue Then
      lsSQL = lsSQL & " AND sBarrCode = '" & txtothers(0).Text & "'"
   Else
      lsSQL = lsSQL & " AND sBarrCode LIKE '" & txtothers(0).Text & "%' "
   End If
            
   lsSQL = lsSQL & " ORDER BY sBarrCode"
   
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      
   If lrs.RecordCount = 1 Then
      oDriver.FieldValue(4) = lrs("sStockIDx")
      txtothers(0).Text = lrs("sBarrCode")
      txtothers(1).Text = lrs("sDescript")
      
      ElseIf lrs.RecordCount > 1 Then
           lsSearch = KwikBrowse(oApp, lrs, _
                             "sBarrcode»" _
                           & "sDescript", _
                             "Bar Code»" _
                           & "Description")
                           
           If lsSearch <> "" Then
               psSelected = Split(lsSearch, "»")
               oDriver.FieldValue(4) = psSelected(0)
               txtothers(0).Text = psSelected(1)
               txtothers(1).Text = psSelected(2)
           End If
   Else
      txtothers(0).Text = ""
      txtothers(1).Text = ""
   End If
   
   Set lrs = Nothing

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
      Call Modified("ELoad_Matrix", "sMatrixID = '" & oDriver.FieldValue(0) & "' ")
   End Select
End Sub

Private Sub txtothers_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   pnindex = Index
   txtothers(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtothers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 0 Then
         SearchBarCode False
         txtothers(Index).Tag = txtothers(Index).Text
         If txtothers(Index).Text <> "" Then SetNextFocus
      End If
      KeyCode = 0
   End If
End Sub

Private Sub Clear_Others()
txtothers(0).Text = ""
txtothers(1).Text = ""
txtothers(1).Enabled = False
End Sub

Private Sub txtothers_LostFocus(Index As Integer)
   txtothers(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtothers_Validate(Index As Integer, Cancel As Boolean)
   txtothers(Index).Text = TitleCase(txtothers(Index).Text)
End Sub

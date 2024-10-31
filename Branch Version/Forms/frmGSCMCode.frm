VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmGSCMCode 
   BorderStyle     =   0  'None
   Caption         =   "Samsung GSCM Code Maintenance"
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3810
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   2265
      Left            =   75
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3995
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1320
         TabIndex        =   3
         Tag             =   "Barcode or Description"
         Text            =   "Text1"
         Top             =   510
         Width           =   3225
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1170
         Width           =   3225
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Barcode or Description"
         Text            =   "Text1"
         Top             =   180
         Width           =   3225
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   840
         Width           =   3225
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
         Left            =   1320
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1650
         Width           =   2340
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   2
         Top             =   570
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "BCode/ Desc."
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   0
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   6
         Top             =   1230
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1395
         Tag             =   "et0;ht2"
         Top             =   1740
         Width           =   2340
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GSCM Code"
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
         Left            =   165
         TabIndex        =   8
         Top             =   1710
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   4
         Top             =   900
         Width           =   1035
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   4020
      TabIndex        =   11
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
      Picture         =   "frmGSCMCode.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   2490
      TabIndex        =   10
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
      Picture         =   "frmGSCMCode.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   3255
      TabIndex        =   12
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
      Picture         =   "frmGSCMCode.frx":0EF4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   1725
      TabIndex        =   13
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
      Picture         =   "frmGSCMCode.frx":166E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   960
      TabIndex        =   14
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
      Picture         =   "frmGSCMCode.frx":1DE8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   4020
      TabIndex        =   15
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
      Picture         =   "frmGSCMCode.frx":2562
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   2490
      TabIndex        =   16
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
      Picture         =   "frmGSCMCode.frx":2CDC
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmGSCMCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmGSCMCode"
Private Const pxeBRAND = "C001006"

Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer
Dim poRS As Recordset
Dim psSQL_Master As String
Dim psBarCode As String
Dim psModelCde As String
Dim psModel As String
Dim psColor As String

Dim pnEditMode As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   'On Error GoTo errProc
   
   Select Case Index
   Case 0 'oDriver.RecordCancelUpdate
      ClearFields
      pnEditMode = xeModeUnknown
   Case 1 'oDriver.BrowseRecord
      If Not BrowseRecord Then
         MsgBox "No record to load.", vbInformation, "Warning"
         GoTo endProc
      End If
   Case 2 'oDriver.RecordSave
      SaveRecord
   Case 3 'oDriver.RecordUpdate
      If pnEditMode = xeModeReady Then
         pnEditMode = xeModeUpdate
      End If
   Case 4 'oDriver.RecordNew
      NewRecord
   Case 5
      Unload Me
      GoTo endProc
   Case 6 'oDriver.RecordDelete
      GoTo endProc
   End Select

   InitButton
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & " )", True
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   'On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If Not bLoaded Then
      bLoaded = True
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   'On Error GoTo errProc
   
   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
   
   psSQL_Master = "SELECT" _
                     & "  a.sModelIDx" _
                     & ", a.sColorIDx" _
                     & ", a.sGSCMCode" _
                     & ", a.cRecdStat" _
                     & ", a.sModified" _
                     & ", a.dModified" _
                     & ", b.sModelNme" _
                     & ", c.sColorNme" _
                     & ", d.sBarrCode" _
                     & ", d.sDescript" _
                     & ", b.sModelCde" _
                  & " FROM CP_Model_GSCM a" _
                     & ", CP_Model b" _
                     & ", Color c" _
                     & ", CP_Inventory d" _
                  & " WHERE a.sModelIDx = b.sModelIDx" _
                     & " AND a.sColorIDx = c.sColorIDx" _
                     & " AND d.sModelIDx = a.sModelIDx" _
                     & " AND d.sColorIDx = a.sColorIDx"
                     
   txtField(0).MaxLength = 16
   ClearFields
   
   pnEditMode = xeModeUnknown
   InitButton
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub


Private Sub InitButton()
   Dim lbShow As Boolean
   
   lbShow = pnEditMode = xeModeAddNew Or pnEditMode = xeModeUpdate
   
   xrFrame1.Enabled = lbShow
   cmdButton(0).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow
   
   txtField(3).Enabled = pnEditMode = xeModeAddNew
   
   If lbShow Then txtField(3).SetFocus
End Sub


Private Function BrowseRecord() As Boolean
   Dim lsSQL As String
   Dim lasSplit() As String
   
   BrowseRecord = False
   lsSQL = psSQL_Master & " GROUP BY a.sModelIDx, a.sColorIDx ORDER BY b.sModelNme, c.sColorNme"
   Debug.Print lsSQL
   lsSQL = KwikSearch(oApp, _
                     lsSQL, _
                     "sGSCMCode»sModelNme»sColorNme", _
                     "Code»Model»Color", _
                     "@»@»@", _
                     "a.GSCMCode»b.sModelNme»c.sColorNme")
                     
   If lsSQL = "" Then GoTo endProc
   lasSplit = Split(lsSQL, "»")
   
   BrowseRecord = OpenRecord(lasSplit(0), lasSplit(1), lasSplit(2))
endProc:
   Exit Function
End Function

Private Function NewRecord()
   Set poRS = New Recordset
   poRS.Open AddCondition(psSQL_Master, "0=1"), oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set poRS.ActiveConnection = Nothing
   
   poRS.AddNew
   poRS("sModelIDx") = ""
   poRS("sColorIDx") = ""
   poRS("sGSCMCode") = ""
   poRS("cRecdStat") = xeRecStateActive
   ClearFields
   
   pnEditMode = xeModeAddNew
   NewRecord = True
End Function

Private Function SaveRecord()
   Dim lsSQL As String
   Dim lors As Recordset

   SaveRecord = False
   If pnEditMode <> xeModeAddNew And pnEditMode <> xeModeUpdate Then GoTo endProc
   
   If poRS("sModelIDx") = "" Or _
      poRS("sColorIDx") = "" Or _
      poRS("sGSCMCode") = "" Then
      
      MsgBox "Invalid field value detected. Please verify your entry.", vbCritical, "Warning"
      txtField(0).SetFocus
      GoTo endProc
   End If
   
   'check if the entry is existing
   lsSQL = AddCondition(psSQL_Master, "a.sModelIDx = " & strParm(poRS("sModelIDx")) & _
                                    " AND a.sColorIDx = " & strParm(poRS("sColorIDx")))
   
   Set lors = New Recordset
   lors.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set lors.ActiveConnection = Nothing
   
   If Not lors.EOF Then
      If MsgBox("GSCM information of this model already encoded." & vbCrLf & _
                  "Do you want to update the current record?", vbQuestion + vbYesNo, "Warning") = vbYes Then
         pnEditMode = xeModeUpdate
      Else
         GoTo endProc
      End If
   End If
   
   Set lors = Nothing
   
   If pnEditMode = xeModeAddNew Then
      lsSQL = ADO2SQL(poRS, _
                        "CP_Model_GSCM", , _
                        oApp.UserID, _
                        oApp.ServerDate, _
                        "sModelNme»sColorNme»sBarrCode»sDescript»sModelCde")
   Else
      lsSQL = ADO2SQL(poRS, _
                        "CP_Model_GSCM", _
                        "sModelIDx = " & strParm(poRS("sModelIDx")) & _
                           " AND sColorIDx = " & strParm(poRS("sColorIDx")), _
                        oApp.UserID, _
                        oApp.ServerDate, _
                        "sModelNme»sColorNme»sBarrCode»sDescript»sModelCde")
   End If
   
   If oApp.Execute(lsSQL, "CP_Model_GSCM", oApp.BranchCode) <= 0 Then
      MsgBox "Unable to modify GSCM information.", vbCritical, "Warning"
      GoTo endProc
   End If
   
   MsgBox "Record saved successfully.", vbInformation, "Warning"
   pnEditMode = xeModeUnknown
   SaveRecord = True
endProc:
   Exit Function
End Function

Private Function OpenRecord(ByVal fsModelIDx As String, _
                              ByVal fsColorIDx As String, _
                              ByVal fsGSCMCode As String) As Boolean
                
   Dim lsSQL As String
   
   OpenRecord = False
   lsSQL = AddCondition(psSQL_Master, "a.sModelIDx = " & strParm(fsModelIDx) & _
                                    " AND a.sColorIDx = " & strParm(fsColorIDx) & _
                                    " AND a.sGSCMCOde = " & strParm(fsGSCMCode))
   
   Set poRS = New Recordset
   poRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
   Set poRS.ActiveConnection = Nothing
   
   If poRS.EOF Then GoTo endProc
   
   txtField(0) = poRS("sGSCMCode")
   txtField(1) = poRS("sModelNme")
   txtField(2) = poRS("sColorNme")
   txtField(3) = poRS("sBarrCode")
   txtField(4) = poRS("sModelCde")
   
   psBarCode = poRS("sBarrCode")
   psModel = poRS("sModelNme")
   psColor = poRS("sColorNme")
   psModelCde = poRS("sModelCde")
   
   pnEditMode = xeModeReady
   OpenRecord = True
endProc:
   Exit Function
End Function

Private Function getCPInventory(ByVal fsValue As String, _
                           ByVal fbByCode As Boolean) As Boolean
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lasSplit() As String
   
   getCPInventory = False
   
   If poRS.EOF Then
      MsgBox "Unable to load info for empty record.", vbCritical, "Warning"
      GoTo endProc
   End If
   
   lsSQL = "SELECT" & _
               "  a.sBarrCode" & _
               ", a.sDescript" & _
               ", b.sModelNme" & _
               ", c.sColorNme" & _
               ", a.sModelIDx" & _
               ", a.sColorIDx" & _
               ", b.sModelCde" & _
            " FROM CP_Inventory a" & _
               ", CP_Model b" & _
               ", Color c" & _
            " WHERE a.sCategID1 = 'C001001'" & _
               " AND a.sBrandIDx = " & strParm(pxeBRAND) & _
               " AND a.sModelIDx = b.sModelIDx" & _
               " AND a.sColorIDx = c.sColorIDx" & _
            " ORDER BY b.sModelNme, a.sDescript, c.sColorNme"
            
   If fbByCode Then
      lsSQL = AddCondition(lsSQL, "a.sBarrCode = " & strParm(fsValue))
   Else
      lsSQL = AddCondition(lsSQL, "b.sModelNme LIKE " & strParm(fsValue & "%"))
   End If
   
   Set lors = New Recordset
   Debug.Print lsSQL
   lors.Open lsSQL, oApp.Connection, , , adCmdText
   Set lors.ActiveConnection = Nothing
   
   psBarCode = ""
   psModel = ""
   psColor = ""
   psModelCde = ""
   poRS("sModelIDx") = ""
   poRS("sColorIDx") = ""
   
   If lors.RecordCount = 0 Then
      txtField(3).SetFocus
      GoTo endProc
   ElseIf lors.RecordCount = 1 Then
      psBarCode = lors("sBarrCode")
      psModel = lors("sModelNme")
      psColor = lors("sColorNme")
      poRS("sModelIDx") = lors("sModelIDx")
      poRS("sColorIDx") = lors("sColorIDx")
   Else
      lsSQL = KwikBrowse(oApp, _
                           lors, _
                           "sDescript»sModelNme»sColorNme»sModelCde", _
                           "Description»Model»Color»Code")
      
      If lsSQL = "" Then GoTo endProc
      
      lasSplit = Split(lsSQL, "»")
      
      psBarCode = lasSplit(0)
      psModel = lasSplit(2)
      psColor = lasSplit(3)
      psModelCde = lasSplit(6)
      poRS("sModelIDx") = lasSplit(4)
      poRS("sColorIDx") = lasSplit(5)
   End If
   
   getCPInventory = True
endProc:
   Exit Function
End Function

Private Sub ClearFields()
   txtField(0) = ""
   txtField(1) = ""
   txtField(2) = ""
   txtField(3) = ""
   txtField(4) = ""
   
   psModel = ""
   psColor = ""
   psBarCode = ""
   psModelCde = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set poRS = Nothing
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("HT1")
   End With
   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 3 Then
      If KeyCode = vbKeyReturn Then
         Call getCPInventory(txtField(3), True)
         txtField(3) = psBarCode
         txtField(1) = psModel
         txtField(2) = psColor
         txtField(4) = psModelCde
      ElseIf KeyCode = vbKeyF3 Then
         Call getCPInventory(txtField(3), False)
         txtField(3) = psBarCode
         txtField(1) = psModel
         txtField(2) = psColor
         txtField(4) = psModelCde
      End If
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   'On Error GoTo errProc
   
   If Index = 0 Then
      If pnEditMode = xeModeAddNew Or pnEditMode = xeModeUpdate Then
         txtField(Index).Text = UCase(txtField(Index).Text)
         poRS("sGSCMCode") = txtField(Index).Text
      End If
   End If
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
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

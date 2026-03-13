VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmPOS_Credit_Register 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Credit Card Transaction Register"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2205
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame1 
      Height          =   690
      Index           =   0
      Left            =   120
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   4890
      _ExtentX        =   8625
      _ExtentY        =   1217
      BackColor       =   7716603
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   1
         Top             =   75
         Width           =   3045
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   3
         Top             =   330
         Width           =   3045
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card Name"
         Height          =   285
         Index           =   5
         Left            =   135
         TabIndex        =   0
         Top             =   75
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Card Number"
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   2
         Top             =   330
         Width           =   1260
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   3915
      TabIndex        =   5
      Top             =   1425
      Width           =   1095
      _ExtentX        =   1931
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
      Picture         =   "frmPOS_Credit_Register.frx":0000
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   2805
      TabIndex        =   4
      Top             =   1425
      Width           =   1095
      _ExtentX        =   1931
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
      Picture         =   "frmPOS_Credit_Register.frx":077A
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
End
Attribute VB_Name = "frmPOS_Credit_Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private pbnewitem As Boolean
Private psSelected() As String
Dim lnrow As Long
Dim lsSQL As String

Private Sub cmdButton_Click(Index As Integer)
Dim lnCtr As Integer

   Select Case Index
      Case 0 'Save
         If txtfield(0).Tag <> "" And txtfield(1).Text <> "" Then
            Update_Data
         Else
            MsgBox "Incomplete Data!!!", vbCritical, "Warning"
         End If
      Case 1 'Cancel
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   If bLoaded = False Then
      bLoaded = True
      ShowData
      txtfield(0).SetFocus
   End If
End Sub

Private Sub ShowData()
Dim lrs As New ADODB.Recordset

   Set lrs = New ADODB.Recordset
   lsSQL = "SELECT" _
            & " a.sTransNox, " _
            & " a.sAcctNmbr, " _
            & " b.sCreditNm, " _
            & " a.sCreditID  " _
         & " FROM CP_SO_Credit a " _
            & " LEFT JOIN Credit_Card b " _
               & " ON a.sCreditID = b.sCreditID " _
         & " WHERE a.sTransNox  = '" & frmPOS_Register.txtfield(0).Text & "' "
   
   If lrs.State = adStateOpen Then lrs.Close
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   txtfield(0).Text = IIf(IsNull(lrs("sCreditNm")), "", lrs("sCreditNm"))
   txtfield(0).Tag = IIf(IsNull(lrs("sCreditID")), "", lrs("sCreditID"))
   txtfield(1).Text = lrs("sAcctNmbr")

End Sub

Private Sub Form_Load()
Dim lsSQL As String

   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
   txtfield(0).Tag = ""
   txtfield(0).Text = ""
   txtfield(1).Text = ""
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtfield(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
      If Index = 0 Then
         SearchCard False
         If txtfield(Index).Text <> "" Then SetNextFocus
      End If
   End If
   KeyCode = 0
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

Private Sub SearchCard(ByVal SearchValue As Boolean)
   Dim lsSearch As String
   Dim lrs As ADODB.Recordset
   Dim lsSQL As String
   Dim temp As Long
   
   Set lrs = New ADODB.Recordset
   
   lsSQL = "SELECT" _
            & " sCreditID, " _
            & " sCreditNm, " _
            & " nPercentx " _
         & " FROM Credit_Card " _
         & " WHERE cRecdStat = 1 " _

   If SearchValue Then
      lsSQL = lsSQL & " AND sCreditNm = '" & txtfield(0).Text & "'"
   Else
      lsSQL = lsSQL & " AND sCreditNm LIKE '%" & txtfield(0).Text & "%' "
   End If
            
   lsSQL = lsSQL & " ORDER BY sCreditNm"
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      
   If lrs.RecordCount = 1 Then
      txtfield(0).Text = lrs("sCreditNm")
      txtfield(0).Tag = lrs("sCreditID")

   ElseIf lrs.RecordCount > 1 Then
      lsSearch = KwikBrowse(oApp, lrs, _
                        "sCreditID»" _
                      & "sCreditNm»" _
                      & "nPercentx", _
                        "Card ID»" _
                      & "Card Name»" _
                      & "% Charge")
                      
      If lsSearch <> "" Then
          psSelected = Split(lsSearch, "»")
          txtfield(0).Text = psSelected(1)
          txtfield(0).Tag = psSelected(0)
      End If
   
   Else
      txtfield(0).Text = ""
      txtfield(0).Tag = ""
      txtfield(0).SetFocus
   End If
   Set lrs = Nothing
End Sub
Private Sub txtField_LostFocus(Index As Integer)
   txtfield(Index).BackColor = &HFFFFFF
End Sub

Private Function Update_Data() As Boolean
Dim lnrow As Long
Dim lnCtr As Integer
Dim lrs As New ADODB.Recordset

Update_Data = True
On Error GoTo errProc
       
   lsSQL = "UPDATE CP_SO_Credit SET" _
               & " sCreditID = '" & txtfield(0).Tag & "'," _
               & " dModified = getdate() " _
         & " WHERE sTransNox = '" & frmPOS_Register.txtfield(0).Text & "' "
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
         
   If lnrow <= 0 Then
      MsgBox "Unable to Update Record!!!" & vbCrLf & vbCrLf & _
      "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
      Update_Data = False
      GoTo endProc
   End If
   
   MsgBox "Record Successfully Updated!!!", vbInformation, "Information"
   Unload Me
   
endProc:
   Exit Function
errProc:
   Update_Data = False
   MsgBox Err.Description, vbCritical, "Warning"

End Function



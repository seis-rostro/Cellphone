VERSION 5.00
Object = "{0A7B56A6-35D0-4533-91DA-1715D3A0DD3E}#1.1#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Serial_Transfer 
   BorderStyle     =   0  'None
   Caption         =   "CP Serial Transfer"
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame2 
      Height          =   5610
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   9895
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   1
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   5460
         Left            =   60
         TabIndex        =   0
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   60
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   9631
         AllowBigSelection=   -1  'True
         AutoAdd         =   -1  'True
         AutoNumber      =   -1  'True
         BACKCOLOR       =   -2147483643
         BACKCOLORBKG    =   8421504
         BACKCOLORFIXED  =   -2147483633
         BACKCOLORSEL    =   -2147483635
         BORDERSTYLE     =   1
         COLS            =   2
         FILLSTYLE       =   0
         FIXEDCOLS       =   1
         FIXEDROWS       =   1
         FOCUSRECT       =   2
         EDITORBACKCOLOR =   -2147483643
         EDITORFORECOLOR =   -2147483640
         FORECOLOR       =   -2147483640
         FORECOLORFIXED  =   -2147483630
         FORECOLORSEL    =   -2147483634
         FORMATSTRING    =   ""
         Object.HEIGHT          =   5460
         GRIDCOLOR       =   12632256
         GRIDCOLORFIXED  =   0
         BeginProperty GRIDFONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GRIDLINES       =   1
         GRIDLINESFIXED  =   2
         GRIDLINEWIDTH   =   1
         MOUSEICON       =   "frmCP_Serial_Transfer.frx":0000
         MOUSEPOINTER    =   0
         REDRAW          =   -1  'True
         RIGHTTOLEFT     =   0   'False
         ROWS            =   2
         SCROLLBARS      =   3
         SCROLLTRACK     =   0   'False
         SELECTIONMODE   =   0
         Object.TOOLTIPTEXT     =   ""
         WORDWRAP        =   0   'False
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   375
      Index           =   0
      Left            =   4575
      TabIndex        =   1
      Top             =   1650
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      Caption         =   "&OK"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCP_Serial_Transfer.frx":001C
      CaptionAlign    =   0
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   4575
      TabIndex        =   2
      Top             =   2040
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmCP_Serial_Transfer.frx":0796
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmCP_Serial_Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Programmed By  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Rosalyn Lazo Descallar  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Started  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 23, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset

Dim lsSQL As String
Dim pnCtr As Integer
Dim psSelected() As String

Private Sub cmdButton_Click(Index As Integer)
Dim Cancel As Boolean

   Select Case Index
      Case 0 'OK
         If Cancel Then Exit Sub
         Cancel = Not Save_Dummy_Serial
      Case 1 'Cancel
         Me.Hide
      End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
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
   oSkin.ApplySkin xeFormTransDetail
   
   InitGrid
            
End Sub

Private Sub InitGrid()
    With GridEditor1
      .Cols = 3
      .Font = "MS Sans Serif"
    
      'column title
      .TextMatrix(0, 1) = "Cellphone IMEI No."
      .TextMatrix(0, 2) = "Stock ID"
      .Row = 0
      
      .ColAlignment(1) = 1
      
      'column width
      .ColWidth(0) = 500
      
      .ColEnabled(2) = False
      .ColFormat(2) = ">"
      
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .Rows = .Rows Then
         Cancel = True
      End If
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Search_Serial
End Sub

Private Sub Search_Serial()
Dim lsSQL As String
Dim lsCondition As String
Dim lsSearch As String
   
   With GridEditor1
         lsSQL = "SELECT" _
               & " a.sIMEINoxx, " _
               & " a.sStockIDx, " _
               & " b.sBarrCode  " _
            & " FROM CP_Serial_Master a " _
               & " LEFT JOIN CP_Inventory b " _
                  & " ON a.sStockIDx = b.sStockIDx " _
            & " WHERE a.sStockIDx = '" & .TextMatrix(.Row, 2) & "'"
         If oRS.State = adStateOpen Then oRS.Close
         oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

         If Not oRS.EOF Then
            If oRS.RecordCount = 1 Then
               .TextMatrix(.Row, 1) = oRS(0)
            Else
               lsSearch = KwikSearch(oApp, lsSQL, _
                          "sIMEINoxx»sBarrCode", _
                          "IMEI No.»Bar Code")
               If lsSearch <> "" Then
                  psSelected = Split(lsSearch, "»")
                  .TextMatrix(.Row, 1) = psSelected(0)
               End If
            End If
            .SetFocus
            .Refresh
         Else
            MsgBox "IMEI No. Not Existing!!!", vbCritical, "Warning"
            .Row = ""
            .Row = .Row - 1
         End If
      Set oRS = Nothing
   End With
   
End Sub
Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Or KeyCode = vbKeyF3 Then
      Search_Serial
   End If
End Sub

Private Sub GridEditor1_GotFocus()
   GridEditor1.Col = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hWnd Then Exit Sub
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Function Save_Dummy_Serial() As Boolean
Dim lnrow As Long
Dim lnCtr As Integer

Save_Dummy_Serial = True
On Error GoTo errProc
   
   With GridEditor1
      lsSQL = "SELECT * " _
            & " FROM CP_Serial_Dummy "
      If oRS.State = adStateOpen Then oRS.Close
      oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
         If oRS.RecordCount <> 0 Then
            For pnCtr = 1 To .Rows - 1
               If .TextMatrix(pnCtr, 1) = "" Then
                  MsgBox "Incomplete Data!!!" & vbCrLf & _
                  "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
                  .SetFocus
                  GoTo endProc
               Else
                  oApp.Connection.Execute "DELETE CP_Serial_Dummy " _
                     & " WHERE sStockIDx = '" & .TextMatrix(pnCtr, 2) & "'"
               End If
            Next
         End If
      'Delete Dummy Content
      
      For lnCtr = 1 To .Rows - 1
         lsSQL = "INSERT INTO CP_Serial_Dummy" _
               & "( sStockIDx, " _
               & "  sIMEINoxx )" _
            & " VALUES " _
               & "('" & .TextMatrix(lnCtr, 2) & "'," _
               & "'" & .TextMatrix(lnCtr, 1) & "')"
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
         
         If lnrow <= 0 Then
            MsgBox "Unable to Save Dummy_Serial!!!", vbCritical, "Warning"
            GoTo endProc
         Else
            Me.Hide
         End If
      Next
   End With

endProc:
   Set oRS = Nothing
   Exit Function
errProc:
   Save_Dummy_Serial = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function


'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Tested  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 15, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'


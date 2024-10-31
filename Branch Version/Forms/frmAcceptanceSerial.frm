VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmAcceptanceSerial 
   BorderStyle     =   0  'None
   Caption         =   "CP Serial"
   ClientHeight    =   6960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   6300
      Left            =   105
      TabIndex        =   0
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   555
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   11113
      AllowBigSelection=   -1  'True
      AutoAdd         =   0   'False
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
      FOCUSRECT       =   1
      EDITORBACKCOLOR =   -2147483643
      EDITORFORECOLOR =   -2147483640
      FORECOLOR       =   -2147483640
      FORECOLORFIXED  =   -2147483630
      FORECOLORSEL    =   -2147483634
      FORMATSTRING    =   ""
      Object.HEIGHT          =   6300
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
      MOUSEICON       =   "frmAcceptanceSerial.frx":0000
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
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   5025
      TabIndex        =   2
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmAcceptanceSerial.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   5025
      TabIndex        =   1
      Top             =   555
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Ok"
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
      Picture         =   "frmAcceptanceSerial.frx":0796
   End
End
Attribute VB_Name = "frmAcceptanceSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmAcceptanceSerial"

Private oSkin As clsFormSkin
Private oRS As Recordset

Dim pnCancel As Integer, pnCtr As Integer
Dim pbEntryNo As Boolean
Dim pbGridValidate As Boolean

Property Let EntryNo(bEntry As Boolean)
   pbEntryNo = bEntry
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc
   
   With GridEditor1
      Select Case Index
      Case 0
         For pnCtr = 1 To .Rows - 1
            If .TextMatrix(pnCtr, 1) = "" Then Exit Sub
         Next
         
         Me.Hide
         pnCancel = Index
      Case 1
         .Rows = 2
            
         .TextMatrix(1, 1) = ""
         .TextMatrix(1, 2) = ""

         Me.Hide
         pnCancel = Index
      End Select
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Property Get Cancel() As Integer
   Cancel = pnCancel
End Property

Private Sub Form_Activate()
   'column width
   With GridEditor1
      If .Rows > 26 Then
         If Not pbEntryNo Then
            .ColWidth(1) = 4150
         Else
            .ColWidth(1) = 360
            .ColWidth(2) = 3800
         End If
      Else
         If Not pbEntryNo Then
            .ColWidth(1) = 4250
         Else
            .ColWidth(1) = 360
            .ColWidth(2) = 3900
         End If
      End If
      
      If Not pbEntryNo Then
         .ColEnabled(1) = True
      End If
            
      .Row = 1
      .Col = 1
      .SetFocus
   End With
   pnCancel = 1
   pbGridValidate = False
   
   Call InitDummySerial
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

Private Sub Form_Load()
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormTransDetail
End Sub

Public Sub InitGrid1()
   With GridEditor1
      .Cols = 2
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Serial No"
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next
      
      .ColWidth(0) = 330
      .ColLimit(1) = 25
      .ColEnabled(1) = True
      
      'column format
      .ColFormat(1) = ">"
      
      'column alignment
      .ColAlignment(1) = 2
   End With
End Sub

Public Sub InitGrid2()
   With GridEditor1
      .Cols = 3
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "No."
      .TextMatrix(0, 2) = "Serial No."
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next
      
      .ColWidth(0) = 330
      .ColEnabled(1) = False
      .ColEnabled(2) = False
      
      'column format
      .ColFormat(2) = ">"
      
      'column alignment
      .ColAlignment(2) = 2
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "GridEditor1_EditorValidate"
   On Error GoTo errProc

   With GridEditor1
      If Trim(.TextMatrix(.Row, .Col)) = "" Then Exit Sub
      If pbGridValidate Then
         pbGridValidate = False
         Exit Sub
      End If

      oRS.Find "sSerialNo = " & strParm(.TextMatrix(.Row, .Col)), 0, adSearchForward, 1
      If Not oRS.EOF Then
         If oRS("nEntryNox") <> .Row Then
            MsgBox "Duplicate Serial No. " & oRS("nEntryNox") & " Detected!!!" & vbCrLf & _
                     "Please Verify Row No." & " " & oRS("nEntryNox"), vbCritical, "Warning"
            .TextMatrix(.Row, .Col) = ""
            Cancel = True
            Exit Sub
         End If
      Else
         oRS.Move .Row - 1, adBookmarkFirst
         oRS("nEntryNox") = .TextMatrix(.Row, 0)
         oRS("sSerialNo") = .TextMatrix(.Row, 1)
      End If
   End With
   pbGridValidate = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )", True
End Sub

Private Sub InitDummySerial()
   Dim lnCtr As Integer
   
   Set oRS = New Recordset
   
   With GridEditor1
      oRS.Fields.Append "nEntryNox", adInteger, 4
      oRS.Fields.Append "sSerialNo", adVarChar, 20
      oRS.Open
      
      For lnCtr = 1 To .Rows - 1
         oRS.AddNew
         oRS("nEntryNox") = .TextMatrix(lnCtr, 0)
         oRS("sSerialNo") = .TextMatrix(lnCtr, 1)
      Next
   End With
End Sub

Private Sub GridEditor1_LostFocus()
   pbGridValidate = False
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

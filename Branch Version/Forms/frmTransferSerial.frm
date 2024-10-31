VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmTransferSerial 
   BorderStyle     =   0  'None
   Caption         =   "CP Serial"
   ClientHeight    =   6960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
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
      MOUSEICON       =   "frmTransferSerial.frx":0000
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
      TabIndex        =   3
      Top             =   1815
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
      Picture         =   "frmTransferSerial.frx":001C
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
      Picture         =   "frmTransferSerial.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   5025
      TabIndex        =   2
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmTransferSerial.frx":0F10
   End
End
Attribute VB_Name = "frmTransferSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmTransSerial"

Private WithEvents oSerial As clsCPTransfer
Attribute oSerial.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnCancel As Integer, pnCtr As Integer
Dim pnEntryNo As Integer
Dim pbEntryNo As Boolean

Property Set SerialTrans(loSerial As clsCPTransfer)
   Set oSerial = loSerial
End Property

Property Let EntryNo(bEntry As Boolean)
   pbEntryNo = bEntry
End Property

Property Let LineNo(nEntry As Integer)
   pnEntryNo = nEntry
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   On Error Goto errProc
   
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
      Case 2
         oSerial.searchSerial pnEntryNo, GridEditor1.Row - 1, "sSerialNo", ""
         GridEditor1.Refresh
         GridEditor1.SetFocus
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
            .ColWidth(2) = 3970
         End If
      Else
         If Not pbEntryNo Then
            .ColWidth(1) = 4250
         Else
            .ColWidth(1) = 360
            .ColWidth(2) = 4070
         End If
      End If
      
      If Not pbEntryNo Then .ColEnabled(1) = True
            
      .Row = 1
      .Col = 1
      .SetFocus
   End With
   pnCancel = 1
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
   'On Error Goto errProc
   
   With GridEditor1
      oSerial.Serial(pnEntryNo, .Row - 1, "sSerialNo") = .TextMatrix(.Row, .Col)
      .TextMatrix(.Row, .Col) = oSerial.Serial(pnEntryNo, .Row - 1, "sSerialNo")
   End With

endProc:
   GridEditor1.Refresh
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )", True
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "GridEditor1_KeyDown"
   'On Error Goto errProc
   
   If KeyCode = vbKeyF3 Then
      With GridEditor1
         oSerial.searchSerial pnEntryNo, .Row - 1, "sSerialNo", .TextMatrix(.Row, 1)
      
         .Refresh
         .SetFocus
         .TextMatrix(.Row, 1) = oSerial.Serial(pnEntryNo, .Row - 1, "sSerialNo")
         KeyCode = 0
      End With
   End If
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & KeyCode _
                       & ", " & Shift _
                       & " )", True
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

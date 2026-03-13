VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmLoadSerial 
   BorderStyle     =   0  'None
   Caption         =   "MC Serial"
   ClientHeight    =   7665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7665
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   6945
      Left            =   120
      TabIndex        =   0
      Top             =   570
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   12250
      AllowBigSelection=   -1  'True
      AutoAdd         =   0   'False
      AutoNumber      =   0   'False
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
      Object.HEIGHT          =   6945
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
      MOUSEICON       =   "frmLoadSerial.frx":0000
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
      Index           =   0
      Left            =   4140
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmLoadSerial.frx":001C
   End
End
Attribute VB_Name = "frmLoadSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'she 12-9-2014
'show imei per transaction
Option Explicit
Private Const pxeMODULENAME = "frmLoadSerial"

Private oSkin As clsFormSkin
Private oRS As ADODB.Recordset

Dim p_sTransNox As String

Property Let TransNox(lsTransNox As String)
   p_sTransNox = lsTransNox
End Property

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
      Case 0
         Unload Me
   End Select
   
End Sub

Private Sub Form_Activate()
   Dim lnCtr As Integer

   Set oRS = New ADODB.Recordset
   oRS.Open "SELECT " & _
               " b.sSerialNo" & _
               ", d.sModelNme" & _
            " FROM CP_Price_Protection_Detail a" & _
            ", CP_Inventory_Serial b" & _
            ", CP_Inventory c" & _
                  " LEFT JOIN CP_Model d" & _
                     " ON c.sModelIDx = d.sModelIDx" & _
            " WHERE a.sSerialId = b.sSerialID" & _
            " AND a.sTransNox = " & strParm(p_sTransNox) & _
            " AND b.sStockIDx = c.sStockIDx" _
   , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
                  
    With GridEditor1
      .Cols = 3
      .Font = "MS Sans Serif"
      
      'column title
      .TextMatrix(0, 1) = "Model"
      .TextMatrix(0, 2) = "SerialNo"
      
      
      'column alignment
      .Row = 0
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 1
      Next
      
      'column width
      .ColWidth(0) = 330
      .ColWidth(1) = 1400
      .ColWidth(2) = 1700
      
      
      .Row = 1
      .Col = 1
   End With
   
   If oRS.EOF Then Exit Sub
   
   With GridEditor1
      .Rows = oRS.RecordCount + 1
      For lnCtr = 0 To oRS.RecordCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = oRS("sModelNme")
         .TextMatrix(lnCtr + 1, 2) = oRS("sSerialNo")
         oRS.MoveNext
      Next
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      
   End With
End Sub

Private Sub Form_Load()
   Dim nCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.DisableClose = True
   oSkin.ApplySkin xeFormTransDetail
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oRS = Nothing
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



VERSION 5.00
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmSOSerialNo 
   BorderStyle     =   0  'None
   Caption         =   "CP Serial"
   ClientHeight    =   6960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   6300
      Left            =   105
      TabIndex        =   0
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   555
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   11113
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
      MOUSEICON       =   "frmSOSerialNo.frx":0000
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
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F5-OK"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   4
      Left            =   6330
      TabIndex        =   4
      Top             =   1050
      Width           =   1275
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Escape"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   3
      Left            =   6330
      TabIndex        =   3
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F8-Del."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   2
      Left            =   6330
      TabIndex        =   2
      Top             =   1545
      Width           =   1275
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F1-Void"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   0
      Left            =   6330
      TabIndex        =   1
      Top             =   555
      Width           =   1275
   End
End
Attribute VB_Name = "frmSOSerialNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmSOSerialNo"

Private WithEvents oTrans As clsCPSales
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pbCancelled As Boolean
Dim pnCtr As Integer
Dim pbGridValidate As Boolean

Property Set SerialTrans(loSerial As clsCPSales)
   Set oTrans = loSerial
End Property

Property Get Cancelled() As Boolean
   Cancelled = pbCancelled
End Property

Private Sub Form_Activate()
   Dim lnCtr As Integer
   
   'column width
   With GridEditor1
      .Rows = IIf(oTrans.AccessoryCount = 0, 2, oTrans.AccessoryCount + 1)
      .ColWidth(2) = 3550
      If .Rows > 26 Then .ColWidth(2) = 3450
      
      For lnCtr = 0 To oTrans.AccessoryCount - 1
         .TextMatrix(lnCtr + 1, 1) = oTrans.Accessory(lnCtr, "sDescript")
         .TextMatrix(lnCtr + 1, 2) = oTrans.Accessory(lnCtr, "sSerialNo")
      Next
            
      .Row = 1
      .Col = 1
'      .SetFocus
   End With
   pbGridValidate = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lnCtr As Integer
   
   With GridEditor1
      Select Case KeyCode
      Case vbKeyF1
         oTrans.WithAccessory = False
         Me.Hide
         pbCancelled = False
      Case vbKeyF5
         oTrans.Accessory(.Row - 1, "sSerialNo") = .TextMatrix(.Row, .Col)
         If .Rows > 2 Then
            lnCtr = 0
            Do While lnCtr < .Rows
               If Trim(.TextMatrix(lnCtr, 1)) = "" Then
                  .Row = lnCtr
                  If oTrans.DeleteAccess(.Row - 1) Then
                     .deleteRow
                  End If
               Else
                  lnCtr = lnCtr + 1
               End If
            Loop
         End If
         
         oTrans.WithAccessory = True
         Me.Hide
         pbCancelled = False
      Case vbKeyEscape
         .Rows = 2
            
         .TextMatrix(1, 1) = ""
         .TextMatrix(1, 2) = ""

         Me.Hide
         pbCancelled = True
      Case vbKeyF8
         If .Rows > 2 Then
            If oTrans.DeleteAccess(.Row - 1) Then
               .deleteRow
               
               If .Rows > 26 Then
                  .ColWidth(2) = 3450
               Else
                  .ColWidth(2) = 3550
               End If
            End If
         End If
      End Select
   End With
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
      .Cols = 3
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Category"
      .TextMatrix(0, 2) = "Serial No."
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      
      .ColWidth(0) = 330
      .ColEnabled(1) = True
      .ColEnabled(2) = True
      
      'column format
      .ColFormat(2) = ">"
      
      'column alignment
      .ColAlignment(2) = 2
      
      'column width
      .ColWidth(1) = 2000
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
'      ElseIf .TextMatrix(.Row, 2) = "" Then
'         Cancel = True
      End If
      
      If Not Cancel Then
          Cancel = Not oTrans.addAccessory(.Row - 1)
      End If
   
      If .Rows > 26 Then .ColWidth(2) = 3450
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "GridEditor1_EditorValidate"
   'On Error GoTo errProc
   
   With GridEditor1
      If pbGridValidate Then
         pbGridValidate = False
         Exit Sub
      End If
      
      Select Case .Col
      Case 1
         oTrans.Accessory(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
         .TextMatrix(.Row, .Col) = oTrans.Accessory(.Row - 1, .Col)
      Case 2
         oTrans.Accessory(.Row - 1, .Col) = Format(.TextMatrix(.Row, .Col), ">")
      End Select
      
      pbGridValidate = True
endProc:
      GridEditor1.Refresh
   End With
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )", True
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "GridEditor1_KeyDown"
   'On Error GoTo errProc
    
   If KeyCode = vbKeyF3 Then
      With GridEditor1
         If .Col = 1 Then
            .Refresh
            .SetFocus
            If oTrans.searchAccessories(.Row - 1, .TextMatrix(.Row, .Col)) Then .Col = 2
            .TextMatrix(.Row, 1) = oTrans.Accessory(.Row - 1, 1)
         End If
      End With
      KeyCode = 0
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

Private Function BranchAutomate(ByVal sBranchCd As String) As Boolean
   Dim lrs As Recordset
   
   Set lrs = New Recordset
   lrs.Open "SELECT * FROM Branch" & _
               " WHERE sBranchCd = " & strParm(sBranchCd) & _
                  " AND cAutomate = " & strParm(xeYes) _
   , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If Not lrs.EOF Then BranchAutomate = True
   Set lrs = Nothing
End Function


Private Sub GridEditor1_LostFocus()
   pbGridValidate = False
End Sub

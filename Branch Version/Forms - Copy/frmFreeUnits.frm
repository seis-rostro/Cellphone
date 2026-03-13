VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmFreeUnits 
   BorderStyle     =   0  'None
   Caption         =   "s"
   ClientHeight    =   6960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   6300
      Left            =   105
      TabIndex        =   0
      Tag             =   "et0;eb0;et0;bc2"
      Top             =   540
      Width           =   8535
      _ExtentX        =   15055
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
      MOUSEICON       =   "frmFreeUnits.frx":0000
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
      Left            =   8895
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
      Picture         =   "frmFreeUnits.frx":001C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   8895
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
      Picture         =   "frmFreeUnits.frx":0796
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   8895
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
      Picture         =   "frmFreeUnits.frx":0F10
   End
End
Attribute VB_Name = "frmFreeUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmFreeUnits"

Private oSkin As clsFormSkin
Private oRS As Recordset

Dim pnCancel As Integer, pnCtr As Integer

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
   pnCancel = 1
   Call InitGrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = GridEditor1.hwnd Then Exit Sub
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

Public Sub InitGrid()
   With GridEditor1
      .Cols = 5
      .Font = "MS Sans Serif"

      'column title
      .TextMatrix(0, 1) = "Barcode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "Brand"
      .TextMatrix(0, 4) = "Serial No"
      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      
      .ColWidth(0) = 330
      .ColWidth(1) = 1900
      .ColWidth(2) = 1800
      .ColWidth(3) = 2500
      .ColWidth(4) = 1900
      
      .ColEnabled(2) = False
      .ColEnabled(3) = False
      
      'column format
      .ColFormat(4) = ">"
      
      'column alignment
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      
      .Col = 1
      .Row = 1
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub searchSerial(Optional SerialNo As Variant)
   Dim lors As ADODB.Recordset
   Dim lsCondition As String, lsBrowse As String
   Dim lsSQL As String, lsOldProc As String
   Dim lnCtr As Integer
   
   lsOldProc = "searchSerial"
   On Error Goto errProc
   
   lsSQL = "SELECT" _
               & "  sSerialNo" _
            & " FROM DummySerial "
   
   With GridEditor1
      If Not IsMissing(SerialNo) Then lsSQL = lsSQL & " WHERE sSerialNo LIKE " & strParm("%" & IFNull(SerialNo))
         
      For lnCtr = 1 To .Rows - 1
         If .TextMatrix(lnCtr, 1) <> "" And .Row <> lnCtr Then
            lsCondition = strParm(.TextMatrix(lnCtr, 1)) & ", "
         End If
      Next

      If lsCondition <> "" Then
         lsCondition = "sSerialNo NOT IN (" & Left(lsCondition, Len(Trim(lsCondition)) - 1) & ")"
         lsSQL = AddCondition(lsSQL, lsCondition)
      End If
      
      lsSQL = lsSQL & " ORDER BY sSerialNo"
      
      Set lors = New Recordset
      lors.Open lsSQL, oApp.Connection, adOpenKeyset, adLockReadOnly, adCmdText
      
      If lors.EOF Then
         .TextMatrix(.Row, 1) = ""
      ElseIf lors.RecordCount = 1 Then
         .TextMatrix(.Row, 1) = lors("sSerialNo")
      Else
         lsBrowse = KwikBrowse(oApp, lors, "sSerialNo", "Serial No")
         
         If lsBrowse <> "" Then
            .TextMatrix(.Row, 1) = Trim(lors("sSerialNo"))
         End If
      End If
   End With
   
endProc:
   lors.Close
errProc:
   ShowError lsOldProc & "( " & IFNull(SerialNo) & " )"
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Dim lsOldProc As String

   lsOldProc = "GridEditor1_EditorValidate"
   On Error Goto errProc

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

      If Not pbVerify Then Exit Sub
      If .TextMatrix(.Row, 1) <> "" Then
         searchSerial .TextMatrix(.Row, 1)
      End If
   End With
   pbGridValidate = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Cancel & " )", True
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   lsOldProc = "GridEditor1_KeyDown"
   On Error Goto errProc
   
   If KeyCode = vbKeyF3 Then
      If pbVerify Then
         With GridEditor1
            searchSerial .TextMatrix(.Row, 1)
            .Refresh
         End With
      End If
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

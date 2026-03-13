VERSION 5.00
Object = "{0A7B56A6-35D0-4533-91DA-1715D3A0DD3E}#1.1#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmIMEI_Utility 
   BorderStyle     =   0  'None
   Caption         =   "IMEI No. Utility"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   735
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   1296
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   1
         Top             =   90
         Width           =   2535
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1050
         TabIndex        =   5
         Top             =   360
         Width           =   4365
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
         Index           =   2
         Left            =   4485
         MaxLength       =   50
         TabIndex        =   3
         Top             =   90
         Width           =   930
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Bar Code"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   105
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   4
         Top             =   375
         Width           =   840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "QOH"
         Height          =   285
         Index           =   3
         Left            =   3975
         TabIndex        =   2
         Top             =   105
         Width           =   585
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   5640
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1365
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9948
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   5490
         Left            =   45
         TabIndex        =   6
         Tag             =   "et0;eb0;et0;bc2"
         Top             =   60
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   9684
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
         Object.HEIGHT          =   5490
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
         MOUSEICON       =   "frmIMEI_Utility.frx":0000
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
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   5865
      TabIndex        =   8
      Top             =   960
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmIMEI_Utility.frx":001C
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   3
      Left            =   5865
      TabIndex        =   10
      Top             =   1800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmIMEI_Utility.frx":0796
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   5865
      TabIndex        =   7
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
      Picture         =   "frmIMEI_Utility.frx":0F10
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   5865
      TabIndex        =   9
      Top             =   1380
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmIMEI_Utility.frx":168A
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmIMEI_Utility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Programmed By  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Rosalyn Lazo Descallar  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Started  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 14, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset

Dim pbnewitem As Boolean
Dim psSelected() As String

Dim pnindex As Integer
Dim pnCtr As Integer

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
      Case 0 'save
         If txtField(0).Tag = "" Then
            MsgBox "Invalid BarCode!!!", vbCritical, "Warning"
         Else
            Save_CP_Serial
         End If
      Case 1 'search
         SearchBarCode
      Case 2 'New
         ClearFields
         EmptyGrid
      Case 3 'close
         Unload Me
      End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded = False Then
      ClearFields
      bLoaded = True
   End If
   
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .Rows <> txtField(2).Text Then
         Cancel = True
      End If
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
Dim lsSQL As String
   
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         MsgBox "Invalid IMEI No."
         .Row = .Row - 1
      End If
   End With

End Sub
Private Sub ShowGrid()
Dim lsSQL As String
Dim lnCtr As Integer
Dim lnrow As Long
                
   lsSQL = "SELECT" _
               & " sStockIDx, " _
               & " sIMEINoxx  " _
         & " FROM CP_Serial_Master " _
         & " WHERE sStockIDx = '" & txtField(0).Tag & " '" _
         & " ORDER BY sSerialID "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If oRS.RecordCount <> 0 Then
      With GridEditor1
         .Rows = oRS.RecordCount + 1
         For lnCtr = 0 To oRS.RecordCount - 1
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = oRS("sIMEINoxx")
            oRS.MoveNext
            .ColEnabled(lnCtr) = False
         Next
      End With
   Else
      Exit Sub
   End If
   Set oRS = Nothing

End Sub

Private Sub Form_Load()
Dim lsSQL As String

   CenterChildForm mdiMain, Me
   bLoaded = False
    
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   InitGrid
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
    
   EmptyGrid

End Sub

Private Sub InitGrid()
    
    With GridEditor1
        .Rows = 2
        .Cols = 2
        .Font = "MS Sans Serif"
        
        'column title
        .TextMatrix(0, 1) = "IMEI NO."
        .Row = 0
        
        'column allignment
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        
        'column width
        .ColWidth(0) = 500
        .ColWidth(1) = 4800
                        
        .Row = 1
    End With
    
End Sub

Private Sub ClearFields()
   For pnCtr = 0 To 2
      txtField(pnCtr).Text = ""
      txtField(pnCtr).Tag = ""
   Next
   txtField(1).Enabled = False
   txtField(2).Enabled = False
End Sub

Private Sub EmptyGrid()
   With GridEditor1
      .Rows = 2
      For pnCtr = 1 To .Cols - 1
         .TextMatrix(1, pnCtr) = ""
      Next
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub GridEditor1_GotFocus()
   GridEditor1.Col = 1
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   With GridEditor1
      If KeyCode = vbKeyF3 Or KeyCode = 13 Then
         If .Col = 1 Then
            SearchBarCode
         End If
      End If
   End With
End Sub

Private Sub SearchBarCode()
Dim lsSQL As String
Dim lsCondition As String
Dim lsSearch As String
   
   With GridEditor1
      lsSQL = "SELECT" _
             & " a.sBarrcode, " _
             & " a.sStockIDx, " _
             & " b.sBrandNme, " _
             & " c.sModelNme, " _
             & " a.sDescript, " _
             & " d.sColorNme, " _
             & " a.nSelPrice, " _
             & " e.nQtyOnHnd  " _
         & " FROM CP_Inventory a " _
             & " LEFT JOIN Brand b " _
               & " ON a.sBrandIdx = b.sBrandIdx " _
             & " LEFT JOIN Model c " _
               & " ON a.sModelIdx = c.sModelIdx " _
             & " LEFT JOIN Color d " _
               & " ON a.sColorIDx = d.sColorIDx " _
            & " LEFT JOIN CP_Inventory_Master e " _
               & " ON a.sStockIDx = e.sStockIDx " _
         & " WHERE a.sBarrcode like  '" & txtField(0).Text & "%' " _
            & " AND (sCategIDx = '01001' or sCategIDx = '01002' or sCategIDx = '01003')"
         If oRS.State = adStateOpen Then oRS.Close
         oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

      If Not oRS.EOF Then
         If oRS.RecordCount = 1 Then
            txtField(0).Text = oRS(0)
            txtField(0).Tag = oRS(1)
            txtField(1).Text = Trim(IIf(IsNull(oRS(2)), "", oRS(2)) & " " & _
                              IIf(IsNull(oRS(3)), "", oRS(3)) & " " & _
                              IIf(IsNull(oRS(4)), "", oRS(4)) & " " & _
                              IIf(IsNull(oRS(5)), "", oRS(5)))
            txtField(2).Text = oRS(7) + 1
            .Rows = oRS(7)
            .Refresh
            .SetFocus
         Else
            lsSearch = KwikSearch(oApp, lsSQL, _
                       "sBarrcode»sBrandNme»sModelNme»sDescript»sColorNme", _
                       "Bar Code»Brand»Model»Description»Color")
            If lsSearch <> "" Then
               psSelected = Split(lsSearch, "»")
               txtField(0).Text = psSelected(0)
               txtField(0).Tag = psSelected(1)
               txtField(1).Text = Trim(IIf(IsNull(psSelected(2)), "", psSelected(2)) & " " & _
                                 IIf(IsNull(psSelected(3)), "", psSelected(3)) & " " & _
                                 IIf(IsNull(psSelected(4)), "", psSelected(4)) & " " & _
                                 IIf(IsNull(psSelected(5)), "", psSelected(5)))
               txtField(2).Text = psSelected(7)
               .Rows = psSelected(7) + 1
               .Refresh
               .SetFocus
            End If
         End If
      Else
         MsgBox "Bar Code Not Existing!!!", vbInformation, "Information"
      End If
      Set oRS = Nothing
   End With
   
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtField(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      If Index = 0 Then
         SearchBarCode
         If txtField(Index).Text <> "" Then SetNextFocus
      End If
      KeyCode = 0
   End If
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

Private Function Save_CP_Serial() As Boolean
Dim lsSQL As String
Dim lnrow As Long
Dim temp As String
Dim lrs As ADODB.Recordset
   
Save_CP_Serial = True
On Error GoTo errProc
   
   If txtField(0).Tag = "" Then Exit Function
   With GridEditor1
      For pnCtr = 1 To .Rows - 1
         temp = getNextCode("CP_Serial_Master", "sSerialID", True, oApp.Connection, True, oApp.BranchCode)
         'CP_Serial_Master
         lsSQL = "INSERT INTO CP_Serial_Master " _
               & "( sSerialID, " _
               & "  sBranchCd, " _
               & "  sIMEINoxx, " _
               & "  sStockIDx, " _
               & "  cSoldStat, " _
               & "  cLocation, " _
               & "  sClientID, " _
               & "  sModified, " _
               & "  dModified) " _
         & "VALUES " _
               & "('" & temp & "', " _
               & "'" & oApp.BranchCode & "', " _
               & "'" & .TextMatrix(pnCtr, 1) & "', " _
               & "'" & txtField(0).Tag & "', " _
               & "'0'," _
               & "'1'," _
               & "'', " _
               & "'" & Encrypt(oApp.UserID) & "', " _
               & " getdate())"
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
      If lnrow <= 0 Then
         MsgBox "Unable to Save CP_Serial_Master!!!", vbCritical, "Warning"
         Save_CP_Serial = False
         GoTo endProc
      End If
      
      Next
      MsgBox "Save Successful!!!", vbInformation, "Information"
   End With

endProc:
   Set oRS = Nothing
   Exit Function
errProc:
   Save_CP_Serial = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Tested  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 16, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Version 1    ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Finished  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 17, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'




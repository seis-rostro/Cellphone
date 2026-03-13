VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmSuppliesTransfer 
   BorderStyle     =   0  'None
   Caption         =   "Supplies Transfer"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   15180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      Height          =   345
      Index           =   4
      Left            =   7365
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2145
      Width           =   4380
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   615
      Index           =   2
      Left            =   105
      TabIndex        =   0
      Top             =   1800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1085
      Caption         =   "&Del. Row"
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
      Picture         =   "frmSupplies_Transfer.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   615
      Index           =   0
      Left            =   105
      TabIndex        =   1
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1085
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
      Picture         =   "frmSupplies_Transfer.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   615
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1085
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
      Picture         =   "frmSupplies_Transfer.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   615
      Index           =   3
      Left            =   105
      TabIndex        =   3
      Top             =   2430
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1085
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
      Picture         =   "frmSupplies_Transfer.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   615
      Index           =   4
      Left            =   105
      TabIndex        =   4
      Top             =   540
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1085
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
      Picture         =   "frmSupplies_Transfer.frx":1DE8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   615
      Index           =   6
      Left            =   105
      TabIndex        =   5
      Top             =   2430
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1085
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
      Picture         =   "frmSupplies_Transfer.frx":2562
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   615
      Index           =   5
      Left            =   105
      TabIndex        =   6
      Top             =   1800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1085
      Caption         =   "&Print"
      AccessKey       =   "P"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSupplies_Transfer.frx":2CDC
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   615
      Index           =   7
      Left            =   105
      TabIndex        =   7
      Top             =   1170
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1085
      Caption         =   "&Save To"
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
      Picture         =   "frmSupplies_Transfer.frx":3456
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   7065
      Index           =   1
      Left            =   1545
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   12462
      BackColor       =   12632256
      BorderStyle     =   1
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   3090
         Left            =   135
         TabIndex        =   8
         Top             =   3390
         Width           =   10320
         _ExtentX        =   18203
         _ExtentY        =   5450
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
         Object.HEIGHT          =   3090
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
         MOUSEICON       =   "frmSupplies_Transfer.frx":3BD0
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
      Begin xrControl.xrFrame xrFrame2 
         Height          =   2010
         Left            =   120
         Top             =   105
         Width           =   10320
         _ExtentX        =   18203
         _ExtentY        =   3545
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1125
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   795
            Width           =   5070
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1080
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   135
            Width           =   2265
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   8205
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   780
            Width           =   2010
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   1125
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   1125
            Width           =   5070
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            Height          =   300
            Index           =   1
            Left            =   4290
            TabIndex        =   19
            Top             =   1560
            Width           =   915
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destination"
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   16
            Top             =   825
            Width           =   795
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Trans. No."
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
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   165
            Width           =   1110
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date "
            Height          =   300
            Index           =   0
            Left            =   7440
            TabIndex        =   14
            Top             =   825
            Width           =   915
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   405
            Left            =   1155
            Tag             =   "et0;ht2"
            Top             =   255
            Width           =   2325
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Index           =   2
            Left            =   345
            TabIndex        =   13
            Top             =   1170
            Width           =   570
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   1080
         Left            =   135
         Top             =   2205
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   1905
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   8
            Left            =   7770
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   540
            Width           =   2460
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   7
            Left            =   4380
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   540
            Width           =   2460
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   6
            Left            =   1065
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   540
            Width           =   2460
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   5
            Left            =   1065
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   105
            Width           =   3810
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entry No"
            Height          =   195
            Index           =   3
            Left            =   210
            TabIndex        =   23
            Top             =   645
            Width           =   615
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   22
            Top             =   225
            Width           =   795
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity"
            Height          =   195
            Index           =   6
            Left            =   3750
            TabIndex        =   21
            Top             =   645
            Width           =   585
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Price"
            Height          =   195
            Index           =   1
            Left            =   6960
            TabIndex        =   20
            Top             =   615
            Width           =   690
         End
      End
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Index           =   4
      Left            =   6765
      TabIndex        =   18
      Top             =   1950
      Width           =   570
   End
End
Attribute VB_Name = "frmSuppliesTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Const pxeMODULENAME = "frmCP_Branch_Transfer"
'
'Private WithEvents oTrans As ggcSuppliesTransfer
'Private oSkin As clsFormSkin
'
'Dim pnIndex As Integer
'Dim pbGridFocus As Boolean
'Dim pnCtr As Integer
'Dim pbSave As Boolean
'Dim pbGridValidate As Boolean
'Dim pbClosedTrans As Boolean
'
'Private Sub chkField_Click()
''   oTrans.DiskTransaction = IIf(chkField.Value = 1, True, False)
'End Sub
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lnRep As String
'
'   lsOldProc = "cmdButton_Click"
'   'On Error GoTo errProc
'
'   With GridEditor1
'      Select Case Index
'      Case 0
'         If .Rows > 2 Then
'            pnCtr = 0
'            Do While pnCtr < .Rows
'               If Trim(.TextMatrix(pnCtr, 1)) = "" Then
'                  .Row = pnCtr
'                  If oTrans.deleteDetail(.Row - 1) Then .DeleteRow
'               Else
'                  pnCtr = pnCtr + 1
'               End If
'            Loop
'
'            .ColWidth(3) = 3100
'            If .Rows > 16 Then .ColWidth(3) = 2850
'         End If
'
'         If isEntryOK Then
'            If oTrans.SaveTransaction Then
'               MsgBox "Record Updated Successfully!!!", vbInformation, "Confirm"
'               If Not BranchAutomate(oTrans.Master("sDestinat")) Then
'                  If Not oTrans.AcceptDelivery(oTrans.Master("dTransact")) Then
'                     MsgBox "Automatic Posting encountered error!!!" & vbCrLf & _
'                              "Please contact GMC/GGC SEG for assistance!!!", vbCritical, "Warning"
'                  End If
'               End If
'
'               InitButton xeModeReady
'               lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'               If lnRep = vbYes Then
'                  If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'               End If
'               pbSave = True
'            Else
'               MsgBox "Unable to Update Record!!!", vbCritical, "Warning"
'            End If
'         End If
'      Case 1
'         If pbGridFocus Then
'            If oTrans.searchDetail(.Row - 1, 1) Then .Col = 1
'            .Refresh
'            .SetFocus
'         Else
'            oTrans.SearchMaster pnIndex
'         End If
'      Case 2
'         If .Rows > 2 Then
'            If oTrans.deleteDetail(.Row - 1) Then .DeleteRow
'
'            For pnCtr = 1 To .Rows - 1
'               .TextMatrix(pnCtr, 0) = pnCtr
'            Next
'
'            .ColWidth(3) = 3100
'            If .Rows > 16 Then .ColWidth(3) = 2850
'         End If
'      Case 3
'         lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
'                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'
'         If lnRep = vbYes Then
'            oTrans.NewTransaction
'            ClearFields
'            InitButton xeModeReady
'         Else
'            txtField(pnIndex).SetFocus
'         End If
'         pbSave = False
'      Case 4
'         oTrans.NewTransaction
'         ClearFields
'         InitButton xeModeAddNew
'         txtField(1).SetFocus
'      Case 5
'         If pbSave Then
'            lnRep = MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'            If lnRep = vbYes Then
'               If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'            End If
'         Else
'            MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
'         End If
'      Case 6
'         Unload Me
'      Case 7
''         If chkField.Value = 1 Then
''            oTrans.DiskTransaction = True
''            If oTrans.CreateDiskTransfer Then
''               MsgBox "Transaction was Successfully Save to Mobile Disk!!!", vbInformation, "Notice"
''            Else
''               MsgBox "Unable to Save Transaction to Mobile Disk!!!", vbCritical, "Warning"
''            End If
''         Else
''            MsgBox "MC Delivery Capture Was Not Yet Set!!!" & vbCrLf & _
''               "Please Checked 'Save to Mobile Disk' then Try Exporting Delivery Again!!!", _
''               vbCritical, "Warning"
''         End If
'      End Select
'   End With
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Index & " )", True
'End Sub
'
'Private Sub Form_Activate()
'   GridEditor1.Refresh
'
'   oApp.MenuName = Me.Tag
'   Me.ZOrder 0
'End Sub
'
'Private Sub Form_Load()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
'   'On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   Set oTrans = New ggcSuppliesTransfer
'   Set oTrans.AppDriver = oApp
'
''   oTrans.DiskTransaction = False
'   oTrans.InitTransaction
''   oTrans.NewTransaction
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'   oSkin.ApplySkin xeFormTransaction
'
'   InitGrid
'   ClearFields
'   InitButton xeModeAddNew
'
''   txtField(3).MaxLength = oTrans.MasFldSize(3)
''   txtField(4).MaxLength = oTrans.MasFldSize(4)
'
'   pbGridValidate = False
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oTrans = Nothing
'   Set oSkin = Nothing
'End Sub
'
'Private Sub GridEditor1_AddingRow(Cancel As Boolean)
'   With GridEditor1
'      If Trim(.TextMatrix(.Row, 1)) = "" Then
'         Cancel = True
'      ElseIf .TextMatrix(.Row, 6) = "0" Then
'         Cancel = True
'      End If
'
'      If Not Cancel Then
'         If .Row = .Rows - 1 Then oTrans.addDetail
'      End If
'
'      If .Rows > 16 Then .ColWidth(3) = 2850
'   End With
'End Sub
'
'Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
'   Dim lsOldProc As String
'
'   lsOldProc = "GridEditor1_EditorValidate"
'   'On Error GoTo errProc
'
'   With GridEditor1
'      If pbGridValidate Then
'         pbGridValidate = False
'         Exit Sub
'      End If
'
'      If .Col = 1 Or .Col = 2 Then
'         .TextMatrix(.Row, .Col) = compareSerial(.TextMatrix(.Row, .Col), .Row)
'      End If
'
'      oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
'      If oTrans.Detail(.Row - 1, "cHsSerial") = xeYes Then
'         Select Case .Col
'         Case 1, 2
'            If .TextMatrix(.Row, .Col) <> "" Then
'               oTrans.Detail(.Row - 1, "nQuantity") = 1
'               .TextMatrix(.Row, 6) = oTrans.Detail(.Row - 1, "nQuantity")
'               If .Row = .Rows - 1 Then
'                  .Rows = .Rows + 1
'                  oTrans.addDetail
'                  .Col = 0
'               End If
'
'               .Row = .Rows - 1
'            End If
'         Case 6
''            If CDbl(.TextMatrix(.Row, 6)) > CDbl(.TextMatrix(.Row, 5)) Then
''               .TextMatrix(.Row, .Col) = 0
''            End If
'
'            If CDbl(.TextMatrix(.Row, .Col)) > 1 Then .TextMatrix(.Row, .Col) = 1
'         End Select
'      End If
'
'      If .Rows > 16 Then
'         .TopRow = .Rows - 1
'         .ColWidth(3) = 2850
'      End If
'   End With
'   pbGridValidate = True
'
'endProc:
'   GridEditor1.Refresh
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & Cancel & " )", True
'End Sub
'
'Private Sub GridEditor1_GotFocus()
'   With GridEditor1
'      .EditorBackColor = oApp.getColor("HT1")
'   End With
'   pbGridFocus = True
'End Sub
'
'Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "GridEditor1_KeyDown"
'   'On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Then
'      With GridEditor1
'         If oTrans.searchDetail(.Row - 1, .Col, .TextMatrix(.Row, .Col)) Then
'            If oTrans.Detail(.Row - 1, "cHsSerial") = xeYes Then
'               .TextMatrix(.Row, 6) = 1
'               oTrans.Detail(.Row - 1, "nQuantity") = 1
'               If .Row = .Rows - 1 Then
'                  .Rows = .Rows + 1
'                  oTrans.addDetail
'               End If
'
'               .Row = .Rows - 1
'               .Col = 1
'            Else
'               .Col = 6
'            End If
'         Else
'            oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
'            .Col = 1
'         End If
'
'         .Refresh
'         .SetFocus
'         If .Rows > 16 Then
'            .TopRow = .Rows - 1
'            .ColWidth(3) = 2850
'         End If
'         KeyCode = 0
'      End With
'   End If
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & KeyCode _
'                       & ", " & Shift _
'                       & " )", True
'End Sub
'
'Private Sub GridEditor1_LostFocus()
'   With GridEditor1
'      .EditorBackColor = oApp.getColor("EB")
'      If cmdButton(0).Visible Then oTrans.Detail(.Row - 1, .Col) = .TextMatrix(.Row, .Col)
'      If .Rows > 16 Then .TopRow = .Rows - 1
'   End With
'
'   pbGridValidate = False
'End Sub
'
''Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
''   With GridEditor1
''      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
''   End With
''End Sub
''
''Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
''   txtField(Index).Text = oTrans.Master(Index)
''End Sub
'
'Private Sub InitGrid()
'   With GridEditor1
'      .Rows = 2
'      .Cols = 5
'      .Font = "MS Sans Serif"
'
'      'column title
'      .TextMatrix(0, 1) = "Entry No"
'      .TextMatrix(0, 2) = "Description"
'      .TextMatrix(0, 3) = "Quantity"
'      .TextMatrix(0, 4) = "Unit Price"
'
'      .Row = 0
'
'      'column alignment
'      For pnCtr = 0 To .Cols - 1
'         .Col = pnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'      Next
'
'      'column width
'      .ColWidth(0) = 330
'      .ColWidth(1) = 2600
'      .ColWidth(2) = 2500
'      .ColWidth(4) = 1020
'
'
'      .ColFormat(4) = "#,##0.00"
'      .ColFormat(5) = "#,##0"
'      .ColFormat(6) = "#,##0"
'      .ColNumberOnly(6) = True
'      .ColDefault(4) = 0#
'      .ColDefault(5) = 0
'      .ColDefault(6) = 0
'
'      .ColAlignment(1) = 1
'      .ColAlignment(2) = 1
'      .ColAlignment(3) = 1
'      .ColAlignment(4) = 6
'      .ColAlignment(5) = 6
'      .ColAlignment(6) = 6
'
'      .ColEnabled(3) = False
'      .ColEnabled(4) = False
'      .ColEnabled(5) = False
'
'      .EditorBackColor = oApp.getColor("HT1")
'
'      .Row = 1
'      .Col = 1
'   End With
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'      If Index = 1 Then .Text = Format(.Text, "MM/DD/YY")
'
'      .SelStart = 0
'      .SelLength = Len(.Text)
'      .BackColor = oApp.getColor("HT1")
'   End With
'
'   pbGridFocus = False
'   pnIndex = Index
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_KeyDown"
'   'On Error GoTo errProc
'
'   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'      With txtField(Index)
'         If KeyCode = vbKeyF3 Then
'            oTrans.SearchMaster Index, .Text
'            If .Text <> "" Then SetNextFocus
'         Else
'            If .Text <> "" Then oTrans.SearchMaster Index, .Text
'         End If
'      End With
'      KeyCode = 0
'   End If
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " _
'                       & "  " & Index _
'                       & ", " & KeyCode _
'                       & ", " & Shift _
'                       & " )", True
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'   Case vbKeyReturn, vbKeyUp, vbKeyDown
'      Select Case KeyCode
'      Case vbKeyReturn, vbKeyDown
'         If GetFocus = GridEditor1.hwnd Then Exit Sub
'         SetNextFocus
'      Case vbKeyUp
'         SetPreviousFocus
'      End Select
'   End Select
'End Sub
'
'Private Sub InitButton(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   cmdButton(4).Visible = Not lbShow
'   cmdButton(5).Visible = Not lbShow
'   cmdButton(6).Visible = Not lbShow
'   cmdButton(7).Visible = Not lbShow
'
'   cmdButton(0).Visible = lbShow
'   cmdButton(1).Visible = lbShow
'   cmdButton(2).Visible = lbShow
'   cmdButton(3).Visible = lbShow
'
'   For pnCtr = 1 To txtField.Count - 1
'      txtField(pnCtr).Enabled = lbShow
'   Next
'
'   With GridEditor1
'      .ColEnabled(1) = lbShow
'      .ColEnabled(2) = lbShow
'      .ColEnabled(6) = lbShow
'   End With
'
'   If Not lbShow Then cmdButton(4).SetFocus
'End Sub
'
'Private Function PrintTrans() As Boolean
'   Dim loreport As frmRepViewer
'   Dim lrs As Recordset
'   Dim lors As Recordset
'   Dim lnCtr As Integer
'   Dim lsOldProc As String
'   Dim lnTotlWSerial As Double
'   Dim lnTotlWOSerial As Double
'
'   lsOldProc = "PrinTrans"
'   'On Error GoTo errProc
'
'   PrintTrans = True
'
'   Set lrs = New ADODB.Recordset
'
'   lrs.Fields.Append "nField01", adInteger, 3
'   lrs.Fields.Append "nField02", adChar, 1
'   lrs.Fields.Append "sField01", adVarChar, 20
'   lrs.Fields.Append "sField02", adVarChar, 128
'   lrs.Fields.Append "sField03", adVarChar, 20
'   lrs.Fields.Append "sField04", adVarChar, 12
'   lrs.Fields.Append "sField05", adVarChar, 100
'   lrs.Open
'
'   With oTrans
'      lnTotlWOSerial = 0
'      lnTotlWSerial = 0
'
'      For lnCtr = 0 To .ItemCount - 1
'         lrs.AddNew
'         lrs.Fields("nField01") = oTrans.Detail(lnCtr, "nQuantity")
'         lrs.Fields("nField02") = oTrans.Detail(lnCtr, "cHsSerial")
'         lrs.Fields("sField01") = oTrans.Detail(lnCtr, "sBarrCode")
'         lrs.Fields("sField02") = oTrans.Detail(lnCtr, "sDescript")
'         lrs.Fields("sField03") = oTrans.Detail(lnCtr, "sSerialNo")
'         lrs.Fields("sField04") = oTrans.Detail(lnCtr, "sStockIDx")
'         lrs.Fields("sField05") = oTrans.Detail(lnCtr, "sBrandNme")
'         If oTrans.Detail(lnCtr, "cHsSerial") = xeYes Then
'            lnTotlWSerial = lnTotlWSerial + 1
'         Else
'            lnTotlWOSerial = lnTotlWOSerial + CDbl(oTrans.Detail(lnCtr, "nQuantity"))
'         End If
'      Next
'      lrs.Sort = "nField02 DESC,sField05,sField05,sField03"
'   End With
'
'   'assign important info to the report
'   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Transfer.rpt")
'   oReport.DiscardSavedData
'   oReport.FieldMappingType = crAutoFieldMapping
'   oReport.Database.SetDataSource lrs
'
'   Set lors = New ADODB.Recordset
'   If lors.State = adStateOpen Then lors.Close
'
'   lors.Open "SELECT" _
'               & "  a.sAddressx" _
'               & ", CONCAT(b.sTownName, ', ' , c.sProvName, ' ' , b.sZippCode) xTownName" _
'               & ", a.sBranchNm" _
'            & " FROM Branch a" _
'               & " LEFT JOIN TownCity b" _
'                  & " LEFT JOIN Province c" _
'                     & " ON b.sProvIDxx = c.sProvIDxx" _
'                  & " ON a.sTownIDxx = b.sTownIDxx" _
'            & " WHERE a.sBranchCd = " & strParm(oTrans.Master("sDestinat")) _
'            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'
'   oReport.Sections("PHa").ReportObjects("txtRefNo").SetText "CP" & "-" & Right(oTrans.Master("sTransNox"), 10)
'   oReport.Sections("PHa").ReportObjects("txtDate").SetText txtField(1).Text
'   oReport.Sections("PHb").ReportObjects("txtTo").SetText lors("sBranchNm")
'   oReport.Sections("PHb").ReportObjects("txtToAddress").SetText lors("sAddressx") & IFNull(lors("xTownName"), "")
'   oReport.Sections("PHb").ReportObjects("txtFrom").SetText oApp.ClientName
'   oReport.Sections("PHb").ReportObjects("txtFromAddress").SetText oApp.Address & ", " & oApp.TownCity & ", " & oApp.Province & " " & oApp.ZippCode
'   oReport.Sections("RFb").ReportObjects("txtRemarks").SetText txtField(4).Text
'   oReport.Sections("RFb").ReportObjects("txtWithSerial").SetText IIf(lnTotlWSerial = 0, "", Format(lnTotlWSerial, "#,##0"))
'   oReport.Sections("RFb").ReportObjects("txtWOutSerial").SetText IIf(lnTotlWOSerial = 0, "", Format(lnTotlWOSerial, "#,##0"))
'   oReport.Sections("PF").ReportObjects("txtRptUser").SetText oApp.UserName
'
'   Set loreport = New frmRepViewer
'   Set loreport.ReportSource = oReport
'   loreport.Show
'
'   PrintTrans = True
'
'endPoc:
'   If Not pbClosedTrans Then
'      If Not BranchAutomate(oTrans.Master("sDestinat")) Then
'         If oTrans.CloseTransaction(oTrans.Master(0)) Then pbClosedTrans = True
'      End If
'   End If
'   Set loreport = Nothing
'   Set oReport = Nothing
'   Set lrs = Nothing
'   Set lors = Nothing
'   Exit Function
'errProc:
'   PrintTrans = False
'   ShowError lsOldProc & "( " & " )"
'End Function
'
'Private Sub ClearFields()
'   For pnCtr = 0 To txtField.Count - 1
'      Select Case pnCtr
'      Case 0
'         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "@@@@-@@@@@@")
'      Case 2
'         txtField(pnCtr).Text = oTrans.Master(pnCtr)
'      Case 1
'         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
'      Case Else
'        txtField(pnCtr).Text = Empty
'      End Select
'   Next
'
'   With GridEditor1
'      .Rows = 2
'      .ColWidth(3) = 3100
'
'      'empty row
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = ""
'      .TextMatrix(1, 4) = "0.00"
'      .TextMatrix(1, 5) = "0"
'      .TextMatrix(1, 6) = "0"
'   End With
'
''   chkField.Value = 0
''   pbSave = False
''   pbClosedTrans = False
'End Sub
'
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   With txtField(Index)
'      .Text = TitleCase(.Text)
'      Select Case Index
'      Case 1
'         If Not IsDate(.Text) Then .Text = oApp.ServerDate
'         .Text = Format(.Text, "MMMM DD, YYYY")
'      Case 3
'         .Text = Format(.Text, ">")
'      End Select
'
'      oTrans.Master(Index) = .Text
'   End With
'End Sub
'
'Private Function isEntryOK() As Boolean
'   If txtField(2).Text = "" Then
'      MsgBox "Destination not found!!!" & vbCrLf & _
'             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
'      txtField(2).SetFocus
'      GoTo EntryNotOK
'   End If
'
'   With GridEditor1
'      If Trim(.TextMatrix(1, 1)) = "" Then
'         MsgBox "Detail is required!!!" & vbCrLf & _
'                "Please Verify your Entry then Try Again", vbCritical, "Warning"
'         .SetFocus
'         .Row = 1
'         .Col = 1
'         GoTo EntryNotOK
'      End If
'   End With
'
'EntryOK:
'   isEntryOK = True
'   Exit Function
'EntryNotOK:
'   isEntryOK = False
'End Function
'
'Private Function compareSerial(Value As String, Row As Integer) As String
'   Dim lnRep As Integer
'   Dim lnCtr As Integer
'   Dim lsValue As String
'   Dim lnValue As Integer
'
'   If Trim(Value) = "" Then
'      compareSerial = ""
'      Exit Function
'   End If
'
'   With GridEditor1
'      For lnCtr = 1 To .Rows - 1
'         If .TextMatrix(lnCtr, 1) = Value And lnCtr <> Row Then
'            If oTrans.Detail(lnCtr - 1, "cHsSerial") = xeYes Then
'               MsgBox "Duplicate Serial No!!!" & vbCrLf & _
'                        "Please Verify your entry then try again!!!", vbCritical, "Warning"
'            Else
'               lnRep = MsgBox("Duplicate Serial No!!!" & vbCrLf & _
'                                 "Item automatically add from existing serial!!!", vbYesNo + vbQuestion, "CONFIRMATION")
'               If lnRep = vbYes Then
'                  lsValue = InputBox("Please specify quantity for serial " & Value & vbCrLf & _
'                                       .TextMatrix(lnCtr, 2) & vbCrLf & _
'                                       .TextMatrix(lnCtr, 3), "Quantity", 0)
'                  lnValue = IIf(lsValue = "", 0, lsValue)
'
'                  .TextMatrix(lnCtr, 6) = .TextMatrix(lnCtr, 6) + lnValue
'                  oTrans.Detail(lnCtr - 1, "nQuantity") = CDbl(.TextMatrix(lnCtr, 6))
'               End If
'            End If
'            compareSerial = ""
'         Else
'            compareSerial = Value
'         End If
'      Next
'   End With
'End Function
'
'Private Function BranchAutomate(ByVal sBranchCd As String) As Boolean
'   Dim lrs As Recordset
'
'   Set lrs = New Recordset
'   lrs.Open "SELECT * FROM Branch" & _
'               " WHERE sBranchCd = " & strParm(sBranchCd) & _
'                  " AND cAutomate = " & strParm(xeYes) _
'   , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'
'   If Not lrs.EOF Then BranchAutomate = True
'   Set lrs = Nothing
'End Function
'
'Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
'   With oApp
'      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
'      If bEnd Then
'         .xShowError
'         End
'      Else
'         With Err
'            .Raise .Number, .Source, .Description
'         End With
'      End If
'   End With
'End Sub
'
'
Private Sub lblField_Click(Index As Integer)

End Sub

VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{0B46E70A-7573-4847-A71B-876F1A303D14}#1.0#0"; "xrGridControl.ocx"
Begin VB.Form frmCPConsignment 
   BorderStyle     =   0  'None
   Caption         =   "CP Consignment Tagging"
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   120
      TabIndex        =   33
      Top             =   1815
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Post"
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
      Picture         =   "frmCPConsigment.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   135
      TabIndex        =   24
      Top             =   585
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCPConsigment.frx":077A
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3135
      Index           =   1
      Left            =   1620
      Tag             =   "wt0;fb0"
      Top             =   4455
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   5530
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin xrGridEditor.GridEditor GridEditor1 
         Height          =   3000
         Left            =   45
         TabIndex        =   21
         Top             =   60
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   5292
         AllowBigSelection=   -1  'True
         AutoAdd         =   -1  'True
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
         Object.HEIGHT          =   3000
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
         MOUSEICON       =   "frmCPConsigment.frx":0EF4
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
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3870
      Index           =   0
      Left            =   1605
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   6826
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   11
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2430
         Width           =   2505
      End
      Begin VB.CheckBox chk 
         Caption         =   "VATable"
         Height          =   195
         Index           =   10
         Left            =   1410
         TabIndex        =   8
         Tag             =   "wt0;fb0"
         Top             =   2790
         Width           =   1530
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2085
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1740
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   4695
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1050
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1425
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1050
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   660
         Index           =   5
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   3075
         Width           =   5955
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1425
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   705
         Width           =   2505
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1425
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   165
         Width           =   1950
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1395
         Width           =   5790
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "With Held"
         Height          =   285
         Index           =   8
         Left            =   165
         TabIndex        =   31
         Top             =   2475
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Add Discount"
         Height          =   285
         Index           =   7
         Left            =   165
         TabIndex        =   30
         Top             =   2130
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         Height          =   285
         Index           =   6
         Left            =   165
         TabIndex        =   29
         Top             =   1770
         Width           =   990
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Thru"
         Height          =   285
         Index           =   3
         Left            =   3930
         TabIndex        =   18
         Top             =   1095
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   17
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   5
         Left            =   135
         TabIndex        =   13
         Top             =   3105
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   285
         Index           =   0
         Left            =   165
         TabIndex        =   10
         Top             =   195
         Width           =   1200
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1515
         Tag             =   "et0;ht2"
         Top             =   270
         Width           =   1950
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. Date"
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   11
         Top             =   735
         Width           =   1200
      End
      Begin VB.Shape Shape3 
         Height          =   705
         Index           =   0
         Left            =   7875
         Top             =   180
         Width           =   2580
      End
      Begin VB.Shape Shape4 
         Height          =   600
         Index           =   0
         Left            =   7920
         Top             =   225
         Width           =   2490
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "unknown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Index           =   12
         Left            =   7905
         TabIndex        =   16
         Tag             =   "eb0;et0"
         Top             =   360
         Width           =   2520
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         Height          =   285
         Index           =   10
         Left            =   165
         TabIndex        =   12
         Top             =   1425
         Width           =   1200
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   525
         Index           =   0
         Left            =   7980
         Tag             =   "et0;et0"
         Top             =   270
         Width           =   2400
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   705
      Index           =   2
      Left            =   1620
      Tag             =   "wt0;fb0"
      Top             =   7635
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   1244
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   6
         Left            =   2100
         MaxLength       =   50
         TabIndex        =   20
         Tag             =   "ht0;ft0"
         Top             =   105
         Width           =   1695
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   7
         Left            =   7680
         MaxLength       =   50
         TabIndex        =   15
         Tag             =   "ht0;ft0"
         Text            =   "0.00"
         Top             =   120
         Width           =   3435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL QTY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   180
         TabIndex        =   19
         Top             =   165
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   6525
         TabIndex        =   14
         Top             =   180
         Width           =   1455
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   135
      TabIndex        =   25
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Search"
      AccessKey       =   "Search"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPConsigment.frx":0F10
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   135
      TabIndex        =   26
      Top             =   585
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Browse"
      AccessKey       =   "Browse"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPConsigment.frx":168A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   615
      Index           =   5
      Left            =   120
      TabIndex        =   27
      Top             =   3675
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1085
      Caption         =   "Close"
      AccessKey       =   "Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPConsigment.frx":1E04
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   135
      TabIndex        =   23
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCPConsigment.frx":257E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   120
      TabIndex        =   28
      Top             =   2445
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Pay"
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
      Picture         =   "frmCPConsigment.frx":2CF8
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   11
      Left            =   120
      TabIndex        =   32
      Top             =   3060
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Register"
      AccessKey       =   "R"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPConsigment.frx":3472
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   120
      TabIndex        =   22
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
      Picture         =   "frmCPConsigment.frx":3B6C
   End
End
Attribute VB_Name = "frmCpConsignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'she 2021-06-21 12:56pm
Option Explicit
Private Const pxeMODULENAME = "frmCPConsignment"

Private WithEvents oTrans As ggcCPPurchasing.clsConsignmentTagging
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pnCtr As Integer
Dim pbSave As Boolean
Dim pbGridValidate As Boolean
Dim pbPosted As Boolean

Private Sub chk_Click(Index As Integer)
   Select Case Index
   Case 10
      If chk(Index).Value = Checked Then
         oTrans.Master("cVATaxabl") = 1
      Else
         oTrans.Master("cVATaxabl") = 0
      End If
   End Select
End Sub

Private Sub cmdButton_Click(Index As Integer)
   
   Dim lsOldProc As String
   Dim lnRep As Integer
   Dim lnMsg As String

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   txtField_LostFocus pnIndex
      Select Case Index
         Case 0
            If isEntryOk Then
               If oTrans.SaveTransaction() Then
                  MsgBox "Record successfully saved!.", vbInformation, "Information!"
                  Call ClearFields
                  Call initButton(xeModeReady)
               Else
                  MsgBox "Unable to save transaction.", vbCritical, pxeMODULENAME
               End If
            End If
         Case 1 ' Search
            oTrans.SearchMaster pnIndex
         Case 8 ' Cancel
           lnMsg = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                           "Do you want to revert transaction???", vbYesNo + vbQuestion, "Confirm")
           If lnMsg = vbYes Then
             Call initButton(xeModeReady)
            ClearFields
            Call NewRecord
             setTransTat (-1)
           End If
         Case 4 ' New
            If oTrans.NewTransaction() Then
              ClearFields
              NewRecord
             Call initButton(xeModeAddNew)
            End If
         Case 5 'Closed
            Unload Me
         Case 3 'Browse
               If oTrans.SearchTransaction() = True Then
                  ClearFields
                  LoadMaster
                  loadMasterDetail
               End If
         Case 7
         If oTrans.Master("sCompnyNm") = "" Then Exit Sub
            If oTrans.UpdateTransaction Then
               initButton (xeModeUpdate)
            End If
         Case 6 'pay
         Case 11 'register
            frmConsignmentRegister.Show
         Case 2 ' confirm
      End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub NewRecord()
   Call LoadMaster
   Call ClearDetail
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      loTxt = ""
   Next
     
   ClearDetail
   setTransTat (-1)
End Sub

Private Sub ClearDetail()
Dim lnCtr As Integer
   With GridEditor1
      .Rows = 2
      
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "S.I"
      .TextMatrix(0, 3) = "Barrcode"
      .TextMatrix(0, 4) = "Descript"
      .TextMatrix(0, 5) = "Qty"
      .TextMatrix(0, 6) = "Refer#"
      .TextMatrix(0, 7) = "U.Price"
      .TextMatrix(0, 8) = "Tag"
      .TextMatrix(0, 9) = "trans#"
      .TextMatrix(0, 10) = "sstockIDx"
      
      .TextMatrix(1, 0) = "1"
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = ""
      .TextMatrix(1, 6) = ""
      .TextMatrix(1, 7) = ""
      .TextMatrix(1, 8) = ""
      .TextMatrix(1, 9) = ""
      .TextMatrix(1, 10) = ""
      
      For lnCtr = 1 To .Cols - 1
         .Col = lnCtr
         .CellBackColor = &HFFFFFF
         .CellFontBold = False
      Next
      
      .Row = 1
      .Col = 1
   End With
End Sub

Private Sub LoadMaster()
   With oTrans
      txtField(0).Text = IFNull(.Master("sTransNox"), "")
      txtField(1).Text = Format(IFNull(.Master("dTransact"), oApp.ServerDate), "MMMM DD, YYYY")
      txtField(2).Text = IFNull(.Master("sCompnyNm"), "")
      txtField(3).Text = Format(IFNull(.Master("dDateFrom"), oApp.ServerDate), "MMMM DD, YYYY")
      txtField(4).Text = Format(IFNull(.Master("dDateThru"), oApp.ServerDate), "MMMM DD, YYYY")
      txtField(5).Text = IFNull(.Master("sRemarksx"), "")
      txtField(7).Text = Format(IFNull(.Master("nTranTotl"), 0), "#,##0.00")
      txtField(8).Text = Format(IFNull(.Master("nDiscount"), 0), "#,##0.00")
      txtField(9).Text = Format(IFNull(.Master("nAddDiscx"), 0), "#,##0.00")
      txtField(11).Text = Format(IFNull(.Master("nTWithHld"), 0), "#,##0.00")
      If .Master("cVATaxabl") = 1 Then
         chk(10).Value = vbChecked
      Else
         chk(10).Value = vbUnchecked
      End If
      setTransTat (.Master("cTranStat"))
   End With
End Sub


Private Function isEntryOk()
   Dim lnCtr As Integer
   With oTrans
      If IFNull(.Master("sCompnyNm"), "") = "" Then
         MsgBox "No Supplier entry detected!!!" & vbCrLf & _
             "Please check entry and try again!", vbCritical, "Warning"
              txtField(2).SetFocus
         GoTo EntryNotOK
        End If

        If txtField(7).Text <= 0 Then
         MsgBox "Please select item to pay!!!" & vbCrLf & _
             "Please check detail and try again!", vbCritical, "Warning"
         GoTo EntryNotOK
        End If
   End With
   
   For lnCtr = 1 To GridEditor1.Rows - 1
      If Trim(GridEditor1.TextMatrix(lnCtr, 8)) = "Yes" Then
         oTrans.Detail(oTrans.ItemCount - 1, "sReferNox") = Trim(GridEditor1.TextMatrix(lnCtr, 6))
         oTrans.Detail(oTrans.ItemCount - 1, "sSourceNo") = Trim(GridEditor1.TextMatrix(lnCtr, 9))
         oTrans.Detail(oTrans.ItemCount - 1, "sStockIDx") = Trim(GridEditor1.TextMatrix(lnCtr, 10))
         oTrans.Detail(oTrans.ItemCount - 1, "nItemQtyx") = CInt(GridEditor1.TextMatrix(lnCtr, 5))
         oTrans.Detail(oTrans.ItemCount - 1, "nUnitPrce") = CDbl(GridEditor1.TextMatrix(lnCtr, 7))
         If oTrans.ItemCount >= 1 Then
            oTrans.addDetail
         End If
      End If
   Next
   
   For lnCtr = 0 To oTrans.ItemCount - 1
     If Trim(oTrans.Detail(lnCtr, "sStockIDx")) = "" Then
         If oTrans.deleteDetail(lnCtr) Then
         End If
      End If
   Next

EntryOK:
   isEntryOk = True
   Exit Function
EntryNotOK:
   isEntryOk = False
   
End Function

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
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
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   '''On Error GoTo errProc

   CenterChildForm mdiMain, Me
   
   Set oTrans = New ggcCPPurchasing.clsConsignmentTagging
   Set oTrans.AppDriver = oApp
   
   oTrans.TransStatus = xeStateOpen
   oTrans.InitTransaction
   oTrans.NewTransaction
    initButton (xeModeReady)

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   InitGrid
   InitForm
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer

   With GridEditor1
      .Cols = 11
      .Font = "MS Sans Serif"

      'Column Title
      .TextMatrix(0, 0) = "No"
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "S.I"
      .TextMatrix(0, 3) = "Barrcode"
      .TextMatrix(0, 4) = "Descript"
      .TextMatrix(0, 5) = "Qty"
      .TextMatrix(0, 6) = "Refer#"
      .TextMatrix(0, 7) = "U.Price"
      .TextMatrix(0, 8) = "Tag"
      .TextMatrix(0, 9) = "trans#"
      .TextMatrix(0, 10) = "sStockIDx"
      .Row = 0

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 4
      Next

      .ColWidth(0) = 450
      .ColWidth(1) = 2500
      .ColWidth(2) = 800
      .ColWidth(3) = 2000
      .ColWidth(4) = 2500
      .ColWidth(5) = 500
      .ColWidth(6) = 1300
      .ColWidth(7) = 800
      .ColWidth(8) = 450
      .ColWidth(9) = 0
      .ColWidth(10) = 0

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 1
      .ColAlignment(8) = 1
      
      .Row = .Rows - 1
      .Col = 1
      
   End With
   
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(11).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow
   cmdButton(2).Visible = Not lbShow
   
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(8).Visible = lbShow

   xrFrame1(0).Enabled = lbShow
   xrFrame1(1).Enabled = lbShow
   
   If lbShow Then txtField(1).SetFocus
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

Private Sub GridEditor1_DblClick()
   Dim lnCtr As Integer
   With GridEditor1
   
      If Trim(.TextMatrix(.Row, 6)) = "" Then
         .Row = .Row
         .Col = 6
'         .SetFocus
         Exit Sub
      End If
   
      If .TextMatrix(.Row, 8) = "No" Then
         .TextMatrix(.Row, 8) = "Yes"
         oTrans.Detail(.Row, "nUnitPrce") = CDbl(.TextMatrix(.Row, 7))
'         oTrans.addDetail
      Else
         .TextMatrix(.Row, 8) = "Yes"
         .TextMatrix(.Row, 8) = "No"
'         oTrans.deleteDetail
      End If
      
      'highlight selected rows
      For lnCtr = 1 To .Cols - 1
         .Col = lnCtr
         If .TextMatrix(.Row, 8) = "No" Then
            .CellBackColor = &HFFFFFF
            .CellFontBold = False
         Else
            .CellBackColor = &HFF80FF
            .CellFontBold = True
            
         End If
      Next
   End With
   Call ComputeTotal
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   Call ComputeTotal
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   Select Case Index
   Case 15
      txtField(2).Text = oTrans.Master(Index)
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   txtField(Index).BackColor = &HC0FFFF
   Select Case Index
      Case 1, 3, 4
      txtField(Index).Text = Format(txtField(Index).Text, "MM/DD/YY")
   End Select
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 2 Then
      If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
         oTrans.SearchMaster 2, txtField(2).Text
         Call LoadDetail("", "", "")
      End If
   End If
End Sub

Public Function setTransTat(nStat As Integer) As String
    Select Case nStat
       Case 0
          lblField(12) = "OPEN"
       Case 1
          lblField(12) = "CLOSED"
       Case 2
          lblField(12) = "POSTED"
       Case 3
          lblField(12) = "CANCELLED"
       Case 4
          lblField(12) = "VOID"
       Case Else
          lblField(12) = "UNKNOWN"
    End Select
End Function

Private Sub txtField_LostFocus(Index As Integer)
   txtField(Index).BackColor = &HFFFFFF
End Sub

Private Sub LoadDetail(lsClientID As String, lsDateFrom As String, lsDateThru As String)
   Dim lnCtr As Integer
   
   If Not oTrans.DetailCount Then
      With GridEditor1
         .Rows = oTrans.DetailCount + 1
         For lnCtr = 0 To oTrans.DetailCount - 1
            .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
            .TextMatrix(lnCtr + 1, 1) = IFNull(oTrans.Others(lnCtr, "sBranchNm"), "")
            .TextMatrix(lnCtr + 1, 2) = IFNull(oTrans.Others(lnCtr, "sSalesInv"), "")
            .TextMatrix(lnCtr + 1, 3) = IFNull(oTrans.Others(lnCtr, "sBarrCode"), "")
            .TextMatrix(lnCtr + 1, 4) = IFNull(oTrans.Others(lnCtr, "sDescript"), "")
            .TextMatrix(lnCtr + 1, 5) = IFNull(oTrans.Others(lnCtr, "nItemQtyx"), 0)
            .TextMatrix(lnCtr + 1, 6) = IFNull(oTrans.Others(lnCtr, "sReferNox"), "")
            .TextMatrix(lnCtr + 1, 7) = Format(IFNull(oTrans.Others(lnCtr, "nUnitPrce"), 0), "#,##0.00")
            .TextMatrix(lnCtr + 1, 8) = "No"
            .TextMatrix(lnCtr + 1, 9) = IFNull(oTrans.Others(lnCtr, "sSourceNo"), "")
            .TextMatrix(lnCtr + 1, 10) = IFNull(oTrans.Others(lnCtr, "sStockIDx"), "")
         Next
      End With
   End If
End Sub

Private Sub loadMasterDetail()
   Dim lnCtr As Integer
   Dim lsSQL As String
   Dim lorec As Recordset
   
   With GridEditor1
   If oTrans.ItemCount = 0 Then
      ClearDetail
      Else
         .Rows = 1
            For pnCtr = 0 To oTrans.ItemCount - 1
               .Rows = .Rows + 1
               .TextMatrix(.Rows - 1, 0) = .Rows - 1
               
               lsSQL = "SELECT b.sBranchNm, a.sSalesInv" & _
                        " FROM CP_SO_Master a" & _
                           " LEFT JOIN Branch b ON LEFT(a.sTransNox,4) = b.sBranchCd" & _
                        " WHERE a.sTransNox = " & strParm(IFNull(oTrans.Detail(pnCtr, "sSourceNo"), ""))
               Set lorec = New Recordset
               lorec.Open lsSQL, oApp.Connection, , , adCmdText
               
               .TextMatrix(.Rows - 1, 0) = .Rows - 1
               If Not lorec.EOF Then
                  .TextMatrix(.Rows - 1, 1) = lorec("sBranchNm")
                  .TextMatrix(.Rows - 1, 2) = lorec("sSalesInv")
               Else
                  .TextMatrix(.Rows - 1, 1) = ""
                  .TextMatrix(.Rows - 1, 2) = ""
               End If
               
               .TextMatrix(.Rows - 1, 3) = IFNull(oTrans.Detail(pnCtr, "sBarrCode"), "")
               .TextMatrix(.Rows - 1, 4) = Format(oTrans.Detail(pnCtr, "sDescript"), "")
               .TextMatrix(.Rows - 1, 5) = IFNull(oTrans.Detail(pnCtr, "nItemQtyx"), 0)
               .TextMatrix(.Rows - 1, 6) = IFNull(oTrans.Detail(pnCtr, "sReferNox"), "")
               .TextMatrix(.Rows - 1, 7) = Format(IFNull(oTrans.Detail(pnCtr, "nUnitPrce"), 0), "#,##0.00")
               .TextMatrix(.Rows - 1, 8) = "Yes"
               .TextMatrix(.Rows - 1, 9) = IFNull(oTrans.Detail(pnCtr, "sSourceNo"), "")
               .TextMatrix(.Rows - 1, 10) = IFNull(oTrans.Detail(lnCtr, "sStockIDx"), "")
            Next
   End If
   End With
   ComputeTotal
End Sub

Private Sub InitForm()
   Dim lnCtr As Integer
   For lnCtr = 0 To txtField.Count - 1
      Select Case lnCtr
      Case 0
         txtField(lnCtr).Text = GetNextCode("CP_Consignment_Payment_Master", "sTransNox", True, oApp.Connection, True, oApp.BranchCode)
      Case 1, 3, 4
         txtField(lnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 2, 5
         txtField(lnCtr).Text = ""
      Case 6
          If Not IsNumeric(txtField(lnCtr)) Then
            txtField(lnCtr).Text = 0
         End If
      Case 7, 8, 9
         If Not IsNumeric(txtField(lnCtr)) Then
            txtField(lnCtr).Text = Format(0, "#,##0.00")
         End If
            
      txtField(11).Text = Format(0, "#,##0.00")
      End Select
   Next
   
'   CreateTempTable
End Sub

Private Sub ComputeTotal()
   Dim lnCtr As Integer
   Dim lnTotalQty As Double
   Dim lnTotalAmt As Currency
   
   lnTotalQty = 0
   lnTotalAmt = 0#
   With GridEditor1
      For lnCtr = 1 To .Rows - 1
         If .TextMatrix(lnCtr, 8) = "Yes" Then
            lnTotalQty = lnTotalQty + .TextMatrix(lnCtr, 5)
            lnTotalAmt = lnTotalAmt + (.TextMatrix(lnCtr, 5) * .TextMatrix(lnCtr, 7))
         End If
      Next
   End With
   txtField(6).Text = CInt(lnTotalQty)
   txtField(7).Text = Format(lnTotalAmt, "#,##0.00")
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
With txtField(Index)
   Select Case Index
      Case 2
      Case 6
      Case 12
      Case 1, 3, 4
          If Not IsDate(.Text) Then .Text = oApp.ServerDate
            .Text = Format(.Text, "MMMM DD, YYYY")
      
            oTrans.Master(Index) = CDate(.Text)
      Case 5
         If Trim(.Text) <> "" Then
            .Text = Replace(TitleCase(.Text), vbCrLf, " ")
         End If
         oTrans.Master(Index) = .Text
      Case 8, 9, 11
         If Not IsNumeric(.Text) Then .Text = 0#
            .Text = Format(.Text, "#,##0.00")
            oTrans.Master(Index) = CDbl(.Text)
      Case Else
            oTrans.Master(Index) = .Text
      End Select
   End With
End Sub

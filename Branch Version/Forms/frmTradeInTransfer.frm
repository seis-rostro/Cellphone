VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmTradeInTransfer 
   BorderStyle     =   0  'None
   Caption         =   "CP Trade In Transfer"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   105
      TabIndex        =   36
      Top             =   1815
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
      Picture         =   "frmTradeInTransfer.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   105
      TabIndex        =   35
      Top             =   2430
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
      Picture         =   "frmTradeInTransfer.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   105
      TabIndex        =   32
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
      Picture         =   "frmTradeInTransfer.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   105
      TabIndex        =   34
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
      Picture         =   "frmTradeInTransfer.frx":166E
   End
   Begin xrControl.xrFrame xrFrame 
      Height          =   5895
      Index           =   1
      Left            =   1635
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   10398
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.Frame Frame3 
         Caption         =   "Transaction Details"
         Height          =   3030
         Left            =   105
         TabIndex        =   21
         Tag             =   "wt0;fb0"
         Top             =   75
         Width           =   3780
         Begin VB.TextBox txtField 
            Height          =   885
            Index           =   4
            Left            =   1275
            MultiLine       =   -1  'True
            TabIndex        =   4
            Text            =   "frmTradeInTransfer.frx":1DE8
            Top             =   2040
            Width           =   2265
         End
         Begin VB.TextBox txtField 
            Height          =   375
            Index           =   3
            Left            =   1275
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1650
            Width           =   2265
         End
         Begin VB.TextBox txtField 
            Height          =   375
            Index           =   2
            Left            =   1275
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   1245
            Width           =   2265
         End
         Begin VB.TextBox txtField 
            Height          =   375
            Index           =   1
            Left            =   1275
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   855
            Width           =   2265
         End
         Begin VB.TextBox txtField 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   0
            Left            =   1275
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   270
            Width           =   2265
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   390
            Left            =   1335
            Tag             =   "et0;ht2"
            Top             =   345
            Width           =   2280
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   270
            TabIndex        =   27
            Top             =   2055
            Width           =   1050
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Destination:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   45
            TabIndex        =   25
            Top             =   1305
            Width           =   1200
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   570
            TabIndex        =   24
            Top             =   885
            Width           =   615
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Reference:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   165
            TabIndex        =   23
            Top             =   1695
            Width           =   1050
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Trans No:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   270
            TabIndex        =   22
            Top             =   330
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Product Details"
         Height          =   2670
         Left            =   105
         TabIndex        =   10
         Tag             =   "wt0;fb0"
         Top             =   3090
         Width           =   3780
         Begin VB.TextBox txtDetail 
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
            Index           =   1
            Left            =   1260
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   270
            Width           =   2265
         End
         Begin VB.TextBox txtDetail 
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   1260
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   825
            Width           =   2265
         End
         Begin VB.TextBox txtDetail 
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   1260
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   1245
            Width           =   2265
         End
         Begin VB.TextBox txtDetail 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   5
            Left            =   1260
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   2040
            Width           =   2265
         End
         Begin VB.TextBox txtDetail 
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   1260
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1650
            Width           =   2265
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Imei/Serial:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   60
            TabIndex        =   26
            Top             =   300
            Width           =   900
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "IMEI:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   675
            TabIndex        =   15
            Top             =   -285
            Width           =   495
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Color:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   420
            TabIndex        =   14
            Top             =   1695
            Width           =   630
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Brand:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   390
            TabIndex        =   13
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Price:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   75
            TabIndex        =   12
            Top             =   2085
            Width           =   915
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Model:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   375
            TabIndex        =   11
            Top             =   1260
            Width           =   735
         End
      End
   End
   Begin xrControl.xrFrame xrFrame 
      Height          =   5895
      Index           =   2
      Left            =   5670
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   10398
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.Frame Frame2 
         Caption         =   "Cellphone Trade In"
         Height          =   5295
         Left            =   105
         TabIndex        =   16
         Tag             =   "wt0;fb0"
         Top             =   150
         Width           =   6720
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
            Height          =   4935
            Left            =   75
            TabIndex        =   28
            Top             =   270
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   8705
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            AllowBigSelection=   0   'False
            FocusRect       =   0
            ScrollBars      =   0
            SelectionMode   =   1
         End
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No.:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -6480
         TabIndex        =   20
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "000001"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -5640
         TabIndex        =   19
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GRAND TOTAL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2565
         TabIndex        =   18
         Top             =   5460
         Width           =   2310
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "10,000.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4950
         TabIndex        =   17
         Top             =   5445
         Width           =   1680
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   105
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1815
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmTradeInTransfer.frx":1DEE
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   105
      TabIndex        =   30
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
      Picture         =   "frmTradeInTransfer.frx":2568
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   105
      TabIndex        =   31
      Top             =   1200
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
      Picture         =   "frmTradeInTransfer.frx":2CE2
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   105
      TabIndex        =   33
      Top             =   2430
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmTradeInTransfer.frx":345C
   End
End
Attribute VB_Name = "frmTradeInTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Const pxeMODULENAME = "frmTradeInTransfer"
'Private Declare Function GetFocus Lib "USER32" () As Long
'
'Private WithEvents oTrans As ggcCPSales.clsTITUTransfer
'
'Private oSkin As clsFormSkin
'
'Private pnIndex As Integer
'Private pnActiveRow As Integer
'Private pDIndex As Integer
'Private pnRow As Integer
'Private pnCtr As Integer
'
'Private pbControl As Boolean
'Private pbMasterGotFocus As Boolean
'Private pbGridGotFocus As Boolean
'Private pbFormLoad As Boolean
'Dim pbLoadRecord As Boolean
'
'Private psBranchCd As String
'
'Private Sub InitGrid()
'    Dim lnCtr As Integer
'
'    With MSFlexGrid2
'        .Cols = 4
'        .Rows = 2
'        .Clear
'
'        .TextMatrix(0, 0) = "No."
'        .TextMatrix(0, 1) = "IMEI No."
'        .TextMatrix(0, 2) = "Description"
'        .TextMatrix(0, 3) = "Unit Price"
'
'        .Row = 0
'        .ColWidth(0) = 800
'        .ColWidth(1) = 1700
'        .ColWidth(2) = 2800
'        .ColWidth(3) = 1100
'
'        For lnCtr = 0 To .Cols - 1
'         .Col = lnCtr
'         .CellFontBold = True
'         .CellAlignment = 3
'      Next
'
'      .Row = 1
'      .ColAlignment(0) = flexAlignCenterCenter
'      .ColAlignment(1) = flexAlignLeftCenter
'      .ColAlignment(2) = flexAlignLeftCenter
'      .ColAlignment(3) = flexAlignRightCenter
'
'      .Col = 1
'      .Row = 1
'      .ColSel = .Cols - 1
'   End With
'
'End Sub
'
'Private Sub InitTransaction()
''   Call InitGrid
'   Call ClearFields
'   Call ClearDetail
'
'   oTrans.InitTransaction
'   oTrans.NewTransaction
'   InitButton (0)
'End Sub
'
'Private Sub cmdButton_Click(Index As Integer)
'   Dim lsOldProc As String
'   Dim lnMsg As String
'
'   lsOldProc = "cmdButton_Click"
'   ' 'On Error GoTo errProc
'   With MSFlexGrid2
'      Select Case Index
'         Case 5 'Close
'            Unload Me
'         Case 0 ' ok
'            If oTrans.Master("sDestinat") <> "" And txtField(2).Text <> "" Then
'                  If .Rows > 2 Then
'                     pnCtr = 1
'                     Do While pnCtr < .Rows
'                        If Trim(.TextMatrix(pnCtr, 1)) = "" Then
'                           .Row = pnCtr
'                           If oTrans.deleteDetail(.Row - 1) Then Call deleteGridRow
'                        Else
'                           pnCtr = pnCtr + 1
'                        End If
'                     Loop
'                  End If
'                  With oTrans
'                  If .SaveTransaction Then
'                     MsgBox "Transaction Saved Successfully.", vbInformation, "Notice"
'                     If MsgBox("Do you want to create a printable documents" & vbCrLf & _
'                     "for this transaction?", _
'                        vbQuestion + vbYesNo, "Confirm") = vbYes Then
'                        If PrintTrans = True Then
'                        Else
'                           MsgBox "Unable to create printable document for this Transaction.", vbCritical, "Error"
'                        End If
'                     End If
'                     InitTransaction
'                  Else
'                     MsgBox "Unable to Save Transaction.", vbInformation, "Notice"
'                  End If
'               End With
'            Else
'               MsgBox "Invalid Destination Branch!!!", vbCritical, "Warning"
'               txtField(2).SetFocus
'            End If
'         Case 1 ' Del. Row
'            lnMsg = MsgBox("Do you want to delete this item?", vbYesNo + vbQuestion, "Confirm")
'            With MSFlexGrid2
'               If lnMsg = vbYes Then
'                   If .Rows > 2 Then
'                        If oTrans.deleteDetail(pnActiveRow - 1) Then deleteGridRow
'                        For pnCtr = 1 To .Rows - 1
'                           .TextMatrix(pnCtr, 0) = pnCtr
'                        Next
'                  Else
'                     Call ClearDetail
'                  End If
'                  Label11.Caption = Format(TotalColumn(MSFlexGrid2, 3), "#,##0.00")
'               End If
'            End With
'         Case 2 ' Search
'            If pnIndex = 2 Then
'                oTrans.SearchMaster Index, txtField(Index).Text
'                If txtField(Index).Text <> "" Then SetNextFocus
'               Exit Sub
'            End If
'
'            If pDIndex = 1 Then
'               If Trim(oTrans.Detail(pnActiveRow - 1, "sSerialNo")) <> "" Then
'                    oTrans.searchDetail pnActiveRow - 1, "sSerialNo", txtField(Index).Text
'                    If txtField(Index).Text <> "" Then
'                        SetNextFocus
'                        Call addDetail
'                    Else
'                        If txtField(Index).Text <> "" Then oTrans.SearchMaster Index, txtField(Index).Text
'                  End If
'               End If
'            End If
'
'         Case 6 ' Cancel
'         Dim lnRep As String
'            lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
'                        "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")
'               If lnRep = vbYes Then
'                  InitButton 1
'               End If
'         Case 3 ' Browse
'             If oTrans.SearchTransaction = True Then
'               LoadMaster
'               LoadDetail
'               showdetail
'             End If
'         Case 4 'New
'            InitTransaction
'         Case 7 'Update
'            frmCPTradeInTransferReg.Tag = "mnuTradeInTransfer"
'            frmCPTradeInTransferReg.Show 1
'      End Select
'   End With
'End Sub
'
'Private Sub LoadMaster()
'  Dim lnCtr As Integer
'   With txtField
'      For lnCtr = 0 To 4
'         Select Case lnCtr
'            Case 0
'               txtField(lnCtr).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@-@@@@@@")
'            Case 1
'               txtField(lnCtr).Text = strLongDate(oTrans.Master("dTransact"))
'            Case 2
'               txtField(lnCtr).Text = oTrans.Master(lnCtr)
'            Case 3
'               txtField(lnCtr).Text = oTrans.Master(lnCtr)
'            Case 4
'               txtField(lnCtr).Text = oTrans.Master(lnCtr)
'         End Select
'      Next
'   End With
'End Sub
'
'Private Sub deleteGridRow()
'   Dim lnLastRow As Integer
'   Dim lnCtr As Integer
'
'   With MSFlexGrid2
'      .Rows = .Rows - 1
'
'      lnLastRow = .Rows - 1
'      For pnCtr = 0 To oTrans.ItemCount - 1
'         .TextMatrix(pnCtr + 1, 0) = pnCtr + 1
'         .TextMatrix(pnCtr + 1, 1) = IFNull(oTrans.Detail(pnCtr, "sSerialNo"), "")
'         .TextMatrix(pnCtr + 1, 2) = IFNull(oTrans.Detail(pnCtr, "sBrandNme"), "") + IFNull(oTrans.Detail(pnCtr, "sModelNme"), "") + IFNull(oTrans.Detail(pnCtr, "sColorNme"), "")
'         .TextMatrix(pnCtr + 1, 3) = IFNull(oTrans.Detail(pnCtr, "nUnitPrce"), 0)
'
'         .Row = pnCtr + 1
'         If (pnCtr + 1) Mod 2 = 0 Then
'            For lnCtr = 1 To .Cols - 1
'               .Col = lnCtr
'               .CellBackColor = oApp.getColor("fb0")
'            Next
'         End If
'      Next
'
'      .Row = 1
'      .Col = 1
'      .ColSel = .Cols - 1
'
'      pnActiveRow = .Row
'   End With
'End Sub
'
'Private Sub ClearFields()
'   Dim loTxt As TextBox
'
'   For Each loTxt In txtField
'      loTxt = ""
'   Next
'
'   Dim loDetail As TextBox
'
'   For Each loDetail In txtDetail
'      loDetail = ""
'   Next
'
'   pnIndex = -1
'   pnRow = -1
'   pDIndex = -1
'   Label11.Caption = "0.00"
'   txtDetail(5).Text = "0.00"
'
'   End Sub
'
'Private Sub Form_Activate()
' Dim lsOldProc As String
'
'   lsOldProc = "Form_Activate"
'   ' 'On Error GoTo errProc
'
'   oApp.MenuName = Me.Tag
'
'   Me.ZOrder 0
'
'   If Not pbFormLoad Then
'      pbFormLoad = True
'    End If
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'      Case vbKeyReturn, vbKeyUp, vbKeyDown
'         Select Case KeyCode
'         Case vbKeyReturn, vbKeyDown
'               SetNextFocus
'         Case vbKeyUp
'            SetPreviousFocus
'         End Select
'      End Select
'End Sub
'
'Private Sub Form_Load()
'   Dim lsOldProc As String
'
'   lsOldProc = "Form_Load"
'   ' 'On Error GoTo errProc
'
'   CenterChildForm mdiMain, Me
'
'   Set oTrans = New ggcCPSales.clsTITUTransfer
'   Set oTrans.AppDriver = oApp
'
'   Set oSkin = New clsFormSkin
'   Set oSkin.AppDriver = oApp
'   Set oSkin.Form = Me
'
'   oSkin.ApplySkin xeFormTransEqualLeft
'
'   oTrans.Branch = oApp.BranchCode
'   Call InitGrid
'   Call InitTransaction
'
'
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc
'End Sub
'
'Private Sub showdetail()
'   With MSFlexGrid2
'      If .TextMatrix(pnActiveRow, 1) <> "" Then
'         txtDetail(1).Text = IFNull(oTrans.Detail(pnActiveRow - 1, "sSerialNo"), "")
'         txtDetail(2).Text = IFNull(oTrans.Detail(pnActiveRow - 1, "sBrandNme"), "")
'         txtDetail(3).Text = IFNull(oTrans.Detail(pnActiveRow - 1, "sModelNme"), "")
'         txtDetail(4).Text = IFNull(oTrans.Detail(pnActiveRow - 1, "sColorNme"), "")
'         txtDetail(5).Text = IIf(oTrans.Detail(pnActiveRow - 1, "nUnitPrce") = "", 0#, IFNull(Format(oTrans.Detail(pnActiveRow - 1, "nUnitPrce"), "#,##0.00"), "0"))
'      Else
'         txtDetail(1).Text = ""
'         txtDetail(2).Text = ""
'         txtDetail(3).Text = ""
'         txtDetail(4).Text = ""
'         txtDetail(5).Text = "0.0"
'      End If
'   End With
'End Sub
'
'Private Sub addDetail()
'   Dim lsOldProc As String
'   Dim lnRow As Integer
'   Dim lnCtr As Integer
'
'   lnRow = oTrans.ItemCount
'
'   lsOldProc = pxeMODULENAME & "addDetail"
'   '''On Error GoTo errProc
'
'   With MSFlexGrid2
'      If .Rows - 1 <> .Row Then
'         LoadDetail
'         Exit Sub
'      End If
'
'      If (oTrans.Detail(pnActiveRow - 1, "sSerialNo") = "") Then
'         MsgBox "No IMEI Entry Detected!" & vbCrLf & _
'                    "Pls Verify Entry Then Try Again!!!", vbCritical, "WARNING"
'                    txtDetail(1).SetFocus
'         Exit Sub
'      End If
'
'
'      If oTrans.addDetail Then
'         .Rows = .Rows + 1
'         Call LoadDetail
'      End If
'
'      .TextMatrix(.Rows - 1, 0) = .Rows - 1
'      .Row = .Rows - 1
'      .Col = 1
'      .ColSel = .Cols - 1
'   End With
'   txtDetail(1).SetFocus
'endProc:
'   Exit Sub
'errProc:
'   ShowError lsOldProc & "( " & " )", True
'End Sub
'
'Private Sub MSFlexGrid2_Click()
'Dim lnCtr As Integer
'   With oTrans
'      pnActiveRow = MSFlexGrid2.Row
'      Call showdetail
'   End With
'
'End Sub
'
'Private Sub MSFlexGrid2_GotFocus()
'   pbGridGotFocus = True
'End Sub
'
'Private Sub MSFlexGrid2_RowColChange()
'If Not pbFormLoad Then Exit Sub
'   pnActiveRow = MSFlexGrid2.Row
'   Call showdetail
'End Sub
'
'Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
'   With MSFlexGrid2
'         Select Case Index
'         Case 1
'            txtDetail(Index) = IFNull(oTrans.Detail(pnActiveRow - 1, "sSerialNo"), "")
'         Case 2
'            txtDetail(Index) = IFNull(oTrans.Detail(pnActiveRow - 1, "sBrandNme"), "")
'         Case 3
'            txtDetail(Index) = IFNull(oTrans.Detail(pnActiveRow - 1, "sModelNme"), "")
'         Case 4
'            txtDetail(Index) = IFNull(oTrans.Detail(pnActiveRow - 1, "sColorNme"), "")
'         Case 5
'            txtDetail(Index) = IFNull(Format(oTrans.Detail(pnActiveRow - 1, "nUnitPrce"), "#,##0.00"), 0)
'         End Select
'         LoadDetail
'      End With
'End Sub
'
'Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
'   With txtField
'      Select Case Index
'         Case 0
'            txtField(Index) = Format(oTrans.Master("sTransNox"), "@@@@-@@-@@@@@@")
'         Case 1
'            txtField(Index) = strLongDate(oTrans.Master("dTransact"))
'         Case Else
'            txtField(Index) = oTrans.Master(Index)
'      End Select
'   End With
'End Sub
'
'Private Sub txtDetail_GotFocus(Index As Integer)
'   Select Case Index
'      Case 1
'         Call HighlightOn(Me.txtDetail(Index))
'   End Select
'   pDIndex = Index
'End Sub
'
'Private Sub txtDetail_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    With MSFlexGrid2
'      Select Case Index
'         Case 1
'            Select Case KeyCode
'               Case vbKeyF3
'                  If oTrans.searchDetail(pnActiveRow - 1, Index, txtDetail(Index)) Then
'                     txtDetail(1).SetFocus
'                  End If
'               Case vbKeyReturn
'                     If .Row = .Rows - 1 Then
'                     If oTrans.searchDetail(pnActiveRow - 1, Index, txtDetail(Index)) Then
'                        txtDetail(1).SetFocus
'                     End If
'                        Call addDetail
'                     Else
'                        .Row = .Rows - 1
'                        .Col = 0
'                        .ColSel = .Cols - 1
'                        showdetail
'               End If
'            End Select
'         End Select
'   End With
'End Sub
'
'Private Sub txtDetail_LostFocus(Index As Integer)
'   Select Case Index
'      Case 1
'         Call HighlightOff(Me.txtDetail(Index))
'   End Select
'End Sub
'
'Private Sub txtField_GotFocus(Index As Integer)
'   With txtField(Index)
'       .BackColor = oApp.getColor("HT1")
'       .SelStart = 0
'       .SelLength = Len(.Text)
'   End With
'    pnIndex = Index
'End Sub
'
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set oSkin = Nothing
'   Set oTrans = Nothing
'   pbFormLoad = False
'End Sub
'
'Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
'   With oApp
'      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
'      If bEnd Then
'         .xShowError
'      Else
'         With Err
'            .Raise .Number, .Source, .Description
'         End With
'      End If
'   End With
'End Sub
'
'Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   Dim lsOldProc As String
'
'   lsOldProc = "txtField_KeyDown"
'   ''' 'On Error GoTo errProc
'      If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
'            Select Case Index
'               Case 2
'                  If KeyCode = vbKeyF3 Then
'                     oTrans.SearchMaster Index, txtField(Index).Text
'                     If txtField(Index).Text <> "" Then SetNextFocus
'                  Else
'                     If txtField(Index).Text <> "" Then oTrans.SearchMaster Index, txtField(Index).Text
'                  End If
'            End Select
'         KeyCode = 0
'      End If
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
'Private Sub txtField_LostFocus(Index As Integer)
'   With txtField(Index)
'      .BackColor = oApp.getColor("EB")
'   End With
'   LoadDetail
'End Sub
'
'Public Sub LoadDetail()
'   Dim lnRow As Integer
'   Dim lnCtr As Integer
'
'   With MSFlexGrid2
'      If oTrans.ItemCount = 0 Then
'         .Rows = 2
'      Else
'         .Rows = oTrans.ItemCount + 1
'      End If
'
'      For lnCtr = 0 To oTrans.ItemCount - 1
'         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
'         .TextMatrix(lnCtr + 1, 1) = IFNull(oTrans.Detail(lnCtr, "sSerialNo"), "")
'         .TextMatrix(lnCtr + 1, 2) = IFNull(oTrans.Detail(lnCtr, "sBrandNme"), "") + " " + IFNull(oTrans.Detail(lnCtr, "sModelNme"), "") + " " + IFNull(oTrans.Detail(lnCtr, "sColorNme"), "")
'         .TextMatrix(lnCtr + 1, 3) = IIf(IFNull(oTrans.Detail(lnCtr, "nUnitPrce"), "") = "", 0, Format(oTrans.Detail(lnCtr, "nUnitPrce"), "#,##0.00"))
'      Next
'   End With
'
'   Label11.Caption = Format(TotalColumn(MSFlexGrid2, 3), "#,##0.00")
'End Sub
'
'Private Function TotalColumn(grid As MSFlexGrid, ByVal ColIndex As Integer) As Integer
'  Dim R As Long
'  Dim Total As Integer
'
'  For R = 0 To grid.Rows - 1
'    If IsNumeric(grid.TextMatrix(R, ColIndex)) Then
'      Total = Total + CDbl(grid.TextMatrix(R, ColIndex))
'    End If
'  Next R
'
'  TotalColumn = Total
'End Function
'
'Private Sub ClearDetail()
'   With MSFlexGrid2
'      .Rows = 2
'      .TextMatrix(1, 0) = 1
'      .TextMatrix(1, 1) = ""
'      .TextMatrix(1, 2) = ""
'      .TextMatrix(1, 3) = 0
'
'      .Col = 1
'      .Row = 1
'      .ColSel = .Cols - 1
'      pnActiveRow = 1
'   End With
'End Sub
'
'Private Sub InitButton(lnStat As Integer)
'   Dim lbShow As Boolean
'
'   lbShow = IIf(lnStat = 0, False, True)
'   cmdButton(0).Visible = Not lbShow
'   cmdButton(2).Visible = Not lbShow
'   cmdButton(1).Visible = Not lbShow
'   cmdButton(6).Visible = Not lbShow
'   xrFrame(1).Enabled = Not lbShow
'
'   MSFlexGrid2.Enabled = Not lbShow
'
'   cmdButton(3).Visible = lbShow
'   cmdButton(4).Visible = lbShow
'   cmdButton(7).Visible = lbShow
'   cmdButton(5).Visible = lbShow
'
'   If lnStat = 0 Then LoadMaster
'
'End Sub
'
'Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
'   With txtField(Index)
'      Select Case Index
'      Case 1
'         If Not IsDate(txtField(Index).Text) Then txtField(Index).Text = oApp.ServerDate
'         oTrans.Master(Index) = txtField(Index).Text
'         txtField(Index).Text = Format(txtField(Index).Text, "MMMM DD, YYYY")
'      Case 4
'         txtField(Index).Text = TitleCase(txtField(Index).Text)
'         oTrans.Master(Index) = txtField(Index).Text
'      Case Else
'         oTrans.Master(Index) = txtField(Index).Text
'      End Select
'   End With
'End Sub
'
'Private Function PrintTrans() As Boolean
'   Dim loreport As frmRepViewer
'   Dim lrs As Recordset
'   Dim lors As Recordset
'   Dim lnCtr As Integer
'   Dim lsOldProc As String
'   Dim lsSourceNo As String
'
'   lsOldProc = "PrinTrans"
'   ''On Error GoTo errProc
'
'   PrintTrans = False
'
'   If oTrans.Master("sDestinat") = "" Then Exit Function
'
'   Set lrs = New ADODB.Recordset
'
'   lrs.Fields.Append "lField01", adCurrency, 20
'   lrs.Fields.Append "sField01", adVarChar, 100
'   lrs.Fields.Append "sField02", adVarChar, 100
'   lrs.Fields.Append "sField03", adVarChar, 100
'   lrs.Fields.Append "sField04", adVarChar, 100
'   lrs.Fields.Append "sField05", adVarChar, 100
'   lrs.Open
'
'   lsSourceNo = IFNull(oTrans.Master("sTransNox"), "")
'
'   With oTrans
'
'      For lnCtr = 0 To .ItemCount - 1
'         lrs.AddNew
'         lrs.Fields("sField01") = IFNull(oTrans.Detail(lnCtr, "sBrandNme"), "")
'         lrs.Fields("sField02") = IFNull(oTrans.Detail(lnCtr, "sColorNme"), "")
'         lrs.Fields("sField03") = IFNull(oTrans.Detail(lnCtr, "sSerialNo"), "")
'         lrs.Fields("sField04") = IFNull(oTrans.Detail(lnCtr, "sSerialNo"), "")
'         lrs.Fields("sField05") = IFNull(oTrans.Detail(lnCtr, "sModelNme"), "")
'         lrs.Fields("lField01") = IFNull(Format(oTrans.Detail(lnCtr, "nUnitPrce"), "#,##0.00"), "0")
'      Next
'      lrs.Sort = "lField01 DESC,sField05,sField05,sField03"
'   End With
'
'   'assign important info to the report
'   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Transfer_TradeIn.rpt")
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
'   oReport.Sections("RFb").ReportObjects("txtWithSerial").SetText oTrans.ItemCount
'   oReport.Sections("PF").ReportObjects("txtRptUser").SetText oApp.UserName
'
'   Set loreport = New frmRepViewer
'   Set loreport.ReportSource = oReport
'   loreport.Show
'
'   PrintTrans = True
'
'endPoc:
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

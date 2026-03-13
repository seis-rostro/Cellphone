VERSION 5.00
Object = "{0A7B56A6-35D0-4533-91DA-1715D3A0DD3E}#1.1#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransfer_Serial_Reg 
   BorderStyle     =   0  'None
   Caption         =   "Branch Transfer Transaction  (w/  IMIE No.) "
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Tag             =   "et0;eb0;et0;bc2"
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4980
      Left            =   150
      TabIndex        =   18
      Top             =   3360
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   8784
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
      Object.HEIGHT          =   4980
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
      MOUSEICON       =   "frmTransfer_Serial_Reg.frx":0000
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
      Height          =   405
      Index           =   4
      Left            =   10260
      TabIndex        =   26
      Top             =   4860
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Browse"
      AccessKey       =   "B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTransfer_Serial_Reg.frx":001C
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   10260
      TabIndex        =   27
      Top             =   4860
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
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
      Picture         =   "frmTransfer_Serial_Reg.frx":0796
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   10260
      TabIndex        =   24
      Top             =   4440
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
      Picture         =   "frmTransfer_Serial_Reg.frx":0F10
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   7
      Left            =   10260
      TabIndex        =   28
      Top             =   5280
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
      Picture         =   "frmTransfer_Serial_Reg.frx":168A
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   5
      Left            =   10260
      TabIndex        =   25
      Top             =   4440
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Update"
      AccessKey       =   "U"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTransfer_Serial_Reg.frx":1E04
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   3
      Left            =   10260
      TabIndex        =   29
      Top             =   5280
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
      Picture         =   "frmTransfer_Serial_Reg.frx":257E
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   6
      Left            =   10260
      TabIndex        =   21
      Top             =   3600
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
      Picture         =   "frmTransfer_Serial_Reg.frx":2CF8
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   5130
      Index           =   2
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   3285
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   9049
      BackColor       =   12632256
      ClipControls    =   0   'False
      BorderStyle     =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1695
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1560
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   2990
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   5040
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "frmTransfer_Serial_Reg.frx":3472
         Top             =   390
         Width           =   2460
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1335
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   390
         Width           =   2460
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   900
         Index           =   4
         Left            =   1335
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "frmTransfer_Serial_Reg.frx":3478
         Top             =   645
         Width           =   8460
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1335
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   135
         Width           =   2460
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   5040
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmTransfer_Serial_Reg.frx":347E
         Top             =   135
         Width           =   2460
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UNKNOWN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7815
         TabIndex        =   19
         Tag             =   "eb0;wb0"
         Top             =   135
         Width           =   1965
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Requested By "
         Height          =   285
         Index           =   0
         Left            =   3945
         TabIndex        =   14
         Top             =   405
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   645
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   405
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   8
         Top             =   135
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By "
         Height          =   285
         Index           =   2
         Left            =   3960
         TabIndex        =   12
         Top             =   135
         Width           =   1200
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   495
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1005
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   873
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1335
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   105
         Width           =   2445
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   5040
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   105
         Width           =   4755
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transmittal No."
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   105
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   285
         Index           =   19
         Left            =   3975
         TabIndex        =   6
         Top             =   105
         Width           =   975
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   10260
      TabIndex        =   22
      Top             =   3600
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
      Picture         =   "frmTransfer_Serial_Reg.frx":3484
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   8
      Left            =   10260
      TabIndex        =   23
      Top             =   4020
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "Pre&view"
      AccessKey       =   "v"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTransfer_Serial_Reg.frx":3BFE
      CaptionAlign    =   0
   End
   Begin MSComCtl2.Animation Progress 
      Height          =   720
      Left            =   10440
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "wt0;wb0"
      Top             =   555
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1270
      _Version        =   393216
      BackColor       =   1842204
      FullWidth       =   61
      FullHeight      =   48
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   9
      Left            =   10260
      TabIndex        =   20
      ToolTipText     =   "Void Transaction"
      Top             =   3180
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      Caption         =   "&Void"
      AccessKey       =   "V"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmTransfer_Serial_Reg.frx":4D10
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   495
      Index           =   3
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   873
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   5025
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   120
         Width           =   4755
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   1335
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   105
         Width           =   2445
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   285
         Index           =   5
         Left            =   3975
         TabIndex        =   2
         Top             =   105
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Origin"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   0
         Top             =   105
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmTransfer_Serial_Reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Programmed By  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Rosalyn Lazo Descallar  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Started  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 19, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private oRS As New ADODB.Recordset
Private poFileSys As FileSystemObject

Dim Drive As String
Dim Reference As String
Dim oSerial As frmCP_Serial_Info
Dim txtfieldGotfocus As Boolean
Dim txtOthersGotfocus As Boolean
Dim pbnewitem As Boolean

Dim psSelected() As String
Dim pnUserRights As Integer
Dim psUserID As String
Dim psUserName As String

Dim pnindex As Integer
Dim Index As Integer
Dim pnCtr As Integer
Dim lnCtr As Integer
Dim void As Boolean
Dim pbExisting As Boolean

Dim Branch As String
Dim Code As String
Dim Address As String

Private Sub cmdButton_Click(Index As Integer)
Dim response As String
Dim lsApproval As Integer
Dim lsConfirm As Integer
Dim lbEntry As Boolean
   
   Select Case Index
      Case 0   'Save
         void = False
         oDriver.RecordSave
      Case 1   'New
         oDriver.RecordNew
         EmptyGrid
      Case 2 'Delete Row
         With GridEditor1
            If .Rows <> 2 Then
               .DeleteRow
            End If
         End With
      Case 3 'Cancel
         oDriver.RecordCancelUpdate
         EmptyGrid
         ShowGrid
         HideButton
      Case 4  'Browse
         Search_Transmittal
         ShowGrid
      Case 5   'Update
         If oDriver.FieldValue(0) <> "" Then
            If oDriver.FieldValue(7) <> oApp.BranchCode Then
               MsgBox "Update of Branch Transaction Not Permitted!!!", vbCritical, "Warning"
               Exit Sub
            End If
            Select Case oDriver.FieldValue(9)
            Case 1
               MsgBox "Transaction Already Posted!!!" & vbCrLf & _
                     "Update Not Permitted", vbInformation, "Notice"
               Exit Sub
            Case 0
               With GridEditor1
                  For pnCtr = 1 To .Rows - 1
                     If oRS.State = adStateOpen Then oRS.Close
                        oRS.Open "SELECT * From CP_Inventory_Ledger " _
                                 & "WHERE sStockIDx = '" & .TextMatrix(pnCtr, 5) & "'" _
                                    & " AND sBranchCd = '" & oApp.BranchCode & "'" _
                                 & "ORDER by nEntryNox Desc" _
                                 , oApp.Connection, adOpenDynamic, adLockReadOnly, adCmdText
                     
                     If oRS.RecordCount <> 0 Then
                        oRS.MoveFirst
                        If oRS("sSourceNo") <> oDriver.FieldValue(0) Then
                           lbEntry = True
                        Else
                           lbEntry = False
                        End If
                     End If
                  Next
               End With
               
               Select Case oApp.UserLevel
                  Case xeEncoder, xeSupervisor
                     lsApproval = MsgBox("User Not Allowed to Update!!!" & vbCrLf & _
                                 "Seek for Approval?", vbQuestion + vbYesNo, "Confirm")
                     If lsApproval = vbYes Then
                        If Not GetApproval(oApp, pnUserRights, psUserID, psUserName) Then Exit Sub
                        If pnUserRights < xeManager Then
                           MsgBox "Approving User is Not Authorized!!!", vbCritical, "Warning"
                           Exit Sub
                        End If
                     Else
                        Exit Sub
                     End If
                  Case xeManager
                     If lbEntry = True Then
                        MsgBox "Item Has other Transactions!!!" & vbCrLf & _
                              "Update Not Permitted!!!" & vbCrLf & vbCrLf & _
                              "Notify ROSALYN LAZO DESCALLAR for Assistance!!!", vbInformation, "Notice"

                        lsApproval = MsgBox("Seek for Approval?", vbQuestion + vbYesNo, "Confirm")
                        If lsApproval = vbYes Then
                           If Not GetApproval(oApp, pnUserRights, psUserID, psUserName) Then Exit Sub
                           If pnUserRights < xeEngineer Then
                              MsgBox "Approving User is Not Authorized!!!", vbCritical, "Warning"
                              Exit Sub
                           End If
                        Else
                           Exit Sub
                        End If
                     Else
                        lsApproval = MsgBox("Are you sure you to Cancel this Transaction?" & vbCrLf _
                                    , vbQuestion + vbYesNo, "Confirm")
                        If lsApproval <> vbYes Then Exit Sub
                     End If
                  Case xeEngineer
                     lsApproval = MsgBox("Item Has other Transactions!!!" & vbCrLf & _
                                 "Do you want to continue?", vbQuestion + vbYesNo, "Confirm")
                     If lsApproval <> vbYes Then Exit Sub
               End Select
      
               oDriver.RecordUpdate
               ShowButton
               GridEditor1.ColEnabled(1) = True
               oDriver.DisableTextbox 0
               oDriver.DisableTextbox 1
            End Select
         Else
            MsgBox "No Active Transaction!!!", vbInformation, "Notice"
         End If
         
      Case 6   'Search Destination
         If txtfieldGotfocus And pnindex = 2 Then oDriver.RecordSearch txtfield(pnindex).Text
      Case 7
         Unload Me
      Case 8   'Print Transaction
         If oDriver.FieldValue(0) = "" Then Exit Sub
         Print_Transaction
      Case 9   'Void Transaction
         void = True
         If oDriver.FieldValue(0) = "" Then Exit Sub
         If oDriver.FieldValue(7) <> oApp.BranchCode Then
            MsgBox "Update of Branch Transaction Not Permitted!!!" & vbCrLf & vbCrLf & _
            "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
            Exit Sub
         End If
            Select Case oDriver.FieldValue(9)
            Case 2
               MsgBox "Transaction Already Cancelled!!!" & vbCrLf & _
                     "Void Not Permitted", vbInformation, "Notice"
               Exit Sub
            Case 1
               MsgBox "Transaction Already Posted!!!" & vbCrLf & _
                     "Update Not Permitted", vbInformation, "Notice"
               Exit Sub
            Case 0
               If oApp.UserLevel = xeEncoder Or oApp.UserLevel = xeSupervisor Then
                  lsApproval = MsgBox("User Not Allowed to Update!!!" & vbCrLf & _
                              "Seek for Approval?", vbQuestion + vbYesNo, "Confirm")
                  If lsApproval = vbYes Then
                     If Not GetApproval(oApp, pnUserRights, psUserID, psUserName) Then Exit Sub
                     If pnUserRights < xeManager Then
                        MsgBox "Approving User is Not Authorized!!!", vbCritical, "Warning"
                        Exit Sub
                     End If
                  Else
                     Exit Sub
                  End If
               End If
            End Select
            
            lsConfirm = MsgBox("Are you sure you want to Cancel this Transaction!!!" & vbCrLf _
                     , vbQuestion + vbYesNo, "Confirm")
            If lsConfirm = vbYes Then
               Delete_Transaction
            Else
               Exit Sub
            End If
            MsgBox "Transaction Cancelled!!!", vbInformation, "Information"
   End Select
   
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   If bLoaded = False Then
      oDriver.RecordNew
      HideButton
      oDriver.ShowButton 6
      bLoaded = True
   End If
   GridEditor1.Refresh
End Sub

Private Sub HideButton()
   oDriver.ShowButton 4
   oDriver.ShowButton 5
   oDriver.ShowButton 7
   oDriver.ShowButton 9
   oDriver.HideButton 0
   oDriver.HideButton 3
   oDriver.HideButton 6
   
   For lnCtr = 1 To 6
      txtfield(lnCtr).Enabled = True
   Next
End Sub

Private Sub ShowButton()
   oDriver.HideButton 4
   oDriver.HideButton 5
   oDriver.HideButton 7
   oDriver.HideButton 9
   oDriver.ShowButton 0
   oDriver.ShowButton 3
   oDriver.ShowButton 6
      
   GridEditor1.SetFocus
End Sub

Private Sub Search_Transmittal()
Dim orig As String
Dim lsSQL As String
Dim lsCondition As String

   orig = oDriver.BrowseQuery
   Select Case pnindex
      Case 0
         lsCondition = " a.sTransNox like '%" & txtfield(0).Text & "%' " _
                        & " AND (sOriginxx = '" & oApp.BranchCode & "'" _
                        & " OR sDestinat = '" & oApp.BranchCode & "')"
         lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
         oDriver.BrowseQuery = lsSQL
      Case 1
         lsCondition = " a.sReferNox like '%" & txtfield(1).Text & "%'" _
                        & " AND (sOriginxx = '" & oApp.BranchCode & "'" _
                        & " OR sDestinat = '" & oApp.BranchCode & "')"
         lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
         oDriver.BrowseQuery = lsSQL
   End Select
   oDriver.BrowseQuery = lsSQL
   oDriver.BrowseRecord
   oDriver.BrowseQuery = orig

End Sub

Private Sub Form_Deactivate()
   Progress.Stop
   Progress.Close
End Sub

Private Sub Form_Load()
Dim lsSQL As String
Dim lnCtr As Integer

   CenterChildForm mdiMain, Me
   bLoaded = False
    
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   InitGrid
   InitTxtField
      
   Set oSerial = New frmCP_Serial_Info
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
   
   oDriver.RecQuery = "SELECT" _
                        & " sTransNox, " _
                        & " sReferNox, " _
                        & " sDestinat, " _
                        & " dTransact, " _
                        & " sRemarksx, " _
                        & " sRequestx, " _
                        & " sApproved, " _
                        & " sOriginxx, " _
                        & " nEntryNox, " _
                        & " cTranStat, " _
                        & " cReceived, " _
                        & " dReceived, " _
                        & " sModified, " _
                        & " dModified, " _
                        & " vTimeStmp  " _
                  & " FROM CP_Serial_Transfer_Master " _

   oDriver.BrowseQuery = "SELECT" _
                  & " Distinct " _
                  & " TOP 1000 " _
                  & " a.sTransNox, " _
                  & " a.sReferNox, " _
                  & " a.dTransact, " _
                  & " b.sBranchNm, " _
                  & " a.cTranStat  " _
            & " FROM CP_Serial_Transfer_Master a " _
               & " LEFT JOIN Branch b " _
                  & " ON a.sDestinat = b.sBranchCd " _
               & " LEFT JOIN CP_Inventory_Ledger c " _
                  & " ON a.sTransNOx = c.sSourceNo " _
            & " ORDER BY a.dTransact Desc "
   
   oDriver.InitRecForm

   oDriver.BrowseColumn(0) = "sTransNox"
   oDriver.BrowseColumn(1) = "sReferNox"
   oDriver.BrowseColumn(2) = "dTransact"
   oDriver.BrowseColumn(3) = "sBranchNm"
   oDriver.BrowseColumn(4) = "cTranstat"
   
   oDriver.BrowseFTitle(0) = "Transaction No"
   oDriver.BrowseFTitle(1) = "Transmittal No"
   oDriver.BrowseFTitle(2) = "Date"
   oDriver.BrowseFTitle(3) = "Destination"
   oDriver.BrowseFTitle(4) = "Status"
   
   oDriver.BrowseFFormat(2) = "MMMM dd, yyyy"
   
   'Branch Destination
   oDriver.LookupQuery(2) = "SELECT" _
                           & " a.sBranchCd, " _
                           & " a.sBranchNm, " _
                           & " a.sAddressx + ' ' + b.sTownName xAddressx " _
                     & " FROM Branch a " _
                        & "LEFT JOIN TownCity b " _
                           & " ON a.sTownIdxx = b.sTownIDxx " _
                     & " WHERE a.cRecdStat = '" & xeRecStateActive & "'" _
                     & " ORDER BY sBranchNm "
   
   oDriver.LookupReference(2) = "sBranchCd»sBranchNm»xAddressx"
   oDriver.LookupColumn(2) = "sBranchNm»xAddressx"
   oDriver.LookupTitle(2) = "Branch»Address"
   
   'Branch Origin
   oDriver.LookupQuery(7) = "SELECT" _
                           & " a.sBranchCd, " _
                           & " a.sBranchNm, " _
                           & " a.sAddressx + ' ' + b.sTownName xAddressx " _
                     & " FROM Branch a " _
                        & "LEFT JOIN TownCity b " _
                           & " ON a.sTownIdxx = b.sTownIDxx " _
                     & " WHERE a.cRecdStat = '" & xeRecStateActive & "'" _
                     & " ORDER BY sBranchNm "
   
   oDriver.LookupReference(7) = "sBranchCd»sBranchNm»xAddressx"
   oDriver.LookupColumn(7) = "sBranchNm»xAddressx"
   oDriver.LookupTitle(7) = "Branch»Address"

   oDriver.FieldStart = 1
   oDriver.FieldFormat(3) = "MMMM dd, yyyy"
   EmptyGrid
   
         
End Sub

Private Sub InitGrid()
    With GridEditor1
      .Rows = 2
      .Cols = 8
      .Font = "MS Sans Serif"
              
      'column title
      .TextMatrix(0, 1) = "IMEI No."
      .TextMatrix(0, 2) = "Bar Code"
      .TextMatrix(0, 3) = "Particulars"
      .TextMatrix(0, 4) = "SRP"
      .TextMatrix(0, 5) = "Stock ID"
      .TextMatrix(0, 6) = "Serial ID"
      .TextMatrix(0, 7) = "Purchase P"
      .Row = 0
            
      'column width
      .ColWidth(0) = 400
      .ColWidth(1) = 1600
      .ColWidth(2) = 1800
      .ColWidth(3) = 4850
      .ColWidth(4) = 1000
      .ColWidth(5) = 0
      .ColWidth(6) = 0
      .ColWidth(7) = 0
              
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 6
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 6

      .ColEnabled(2) = False
      .ColEnabled(3) = False
      .ColEnabled(4) = False
      .ColEnabled(5) = False
      .ColEnabled(6) = False
      .ColEnabled(7) = False

      .Row = 1
    End With
End Sub

Private Sub InitTxtField()
Dim Index As Integer
   For Index = 0 To 5
      txtfield(Index).Text = ""
   Next
End Sub

Private Sub Print_Transaction()
Dim lnCtr As Integer
Dim lrs As New ADODB.Recordset
Dim lsSQL As String
Dim lrsDetail As New ADODB.Recordset

   Set lrs = New ADODB.Recordset
   lrs.Fields.Append "sField06", adVarChar, 150
   lrs.Fields.Append "sField09", adVarChar, 10
   lrs.Fields.Append "sField10", adVarChar, 20
   lrs.Open

   'CP_Serial_Transfer_Master
   lsSQL = "SELECT" _
               & " a.sTransNox, " _
               & " a.sReferNox, " _
               & " a.sApproved, " _
               & " a.sDestinat, " _
               & " b.sBranchNm, " _
               & " b.sAddressx + ', ' + c.sTownName xAddressx " _
         & " FROM CP_Serial_Transfer_Master a " _
            & " LEFT JOIN Branch b " _
               & " ON a.sOriginxx = b.sBranchCd " _
            & " LEFT JOIN TownCity c " _
               & " ON b.sTownIDxx = c.sTownIDxx " _
         & " WHERE a.sTransNox = '" & oDriver.FieldValue(0) & "' "
   
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If oRS.EOF Then
      MsgBox "No Record Found!!!" & vbCrLf & _
             "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      Exit Sub
   End If
      
      Progress.Open App.Path & "\images\FINDCOMP.AVI"
      Progress.Play

      For lnCtr = 0 To oRS.RecordCount - 1
         lrs.AddNew

         lsSQL = "SELECT" _
                  & " a.nEntryNox, " _
                  & " b.sIMEINoxx, " _
                  & " b.sStockIDx, " _
                  & " c.sDescript, " _
                  & " d.sBrandNme, " _
                  & " e.sModelNme, " _
                  & " f.sColorNme  " _
               & " FROM CP_Serial_Transfer_Detail a " _
                  & " LEFT JOIN CP_Serial_Master b " _
                     & " ON a.sSerialID = b.sSerialID " _
                  & " LEFT JOIN CP_Inventory c " _
                     & " ON b.sStockIDx = c.sStockIDx " _
                  & " LEFT JOIN Brand d " _
                     & " ON c.sBrandIdx = d.sBrandIDx " _
                  & " LEFT JOIN Model e  " _
                     & " ON c.sModelIDx = e.sModelIDx " _
                  & " LEFT JOIN Color f  " _
                     & " ON c.sColorIDx = f.sColorIdx " _
               & " WHERE a.sTransNox = '" & oRS("sTransNox") & "' " _
               & " ORDER BY a.nEntryNox "

            If lrsDetail.State = adStateOpen Then lrsDetail.Close
            lrsDetail.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
            Do While Not lrsDetail.EOF
               lrs.AddNew
               lrs("sField06").Value = Trim(IIf(IsNull(lrsDetail("sBrandNme")), "", lrsDetail("sBrandNme")) _
                                       & " " & IIf(IsNull(lrsDetail("sModelNme")), "", lrsDetail("sModelNme")) _
                                       & " " & IIf(IsNull(lrsDetail("sDescript")), "", lrsDetail("sDescript")) _
                                       & " " & IIf(IsNull(lrsDetail("sColorNme")), "", lrsDetail("sColorNme")))
               lrs("sField09").Value = lrsDetail("sStockIDx")
               lrs("sField10").Value = lrsDetail("sIMEINoxx")
               lrsDetail.MoveNext
            Loop

         oRS.MoveNext
      Next
      
      Branch = txtfield(2)
      getBranch Code, Branch, Address
      
      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Transmittal_Serial.rpt")
      oReport.DiscardSavedData
      oReport.FieldMappingType = crAutoFieldMapping
      oReport.Database.SetDataSource lrs
      
      oRS.MoveFirst
      With oReport
         .Sections("PH").ReportObjects("txtReportDate").SetText txtfield(3).Text
         .Sections("PH").ReportObjects("txtTransmittal").SetText txtfield(1).Text
         .Sections("PH").ReportObjects("txtToBranch").SetText Branch
         .Sections("PH").ReportObjects("txtToAddress").SetText Address
         .Sections("PH").ReportObjects("txtFromBranch").SetText oRS("sBranchNm")
         .Sections("PH").ReportObjects("txtFromAddress").SetText oRS("xAddressx")
      
         .Sections("PF").ReportObjects("txtApproved").SetText txtfield(6).Text
         .Sections("PF").ReportObjects("txtPrepared").SetText oApp.UserName
         .Sections("RF").ReportObjects("txtRemarks").SetText txtfield(4).Text
      End With

      Set lrs = Nothing
      Set oRS = Nothing
      Set lrsDetail = Nothing
      
      frmViewer.CRViewer91.ReportSource = oReport
      frmViewer.CRViewer91.ViewReport
      frmViewer.Show

End Sub

Private Sub ShowGrid()
Dim lsSQL As String
Dim lnCtr As Integer
Dim lnrow As Long

   lsSQL = "SELECT" _
               & " DISTINCT " _
               & " a.sSerialID, " _
               & " a.sTransNox, " _
               & " a.nEntryNox, " _
               & " a.nUnitPrce, " _
               & " b.sIMEINoxx, " _
               & " b.sStockIDx, " _
               & " c.sBarrCode, " _
               & " c.sDescript, " _
               & " e.sBrandNme, " _
               & " f.sModelNme, " _
               & " g.sColorNme, " _
               & " c.nSelPrice  "
   lsSQL = lsSQL _
         & " FROM CP_Serial_Transfer_Detail a " _
               & " LEFT JOIN CP_Serial_Master b " _
                  & " ON a.sSerialID = b.sSerialID " _
               & " LEFT JOIN CP_Inventory c " _
                  & " ON b.sStockIDx = c.sStockIDx " _
               & " LEFT JOIN CP_Inventory_Master d " _
                  & " ON b.sStockIDx = d.sStockIDx " _
               & " LEFT JOIN Brand e " _
                  & " ON c.sBrandIDx = e.sBrandIDx " _
               & " LEFT JOIN Model f " _
                  & " ON c.sModelIDx = f.sModelIDx " _
               & " LEFT JOIN Color g " _
                  & " ON c.sColorIDx =  g.sColorIDx " _
               & " LEFT JOIN CP_Inventory_Ledger h " _
                  & " ON a.sTransNox = h.sSourceNo " _
         & " WHERE a.sTransNox = '" & oDriver.FieldValue(0) & "'" _
         & " ORDER BY a.nEntryNox "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If oRS.RecordCount <> 0 Then
      With GridEditor1
         .Rows = oRS.RecordCount + 1
         For lnCtr = 0 To oRS.RecordCount - 1
            .TextMatrix(lnCtr + 1, 0) = oRS("nEntryNox")
            .TextMatrix(lnCtr + 1, 1) = oRS("sIMEINoxx")
            .TextMatrix(lnCtr + 1, 2) = oRS("sBarrCode")
            .TextMatrix(lnCtr + 1, 3) = Trim(IIf(IsNull(oRS("sBrandNme")), "", oRS("sBrandNme")) & " " & _
                                          IIf(IsNull(oRS("sModelNme")), "", oRS("sModelNme")) & " " & _
                                          IIf(IsNull(oRS("sDescript")), "", oRS("sDescript")) & " " & _
                                          IIf(IsNull(oRS("sColorNme")), "", oRS("sColorNme")))
            .TextMatrix(lnCtr + 1, 4) = Format(oRS("nSelPrice"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 5) = oRS("sStockIDx")
            .TextMatrix(lnCtr + 1, 6) = oRS("sSerialID")
            .TextMatrix(lnCtr + 1, 7) = Format(oRS("nUnitPrce"), "#,##0.00")
            oRS.MoveNext
         Next
         If .Rows > 20 Then
            .ColWidth(3) = 4600
         Else
            .ColWidth(3) = 4850
         End If
         .ColEnabled(1) = False
      End With
   Else
      Exit Sub
   End If

   Set oRS = Nothing

End Sub

Private Sub EmptyGrid()
Dim lnCtr As Integer
   With GridEditor1
      .Rows = 2
      For lnCtr = 1 To .Cols - 1
         .TextMatrix(1, lnCtr) = ""
      Next
      .ColEnabled(1) = True
   End With
   GridEditor1.Refresh
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
            Case vbKeyF3
               If pnindex = 2 Then oDriver.RecordSearch txtfield(pnindex).Text
         End Select
      Case 27
         Call Modified("CP_Serial_Transfer_Master", "sTransNox = '" & oDriver.FieldValue(0) & "' ")
   End Select
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         Cancel = True
      ElseIf .TextMatrix(.Row, 5) = "" Then
         Cancel = True
      ElseIf pbExisting = True Then
         Cancel = True
      End If
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         MsgBox "Invalid IMEI No.!!!", vbCritical, "Warning"
         .SetFocus
      End If
   End With
End Sub

Private Sub GridEditor1_RowColChange()
'   With GridEditor1
'      If .Row <> 0 And Trim(.TextMatrix(.Row, 1)) <> "" And _
'         .Col = 1 And .ColEnabled(1) = True Then
'         Search_Serial
'      End If
'   End With
End Sub

Private Sub oDriver_InitValue()
   label.Caption = "UNKNOWN"
   pbExisting = False
   txtothers(0).Text = ""
   txtothers(0).Tag = ""
   oDriver.FieldValue(3) = Date
End Sub

Private Sub oDriver_LoadOtherData()
Dim lrs As ADODB.Recordset
Dim lsSQL As String
   
   Select Case oDriver.FieldValue(9)
      Case 0
         label.Caption = "UNKNOWN"
      Case 1
         label.Caption = "POSTED"
      Case 2
         label.Caption = "CANCELLED"
   End Select
   Set lrs = New ADODB.Recordset
   lsSQL = "SELECT" _
            & " a.sAddressx + ' ' + b.sTownName xAddressx " _
      & " FROM Branch a " _
         & "LEFT JOIN TownCity b " _
            & " ON a.sTownIdxx = b.sTownIDxx " _
      & " WHERE a.sBranchCd = '" & oDriver.FieldValue(7) & "'"
   lrs.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If lrs.EOF = False Then
      txtothers(0).Text = lrs("xAddressx")
   Else
      txtothers(0).Text = ""
   End If
   Set lrs = Nothing
   txtothers(0).Enabled = False
   oDriver.FieldValue(3) = Format(oDriver.FieldValue(3), "m/d/yyyy")
End Sub

Private Sub oDriver_SaveComplete()
   HideButton
   MsgBox "Transaction Successfully Updated!!!", vbInformation, "Information"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   With GridEditor1
      If oDriver.FieldValue(2) = "" Then
         MsgBox "Invalid Destination Detected!!!", vbCritical, "Warning"
         txtfield(2).SetFocus
         Cancel = True
      ElseIf oDriver.FieldValue(3) = "" Then
         MsgBox "Invalid Transaction Date Detected!!!", vbCritical, "Warning"
         txtfield(3).SetFocus
         Cancel = True
      Else
         Time = Format(Now, "hh:nn:ss AM/PM")
         Cancel = Not Delete_Transaction
            If Cancel Then Exit Sub
         Cancel = Not Save_CP_Serial
            If Cancel Then Exit Sub
         Cancel = Not Update_CP_Inventory
            If Cancel Then Exit Sub
         Cancel = Not Save_CP_Inventory_Ledger
            If Cancel Then Exit Sub
         
         oDriver.FieldValue(3) = CDate(txtfield(3).Text) & " " & Time
         oDriver.FieldValue(7) = oApp.BranchCode
         oDriver.FieldValue(8) = .TextMatrix(.Rows - 1, 0)
         oDriver.FieldValue(9) = 0  'cTranStat
         oDriver.FieldValue(10) = 0  'cReceived
         Reference = txtfield(1).Text
         Branch = txtfield(2)
      End If
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   txtfieldGotfocus = True
   txtOthersGotfocus = False
   pnindex = Index
   oDriver.ColumnIndex = Index
   txtfield(Index).BackColor = &HE1FEFF
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim orig As String
Dim lsSQL As String
Dim lsCondition As String

   If KeyCode = vbKeyF3 Or KeyCode = 13 Then
      Select Case Index
      Case 0, 1
         Search_Transmittal
         If oDriver.FieldValue(0) <> "" Then ShowGrid
      Case 2
         oDriver.RecordSearch txtfield(Index).Text
         If txtfield(Index).Text <> "" Then SetNextFocus
      Case 7
         oDriver.RecordSearch txtfield(Index).Text
         cmdButton(4).SetFocus
      End Select
   End If
   KeyCode = 0
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   If Not IsDate(txtfield(3).Text) Then
      txtfield(3).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   Else
      txtfield(3).Text = Format(txtfield(3).Text, "MMMM DD, YYYY")
   End If
   txtfield(Index).BackColor = &HFFFFFF
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   If Not IsDate(txtfield(3).Text) Then
      txtfield(3).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   Else
      txtfield(3).Text = Format(txtfield(3).Text, "MMMM DD, YYYY")
   End If
   txtfield(Index).Text = TitleCase(txtfield(Index).Text)
   Cancel = Not oDriver.ValidateField(Index)
End Sub

Private Function Delete_Transaction() As Boolean
Dim lsSQL As String
Dim lnrow As Long
Dim Entry As Integer
Dim QOH As Integer

Delete_Transaction = True
On Error Goto errProc

   'Roll Back QOH in CP_Inventory_Master
   lsSQL = "SELECT" _
            & " b.sSerialID," _
            & " b.sStockIDx " _
         & " FROM CP_Serial_Transfer_Detail a " _
            & " LEFT JOIN CP_Serial_Master b " _
               & " ON a.sSerialID = b.sSerialID " _
            & " WHERE a.sTransNox = '" & oDriver.FieldValue(0) & "'" _
               & " ORDER BY b.sStockIDx "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If oRS.RecordCount <> 0 Then
      Do While Not oRS.EOF
         lsSQL = "UPDATE CP_Inventory_Master SET" _
               & " nQtyOnHnd = nQtyOnHnd + 1 ," _
               & " dModified = getdate() " _
         & " WHERE sStockIDx = '" & oRS("sStockIDx") & "' " _
               & " And sBranchCd = '" & oApp.BranchCode & "' "
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
         oRS.MoveNext
      Loop
      Set oRS = Nothing
   End If

   'Update CP_Serial_Master
   lsSQL = "SELECT * " _
         & " FROM CP_Serial_Transfer_Detail " _
         & " WHERE sTransNox = '" & oDriver.FieldValue(0) & "'"
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   If oRS.RecordCount <> 0 Then
      Do While Not oRS.EOF
         lsSQL = "UPDATE CP_Serial_Master SET" _
               & " sBranchCd = '" & oApp.BranchCode & "', " _
               & " cLocation = '1', " _
               & " sModified = '" & Encrypt(oApp.UserID) & "', " _
               & " dModified = getdate() " _
         & " WHERE sSerialID = '" & oRS("sSerialID") & "' "
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
         oRS.MoveNext
      Loop
      Set oRS = Nothing
   End If
                              
   'Delete CP_Serial_Ledger
   lsSQL = "DELETE CP_Serial_Ledger " _
                        & " WHERE sSourceNo = '" & oDriver.FieldValue(0) & "'" _
                           & " AND sSourceCd = 'CPDv' "
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
   If lnrow <> 0 Then oApp.RegisDelete lsSQL
   
   'Update CP_Serial_Ledger nEntryNox
   lsSQL = "SELECT" _
            & " sSerialID" _
         & " FROM CP_Serial_Transfer_Detail " _
            & " WHERE sTransNox = '" & oDriver.FieldValue(0) & "'" _
            & " ORDER BY sSerialID "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   Do While Not oRS.EOF
      Call Recalc_Serial("'" & oRS("sSerialID") & "'")
      oRS.MoveNext
   Loop
   Set oRS = Nothing
   
   'Delete CP_Inventory_Ledger
   lsSQL = "DELETE CP_Inventory_Ledger " _
                        & " WHERE sSourceNo = '" & oDriver.FieldValue(0) & "'" _
                           & " AND sSourceCd = 'CPDv' " _
                           & " AND sBranchCd = '" & oApp.BranchCode & "'"
   oApp.Connection.Execute lsSQL, lnrow, adCmdText
   If lnrow <> 0 Then oApp.RegisDelete lsSQL

   'update CP_Inventory_Ledger nEntryNox, nQtyOnHnd
   lsSQL = "SELECT" _
         & " b.sStockIDx " _
      & " FROM CP_Serial_Transfer_Detail a " _
         & " LEFT JOIN CP_Serial_Master b " _
            & " ON a.sSerialID = b.sSerialID " _
         & " WHERE a.sTransNox = '" & oDriver.FieldValue(0) & "'" _
         & " ORDER BY b.sStockIDx "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   Do While Not oRS.EOF
      Call Recalc_Inventory("'" & oApp.BranchCode & "'", "'" & oRS("sStockIDx") & "'")
      oRS.MoveNext
   Loop
   Set oRS = Nothing

   'Update CP_Serial_Transfer_Master
   If void = True Then 'Transaction Cancelled
      lsSQL = "Update CP_Serial_Transfer_Master SET " _
               & " cTranStat = 2, " _
               & " sModified = '" & Encrypt(oApp.UserID) & "', " _
               & " dModified = getdate() " _
               & " WHERE sTransNox = '" & oDriver.FieldValue(0) & "'"
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
      If lnrow <> 0 Then oApp.RegisDelete lsSQL
      label.Caption = "CANCELLED"
   Else
      'Delete CP_Serial_Ledger
      lsSQL = "DELETE CP_Serial_Transfer_Detail " _
               & " WHERE sTransNOx = '" & oDriver.FieldValue(0) & "'"
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
      If lnrow <> 0 Then oApp.RegisDelete lsSQL
   End If

   Set oRS = Nothing
                              
endProc:
   Exit Function
errProc:
   Delete_Transaction = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Function Save_CP_Serial() As Boolean
Dim lsSQL As String
Dim lnrow As Long
Dim lnEntry As Integer
   
Save_CP_Serial = True
On Error GoTo errProc
   
   With GridEditor1
      For pnCtr = 1 To .Rows - 1
         If .TextMatrix(pnCtr, 1) = "" Then Exit For
         
         'Get last Entry No.
         lnEntry = getIMEIEntry("'" & .TextMatrix(pnCtr, 6) & "'")
         
         'Save_CP_Serial_Transfer_Detail
         lsSQL = "INSERT INTO CP_Serial_Transfer_Detail " _
               & "( sTransNox, " _
               & "  nEntryNox, " _
               & "  sSerialID, " _
               & "  nUnitPrce, " _
               & "  dModified) " _
         & "VALUES " _
               & "('" & oDriver.FieldValue(0) & "', " _
               & "'" & .TextMatrix(pnCtr, 0) & "', " _
               & "'" & .TextMatrix(pnCtr, 6) & "', " _
               & "'" & CDbl(.TextMatrix(pnCtr, 7)) & "', " _
               & " getdate())"
         oApp.Connection.Execute lsSQL, lnrow, adCmdText

         'Update CP_Serial_Master
         lsSQL = "UPDATE CP_Serial_Master SET" _
               & " sBranchCd = '" & oDriver.FieldValue(2) & "', " _
               & " cLocation = '0', " _
               & " sModified = '" & Encrypt(oApp.UserID) & "', " _
               & " dModified = getdate() " _
         & " WHERE sSerialID = '" & .TextMatrix(pnCtr, 6) & "' "
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
         
         'CP_Serial_Ledger
         lsSQL = "INSERT INTO CP_Serial_Ledger " _
               & "( sSerialID, " _
               & "  sBranchCd, " _
               & "  dTransact, " _
               & "  nEntryNox, " _
               & "  sSourceCd, " _
               & "  sSourceNo, " _
               & "  cSoldStat, " _
               & "  cLocation, " _
               & "  dModified) " _
         & "VALUES " _
               & "('" & .TextMatrix(pnCtr, 6) & "', " _
               & "'" & oApp.BranchCode & "', " _
               & "'" & CDate(oDriver.FieldValue(3)) & " " & Time & "', " _
               & "'" & lnEntry & "', " _
               & "'CPDv', " _
               & "'" & oDriver.FieldValue(0) & "', " _
               & "'0'," _
               & "'1', " _
               & " getdate())"
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
      Next
      
      If lnrow <= 0 Then
         MsgBox "Unable to Save CP Serial!!!" & vbCrLf & vbCrLf & _
         "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
         Save_CP_Serial = False
         GoTo endProc
      End If
   End With

endProc:
   Exit Function
errProc:
   Save_CP_Serial = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Function Save_CP_Inventory_Ledger() As Boolean
Dim lsSQL As String
Dim lnrow As Long
Dim lnEntry As Integer
Dim QOH As Integer

Save_CP_Inventory_Ledger = True
On Error GoTo errProc
   
   With GridEditor1
      For pnCtr = 1 To .Rows - 1
         If .TextMatrix(pnCtr, 1) = "" Then Exit For
         'Search sSourceNo
         lsSQL = "SELECT" _
                  & " sStockIDx, " _
                  & " sSourceNo  " _
               & " FROM CP_Inventory_Ledger " _
               & " WHERE sStockIdx = '" & .TextMatrix(pnCtr, 5) & "'" _
                  & " AND sSourceNo = '" & oDriver.FieldValue(0) & "'" _
                  & " AND sSourceCd = 'CPDv' " _
                  & " AND sBranchCd = '" & oApp.BranchCode & "'"
         If oRS.State = adStateOpen Then oRS.Close
         oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
         
            'Get QOH
            QOH = getQuantity("'" & .TextMatrix(pnCtr, 5) & "'", "'" & oApp.BranchCode & "'")
            
            If oRS.EOF = False Then
               'Update Record, CP_Inventory_Ledger
               lsSQL = "UPDATE CP_Inventory_Ledger SET" _
                        & " nQtyOutxx = nQtyOutxx + 1 , " _
                        & " nQtyOnHnd = '" & CLng(QOH) & "', " _
                        & " dModified = getdate() " _
                  & " WHERE sStockIdx = '" & .TextMatrix(pnCtr, 5) & "'" _
                     & " AND sSourceNo = '" & oDriver.FieldValue(0) & "'" _
                     & " AND sSourceCd = 'CPDv' " _
                     & " AND sBranchCd = '" & oApp.BranchCode & "'"
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
            
            Else
               'Get last Entry No.
               lnEntry = getEntryNo("CP_Inventory_Ledger", "'" & .TextMatrix(pnCtr, 5) & "'", "'" & oApp.BranchCode & "'")
               
               'Add Record, CP_Inventory_Ledger
               lsSQL = "INSERT INTO CP_Inventory_Ledger " _
                     & "( sStockIDx, " _
                     & "  sBranchCd, " _
                     & "  sLocation, " _
                     & "  sSourceCd, " _
                     & "  sSourceNo, " _
                     & "  nQtyInxxx, " _
                     & "  nQtyOutxx, " _
                     & "  nQtyOnHnd, " _
                     & "  nEntryNox, " _
                     & "  dTransact, " _
                     & "  dModified) " _
               & "VALUES " _
                     & "('" & .TextMatrix(pnCtr, 5) & "', " _
                     & "'" & oApp.BranchCode & "', " _
                     & "'" & oDriver.FieldValue(2) & "', " _
                     & "'CPDv' , " _
                     & "'" & oDriver.FieldValue(0) & "', " _
                     & " 0, " _
                     & "'1', " _
                     & "'" & CLng(QOH) & "', " _
                     & "'" & lnEntry & "', " _
                     & "'" & CDate(oDriver.FieldValue(3)) & " " & Time & "', " _
                     & " getdate())"
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
               
            End If
         Set oRS = Nothing
      Next

      If lnrow <= 0 Then
         MsgBox "Unable to Update Inventory Ledger!!!" & vbCrLf & vbCrLf & _
         "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
         Save_CP_Inventory_Ledger = False
         GoTo endProc
      End If
   End With

endProc:
   Exit Function
errProc:
   Save_CP_Inventory_Ledger = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Function Update_CP_Inventory() As Boolean
Dim lsSQL As String
Dim lnrow As Long

Update_CP_Inventory = True
On Error GoTo errProc
   
   With GridEditor1
         For pnCtr = 1 To .Rows - 1
            If .TextMatrix(pnCtr, 1) = "" Then Exit For
            'Update QOH, CP_Inventory_Master
            lsSQL = "UPDATE CP_Inventory_Master SET" _
                  & " nQtyOnHnd = nQtyOnHnd - 1, " _
                  & " sModified = '" & Encrypt(oApp.UserID) & "', " _
                  & " dModified = getdate() " _
            & " WHERE sStockIDx = '" & .TextMatrix(pnCtr, 5) & "' " _
                  & " AND sBranchCd = '" & oApp.BranchCode & "'"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
         Next
   
         If lnrow <= 0 Then
            MsgBox "Unable to Update Inventory Master!!!" & vbCrLf & vbCrLf & _
            "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
            Update_CP_Inventory = False
            GoTo endProc
         End If
   End With

endProc:
   Exit Function
errProc:
   Update_CP_Inventory = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Sub GridEditor1_GotFocus()
   GridEditor1.Col = 1
End Sub

Private Sub GridEditor1_KeyDown(KeyCode As Integer, Shift As Integer)
   With GridEditor1
      If KeyCode = vbKeyF3 Then
         If .Col = 1 Then
            Search_Serial
         End If
      End If
   End With
End Sub

Private Sub Search_Serial()
Dim lsSQL As String
Dim lsCondition As String
Dim lsSearch As String
   
   With GridEditor1
      lsSQL = "SELECT" _
            & " Distinct " _
            & " f.sIMEINoxx, " _
            & " a.sBarrcode, " _
            & " a.sStockIDx, " _
            & " b.sBrandNme, " _
            & " c.sModelNme, " _
            & " a.sDescript, " _
            & " d.sColorNme, " _
            & " a.nSelPrice, " _
            & " f.sSerialID, " _
            & " a.nPurPrice  "
      lsSQL = lsSQL _
         & " FROM CP_Serial_Master f " _
            & " LEFT JOIN CP_Inventory a " _
               & " ON f.sStockIDx = a.sStockIDx " _
            & " LEFT JOIN CP_Inventory_Master e " _
               & " ON a.sStockIdx = e.sStockIDx " _
            & " LEFT JOIN Brand b " _
               & " ON a.sBrandIdx = b.sBrandIdx " _
            & " LEFT JOIN Model c " _
               & " ON a.sModelIdx = c.sModelIdx " _
            & " LEFT JOIN Color d " _
               & " ON a.sColorIDx = d.sColorIDx " _
         & " WHERE f.sIMEINoxx like  '%" & .TextMatrix(.Row, 1) & "%'" _
            & " AND (sCategIDx = '01001' or sCategIDx = '01002' or sCategIDx = '01003') " _
            & " AND f.cSoldStat = 0 " _
            & " AND ((f.sBranchCd = '" & oDriver.FieldValue(2) & "' AND cLocation = 0) " _
            & " or (f.sBranchCd = '" & oApp.BranchCode & "' AND cLocation = 1)) "

      If oRS.State = adStateOpen Then oRS.Close
      oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
         
         If Not oRS.EOF Then
            If oRS.RecordCount = 1 Then
               .TextMatrix(.Row, 1) = IIf(IsNull(oRS(0)), "", oRS(0))
               .TextMatrix(.Row, 2) = IIf(IsNull(oRS(1)), "", oRS(1))
               .TextMatrix(.Row, 3) = Trim(IIf(IsNull(oRS(3)), "", oRS(3)) & " " & _
                                       IIf(IsNull(oRS(4)), "", oRS(4)) & " " & _
                                       IIf(IsNull(oRS(5)), "", oRS(5)) & " " & _
                                       IIf(IsNull(oRS(6)), "", oRS(6)))
               .TextMatrix(.Row, 4) = IIf(IsNull(oRS(7)), "", Format(oRS(7), "#,##0.00"))
               .TextMatrix(.Row, 5) = IIf(IsNull(oRS(2)), "", oRS(2))
               .TextMatrix(.Row, 6) = IIf(IsNull(oRS(8)), "", oRS(8))
               .TextMatrix(.Row, 7) = IIf(IsNull(oRS(9)), "", Format(oRS(9), "#,##0.00"))
            Else
               lsSearch = KwikSearch(oApp, lsSQL, _
                          "sIMEINoxx»sBarrcode»sBrandNme»sModelNme»sDescript»sColorNme", _
                          "IMEI No.»Bar Code»Brand»Model»Description»Color")
               If lsSearch <> "" Then
                  psSelected = Split(lsSearch, "»")
                  .TextMatrix(.Row, 1) = IIf(IsNull(psSelected(0)), "", psSelected(0))
                  .TextMatrix(.Row, 2) = IIf(IsNull(psSelected(1)), "", psSelected(1))
                  .TextMatrix(.Row, 3) = Trim(IIf(IsNull(psSelected(3)), "", psSelected(3)) & " " & _
                                          IIf(IsNull(psSelected(4)), "", psSelected(4)) & " " & _
                                          IIf(IsNull(psSelected(5)), "", psSelected(5)) & " " & _
                                          IIf(IsNull(psSelected(6)), "", psSelected(6)))
                  .TextMatrix(.Row, 4) = IIf(IsNull(psSelected(7)), "", Format(psSelected(7), "#,##0.00"))
                  .TextMatrix(.Row, 5) = IIf(IsNull(psSelected(2)), "", psSelected(2))
                  .TextMatrix(.Row, 6) = IIf(IsNull(psSelected(8)), "", psSelected(8))
                  .TextMatrix(.Row, 7) = IIf(IsNull(psSelected(9)), "", Format(psSelected(9), "#,##0.00"))
               Else
                  pbExisting = True
                  Exit Sub
               End If
            End If
            
            For pnCtr = 1 To .Rows - 1
               If .Row <> pnCtr Then
                  If .TextMatrix(.Row, 6) = .TextMatrix(pnCtr, 6) Then
                     MsgBox "Duplicate IMEI Entry!!!" & vbCrLf & _
                     "Verify your Entry", vbCritical, "Warning"
                     For lnCtr = 2 To 7
                        .TextMatrix(.Row, lnCtr) = ""
                     Next
                     .SetFocus
                     pbExisting = True
                     Exit Sub
                  End If
               End If
            Next
            .Refresh
            If .TextMatrix(.Row, 1) <> "" Then
               .Rows = .Rows + 1
               .Row = .Row + 1
            End If
            .SetFocus

         Else
            MsgBox "IMEI NO. Not Existing!!!", vbCritical, "Warning"
            For pnCtr = 1 To .Cols
               .TextMatrix(.Row, pnCtr) = ""
            Next
            .Col = 1
            .SetFocus
            .Refresh
         End If
      Set oRS = Nothing
   End With
   
End Sub

Private Sub txtothers_GotFocus(Index As Integer)
   txtOthersGotfocus = True
   txtfieldGotfocus = False
   pnindex = Index
   oDriver.ColumnIndex = Index
   txtothers(Index).BackColor = &HE1FEFF
End Sub


Private Sub txtothers_LostFocus(Index As Integer)
   txtothers(Index).BackColor = &HFFFFFF
End Sub

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Tested  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 20, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'


'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Version 1    ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Finished  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 20, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Version 2    ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Add Export   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Finished  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 27, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Version 2    ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Add Branch   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Finished  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  July 09, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'



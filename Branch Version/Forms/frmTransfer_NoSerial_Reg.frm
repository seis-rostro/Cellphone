VERSION 5.00
Object = "{0A7B56A6-35D0-4533-91DA-1715D3A0DD3E}#1.1#0"; "xrGridControl.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransfer_NoSerial_Reg 
   BorderStyle     =   0  'None
   Caption         =   "Branch Transfer Register  (NO IMIE No.) "
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin xrGridEditor.GridEditor GridEditor1 
      Height          =   4980
      Left            =   150
      TabIndex        =   15
      Tag             =   "et0;eb0;et0;bc2"
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
      MOUSEICON       =   "frmTransfer_NoSerial_Reg.frx":0000
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
         Index           =   6
         Left            =   5040
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frmTransfer_NoSerial_Reg.frx":001C
         Top             =   150
         Width           =   2460
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1335
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   150
         Width           =   2460
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   900
         Index           =   4
         Left            =   1335
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmTransfer_NoSerial_Reg.frx":0022
         Top             =   660
         Width           =   8460
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   3
         Left            =   1335
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   405
         Width           =   2460
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   5040
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmTransfer_NoSerial_Reg.frx":0028
         Top             =   405
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
         Left            =   7830
         TabIndex        =   14
         Tag             =   "eb0;wb0"
         Top             =   150
         Width           =   1965
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivered By "
         Height          =   285
         Index           =   2
         Left            =   3945
         TabIndex        =   10
         Top             =   150
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No"
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   4
         Top             =   150
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   12
         Top             =   675
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Requested By "
         Height          =   285
         Index           =   0
         Left            =   3930
         TabIndex        =   8
         Top             =   420
         Width           =   1200
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   495
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   873
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   2
         Left            =   5040
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   105
         Width           =   4755
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   1
         Left            =   1335
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   105
         Width           =   2445
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   285
         Index           =   19
         Left            =   3960
         TabIndex        =   2
         Top             =   105
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transmittal No."
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   0
         Top             =   105
         Width           =   1200
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   4
      Left            =   10260
      TabIndex        =   21
      Top             =   4830
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
      Picture         =   "frmTransfer_NoSerial_Reg.frx":002E
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
      TabIndex        =   22
      Top             =   4830
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
      Picture         =   "frmTransfer_NoSerial_Reg.frx":07A8
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
      TabIndex        =   19
      Top             =   4410
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
      Picture         =   "frmTransfer_NoSerial_Reg.frx":0F22
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
      TabIndex        =   23
      Top             =   5250
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
      Picture         =   "frmTransfer_NoSerial_Reg.frx":169C
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
      TabIndex        =   20
      Top             =   4410
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
      Picture         =   "frmTransfer_NoSerial_Reg.frx":1E16
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
      TabIndex        =   24
      Top             =   5250
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
      Picture         =   "frmTransfer_NoSerial_Reg.frx":2590
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
      TabIndex        =   16
      Top             =   3570
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
      Picture         =   "frmTransfer_NoSerial_Reg.frx":2D0A
      CaptionAlign    =   0
      BackColor       =   14286077
      BackColorDown   =   8775418
      BorderColorFocus=   8775418
      BorderColorHover=   8775418
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   10260
      TabIndex        =   17
      Top             =   3570
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
      Picture         =   "frmTransfer_NoSerial_Reg.frx":3484
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   8
      Left            =   10260
      TabIndex        =   18
      Top             =   3990
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
      Picture         =   "frmTransfer_NoSerial_Reg.frx":3BFE
      CaptionAlign    =   0
   End
   Begin MSComCtl2.Animation Progress 
      Height          =   720
      Left            =   10425
      TabIndex        =   25
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
      TabIndex        =   26
      ToolTipText     =   "Void Transaction"
      Top             =   3150
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
      Picture         =   "frmTransfer_NoSerial_Reg.frx":4D10
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
      Top             =   1020
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   873
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   7
         Left            =   1335
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   105
         Width           =   2445
      End
      Begin VB.TextBox txtothers 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   5025
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   105
         Width           =   4755
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Origin"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   105
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   285
         Index           =   5
         Left            =   3975
         TabIndex        =   29
         Top             =   105
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmTransfer_NoSerial_Reg"
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

Dim oSerial As frmCP_Serial_Info
Dim txtfieldGotfocus As Boolean
Dim txtOthersGotfocus As Boolean
Dim pbnewitem As Boolean

Dim pnUserRights As Integer
Dim psUserID As String
Dim psUserName As String

Dim psSelected() As String
Dim Reference As String
Dim Drive As String
Dim void As Boolean

Dim pnindex As Integer
Dim Index As Integer
Dim pnCtr As Integer
Dim lnCtr As Integer

Dim Time As String
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
   oDriver.HideButton 0
   oDriver.HideButton 3
   oDriver.HideButton 6
   
'   xrFrame1(0).Enabled = False
   For lnCtr = 1 To 5
      txtfield(lnCtr).Enabled = True
   Next
End Sub

Private Sub ShowButton()
   oDriver.HideButton 4
   oDriver.HideButton 5
   oDriver.HideButton 7
   oDriver.ShowButton 0
   oDriver.ShowButton 3
   oDriver.ShowButton 6
   
'   xrFrame1(0).Enabled = True
   GridEditor1.SetFocus
End Sub

Private Sub Search_Transmittal()
Dim orig As String
Dim lsSQL As String
Dim lsCondition As String

   orig = oDriver.BrowseQuery
   Select Case pnindex
      Case 0
         lsCondition = " a.sTransNox like '%" & txtfield(0).Text & "' "
         lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
         oDriver.BrowseQuery = lsSQL
      Case 1
         lsCondition = " a.sReferNox like '%" & txtfield(1).Text & "'"
         lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
         oDriver.BrowseQuery = lsSQL
      Case 2
         lsCondition = " b.sBranchNm like '" & txtfield(2).Text & "%'"
         lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
         oDriver.BrowseQuery = lsSQL
      Case 7
         If oDriver.FieldValue(7) = "" Then
            lsCondition = " a.sOriginxx = '" & oApp.BranchCode & "'"
            lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
            oDriver.BrowseQuery = lsSQL
         Else
            lsCondition = " a.sOriginxx = '" & oDriver.FieldValue(7) & "'"
            lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
            oDriver.BrowseQuery = lsSQL
         End If
   End Select
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
                  & " FROM CP_Transfer_Master " _

   oDriver.BrowseQuery = "SELECT" _
                  & " Distinct" _
                  & " TOP 1000 " _
                  & " a.sTransNox, " _
                  & " a.sReferNox, " _
                  & " a.dTransact, " _
                  & " b.sBranchNm  " _
            & " FROM CP_Transfer_Master a " _
               & " LEFT JOIN Branch b " _
                  & " ON a.sDestinat = b.sBranchCd " _
               & " LEFT JOIN CP_Inventory_Ledger c " _
                  & " ON a.sTransNOx = c.sSourceNo " _
            & " ORDER BY a.dTransact Desc "
            
   oDriver.InitRecForm
   
   oDriver.BrowseFTitle(0) = "Transaction No"
   oDriver.BrowseFTitle(1) = "Transmittal No"
   oDriver.BrowseFTitle(2) = "Date"
   oDriver.BrowseFTitle(3) = "Destination"
   
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

Private Sub Print_Transaction()
Dim lnCtr As Integer
Dim lrs As New ADODB.Recordset
Dim lsSQL As String
Dim lrsDetail As New ADODB.Recordset

   Set lrs = New ADODB.Recordset
   lrs.Fields.Append "sField01", adVarChar, 150
   lrs.Fields.Append "sField02", adVarChar, 20
   lrs.Fields.Append "nField01", adInteger, 5
   lrs.Fields.Append "nField02", adInteger, 5
   lrs.Open

   'CP_Transfer_Master
   lsSQL = "SELECT" _
               & " a.sTransNox, " _
               & " a.sReferNox, " _
               & " a.sApproved, " _
               & " a.sDestinat, " _
               & " b.sBranchNm, " _
               & " c.sTownIDxx, " _
               & " b.sAddressx + ', ' + c.sTownName xAddressx " _
         & " FROM CP_Transfer_Master a " _
            & " LEFT JOIN Branch b " _
               & " ON a.sOriginxx = b.sBranchCd " _
            & " LEFT JOIN TownCity c " _
               & " ON b.sTownIDxx = c.sTownIDxx " _
         & " WHERE a.sTransNox = '" & oDriver.FieldValue(0) & "' " _

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
                     & " a.nQuantity, " _
                     & " b.sDescript, " _
                     & " c.sBrandNme, " _
                     & " d.sModelNme, " _
                     & " e.sColorNme, " _
                     & " b.sBarrcode  " _
                  & " FROM CP_Transfer_Detail a " _
                     & " LEFT JOIN CP_Inventory b " _
                        & " ON a.sStockIDx = b.sStockIDx " _
                     & " LEFT JOIN Brand c " _
                        & " ON b.sBrandIdx = c.sBrandIDx " _
                     & " LEFT JOIN Model d  " _
                        & " ON b.sModelIDx = d.sModelIDx " _
                     & " LEFT JOIN Color e  " _
                        & " ON b.sColorIDx = e.sColorIDx " _
                  & " WHERE a.sTransNox = '" & oRS("sTransNox") & "' " _
                  & " ORDER BY a.nEntryNox "
            
            If lrsDetail.State = adStateOpen Then lrsDetail.Close
            lrsDetail.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

            Do While Not lrsDetail.EOF
               lrs.AddNew
               lrs("sField01").Value = Trim(IIf(IsNull(lrsDetail("sBrandNme")), "", lrsDetail("sBrandNme")) _
                                       & " " & IIf(IsNull(lrsDetail("sModelNme")), "", lrsDetail("sModelNme")) _
                                       & " " & IIf(IsNull(lrsDetail("sDescript")), "", lrsDetail("sDescript")) _
                                       & " " & IIf(IsNull(lrsDetail("sColorNme")), "", lrsDetail("sColorNme")))
               lrs("sField02").Value = lrsDetail("sBarrCode")
               lrs("nField01").Value = lrsDetail("nEntryNox")
               lrs("nField02").Value = lrsDetail("nQuantity")
               lrsDetail.MoveNext
            Loop
            
         oRS.MoveNext
      Next

      Branch = txtfield(2)
      getBranch Code, Branch, Address
      
      Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Transmittal_NoSerial.rpt")
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

Private Sub InitGrid()
    
   With GridEditor1
      .Rows = 2
      .Cols = 8
      .Font = "MS Sans Serif"
      
      'column title
      .TextMatrix(0, 1) = "Barr Code"
      .TextMatrix(0, 2) = "Particulars"
      .TextMatrix(0, 3) = "SRP"
      .TextMatrix(0, 4) = "Stock ID"
      .TextMatrix(0, 5) = "Qty"
      .TextMatrix(0, 6) = "QOH"
      .TextMatrix(0, 7) = "Purchase P"
   
      'column width
      .ColWidth(0) = 400
      .ColWidth(1) = 2200
      .ColWidth(2) = 4750
      .ColWidth(3) = 1000
      .ColWidth(4) = 0
      .ColWidth(5) = 670
      .ColWidth(6) = 670
      .ColWidth(7) = 0
              
      .ColFormat(1) = ">"
      
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 6
      .ColAlignment(6) = 6
      .ColAlignment(7) = 6
      
      .ColDefault(5) = 0
      .ColNumberOnly(5) = True
      
      .ColEnabled(2) = False
      .ColEnabled(3) = False
      .ColEnabled(4) = False
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

Private Sub ShowGrid()
Dim lsSQL As String
Dim lnCtr As Integer
Dim lnrow As Long

   lsSQL = "SELECT" _
               & " Distinct " _
               & " a.sTransNox, " _
               & " a.nEntryNox, " _
               & " a.sStockIdx, " _
               & " a.nQuantity, " _
               & " a.nUnitPrce, " _
               & " b.sBarrCode, " _
               & " b.sDescript, " _
               & " b.sCategIDx, " _
               & " c.sBrandNme, " _
               & " f.sColorNme, " _
               & " d.sModelNme, " _
               & " e.nQtyOnHnd, " _
               & " b.nSelPrice  "
   lsSQL = lsSQL _
         & " FROM CP_Transfer_Detail a " _
               & " LEFT JOIN CP_Inventory b " _
                  & " ON a.sStockIDx = b.sStockIDx " _
               & " LEFT JOIN CP_Inventory_Master e " _
                  & " ON a.sStockIDx = e.sStockIDx " _
               & " LEFT JOIN Brand c " _
                  & " ON b.sBrandIDx = c.sBrandIDx " _
               & " LEFT JOIN Model d " _
                  & " ON b.sModelIDx = d.sModelIDx " _
               & " LEFT JOIN Color f " _
                  & " ON b.sColorIDx = f.sColorIDx " _
         & " WHERE a.sTransNox = '" & oDriver.FieldValue(0) & " '" _
               & " AND e.sBranchcd = '" & oApp.BranchCode & "'" _
         & " ORDER BY a.nEntryNox "
   If oRS.State = adStateOpen Then oRS.Close
   oRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If oRS.RecordCount <> 0 Then
      With GridEditor1
         .Rows = oRS.RecordCount + 1
         For lnCtr = 0 To oRS.RecordCount - 1
            .TextMatrix(lnCtr + 1, 0) = oRS("nEntryNox")
            .TextMatrix(lnCtr + 1, 1) = oRS("sBarrCode")
            .TextMatrix(lnCtr + 1, 2) = Trim(IIf(IsNull(oRS("sBrandNme")), "", oRS("sBrandNme")) & " " & _
                                          IIf(IsNull(oRS("sModelNme")), "", oRS("sModelNme")) & " " & _
                                          IIf(IsNull(oRS("sDescript")), "", oRS("sDescript")) & " " & _
                                          IIf(IsNull(oRS("sColorNme")), "", oRS("sColorNme")))
            .TextMatrix(lnCtr + 1, 3) = Format(oRS("nSelPrice"), "#,##0.00")
            .TextMatrix(lnCtr + 1, 4) = oRS("sStockIDx")
            .TextMatrix(lnCtr + 1, 5) = oRS("nQuantity")
            .TextMatrix(lnCtr + 1, 6) = oRS("nQtyOnHnd")
            .TextMatrix(lnCtr + 1, 7) = Format(oRS("nUnitPrce"), "#,##0.00")
            oRS.MoveNext
         Next
      .ColEnabled(1) = True
      .ColEnabled(5) = True
      If .Rows > 20 Then
         .ColWidth(2) = 4500
      Else
         .ColWidth(2) = 4750
      End If

      End With
      HideButton
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
      .ColEnabled(5) = True
   End With
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
      Call Modified("CP_Transfer_Master", "sTransNox = '" & oDriver.FieldValue(0) & "' ")
End Select
End Sub

Private Sub GridEditor1_AddingRow(Cancel As Boolean)
   With GridEditor1
      If .TextMatrix(.Row, 1) = "" Then
         MsgBox "Invalid Bar Code!!!", vbCritical, "Warning"
         Cancel = True
      ElseIf .TextMatrix(.Row, 5) = "0" Or .TextMatrix(.Row, 5) = "" Then
         MsgBox "Invalid Quantity!!!", vbCritical, "Warning"
         Cancel = True
      End If
   End With
End Sub

Private Sub GridEditor1_EditorValidate(Cancel As Boolean)
Dim lsSQL As String

   With GridEditor1
      If .Col = 5 Then
         If CLng(.TextMatrix(.Row, 5)) = 0# Then
            MsgBox "Invalid Quantity!!!", vbCritical, "Warning"
            .Col = 5
         End If
      End If
   End With
End Sub

Private Sub GridEditor1_RowColChange()
With GridEditor1
   If .TextMatrix(.Row, 1) <> "" And .TextMatrix(.Row, 4) = "" Then
      .Col = 1
   End If
End With
End Sub

Private Sub oDriver_InitValue()
   label.Caption = "UNKNOWN"
   txtothers(0).Tag = ""
   txtothers(0).Text = ""
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
   oDriver.FieldValue(3) = Format(oDriver.FieldValue(3), "m/d/yyyy")
   txtothers(0).Enabled = False
   Branch = txtfield(2)
   
End Sub

Private Sub oDriver_SaveComplete()
   HideButton
   MsgBox "Transaction Successfully Updated!!!", vbInformation, "Information"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   With GridEditor1
      If txtfield(1).Text = "" Then
         MsgBox "Invalid Transmittal No. Detected!!!", vbCritical, "Warning"
         txtfield(1).SetFocus
         Cancel = True
      ElseIf oDriver.FieldValue(2) = "" Then
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
         Cancel = Not Save_CP_Transfer_Detail
            If Cancel Then Exit Sub
         Cancel = Not Update_CP_Inventory
            If Cancel Then Exit Sub
            
         oDriver.FieldValue(3) = CDate(txtfield(3).Text) & " " & Time
         oDriver.FieldValue(7) = oApp.BranchCode
         oDriver.FieldValue(8) = .TextMatrix(.Rows - 1, 0)
         oDriver.FieldValue(9) = 0  'cTranStat
         oDriver.FieldValue(10) = 0  'cReceived
         Reference = oDriver.FieldValue(1)
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
         Case 2
            oDriver.RecordSearch txtfield(Index).Text
            If txtfield(Index).Text <> "" Then SetNextFocus
         Case 7
            oDriver.RecordSearch txtfield(Index).Text
            cmdButton(4).SetFocus
      End Select
      oDriver.BrowseQuery = lsSQL
      oDriver.BrowseRecord
      oDriver.BrowseQuery = orig
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

Delete_Transaction = True
On Error GoTo errProc

   'Roll Back Quantity
   Call RollBack_Qty("CP_Transfer_Detail", "'" & oDriver.FieldValue(0) & "'")

   'Roll Back EntryNo
   Call Recalc_Ledger("CP_Transfer_Detail", "'" & oDriver.FieldValue(0) & "'", "'CPDv'", "'" & oApp.BranchCode & "'")
   
   'Update CP_Transfer_Master
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
      lsSQL = "DELETE CP_Transfer_Detail " _
               & " WHERE sTransNox = '" & oDriver.FieldValue(0) & "'"
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
      If lnrow <> 0 Then oApp.RegisDelete lsSQL
   End If

               
endProc:
   Exit Function
errProc:
   Delete_Transaction = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Function Save_CP_Transfer_Detail() As Boolean
Dim lsSQL As String
Dim lnrow As Long
   
Save_CP_Transfer_Detail = True
On Error GoTo errProc
   
   With GridEditor1
      'Insert Record
      For pnCtr = 1 To .Rows - 1
         If .TextMatrix(pnCtr, 1) = "" Then Exit For
         lsSQL = "INSERT INTO CP_Transfer_Detail " _
               & "( sTransNox, " _
               & "  nEntryNox, " _
               & "  sStockIDx, " _
               & "  nQuantity, " _
               & "  nUnitPrce, " _
               & "  dModified) " _
         & "VALUES " _
               & "('" & oDriver.FieldValue(0) & "', " _
               & "'" & .TextMatrix(pnCtr, 0) & "', " _
               & "'" & .TextMatrix(pnCtr, 4) & "', " _
               & "'" & CLng(.TextMatrix(pnCtr, 5)) & "', " _
               & "'" & CDbl(.TextMatrix(pnCtr, 7)) & "', " _
               & " getdate())"
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
      Next
      
      If lnrow <= 0 Then
         MsgBox "Unable to Save Transfer Detail!!!" & vbCrLf & vbCrLf & _
         "Notify ROSALYN LAZO DESCALLAR for Assistance", vbCritical, "Warning"
         Save_CP_Transfer_Detail = False
         GoTo endProc
      End If

   End With

endProc:
   Exit Function
errProc:
   Save_CP_Transfer_Detail = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Function Update_CP_Inventory() As Boolean
Dim lsSQL As String
Dim lnrow As Long
Dim lnEntry As Integer
Dim QOH As Integer

Update_CP_Inventory = True
On Error GoTo errProc
   
   With GridEditor1
         For pnCtr = 1 To .Rows - 1
            If .TextMatrix(pnCtr, 1) = "" Then Exit For
            
            'Get last Entry No.
            lnEntry = getEntryNo("CP_Inventory_Ledger", "'" & .TextMatrix(pnCtr, 4) & "'", "'" & oApp.BranchCode & "'")
     
            'Get QOH
            QOH = getQuantity("'" & .TextMatrix(pnCtr, 4) & "'", "'" & oApp.BranchCode & "'") _
                     - .TextMatrix(pnCtr, 5)
            
               'Add Record, CP_Inventroy_Ledger
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
                     & "('" & .TextMatrix(pnCtr, 4) & "', " _
                     & "'" & oApp.BranchCode & "'," _
                     & "'" & oDriver.FieldValue(2) & "', " _
                     & "'CPDv' , " _
                     & "'" & oDriver.FieldValue(0) & "', " _
                     & " 0, " _
                     & "'" & CLng(.TextMatrix(pnCtr, 5)) & "', " _
                     & "'" & CLng(QOH) & "', " _
                     & "'" & lnEntry & "', " _
                     & "'" & CDate(oDriver.FieldValue(3)) & " " & Time & "', " _
                     & " getdate())"
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
            
               'Update QOH, CP_Inventory_Master
               lsSQL = "UPDATE CP_Inventory_Master SET" _
                     & " nQtyOnHnd = '" & CLng(QOH) & "', " _
                     & " sModified = '" & Encrypt(oApp.UserID) & "', " _
                     & " dModified = getdate() " _
               & " WHERE sStockIDx = '" & .TextMatrix(pnCtr, 4) & "' " _
                     & " And sBranchCd = '" & oApp.BranchCode & "' "
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
               
               'Update QOH, CP_Inventory
               lsSQL = "UPDATE CP_Inventory SET" _
                     & " sModified = '" & Encrypt(oApp.UserID) & "', " _
                     & " dModified = getdate() " _
               & " WHERE sStockIDx = '" & .TextMatrix(pnCtr, 4) & "' "
               oApp.Connection.Execute lsSQL, lnrow, adCmdText

               
         Next
         
         If lnrow <= 0 Then
            MsgBox "Unable to Update Inventory Ledger!!!" & vbCrLf & vbCrLf & _
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
            SearchBarCode
         End If
      End If
   End With
End Sub

Private Sub SearchBarCode()
Dim lsSQL As String
Dim lsCondition As String
Dim lsSearch As String
Dim lnCtr As Integer
   
   With GridEditor1
      lsSQL = "SELECT" _
            & " a.sBarrcode, " _
            & " a.sStockIDx, " _
            & " b.sBrandNme, " _
            & " c.sModelNme, " _
            & " a.sDescript, " _
            & " d.sColorNme, " _
            & " a.nSelPrice, " _
            & " e.nQtyOnHnd, " _
            & " a.cWdSerial, " _
            & " a.nPurPrice  " _
         & " FROM CP_Inventory a " _
            & " LEFT JOIN CP_Inventory_Master e " _
               & " ON a.sStockIdx = e.sStockIDx " _
            & " LEFT JOIN Brand b " _
               & " ON a.sBrandIdx = b.sBrandIdx " _
            & " LEFT JOIN Model c " _
               & " ON a.sModelIdx = c.sModelIdx " _
            & " LEFT JOIN Color d " _
               & " ON a.sColorIDx = d.sColorIDx " _
         & " WHERE a.sBarrcode like  '%" & .TextMatrix(.Row, 1) & "%' " _
            & " AND a.cWdSerial = 0 and a.cWalletxx = 0 and a.cCellLoad = 0 " _
            & " AND e.sBranchCd = '" & oApp.BranchCode & "'"
      If oRS.State = adStateOpen Then oRS.Close
      oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

      If Not oRS.EOF Then
         If oRS.RecordCount = 1 Then
            .TextMatrix(.Row, 1) = IIf(IsNull(oRS(0)), "", oRS(0))
            .TextMatrix(.Row, 2) = Trim(IIf(IsNull(oRS(2)), "", oRS(2)) & " " & _
                                    IIf(IsNull(oRS(3)), "", oRS(3)) & " " & _
                                    IIf(IsNull(oRS(4)), "", oRS(4)) & " " & _
                                    IIf(IsNull(oRS(5)), "", oRS(5)))
            .TextMatrix(.Row, 3) = IIf(IsNull(oRS(6)), 0, Format(oRS(6), "#,##0.00"))
            .TextMatrix(.Row, 4) = IIf(IsNull(oRS(1)), "", oRS(1))
            .TextMatrix(.Row, 6) = IIf(IsNull(oRS(7)), 0, oRS(7))
            .TextMatrix(.Row, 7) = IIf(IsNull(oRS(9)), 0, Format(oRS(9), "#,##0.00"))
            .Col = 5
         Else
            lsSearch = KwikSearch(oApp, lsSQL, _
                       "sBarrcode»sBrandNme»sModelNme»sDescript»sColorNme", _
                       "Bar Code»Brand»Model»Description»Color")
            If lsSearch <> "" Then
               psSelected = Split(lsSearch, "»")
               oDriver.LookupValue(0) = psSelected(0)
               .TextMatrix(.Row, 1) = IIf(IsNull(psSelected(0)), "", psSelected(0))
               .TextMatrix(.Row, 2) = Trim(IIf(IsNull(psSelected(2)), "", psSelected(2)) & " " & _
                                    IIf(IsNull(psSelected(3)), "", psSelected(3)) & " " & _
                                    IIf(IsNull(psSelected(4)), "", psSelected(4)) & " " & _
                                    IIf(IsNull(psSelected(5)), "", psSelected(5)))
               .TextMatrix(.Row, 3) = IIf(IsNull(psSelected(6)), 0, Format(psSelected(6), "#,##0.00"))
               .TextMatrix(.Row, 4) = IIf(IsNull(psSelected(1)), "", psSelected(1))
               .TextMatrix(.Row, 6) = IIf(IsNull(psSelected(7)), 0, psSelected(7))
               .TextMatrix(.Row, 7) = IIf(IsNull(psSelected(9)), 0, Format(psSelected(9), "#,##0.00"))
               .Col = 5
            End If
         End If
         .SetFocus
         .Refresh
      Else
         For pnCtr = 1 To .Cols
            .TextMatrix(.Row, pnCtr) = ""
         Next
         .Col = 1
         MsgBox "Bar Code Not Existing!!!", vbInformation, "Information"
      End If
      
      For lnCtr = 1 To .Rows - 1
         If lnCtr <> .Row And .TextMatrix(.Row, 3) <> "" Then
            If .TextMatrix(lnCtr, 3) = .TextMatrix(.Row, 3) Then
               MsgBox "Duplicate Entry!!!" & vbCrLf & vbCrLf & _
               "Update Quantity of Row" & " " & lnCtr, vbCritical, "Warning"
               For pnCtr = 1 To .Cols
                  .TextMatrix(.Row, pnCtr) = ""
               Next
               .SetFocus
            End If
         End If
      Next
      .Refresh
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

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Add Export   ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Finished  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  February 21, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  March 24, 2008  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'Include cWdSerial in Table CP_Inventory




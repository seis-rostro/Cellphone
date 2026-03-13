VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmUtility 
   BackColor       =   &H00FDC9FE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MY UTILITY"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6510
   Begin xrControl.xrButton xrButton 
      Height          =   405
      Index           =   0
      Left            =   195
      TabIndex        =   15
      Top             =   2310
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   714
      Caption         =   "Create Branch Acceptance"
      AccessKey       =   "Create Branch Acceptance"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388736
      ForeColor       =   16777215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FDC9FE&
      Height          =   1200
      Left            =   180
      TabIndex        =   4
      Top             =   645
      Width           =   6135
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   975
         TabIndex        =   13
         Top             =   855
         Width           =   1770
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   975
         TabIndex        =   11
         Top             =   555
         Width           =   1770
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   3915
         TabIndex        =   9
         Top             =   555
         Width           =   2055
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   3915
         TabIndex        =   7
         Top             =   255
         Width           =   2055
      End
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   975
         TabIndex        =   5
         Top             =   255
         Width           =   1770
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "dTransact"
         Height          =   270
         Left            =   75
         TabIndex        =   14
         Top             =   870
         Width           =   780
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "sBranchCd"
         Height          =   270
         Left            =   75
         TabIndex        =   12
         Top             =   585
         Width           =   780
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "nEntryNox"
         Height          =   270
         Left            =   3060
         TabIndex        =   10
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "sTransNox"
         Height          =   270
         Left            =   3030
         TabIndex        =   8
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "sStockIDx"
         Height          =   270
         Left            =   90
         TabIndex        =   6
         Top             =   300
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FDC9FE&
      Height          =   600
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   6150
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   765
         TabIndex        =   1
         Top             =   180
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   2970
         TabIndex        =   3
         Top             =   210
         Width           =   3000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   270
         Left            =   90
         TabIndex        =   2
         Top             =   210
         Width           =   675
      End
   End
   Begin xrControl.xrButton xrButton 
      Height          =   405
      Index           =   1
      Left            =   195
      TabIndex        =   16
      Top             =   2730
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   714
      Caption         =   "Create Branch Delivery "
      AccessKey       =   "Create Branch Delivery "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388736
      ForeColor       =   16777215
   End
   Begin xrControl.xrButton xrButton 
      Height          =   405
      Index           =   2
      Left            =   195
      TabIndex        =   17
      Top             =   3150
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   714
      Caption         =   "Create Inventory Master"
      AccessKey       =   "Create Inventory Master"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388736
      ForeColor       =   16777215
   End
   Begin xrControl.xrButton xrButton 
      Height          =   405
      Index           =   3
      Left            =   195
      TabIndex        =   18
      Top             =   3570
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   714
      Caption         =   "Create Initial Inventory"
      AccessKey       =   "Create Initial Inventory"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388736
      ForeColor       =   16777215
   End
   Begin xrControl.xrButton xrButton 
      Height          =   405
      Index           =   4
      Left            =   195
      TabIndex        =   19
      Top             =   3990
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   714
      Caption         =   "Update Ledger Entry No"
      AccessKey       =   "Update Ledger Entry No"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388736
      ForeColor       =   16777215
   End
   Begin xrControl.xrButton xrButton 
      Height          =   405
      Index           =   6
      Left            =   3270
      TabIndex        =   20
      Top             =   2310
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   714
      Caption         =   "Existing IMEI"
      AccessKey       =   "Existing IMEI"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388736
      ForeColor       =   16777215
   End
   Begin xrControl.xrButton xrButton 
      Height          =   405
      Index           =   7
      Left            =   3270
      TabIndex        =   21
      Top             =   2730
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   714
      Caption         =   "Create Serial Ledger Sales"
      AccessKey       =   "Create Serial Ledger Sales"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388736
      ForeColor       =   16777215
   End
   Begin xrControl.xrButton xrButton 
      Height          =   405
      Index           =   8
      Left            =   3270
      TabIndex        =   22
      Top             =   3150
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   714
      Caption         =   "Update Existing IMEI"
      AccessKey       =   "Update Existing IMEI"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388736
      ForeColor       =   16777215
   End
   Begin xrControl.xrButton xrButton 
      Height          =   405
      Index           =   9
      Left            =   3270
      TabIndex        =   23
      Top             =   3570
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   714
      Caption         =   "Create Inventory Ledger Sales"
      AccessKey       =   "Create Inventory Ledger Sales"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388736
      ForeColor       =   16777215
   End
   Begin xrControl.xrButton xrButton 
      Height          =   405
      Index           =   10
      Left            =   3270
      TabIndex        =   24
      Top             =   3990
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   714
      Caption         =   "Update nQtyOnHnd Accessories"
      AccessKey       =   "Update nQtyOnHnd Accessories"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388736
      ForeColor       =   16777215
   End
   Begin xrControl.xrButton xrButton 
      Height          =   405
      Index           =   12
      Left            =   195
      TabIndex        =   25
      Top             =   1890
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   714
      Caption         =   "Fix Entry No According to Date"
      AccessKey       =   "Fix Entry No According to Date"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388736
      ForeColor       =   16777215
   End
   Begin xrControl.xrButton xrButton 
      Height          =   405
      Index           =   5
      Left            =   195
      TabIndex        =   26
      Top             =   4410
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   714
      Caption         =   "Update Ledger nQtyOnHnd"
      AccessKey       =   "Update Ledger nQtyOnHnd"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388736
      ForeColor       =   16777215
   End
   Begin xrControl.xrButton xrButton 
      Height          =   405
      Index           =   11
      Left            =   3270
      TabIndex        =   27
      Top             =   4410
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   714
      Caption         =   "Exit"
      AccessKey       =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388736
      ForeColor       =   16777215
   End
End
Attribute VB_Name = "frmUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'>>>>>>>>>>>>>>>>Installation Notes<<<<<<<<<<<<<<<<<<<<<<<

'***** First Check if CPDV or CPDl existing *****
' Then Insert CPDl
' Select From CP_Inventory_Ledger where dmodified = getdate and sbranchd in ('01','07')
' Delete -Not to be Added in Inventory Ledger of other Branch, Update Branch Ledger Only

'***** Second If Ledger Existing then Create CP_Inventory_Master *****

'***** Third Update nEntryNox CP_Inventory_Ledger *****

' Select From CP_Inventory_Ledger where Top 1 sSourceCd = CPDv and sBranchcd = text1.text
' if sSourceCd = 'CPDv' then insert Intitial Inventory
' Then Update EntryNo Again

'***** Fourth Update nQtyOnhnd CP_Inventory_Ledger *****

'Update CP_Serial_Master, Make Existing IMEI cLocation = 9

'update CP_Serial_Master of Not Existing IMEI, Add CP_Serial_Ledger sales

'Update CP_Serial_Master, Make Existing IMEI back to cLocation = 1

'Update CP_Inventory_Master nQtyOnHnd, Add CP_Inventory_Ledger sSourceCd = text1.text & 'Captured'
'Based On Existing IMEI

'Accessories and Cards

Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private poFileSys As FileSystemObject

Dim Reference As String
Dim lsSQL As String
Dim rsTarget As ADODB.Recordset
Dim oFolder As Folder
Dim oFiles As Files
Dim oFile As File
Dim oFileObject As New FileSystemObject

Dim rsSource As ADODB.Recordset
Dim lnrow As Long
Dim rsMain As ADODB.Recordset
Dim rsBranch As ADODB.Recordset
Dim Quantity As Integer
Dim Entry As Integer
Dim oRS As New ADODB.Recordset
Dim lrs As New ADODB.Recordset
Dim lnEntry As Integer

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If bLoaded = False Then
      bLoaded = True
      Label2.Caption = Format(Date, "dddd MMMM dd, yyyy")
   End If
End Sub

Private Sub Form_Load()
   CenterChildForm mdiMain, Me
   bLoaded = False

   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me

End Sub

Private Sub ClearFields()
Dim lnCtr As Integer
   For lnCtr = 0 To 4
      txtfield(lnCtr).Text = ""
   Next
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   If Text1.Text = "" Then MsgBox "Invalid Branch"
End Sub

Private Function Insert_CPDl() As Boolean

Insert_CPDl = True
On Error GoTo errProc

    Set rsTarget = New ADODB.Recordset
    lsSQL = "SELECT * From CP_Inventory_Ledger " _
         & "WHERE ssourcecd = 'CPDv' " _
         & " AND slocation = '" & Text1.Text & "'" _
         & "ORDER BY sstockidx "
    If rsTarget.State = adStateOpen Then rsTarget.Close
    rsTarget.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

    Do While Not rsTarget.EOF
        Set rsSource = New ADODB.Recordset
        lsSQL = "SELECT * From CP_Inventory_Ledger " _
             & "WHERE ssourcecd = 'CPDl' " _
             & "AND sBranchcd = '" & Text1.Text & "'" _
             & "AND sStockIDx = '" & rsTarget("sStockIDx") & "'" _
             & "AND sSourceNo = '" & rsTarget("sSourceNo") & "'"
        If rsSource.State = adStateOpen Then rsSource.Close
        rsSource.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

        If rsSource.RecordCount = 0 Then
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
                           & "('" & rsTarget("sstockidx") & "', " _
                           & "'" & Text1.Text & "', " _
                           & "'" & rsTarget("sBranchcd") & "', " _
                           & "'CPDl', " _
                           & "'" & rsTarget("ssourceno") & "', " _
                           & "'" & rsTarget("nQtyoutxx") & "', " _
                           & "'0', " _
                           & "'" & rsTarget("nqtyonhnd") & "', " _
                           & "'" & rsTarget("nentrynox") & "', " _
                           & "'" & rsTarget("dtransact") & "', " _
                           & " '1/19/1979' )"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
        End If
        txtfield(0).Text = rsTarget("sStockIDx")
        txtfield(1).Text = rsTarget("sBranchCd")
        txtfield(2).Text = rsTarget("dTransact")
        txtfield(3).Text = rsTarget("sSourceNo")
        txtfield(4).Text = rsTarget("nEntryNox")
        DoEvents
        rsTarget.MoveNext
    Loop
    MsgBox "Tapos na Po"
    ClearFields
    Set rsTarget = Nothing
    Set rsSource = Nothing
endProc:
   Exit Function
errProc:
   Insert_CPDl = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Function Insert_CPDv() As Boolean

Insert_CPDv = True
On Error GoTo errProc

    Set rsTarget = New ADODB.Recordset
    lsSQL = "SELECT * From CP_Inventory_Ledger " _
         & "WHERE sSourcecd = 'CPDl' " _
         & " AND slocation = '" & Text1.Text & "'" _
         & "ORDER BY sstockidx "
    If rsTarget.State = adStateOpen Then rsTarget.Close
    rsTarget.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

    Do While Not rsTarget.EOF
        Set rsSource = New ADODB.Recordset
        lsSQL = "SELECT * From CP_Inventory_Ledger " _
             & "WHERE ssourcecd = 'CPDv' " _
             & "AND sBranchcd = '" & Text1.Text & "'" _
             & "AND sStockIDx = '" & rsTarget("sStockIDx") & "'" _
             & "AND sSourceNo = '" & rsTarget("sSourceNo") & "'"
        If rsSource.State = adStateOpen Then rsSource.Close
        rsSource.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

        If rsSource.RecordCount = 0 Then
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
                           & "('" & rsTarget("sstockidx") & "', " _
                           & "'" & Text1.Text & "', " _
                           & "'" & rsTarget("sBranchcd") & "', " _
                           & "'CPDv', " _
                           & "'" & rsTarget("ssourceno") & "', " _
                           & "'0', " _
                           & "'" & rsTarget("nQtyinxxx") & "', " _
                           & "'" & rsTarget("nqtyonhnd") & "', " _
                           & "'" & rsTarget("nentrynox") & "', " _
                           & "'" & rsTarget("dtransact") & "', " _
                           & " getdate())"
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
        End If
        txtfield(0).Text = rsTarget("sStockIDx")
        txtfield(1).Text = rsTarget("sBranchCd")
        txtfield(2).Text = rsTarget("dTransact")
        txtfield(3).Text = rsTarget("sSourceNO")
        txtfield(4).Text = rsTarget("nEntryNox")
        DoEvents
        rsTarget.MoveNext
    Loop
    MsgBox "Tapos na Po"
    ClearFields
    Set rsTarget = Nothing
    Set rsSource = Nothing
endProc:
   Exit Function
errProc:
   Insert_CPDv = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'If Ledger Existing then Create CP_Inventory_Master
Private Function Insert_Inventory_Master() As Boolean

Insert_Inventory_Master = True
On Error GoTo errProc

    Set rsTarget = New ADODB.Recordset
    lsSQL = "SELECT sStockIDx " _
         & "From CP_Inventory_Ledger " _
         & "WHERE sBranchcd = '" & Text1.Text & "' " _
         & "GROUP BY sStockIDx " _
         & "ORDER BY sstockidx "
         
    If rsTarget.State = adStateOpen Then rsTarget.Close
    rsTarget.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

    Do While Not rsTarget.EOF
        Set rsSource = New ADODB.Recordset
        lsSQL = "SELECT * From CP_Inventory_Master " _
             & "WHERE sBranchcd = '" & Text1.Text & "' " _
             & "AND sStockIDx = '" & rsTarget("sStockIDx") & "'"
        If rsSource.State = adStateOpen Then rsSource.Close
        rsSource.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

        If rsSource.RecordCount = 0 Then
            Set rsMain = New ADODB.Recordset
            lsSQL = "SELECT * From CP_Inventory_Master " _
                 & "WHERE sBranchcd = '01' " _
                 & "AND sStockIDx = '" & rsTarget("sStockIDx") & "'"
            If rsMain.State = adStateOpen Then rsMain.Close
            rsMain.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

            Set rsBranch = New ADODB.Recordset
            lsSQL = "SELECT * From CP_Inventory_Ledger " _
                 & "WHERE sBranchcd = '" & Text1.Text & "'" _
                 & "AND sStockIDx = '" & rsTarget("sStockIDx") & "'" _
                 & "ORDER BY dTransact desc "
            If rsBranch.State = adStateOpen Then rsBranch.Close
            rsBranch.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

            lsSQL = "INSERT INTO CP_Inventory_Master " _
                     & "( sStockIDx, " _
                     & "  sBranchCd, " _
                     & "  nBegQtyxx, " _
                     & "  nQtyOnHnd, " _
                     & "  nReorderx, " _
                     & "  nMinLevel, " _
                     & "  nMaxLevel, " _
                     & "  dBegInvxx, " _
                     & "  cRecdStat, " _
                     & "  sModified, " _
                     & "  dModified) " _
                         & "VALUES " _
                         & "('" & rsMain("sstockIDx") & "', " _
                         & "'" & Text1.Text & "', " _
                         & "'0', " _
                         & "'" & CLng(rsBranch("nQtyOnhnd")) & "', " _
                         & "'1', " _
                         & "'1', " _
                         & "'1', " _
                         & "'3/15/2007', " _
                         & " '" & xeRecStateActive & "', " _
                         & " '" & Encrypt(oApp.UserID) & "', " _
                         & " getdate())"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
        End If
        txtfield(0).Text = rsTarget("sStockIDx")
        DoEvents
        rsTarget.MoveNext
    Loop
    MsgBox "Tapos na Po"
    ClearFields
    Set rsTarget = Nothing
    Set rsSource = Nothing
    Set rsMain = Nothing
    Set rsBranch = Nothing

endProc:
   Exit Function
errProc:
   Insert_Inventory_Master = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

' Select From CP_Inventory_Ledger where Top 1 sSourceCd = CPDv and sBranchcd = text1.text
' if sSourceCd = 'CPDv' then insert Intitial Inventory
Private Function Insert_Initial_Inventory() As Boolean

Insert_Initial_Inventory = True
On Error GoTo errProc

   Set rsTarget = New ADODB.Recordset
   lsSQL = "SELECT * From CP_Inventory_Master " _
        & "WHERE sBranchcd = '" & Text1.Text & "' " _
        & "ORDER BY sstockidx "
   If rsTarget.State = adStateOpen Then rsTarget.Close
   rsTarget.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   Do While Not rsTarget.EOF
      Set rsSource = New Recordset
      lsSQL = " SELECT " _
            & " sStockIDx, " _
            & " nQtyOutxx  " _
         & " FROM CP_Inventory_Ledger " _
         & " WHERE sBranchCd = '" & Text1.Text & "'" _
            & " AND sStockIDx = '" & rsTarget("sStockIdx") & "'" _
            & " AND nEntryNox = 1 " _
            & " AND sSourceCd = 'CPDv' "
      If rsSource.State = adStateOpen Then rsSource.Close
      rsSource.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
      
      If rsSource.RecordCount = 1 Then
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
                & "('" & rsSource("sStockIDx") & "' ," _
                & "'" & Text1.Text & "', " _
                & "'" & Text1.Text & "', " _
                & " 'CPAd', " _
                & " '99000001', " _
                & "'" & rsSource("nQtyOutxx") & "', " _
                & " '0', " _
                & "'" & rsSource("nQtyOutxx") & "', " _
                & " '1', " _
                & "'3/15/2007'," _
                & " getdate())"
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
      End If
      txtfield(0).Text = rsTarget("sStockIDx")
      txtfield(1).Text = rsTarget("sBranchCd")
      DoEvents
      rsTarget.MoveNext
   Loop
   MsgBox "Tapos na Po"
   ClearFields
   Set rsTarget = Nothing
   Set rsSource = Nothing
endProc:
   Exit Function
errProc:
   Insert_Initial_Inventory = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'Update nEntryNox CP_Inventory_Ledger
Private Function Update_Entry() As Boolean

Update_Entry = True
On Error GoTo errProc
    
    Set rsTarget = New ADODB.Recordset
    lsSQL = "SELECT * From CP_Inventory_Master" _
         & " WHERE sBranchcd = '" & Text1.Text & "' " _
         & " ORDER BY sstockidx "
    If rsTarget.State = adStateOpen Then rsTarget.Close
    rsTarget.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

    Do While Not rsTarget.EOF
        Entry = 0
        Set rsSource = New ADODB.Recordset
        lsSQL = "SELECT * From CP_Inventory_Ledger " _
             & "Where sStockIDx = '" & rsTarget("sStockIDx") & "'" _
               & " AND sBranchcd = '" & rsTarget("sBranchcd") & "'" _
             & "Order by dTransact "
        If rsSource.State = adStateOpen Then rsSource.Close
        rsSource.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
        If rsSource.RecordCount <> 0 Then
            Do While Not rsSource.EOF
               Entry = Entry + 1
               lsSQL = "UPDATE CP_Inventory_Ledger SET " _
                        & " nEntryNox = '" & Entry & "'," _
                        & " dModified = '10/15/2007 9:30:00 AM' " _
                   & " WHERE sStockIDx = '" & rsSource("sStockIDx") & "'" _
                       & " AND sBranchCd = '" & rsSource("sBranchcd") & "'" _
                       & " AND sSourceNo = '" & rsSource("sSourceNo") & "'" _
                       & " AND nEntryNox = '" & rsSource("nEntryNox") & "'"
               oApp.Connection.Execute lsSQL, lnrow, adCmdText
               txtfield(0).Text = rsSource("sStockIDx")
               txtfield(1).Text = rsSource("sBranchCd")
               txtfield(2).Text = rsSource("dTransact")
               txtfield(3).Text = rsSource("sSourceNO")
               txtfield(4).Text = rsSource("nEntryNox")
               DoEvents
               rsSource.MoveNext
            Loop
        End If
        rsTarget.MoveNext
        Entry = 0
   Loop
   MsgBox "Tapos na Po"
   ClearFields
   Set rsTarget = Nothing
   Set rsSource = Nothing
endProc:
   Exit Function
errProc:
   Update_Entry = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'Update nQtyOnhnd CP_Inventory_Ledger
Private Function Update_QOH() As Boolean
Dim Quantity As Double
Dim QTyIn As Double
Dim QtyOut As Double
Dim rsLedger As New ADODB.Recordset

Update_QOH = True
On Error GoTo errProc

   Set rsTarget = New ADODB.Recordset
   lsSQL = "SELECT * From CP_Inventory " _
         & " ORDER BY sstockidx "
   If rsTarget.State = adStateOpen Then rsTarget.Close
   rsTarget.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   Do While Not rsTarget.EOF
       Set rsSource = New ADODB.Recordset
       lsSQL = "SELECT * From CP_Inventory_Ledger " _
            & "WHERE sBranchcd = '" & Text1.Text & "' " _
            & "AND sStockIDx = '" & rsTarget("sStockIDx") & "' " _
            & "Order by nEntryNox "
       If rsSource.State = adStateOpen Then rsSource.Close
       rsSource.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
       If rsSource.RecordCount <> 0 Then
           Quantity = rsSource("nQtyOnHnd")
           rsSource.MoveNext
           Do While Not rsSource.EOF
              lsSQL = "UPDATE CP_Inventory_Ledger SET " _
                       & " nQtyOnhnd = '" & Quantity & "' + nQtyInxxx - nQtyOutxx," _
                       & " dModified = '10/15/2007 9:30:00 AM' " _
                  & " WHERE sStockIDx = '" & rsSource("sStockIDx") & "'" _
                      & " AND sBranchCd = '" & rsSource("sBranchcd") & "'" _
                      & " AND nEntryNox = '" & rsSource("nEntryNox") & "'" _
                      & " AND sSourceCd <> 'CPAd' "
              oApp.Connection.Execute lsSQL, lnrow, adCmdText
              txtfield(0).Text = rsSource("sStockIDx")
              txtfield(1).Text = rsSource("sBranchCd")
              txtfield(2).Text = rsSource("dTransact")
              txtfield(3).Text = rsSource("sSourceNo")
              txtfield(4).Text = rsSource("nEntryNox")
              DoEvents
              Quantity = Quantity + rsSource("nQtyInxxx") - rsSource("nQtyOutxx")
              rsSource.MoveNext
           Loop
       End If
       rsTarget.MoveNext
   Loop
   MsgBox "Tapos na Po"
   ClearFields
   Set rsTarget = Nothing
   Set rsSource = Nothing
endProc:
   Exit Function
errProc:
   Update_QOH = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'Update CP_Serial_Master, Make Existing IMEI cLocation = 9
'Use Form CP_Serial_Status
'Private Sub xrButton7_Click()
'Dim lssql As String
'Dim lnrow As Long
'   'Update_Serial_Master of Existing IMEI
'   lssql = "UPDATE CP_Serial_Master SET" _
'         & " cLocation = 9 " _
'   & " WHERE sIMEINoxx = '" & txtothers(6).Text & "' " _
'         & " And sBranchCd = '" & oApp.BranchCode & "' "
'   oApp.Connection.Execute lssql, lnrow, adCmdText
'   If lnrow > 0 Then
'      MsgBox "Updated"
'   End If
'End Sub

'Update CP_Serial_Master of Not Existing IMEI, Add CP_Serial_Ledger Sales
Private Function Insert_Serial_Ledger() As Boolean

Insert_Serial_Ledger = True
On Error GoTo errProc

Set oRS = New ADODB.Recordset
lsSQL = "SELECT * From CP_Serial_Master " _
         & " WHERE sBranchcd = '" & Text1.Text & "' " _
         & " AND clocation = 1" _
         & "ORDER by sserialid "
If oRS.State = adStateOpen Then oRS.Close
oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

   Do While Not oRS.EOF
      'Get Last Entry No
      lnEntry = getIMEIEntry("'" & oRS("sSerialID") & "'")

      'CP_Serial_Ledger
      lsSQL = "INSERT INTO CP_Serial_Ledger" _
                  & "( sSerialID ," _
                  & "  sBranchcd ," _
                  & "  dTransact ," _
                  & "  nEntryNox ," _
                  & "  sSourceCd ," _
                  & "  sSourceNo ," _
                  & "  cSoldStat ," _
                  & "  cLocation ," _
                  & "  dModified) " _
                     & " VALUES " _
                        & "('" & oRS("sSerialID") & "', " _
                        & " '" & Text1.Text & "', " _
                        & "'9/1/2007 8:30:00 AM', " _
                        & " '" & lnEntry & "', " _
                        & " 'CPSl', " _
                        & " '" & Text1.Text & "' & 'Captured', " _
                        & " '1', " _
                        & " '2', " _
                        & " getdate())"
      oApp.Connection.Execute lsSQL, lnrow, adCmdText

      'Update Location, CP_Serial_Master
      lsSQL = "UPDATE CP_Serial_Master SET" _
            & " cSoldStat = '1', " _
            & " cLocation = '2', " _
            & " sClientID = '0907000001', " _
            & " sModified = '" & Encrypt(oApp.UserID) & "', " _
            & " dModified = getdate() " _
      & " WHERE sSerialID = '" & oRS("sSerialID") & "' "
      oApp.Connection.Execute lsSQL, lnrow, adCmdText

      txtfield(0).Text = oRS("sSerialID")
      txtfield(1).Text = oRS("sStockIdx")
      DoEvents
      oRS.MoveNext
   Loop
   MsgBox "Tapos na Po"
   ClearFields
   Set oRS = Nothing
endProc:
   Exit Function
errProc:
   Insert_Serial_Ledger = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'Update CP_Serial_Master, Make Existing IMEI back to cLocation = 1
Private Function Update_Serial_Master() As Boolean

Update_Serial_Master = True
On Error GoTo errProc

Set oRS = New ADODB.Recordset
lsSQL = "SELECT * From CP_Serial_Master " _
         & " WHERE sBranchcd = '" & Text1.Text & "' " _
         & " AND clocation = 9 " _
         & "ORDER by sserialid "
If oRS.State = adStateOpen Then oRS.Close
oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

   Do While Not oRS.EOF
      'Update_Serial_Master of Existing IMEI
      lsSQL = "UPDATE CP_Serial_Master SET" _
            & " cLocation = 1 " _
      & " WHERE sSerialID = '" & oRS("sSerialID") & "' "
      oApp.Connection.Execute lsSQL, lnrow, adCmdText
      
      txtfield(0).Text = oRS("sSerialID")
      txtfield(1).Text = oRS("sStockIdx")
      DoEvents
      oRS.MoveNext
   Loop
   MsgBox "Tapos na Po"
   ClearFields
   Set oRS = Nothing
endProc:
   Exit Function
errProc:
   Update_Serial_Master = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'Update CP_Inventory_Master nQtyOnHnd, Add CP_Inventory_Ledger sSourceCd = text1.text & 'Captured'
'Based On Existing IMEI
Private Function Insert_Inventory_Ledger() As Boolean

Insert_Inventory_Ledger = True
On Error GoTo errProc

Set oRS = New ADODB.Recordset
lsSQL = "SELECT a.sstockidx, a.nbegqtyxx, a.nqtyonhnd" _
      & " From CP_Inventory_Master a " _
         & " LEFT JOIN CP_Inventory b " _
      & " ON a.sstockidx = b.sstockidx " _
      & " Where a.sbranchcd = '" & Text1.Text & "' " _
         & " AND a.nqtyonhnd <> 0 " _
         & " AND b.sCategIDx = '01001' " _
      & " Order by a.sstockidx "
If oRS.State = adStateOpen Then oRS.Close
oRS.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

   Do While Not oRS.EOF
      Set lrs = New ADODB.Recordset
      lsSQL = "SELECT * " _
            & " From CP_Serial_Master " _
            & " Where sbranchcd = '" & oRS("sBranchcd") & "' " _
            & " And sstockidx = '" & oRS("sStockIDx") & "' " _
            & " AND cLocation = 1 "
      If lrs.State = adStateOpen Then lrs.Close
      lrs.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

      Set rsBranch = New Recordset
      lsSQL = "SELECT TOP 1 nEntryNox, nQtyOnHnd " _
         & " FROM CP_Inventory_Ledger " _
         & " WHERE sStockIDx = '" & oRS("sStockIDx") & "'" _
         & " AND sBranchCd = '" & oRS("sBranchCd") & "'" _
         & " ORDER by nEntryNox Desc "
      If rsBranch.State = adStateOpen Then rsBranch.Close
      rsBranch.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

      Set rsMain = New Recordset
      lsSQL = "SELECT TOP 1 nEntryNox, nQtyOnHnd " _
         & " FROM CP_Inventory_Ledger " _
         & " WHERE sStockIDx = '" & oRS("sStockIDx") & "'" _
         & " AND sBranchCd = '" & oRS("sBranchcd") & "'" _
         & " AND sSourceNo = '" & Text1.Text & "' & 'Captured' " _
         & " ORDER by nEntryNox Desc "
      If rsMain.State = adStateOpen Then rsBranch.Close
      rsMain.Open lsSQL, oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText

      If lrs.RecordCount = 0 Then
         'Update QOH CP_Inventory_Master
         lsSQL = "UPDATE CP_Inventory_Master SET" _
               & " nQtyOnHnd = '0', " _
               & " sModified = '" & Encrypt(oApp.UserID) & "', " _
               & " dModified = getdate() " _
         & " WHERE sStockIDx = '" & oRS("sStockIDx") & "' " _
               & " And sBranchCd = '" & oRS("sBranchCd") & "' "
         oApp.Connection.Execute lsSQL, lnrow, adCmdText

         'Insert CP_Inentory_Ledger
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
               & "('" & oRS("sStockIDx") & "', " _
               & "'" & oRS("sBranchCd") & " ', " _
               & "'" & oRS("sBranchCd") & " ', " _
               & "'CPSl' , " _
               & "'" & Text1.Text & "' & 'Captured', " _
               & "'0', " _
               & "'" & rsBranch("nQtyOnHnd") & "', " _
               & "'0', " _
               & "'" & rsBranch("nEntryNox") + 1 & "', " _
               & "'9/1/2007 8:30:00 AM', " _
               & " getdate())"
         oApp.Connection.Execute lsSQL, lnrow, adCmdText
      ElseIf rsMain.RecordCount = 0 Then
         'Update QOH CP_Inventory_Master
         lsSQL = "UPDATE CP_Inventory_Master SET" _
               & " nQtyOnHnd = '" & lrs.RecordCount & "', " _
               & " sModified = '" & Encrypt(oApp.UserID) & "', " _
               & " dModified = getdate() " _
         & " WHERE sStockIDx = '" & oRS("sStockIDx") & "' " _
               & " And sBranchCd = '" & Text1.Text & "' "
         oApp.Connection.Execute lsSQL, lnrow, adCmdText

         If rsBranch("nQtyOnHnd") <> lrs.RecordCount Then
            'Insert CP_Inentory_Ledger
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
                  & "('" & oRS("sStockIDx") & "', " _
                  & "'" & Text1.Text & " ', " _
                  & "'" & Text1.Text & " ', " _
                  & "'CPSl' , " _
                  & "'" & Text1.Text & "' & 'Captured', " _
                  & "'0', " _
                  & "'" & rsBranch("nQtyOnHnd") - lrs.RecordCount & "', " _
                  & "'" & lrs.RecordCount & "', " _
                  & "'" & rsBranch("nEntryNox") + 1 & "', " _
                  & "'9/1/2007 8:30:00 AM', " _
                  & " getdate())"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
         End If
      Else
         MsgBox "Already Captured"
      End If
      txtfield(0).Text = oRS("sStockIdx")
      DoEvents
      oRS.MoveNext
   Loop
   Set lrs = Nothing
   Set oRS = Nothing
   Set rsBranch = Nothing
   Set rsMain = Nothing
endProc:
   Exit Function
errProc:
   Insert_Inventory_Ledger = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

'Check nEntryNox CP_Inventory_Ledger
Private Function Check_Entry() As Boolean

Check_Entry = True
On Error GoTo errProc
    
   Set rsSource = New ADODB.Recordset
   lsSQL = "SELECT a.sstockidx, a.nqtyonhnd, b.sbarrcode " _
        & "From CP_Inventory_Master a " _
         & " LEFT JOIN CP_Inventory b " _
         & " ON a.sstockidx = b.sstockidx " _
        & "WHERE a.sBranchcd = '" & Text1.Text & "' " _
        & "Order By a.sStockIDx "
   If rsSource.State = adStateOpen Then rsSource.Close
   rsSource.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

    Do While Not rsSource.EOF
      Set rsTarget = New ADODB.Recordset
          lsSQL = "SELECT TOP 1 sstockidx, nQtyOnHnd " _
               & " From CP_Inventory_Ledger " _
               & " WHERE sBranchcd = '" & Text1.Text & "' " _
               & " AND sStockIDx = '" & rsSource("sStockIdx") & "'" _
               & " Order by nEntryNox desc "
      If rsTarget.State = adStateOpen Then rsTarget.Close
      rsTarget.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

         If rsTarget.RecordCount <> 0 Then
            If rsTarget("nQtyOnHnd") <> rsSource("nQtyOnhnd") Then
               MsgBox rsSource("sStockIDx") & " " & rsSource("sbarrcode")
            End If
         End If
         txtfield(0).Text = rsSource("sStockIDx")
         DoEvents
         rsSource.MoveNext
   Loop
'   Set rsSource = New ADODB.Recordset
'   lsSQL = "SELECT * " _
'        & "From CP_Inventory_Ledger " _
'        & "WHERE (sBranchcd = '" & Text1.Text & "' " _
'            & " or sLocation = '" & Text1.Text & "')" _
'            & " AND sSourceCd = 'CPDv'" _
'        & "Order By sStockIDx,sSourceNo "
'   If rsSource.State = adStateOpen Then rsSource.Close
'   rsSource.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'    Do While Not rsSource.EOF
'      Set rsTarget = New ADODB.Recordset
'          lsSQL = "SELECT * " _
'               & " From CP_Inventory_Ledger " _
'               & " WHERE sStockIDx = '" & rsSource("sStockIdx") & "'" _
'               & " AND sSourceNo = '" & rsSource("sSourceNo") & "'" _
'               & " AND sSourceCd = 'CPDl'" _
'               & " AND sBranchCd = '" & rsSource("sLocation") & "'" _
'               & " AND sLocation = '" & rsSource("sBranchcd") & "'"
'      If rsTarget.State = adStateOpen Then rsTarget.Close
'      rsTarget.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
'         If rsTarget.RecordCount <> 0 Then
''            MsgBox rsTarget("sstockidx"), , rsSource("sstockidx")
''            MsgBox rsTarget("ssourceno"), , rsSource("ssourceno")
''            MsgBox rsTarget("ssourcecd"), , rsSource("ssourcecd")
'            If rsTarget("nQtyInxxx") <> rsSource("nQtyOutxx") Then
'               MsgBox rsSource("sStockIDx") & " " & rsSource("sSourceNo")
'            End If
'         End If
'         txtfield(0).Text = rsSource("sStockIDx")
'         DoEvents
'         rsSource.MoveNext
'   Loop
   MsgBox "Tapos na Po"
   ClearFields
   Set rsTarget = Nothing
   Set rsSource = Nothing
endProc:
   Exit Function
errProc:
   Check_Entry = False
   MsgBox Err.Description, vbCritical, "Warning"
End Function

Private Sub xrButton_Click(Index As Integer)
   If Text1.Text = "" And Index <> 11 Then
      MsgBox "Invalid Branch"
      Text1.SetFocus
      Exit Sub
   End If
   
   Select Case Index
      '***** First Check if CPDV or CPDl existing *****
      '********** Insert Branch Ledger Only ***********
      Case 0   'Insert CPDl
         Insert_CPDl
         
      Case 1   'Insert CPDv
         Insert_CPDv
         
      Case 2   'Insert CP_Inventory_Master
         Insert_Inventory_Master
         
      Case 3   'Insert CPAd
         Insert_Initial_Inventory
         
      Case 4   'Update Entry No
         Update_Entry
         
      Case 5   'Update QOH of Inventory Ledger
         Update_QOH
         
      Case 6   'Use frmCP_Serial_Status
      
      'Insert Branch Name as Customer for Previous Sales
      Case 7   'Insert Serial Ledger of Previous Sales
         Insert_Serial_Ledger
         
      Case 8   'Update Existing IMEI from 9 to 1
         Update_Serial_Master
         
      'Update CP_Inventory_Master nQtyOnHnd, Add CP_Inventory_Ledger sSourceCd = text1.text & 'Captured'
      'Based On Existing IMEI
      Case 9   'Insert Inventory Ledger of Previous Sales, Update QOH of Units Only
         Insert_Inventory_Ledger
      Case 10  'Update Accessories and Cards
         frmAccessories.Branch = Text1.Text
         frmAccessories.Show
      Case 11
         Unload Me
         Unload frmAccessories
      Case 12
         Check_Entry
   End Select
End Sub


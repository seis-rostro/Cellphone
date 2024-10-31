VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmImport_Branch 
   BorderStyle     =   0  'None
   Caption         =   "Import Utility"
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "wt0;wb0"
      Top             =   2385
      Width           =   2415
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   4365
      TabIndex        =   2
      Top             =   1965
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Import"
      AccessKey       =   "I"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmImport_Branch.frx":0000
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   5130
      TabIndex        =   3
      Top             =   1965
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmImport_Branch.frx":077A
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   600
      Left            =   1770
      Tag             =   "wt0;fb0"
      Top             =   1140
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   1058
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1515
         TabIndex        =   4
         Top             =   120
         Width           =   2430
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Drive Source"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   150
         Width           =   1260
      End
   End
   Begin xrControl.xrFrame xrFrame2 
      Height          =   570
      Left            =   1770
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   1005
      Begin VB.TextBox txtfield 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1515
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   105
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No."
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   135
         Width           =   1260
      End
   End
   Begin MSComCtl2.Animation Progress 
      Height          =   720
      Left            =   105
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "wt0;wb0"
      Top             =   1965
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1270
      _Version        =   393216
      BackColor       =   1842204
      FullWidth       =   61
      FullHeight      =   48
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Very Accurate..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   1
      Top             =   795
      Width           =   1395
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Be "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   16
      Left            =   135
      TabIndex        =   0
      Top             =   585
      Width           =   930
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   90
      Picture         =   "frmImport_Branch.frx":0EF4
      Stretch         =   -1  'True
      Top             =   540
      Width           =   1635
   End
End
Attribute VB_Name = "frmImport_Branch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Programmed By  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Rosalyn Lazo Descallar  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Started  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  July 05, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'

Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Private poFileSys As FileSystemObject

Private pnCtr As Integer
Dim oRS As New ADODB.Recordset
Dim lrs As New ADODB.Recordset

Private Sub cmdButton_Click(Index As Integer)
   Select Case Index
      Case 0
         If txtfield(0).Text <> "" Then
            Import_Data
         Else
            MsgBox "Invalid Reference No.!!!", vbCritical, "Warning"
         End If
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   If bLoaded = False Then
      txtfield(0).Text = ""
      bLoaded = True
   End If
End Sub

Private Sub Form_Load()
Dim Table As String

   CenterChildForm mdiMain, Me
   bLoaded = False

   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me

   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin

End Sub

Private Function Import_Data() As Boolean
Dim oFileObject As New FileSystemObject
Dim oFolder As Folder
Dim oFiles As Files
Dim oFile As File
Dim Reference As String
Dim rsSource As ADODB.Recordset
Dim rsTarget As ADODB.Recordset
Dim rsNew As ADODB.Recordset
Dim lnrow As Long
Dim Load As Long
Dim lsSQL As String
Dim Branch As String
Dim Deleted As String
Dim lnCtr As Integer
Dim QOH As Integer
Dim ctr As Integer

Import_Data = True
oApp.Connection.BeginTrans
On Error GoTo errProc
   
   Set poFileSys = New FileSystemObject
   Set oFolder = oFileObject.GetFolder(Drive1 & "\")
   Set oFiles = oFolder.Files

   If Not poFileSys.DriveExists(Drive1) Then
      MsgBox "Drive Does not Exist!!!" & vbCrLf & _
            "Please Insert Mobile Disk then Try again.", vbCritical, "Notice"
      Exit Function
   End If
   
   Progress.Open App.Path & "\images\FILECOPY.AVI"
   Progress.Play
   
   Deleted = ""
   For Each oFile In oFiles
      
      If Trim(Right((Left(oFile, 18)), 15)) = Trim(txtfield(0).Text) Then
         Set rsSource = New ADODB.Recordset
         rsSource.Open "" & oFile & ""

         Reference = Right(oFile, Len(oFile) - 19)
         Set rsNew = New ADODB.Recordset
         rsNew.CursorLocation = adUseClient
         rsNew.Open Reference, _
         oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdTable
         
         Do While Not rsSource.EOF
            Set rsTarget = New ADODB.Recordset
               lsSQL = "SELECT * " _
                     & " FROM " & Reference & ""

            Select Case Reference
               Case "CP_Serial_Transfer_Master", "CP_Transfer_Master"
                  lsSQL = lsSQL & " WHERE sTransNox = '" & rsSource("sTransNox") & "'"
                  rsTarget.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
                  If rsTarget.RecordCount = 0 Then
                     rsNew.AddNew
                     For lnCtr = 0 To rsSource.Fields.Count - 2
                        rsNew(lnCtr) = rsSource(lnCtr)
                     Next
                     rsNew.MoveNext
                  ElseIf DateDiff("s", rsTarget("dModified"), rsSource("dModified")) > 0 _
                        And rsTarget("cTranStat") = 0 Then
                        
                     If rsTarget("nEntryNox") <> rsSource("nEntryNox") Then
                        'Delete Detail
                        If Reference = "CP_Serial_Transfer_Master" Then
                           lsSQL = "DELETE CP_Serial_Transfer_Detail " _
                                 & " WHERE sTransNox = '" & rsSource("sTransNox") & "'"
                           oApp.Connection.Execute lsSQL, lnrow, adCmdText
                           If lnrow <> 0 Then oApp.RegisDelete lsSQL
                        ElseIf Reference = "CP_Transfer_Master" Then
                           lsSQL = "DELETE CP_Transfer_Detail " _
                                 & " WHERE sTransNox = '" & rsSource("sTransNox") & "'"
                           oApp.Connection.Execute lsSQL, lnrow, adCmdText
                           If lnrow <> 0 Then oApp.RegisDelete lsSQL
                        End If
                     End If
                        
                     For lnCtr = 1 To rsSource.Fields.Count - 2
                        rsTarget(lnCtr) = rsSource(lnCtr)
                     Next
                     rsTarget.Update
                  End If
               Case "Banks", "Branch", "Card", "Color", "Credit_Card", _
                     "CP_Inventory", "Model", "CP_Serial_Transfer_Detail", _
                     "CP_Transfer_Detail", "CP_Serial_Ledger", "CP_Serial_Master", _
                     "Client_Master", "Sales_Person", "Brand", "Supplier", "ELoad_Matrix", _
                     "xxxSysUser", "Category", "Category_Master"
                     
                  Select Case Reference
                     Case "Banks"
                        lsSQL = lsSQL & " WHERE sBankIDxx = '" & rsSource("sBankIDxx") & "'"
                     Case "Branch"
                        lsSQL = lsSQL & " WHERE sBranchCd = '" & rsSource("sBranchCd") & "'"
                     Case "Card"
                        lsSQL = lsSQL & " WHERE sCardIDxx = '" & rsSource("sCardIDxx") & "'"
                     Case "Color"
                        lsSQL = lsSQL & " WHERE sColorIDx = '" & rsSource("sColorIDx") & "'"
                     Case "Credit_Card"
                        lsSQL = lsSQL & " WHERE sCreditID = '" & rsSource("sCreditID") & "'"
                     Case "CP_Inventory"
                        lsSQL = lsSQL & " WHERE sStockIDx = '" & rsSource("sStockIDx") & "'"
                     Case "Model"
                        lsSQL = lsSQL & " WHERE sModelIDx = '" & rsSource("sModelIDx") & "'"
                     Case "CP_Serial_Transfer_Detail", "CP_Transfer_Detail"
                        lsSQL = lsSQL & " WHERE sTransNox = '" & rsSource("sTransNox") & "'" _
                                       & " AND nEntryNox = '" & rsSource("nEntryNox") & "'"
                     Case "CP_Serial_Ledger"
                        lsSQL = lsSQL & " WHERE nEntryNox = '" & rsSource("nEntryNox") & "'" _
                                 & " AND sSerialID = '" & rsSource("sSerialID") & "'"
                     Case "CP_Serial_Master"
                        lsSQL = lsSQL & " WHERE sSerialID = '" & rsSource("sSerialID") & "'"
                     Case "Client_Master"
                        lsSQL = lsSQL & " WHERE sClientID = '" & rsSource("sClientID") & "'"
                     Case "Sales_Person"
                        lsSQL = lsSQL & " WHERE sEmployID = '" & rsSource("sEmployID") & "'"
                     Case "Brand"
                        lsSQL = lsSQL & " WHERE sBrandIDx = '" & rsSource("sBrandIDx") & "'"
                     Case "Supplier"
                        lsSQL = lsSQL & " WHERE sSupplyID = '" & rsSource("sSupplyID") & "'"
                     Case "ELoad_Matrix"
                        lsSQL = lsSQL & " WHERE sMatrixID = '" & rsSource("sMatrixID") & "'"
                     Case "xxxSysUser"
                        lsSQL = lsSQL & " WHERE sUserIDxx = '" & rsSource("sUserIDxx") & "'"
                     Case "Category"
                        lsSQL = lsSQL & " WHERE sCategIDx = '" & rsSource("sCategIDx") & "'"
                     Case "Category_Master"
                        lsSQL = lsSQL & " WHERE sCategory = '" & rsSource("sCategory") & "'"
                  End Select

                  rsTarget.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
                   If rsTarget.RecordCount = 0 Then
                     rsNew.AddNew
                     For lnCtr = 0 To rsSource.Fields.Count - 2
                        rsNew(lnCtr) = rsSource(lnCtr)
                     Next
                     rsNew.MoveNext
                  ElseIf DateDiff("s", rsTarget("dModified"), rsSource("dModified")) > 0 Then
                     For lnCtr = 1 To rsSource.Fields.Count - 2
                        rsTarget(lnCtr) = rsSource(lnCtr)
                     Next
                     rsTarget.Update
                  End If
                  
            End Select

            If Reference = "CP_Inventory_Ledger" Then
               Text1.Text = Reference & " " & rsSource("sstockidx")
               DoEvents
            Else
               Text1.Text = Reference
               DoEvents
            End If

            rsSource.MoveNext
         Loop
         rsSource.Close
         rsNew.Close
         Set rsTarget = Nothing
      End If
   Next

   Progress.Close
   Progress.Stop
   Text1.Text = ""
   MsgBox "Data Import Successfully Completed!!!", vbInformation, "Information"

endProc:
   oApp.Connection.CommitTrans
   Exit Function
errProc:
   oApp.Connection.RollbackTrans
   Import_Data = False
   MsgBox Err.Description, vbCritical, "Warning"

End Function

Private Sub oDriver_InitValue()
   Text1.Text = ""
End Sub

''¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Tested  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
''¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  July xx, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
'
''¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤    Version 1    ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
''¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  Date Finished  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'
''¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤  July 08, 2007  ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤'


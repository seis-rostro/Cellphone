VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmServicePhonePosting 
   BorderStyle     =   0  'None
   Caption         =   "Service Phone Posting"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3000
      Index           =   2
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1950
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   5292
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3000
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   5292
         _Version        =   393216
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   900
      Index           =   0
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1020
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1588
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1800
         TabIndex        =   11
         Text            =   "mmmm dd yyyy"
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   6000
         TabIndex        =   9
         Top             =   90
         Width           =   3375
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   88
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "C00121-000001"
         Top             =   -960
         Width           =   2775
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   4
         Left            =   6000
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   2775
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1410
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source/Origin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   4800
         TabIndex        =   10
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   90
         TabIndex        =   5
         Top             =   120
         Width           =   1485
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   4800
         TabIndex        =   4
         Top             =   480
         Width           =   765
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   9780
      TabIndex        =   6
      Top             =   1740
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmServicePhonePosting.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   9780
      TabIndex        =   7
      Top             =   480
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmServicePhonePosting.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   9780
      TabIndex        =   8
      Top             =   1110
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
      Picture         =   "frmServicePhonePosting.frx":0EF4
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   480
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   510
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   847
      Begin VB.TextBox txtField 
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
         Height          =   315
         Index           =   5
         Left            =   4180
         MaxLength       =   50
         TabIndex        =   14
         Top             =   90
         Width           =   2370
      End
      Begin VB.TextBox txtField 
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
         Height          =   315
         Index           =   0
         Left            =   1245
         TabIndex        =   13
         Top             =   90
         Width           =   1950
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "unknown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   6840
         TabIndex        =   17
         Tag             =   "eb0;et0"
         Top             =   90
         Width           =   2505
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   285
         Index           =   0
         Left            =   6880
         Tag             =   "et0;et0"
         Top             =   90
         Width           =   2400
      End
      Begin VB.Shape Shape3 
         Height          =   375
         Index           =   0
         Left            =   6840
         Top             =   45
         Width           =   2490
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Branch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   3445
         TabIndex        =   16
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Transact. No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   30
         TabIndex        =   15
         Top             =   120
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmServicePhonePosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmServicePhonePosting"

Private WithEvents oTrans As clsCPServicePhoneTransfer
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pnIndex As Integer
Dim pbGridFocus As Boolean
Dim pnCtr As Integer
Dim pbLoad As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsRep As String
   
   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc
   
   txtField_LostFocus pnIndex
   
   Select Case Index
   Case 0
      If oTrans.SearchAcceptance Then
         LoadMaster
         LoadDetail
         pbLoad = True
     Else
         pbLoad = False
         If txtField(0).Text <> "" Then pbLoad = True
      End If
         'txtField(0).SetFocus
   Case 1
      If pbLoad Then
         
         lsRep = MsgBox("Post Transaction!!!", vbYesNo + vbQuestion, "Confrim")
         Call txtField_Validate(5, True)
               
         If lsRep = vbYes Then
            Select Case Label2
               Case "UNKNOWN"
                  MsgBox "Unable to Post Transaction!!!", vbCritical, "Warning"
               Case "OPEN"
                  MsgBox "Unable to Post OPEN Transaction!!!", vbCritical, "Warning"
               Case "CLOSED"
                  If Not oTrans.AcceptDelivery(CDate(txtField(2).Text)) Then
                     MsgBox "Unable to Post Transaction!!!", vbCritical, "Warning"
                  Else
                     MsgBox "Transaction Post Successfully!!!", vbInformation, "Notice"
                     ClearFields
                  End If
             End Select
         End If
      Else
         MsgBox "Unable to Post Transaction!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   Case 2
      Unload Me
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   MSFlexGrid1.Refresh

   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"

   ''On Error GoTo errProc
   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPServicePhoneTransfer
   Set oTrans.AppDriver = oApp
   
   oTrans.TransStatus = 10
   oTrans.Destination = oApp.BranchCode
   oTrans.InitTransaction


   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
  
   InitGrid
   ClearFields
   xrFrame1(0).Enabled = False
   
'endProc:
 '  Exit Sub
'errProc:
 '  ShowError lsOldProc & "( " & " )", True

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub MSFlexGrid1_GotFocus()
   pbGridFocus = True
End Sub

Private Sub oTrans_DetailRetrieved(ByVal Index As Integer)
   With MSFlexGrid1
      .TextMatrix(.Row, Index) = oTrans.Detail(.Row - 1, Index)
   End With
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   'txtSearch(0).Text = oTrans.Master()
   'txtSearch(0).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@-@@@@@@")
End Sub

Private Sub ClearFields()
 Dim loTxt As TextBox

   For Each loTxt In txtField
      loTxt = ""
   Next
   Label2.Caption = "UNKNOWN"
   
   With MSFlexGrid1
      .Rows = 2
      .ColWidth(3) = 3100
      
      'empty row
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
   End With
End Sub
Private Sub InitGrid()
      Dim lnCtr As Integer

   With MSFlexGrid1
      .Cols = 5
      .Rows = 2
      .Clear

      'pnActiveRow = 0
      .Row = 0
      .TextMatrix(0, 0) = "No."
      .TextMatrix(0, 1) = "BRAND"
      .TextMatrix(0, 2) = "MODEL"
      .TextMatrix(0, 3) = "IMEI"
      .TextMatrix(0, 4) = "QTY."
      
      .Row = 0
      'Column Width
      .ColWidth(0) = 600
      .ColWidth(1) = 2480
      .ColWidth(2) = 2480
      .ColWidth(3) = 3200
      .ColWidth(4) = 700

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next
      .Row = 1
      'Column Alignment
      .TextMatrix(1, 0) = 1

      .ColAlignment(0) = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignLeftCenter

      .Col = 0
      .ColSel = .Cols - 1
      
   End With
End Sub

Private Sub LoadMaster()
   For pnCtr = 0 To txtField.Count - 1
      Select Case pnCtr
      Case 1, 0
         txtField(pnCtr).Text = Format(oTrans.Master("sTransNox"), "@@@@-@@-@@@@@@")
         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 2
         txtField(pnCtr).Text = Format(oTrans.Master("dTransact"), "MMMM DD, YYYY")
      Case 3, 5
         txtField(pnCtr).Text = oTrans.Master(7)

         txtField(pnCtr).Tag = txtField(pnCtr).Text
      Case 4
         txtField(pnCtr).Text = oTrans.Master("sRemarksx")
      'Case Else
         'txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next
   
   Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), "1")
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Rows = IIf(oTrans.ItemCount = 0, 2, oTrans.ItemCount + 1)
      
      .ColWidth(3) = 3100
      If .Rows > 20 Then .ColWidth(3) = 2850
            
      For pnCtr = 0 To oTrans.ItemCount - 1
         For lnCtr = 1 To 4
            .TextMatrix(pnCtr + 1, lnCtr) = oTrans.Detail(pnCtr, lnCtr)
         Next
         .TextMatrix(pnCtr + 1, 1) = oTrans.Detail(pnCtr, "sBrandNme")
         .TextMatrix(pnCtr + 1, 2) = oTrans.Detail(pnCtr, "sModelNme")
         .TextMatrix(pnCtr + 1, 3) = oTrans.Detail(pnCtr, "sSerialNo")
         .TextMatrix(pnCtr + 1, 4) = oTrans.Detail(pnCtr, "nQuantity")
      Next
   End With
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      .Text = TitleCase(.Text)
      Select Case Index
      Case 5, 0
         If .Text = "" Then
            ClearFields
            Exit Sub
         End If
                           
         If Trim(LCase(.Text)) <> Trim(LCase(.Tag)) Then
            If oTrans.SearchAcceptance _
            (IIf(Index = 5, CodeFormat(oApp.BranchCode, .Text), .Text) _
            , IIf(Index = 5, True, False)) Then
               LoadMaster
               LoadDetail
            Else
               ClearFields
               .SetFocus
            End If
        End If
      Case 7
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
      End Select
   End With
   
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

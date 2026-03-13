VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPWholeSaleReturnReg 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "CP Whole Sale"
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   525
      Index           =   1
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   926
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
         Index           =   80
         Left            =   1470
         TabIndex        =   1
         Top             =   90
         Width           =   2130
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
         Index           =   81
         Left            =   5340
         MaxLength       =   50
         TabIndex        =   0
         Top             =   90
         Width           =   4830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact. No"
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
         Index           =   9
         Left            =   165
         TabIndex        =   3
         Top             =   150
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
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
         Index           =   0
         Left            =   3840
         TabIndex        =   2
         Top             =   150
         Width           =   1320
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   8
      Left            =   10860
      TabIndex        =   4
      Top             =   1800
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Print"
      AccessKey       =   "Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPWholeSaleReturnReg.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   10860
      TabIndex        =   5
      Top             =   2430
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
      Picture         =   "frmCPWholeSaleReturnReg.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   10860
      TabIndex        =   6
      Top             =   540
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
      Picture         =   "frmCPWholeSaleReturnReg.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   10860
      TabIndex        =   7
      Top             =   1170
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
      Picture         =   "frmCPWholeSaleReturnReg.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   10860
      TabIndex        =   8
      Top             =   540
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
      Picture         =   "frmCPWholeSaleReturnReg.frx":1DE8
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   6405
      Index           =   2
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   11298
      BackColor       =   12632256
      Enabled         =   0   'False
      BorderStyle     =   1
      Begin xrControl.xrFrame xrFrame1 
         Height          =   2385
         Index           =   0
         Left            =   75
         Tag             =   "wt0;fb0"
         Top             =   90
         Width           =   10350
         _ExtentX        =   18256
         _ExtentY        =   4207
         BackColor       =   12632256
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   540
            Index           =   4
            Left            =   1290
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   1365
            Width           =   5520
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   3
            Left            =   1290
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   1020
            Width           =   5520
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   2
            Left            =   1290
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   675
            Width           =   5520
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
            Height          =   315
            Index           =   0
            Left            =   1290
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   120
            Width           =   2265
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Index           =   1
            Left            =   7935
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   675
            Width           =   2265
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   5
            Left            =   1290
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   1920
            Width           =   8910
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   195
            Index           =   4
            Left            =   105
            TabIndex        =   23
            Top             =   1440
            Width           =   570
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name"
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   22
            Top             =   1095
            Width           =   1125
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company"
            Height          =   195
            Index           =   2
            Left            =   105
            TabIndex        =   21
            Top             =   750
            Width           =   660
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
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
            Index           =   0
            Left            =   105
            TabIndex        =   20
            Top             =   180
            Width           =   915
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   315
            Left            =   1380
            Tag             =   "et0;ht2"
            Top             =   225
            Width           =   2265
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Index           =   1
            Left            =   7410
            TabIndex        =   19
            Top             =   743
            Width           =   345
         End
         Begin VB.Label lblNetTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   7935
            TabIndex        =   18
            Tag             =   "et0;hb0"
            Top             =   1365
            Width           =   2265
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   8
            Left            =   7050
            TabIndex        =   17
            Top             =   1455
            Width           =   705
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            Height          =   195
            Index           =   12
            Left            =   105
            TabIndex        =   16
            Top             =   1988
            Width           =   630
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
            Height          =   240
            Left            =   7755
            TabIndex        =   15
            Tag             =   "eb0;et0"
            Top             =   135
            Width           =   2385
         End
         Begin VB.Shape Shape3 
            Height          =   390
            Index           =   0
            Left            =   7695
            Top             =   75
            Width           =   2505
         End
         Begin VB.Shape Shape4 
            Height          =   330
            Index           =   0
            Left            =   7725
            Top             =   105
            Width           =   2445
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   285
            Index           =   0
            Left            =   7755
            Tag             =   "et0;et0"
            Top             =   135
            Width           =   2400
         End
      End
      Begin xrControl.xrFrame xrFrame3 
         Height          =   900
         Left            =   75
         Tag             =   "wt0;fb0"
         Top             =   2490
         Width           =   10350
         _ExtentX        =   18256
         _ExtentY        =   1588
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtOthers 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   6735
            TabIndex        =   27
            Text            =   "0"
            Top             =   480
            Width           =   945
         End
         Begin VB.TextBox txtOthers 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   6735
            TabIndex        =   26
            Text            =   "0.00"
            Top             =   150
            Width           =   945
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1305
            TabIndex        =   25
            Top             =   450
            Width           =   3870
         End
         Begin VB.TextBox txtOthers 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1305
            TabIndex        =   24
            Top             =   120
            Width           =   2130
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity"
            Height          =   195
            Index           =   13
            Left            =   5910
            TabIndex        =   31
            Top             =   540
            Width           =   585
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "U. Price"
            Height          =   195
            Index           =   14
            Left            =   5925
            TabIndex        =   30
            Top             =   210
            Width           =   570
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   29
            Top             =   510
            Width           =   795
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bar Code"
            Height          =   195
            Index           =   17
            Left            =   90
            TabIndex        =   28
            Top             =   180
            Width           =   660
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2850
         Left            =   75
         TabIndex        =   32
         Top             =   3420
         Width           =   10350
         _ExtentX        =   18256
         _ExtentY        =   5027
         _Version        =   393216
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmCPWholeSaleReturnReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCPWholeSaleReturnReg"

Private WithEvents oTrans As clsCPWholeSaleReturn
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pbGridGotFocus As Boolean
Dim pnIndex As Integer
Dim pbSave As Boolean

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lsRep As String

   lsOldProc = "cmdButton_Click"
   ''On Error GoTo errProc

   With MSFlexGrid1
      Select Case Index
      Case 4
         If oTrans.SearchTransaction() = True Then
            LoadMaster
            LoadDetail
         End If
         txtField(81).SetFocus
      Case 6
         If oTrans.Master("cTranStat") < xeStatePosted Then
            If oTrans.CancelTransaction = True Then
               Label2.Caption = TransStat(oTrans.Master("cTranStat"))
               MsgBox "Transaction was cancelled successfuly.", vbInformation, "Notice"
            End If
         Else
            MsgBox "Unable to cancel Posted/Cancelled transaction.", vbInformation, "Notice"
         End If
      Case 7
         Unload Me
      Case 8
         If oTrans.Master("sTransNox") <> "" And oTrans.Master("cTranStat") <> xeStateCancelled Then
            Call PrintTrans
         End If
      End Select
   End With

endProc:
     Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPWholeSaleReturn
   Set oTrans.AppDriver = oApp

   oTrans.InitTransaction

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance

   ClearFields
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oTrans = Nothing
   Set oSkin = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         If GetFocus = txtOthers(3).hwnd Then
            txtOthers(1).SetFocus
         Else
            SetNextFocus
         End If
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = Not lbShow
   cmdButton(6).Visible = Not lbShow

   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow
   xrFrame1(2).Enabled = lbShow

   If Not lbShow Then cmdButton(4).SetFocus
End Sub

Private Sub ClearFields()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      Select Case loTxt.Index
      Case 6, 7, 9, 10
         loTxt.Text = "0.00"
      Case Else
         loTxt.Text = ""
      End Select
   Next
      
   txtField(80) = txtField(0)
   txtField(81) = txtField(2)
   Label2.Caption = TransStat(IFNull(oTrans.Master("cTranStat"), -1))
   
   txtOthers(1) = ""
   txtOthers(2) = ""
   txtOthers(3) = "0"
   txtOthers(4) = "0.00"
   
   lblNetTotal = "0.00"

   Call InitForm
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

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      .Col = 0
      .ColSel = .Cols - 1
      
      setFieldInfo
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If .Text <> Empty Then
         If Index = 80 Then
            .Text = Replace(.Text, "-", "")
         End If
      
         .SelStart = 0
         .SelLength = Len(.Text)
      End If
      
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
   ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         If KeyCode = vbKeyF3 Then
            Select Case Index
            Case 80, 81
               If oTrans.SearchTransaction(IIf(Index = 80, CodeFormat(oApp.BranchCode, .Text), .Text), IIf(Index = 80, True, False)) Then
                  LoadMaster
                  LoadDetail
               Else
                  ClearFields
               End If
               
               .SelStart = 0
               .SelLength = Len(.Text)
               .SetFocus
            End Select
         End If
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True

End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      If Index = 80 Then
         If Len(.Text) = 12 Then .Text = Format(.Text, "@@@@@@-@@@@@@")
      End If
   
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub InitForm()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Cols = 9
      .Rows = 2

      'Column Title
      .TextMatrix(0, 1) = "Barcode"
      .TextMatrix(0, 2) = "Description"
      .TextMatrix(0, 3) = "SN"
      .TextMatrix(0, 4) = "Model"
      .TextMatrix(0, 5) = "QH"
      .TextMatrix(0, 6) = "UPrice"
      .TextMatrix(0, 7) = "Qty"
      .TextMatrix(0, 8) = "Total"

      .Row = 0

      'Column Alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 3
      Next

      'Column Width
      .ColWidth(0) = 350
      .ColWidth(1) = 2350
      .ColWidth(2) = 3175
      .ColWidth(3) = 415
      .ColWidth(4) = 1100
      .ColWidth(5) = 500
      .ColWidth(6) = 850
      .ColWidth(7) = 550
      .ColWidth(8) = 1000

      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1

      .TextMatrix(1, 0) = "1"
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = "0"
      .TextMatrix(1, 6) = "0.00"
      .TextMatrix(1, 7) = "0"
      .TextMatrix(1, 8) = "0.00"

      .Col = 1
      .Row = 1
      
      .Col = 0
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub LoadDetail()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Rows = oTrans.ItemCount + 1
      
      For lnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, "xReferNox")
         .TextMatrix(lnCtr + 1, 2) = oTrans.Detail(lnCtr, "sDescript")
         .TextMatrix(lnCtr + 1, 3) = IIf(oTrans.Detail(lnCtr, "cHsSerial") = "0", "No", "Yes")
         .TextMatrix(lnCtr + 1, 4) = IFNull(oTrans.Detail(lnCtr, "sModelNme"), "N-O-N-E")
         .TextMatrix(lnCtr + 1, 5) = oTrans.Detail(lnCtr, "nQtyOnHnd")
         .TextMatrix(lnCtr + 1, 6) = Format(oTrans.Detail(lnCtr, "nUnitPrce"), "#,##0.00")
         .TextMatrix(lnCtr + 1, 7) = oTrans.Detail(lnCtr, "nQuantity")
         
         ComputeTotal lnCtr + 1
      Next
   
      .Row = .Rows - 1
      
      .Col = 0
      .ColSel = .Cols - 1
      
      setFieldInfo
   End With
End Sub

Private Sub setFieldInfo()
   With MSFlexGrid1
      txtOthers(1) = .TextMatrix(.Row, 1)
      txtOthers(2) = .TextMatrix(.Row, 2)
      txtOthers(3) = .TextMatrix(.Row, 7)
      txtOthers(4) = .TextMatrix(.Row, 6)
   End With
End Sub

Private Sub ComputeTotal(lnRow As Integer)
   With MSFlexGrid1
      .TextMatrix(lnRow, 8) = Format(Round(CDbl(.TextMatrix(lnRow, 7)) * CDbl(.TextMatrix(lnRow, 6)), 2), "#,##0.00")
   End With
   
   computeNet
End Sub

Private Sub computeNet()
   Dim lnCtr As Integer
   Dim lnTotal As Double

   With MSFlexGrid1
      lnTotal = 0#
      For lnCtr = 1 To .Rows - 1
         lnTotal = CDbl(lnTotal) + CDbl(IIf(.TextMatrix(lnCtr, 8) = "", 0, .TextMatrix(lnCtr, 8)))
      Next
   End With

   lblNetTotal.Caption = FormatNumber(lnTotal, 2)
End Sub

Private Sub LoadMaster()
   Dim loTxt As TextBox
   
   For Each loTxt In txtField
      Select Case loTxt.Index
      Case 0
         loTxt.Text = Format(oTrans.Master(loTxt.Index), IIf(Len(oApp.BranchCode) = 2, "@@@@-@@@@@@", "@@@@@@-@@@@@@"))
      Case 1
         loTxt.Text = Format(oApp.ServerDate, "Mmm dd, yyyy")
      Case 3, 5, 6, 7
         loTxt.Text = Format(oTrans.Master(loTxt.Index), "#,##0.00")
      Case 80, 81
      Case Else
         loTxt.Text = IIf(IsNull(oTrans.Master(loTxt.Index)), "", oTrans.Master(loTxt.Index))
      End Select
   Next
   
   txtField(80) = txtField(0)
   txtField(81) = txtField(2)
   
   Label2.Caption = TransStat(oTrans.Master("cTranStat"))
   txtField(81).SetFocus
End Sub

Function PrintTrans() As Boolean
   
   Dim lrs As New ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsIMEI As String
   Dim lsOldProc As String
   
   Dim lrsCOInv As Recordset
   Dim lsSQL As String
   Dim loModel As Recordset
   
   lsOldProc = "printTrans"
'   ''On Error GoTo errProc

   PrintTrans = False
   
   Set lrs = New ADODB.Recordset
   
   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "sField01", adVarChar, 50
   lrs.Fields.Append "sField02", adVarChar, 60
   lrs.Fields.Append "sField03", adVarChar, 30
   lrs.Fields.Append "nField02", adInteger, 10
   lrs.Fields.Append "nField03", adInteger, 10
   lrs.Fields.Append "lField01", adCurrency
   lrs.Fields.Append "lField02", adCurrency
   lrs.Fields.Append "lField03", adCurrency
   lrs.Fields.Append "lField04", adCurrency
   lrs.Open
   
   lsSQL = "SELECT a.sTransNox" & _
            ", b.sBarrCode" & _
            ", c.sModelNme" & _
            ", c.sModelCde" & _
            ", d.sColorNme" & _
            ", a.nUnitPrce" & _
            ", a.nQuantity" & _
            ", b.sDescript" & _
            ", b.cHsSerial" & _
            ", a.sSerialID" & _
            ", e.sSerialNo"

    lsSQL = lsSQL & _
         " FROM CP_WSO_Return_Detail a" & _
               " LEFT JOIN CP_Inventory_Serial e" & _
                  " ON a.sSerialID = e.sSerialID" & _
            ", CP_Inventory b" & _
               " LEFT JOIN Color d" & _
                  " ON b.sColorIDx = d.sColorIDx" & _
            ", CP_Model c" & _
            ", CP_WSO_Return_Master f" & _
         " WHERE a.sTransNox = " & strParm(oTrans.Master("sTransNox")) & _
            " AND a.sTransNox = f.sTransNox" & _
            " AND a.sStockIDx = b.sStockIDx" & _
            " AND b.sModelIDx = c.sModelIDx" & _
            " AND f.cTranStat <> 3" & _
         " ORDER BY a.nEntryNox"
      
      Set loModel = New Recordset
      loModel.Open lsSQL, oApp.Connection, , adCmdText
   Debug.Print lsSQL
   

   For lnCtr = 0 To loModel.RecordCount - 1
      lrs.AddNew
      lrs("sField01").Value = loModel("sBarrCode")
      lrs("sField02").Value = loModel("sModelNme") & " " & loModel("sColorNme")
      lrs("sField03").Value = IFNull(loModel("sSerialNo"), "")
      lrs("nField01").Value = loModel("nQuantity")
      lrs("lField01").Value = loModel("nUnitPrce")
      lrs("lField02").Value = loModel("nUnitPrce") * loModel("nQuantity")
      loModel.MoveNext
   Next
  
 ' assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CPWholeSaleReturn.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   With oTrans
      oReport.Sections("PH").ReportObjects("txtCustomerName").SetText txtField(2)
      oReport.Sections("RH").ReportObjects("txtDate").SetText Format(txtField(1), "MMM-DD-YYYY")
      oReport.Sections("PH").ReportObjects("txtToAddress").SetText Trim(txtField(3))
      oReport.Sections("RH").ReportObjects("txtRefNo").SetText txtField(0)
      oReport.Sections("RF").ReportObjects("txtNote").SetText oTrans.Master("sRemarksx")
   End With
   
   oReport.PrintOutEx False, 1
   lrs.Close
   PrintTrans = True

endProc:
   If oTrans.Master("cTranStat") = xeStateOpen Then
      oTrans.CloseTransaction oTrans.Master(0)
   End If
   Set oReport = Nothing
   Set lrs = Nothing
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )"
End Function

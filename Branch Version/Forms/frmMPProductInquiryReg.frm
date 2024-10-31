VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMPProductInquiryReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP Product Inquiry"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10516.58
   ScaleMode       =   0  'User
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2505
      Left            =   120
      TabIndex        =   23
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   7635
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1560
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   780
         Width           =   5910
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1560
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   315
         Width           =   2355
      End
      Begin VB.TextBox txtField 
         Height          =   690
         HideSelection   =   0   'False
         Index           =   2
         Left            =   1575
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1530
         Width           =   5880
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
         Height          =   240
         Index           =   11
         Left            =   135
         TabIndex        =   29
         Top             =   1545
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   1665
         Tag             =   "et0;ht2"
         Top             =   405
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   240
         Index           =   10
         Left            =   120
         TabIndex        =   26
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact No"
         Height          =   240
         Index           =   9
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   1215
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   8070
      TabIndex        =   0
      Top             =   2520
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Cl&ose"
      AccessKey       =   "o"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMPProductInquiryReg.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   8085
      TabIndex        =   1
      Top             =   1905
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
      Picture         =   "frmMPProductInquiryReg.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   8070
      TabIndex        =   2
      Top             =   3135
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
      Picture         =   "frmMPProductInquiryReg.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   8085
      TabIndex        =   3
      Top             =   660
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
      Picture         =   "frmMPProductInquiryReg.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   8100
      TabIndex        =   4
      Top             =   1290
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
      Picture         =   "frmMPProductInquiryReg.frx":1DE8
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4800
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   3135
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   8467
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   83
         Left            =   1590
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   2670
         Width           =   3000
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   645
         Index           =   7
         Left            =   1590
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   3825
         Width           =   5850
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   82
         Left            =   1575
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2340
         Width           =   3000
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   80
         Left            =   1590
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1215
         Width           =   5850
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1575
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   855
         Width           =   2355
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1590
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   150
         Width           =   2355
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   645
         Index           =   81
         Left            =   1590
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1560
         Width           =   5850
      End
      Begin VB.ComboBox cmbField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         ItemData        =   "frmMPProductInquiryReg.frx":2562
         Left            =   1605
         List            =   "frmMPProductInquiryReg.frx":256C
         TabIndex        =   6
         Top             =   3375
         Width           =   2370
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   1605
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3015
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   22
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   240
         Index           =   6
         Left            =   180
         TabIndex        =   21
         Top             =   3885
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Transact No"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   240
         Index           =   5
         Left            =   180
         TabIndex        =   19
         Top             =   2700
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   240
         Index           =   4
         Left            =   180
         TabIndex        =   18
         Top             =   2355
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   1245
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1695
         Tag             =   "et0;ht2"
         Top             =   270
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Top             =   1590
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Type"
         Height          =   240
         Index           =   7
         Left            =   195
         TabIndex        =   15
         Top             =   3420
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Date"
         Height          =   240
         Index           =   8
         Left            =   195
         TabIndex        =   14
         Top             =   3060
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMPProductInquiryReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oTrans As clsMPProductInquiry
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Dim pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)

'Unload Me

Dim lsOldProc As String
Dim lnRep As Integer

lsOldProc = "cmdButton_Click"
'On Error GoTo errProc

Select Case Index
Case 0   'browse
 If pnIndex = 0 Or pnIndex = 1 Then
            If pnIndex = 0 Then
               If oTrans.SearchTransaction(txtSearch(pnIndex).Text, False) Then
                  ClearFields
                  LoadMaster
               End If
            Else
               If oTrans.SearchTransaction(txtSearch(pnIndex).Text) Then
                 ClearFields
                  LoadMaster
               End If
            End If
            pnIndex = 3
         Else
            If oTrans.SearchTransaction("") Then
               ClearFields
               LoadMaster
            End If
         End If

 Case 1     'close
             Unload Me
 End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True



End Sub
Private Sub Form_Activate()

oApp.MenuName = Me.Tag
   Me.ZOrder 0

End Sub
Private Sub Form_Load()
 Dim lsOldProc As String

      lsOldProc = "Form_Load"
      'On Error GoTo errProc

   '    CenterChildForm mdiMain, Me

       Set oTrans = New clsMPProductInquiry
       Set oTrans.AppDriver = oApp

       oTrans.Branch = oApp.BranchCode
       oTrans.InitTransaction

 Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight


   ClearFields

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True


 End Sub

Private Sub ClearFields()
       Dim loTxt As TextBox

          For Each loTxt In txtField
            loTxt = ""
          Next
          txtSearch(0) = ""
          txtSearch(1) = ""


End Sub
Private Sub LoadMaster()
       Dim loTxt As TextBox

       For Each loTxt In txtField
          Select Case loTxt.Index
         Case 1, 3
                loTxt.Text = strLongDate(oTrans.Master(loTxt.Index))
             Case Else
                loTxt.Text = oTrans.Master(loTxt.Index)
          End Select
       Next

       txtSearch(0) = txtField(0)
      txtSearch(1) = txtField(2)

End Sub
Private Sub txtSearch_LostFocus(Index As Integer)
       With txtSearch(Index)
          .BackColor = oApp.getColor("EB")
       End With

       pnIndex = Index
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
       With txtSearch(Index)
          .BackColor = oApp.getColor("HT1")
          .SelStart = 0
          .SelLength = Len(.Text)
       End With

       pnIndex = Index
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
       If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
          Select Case Index
      Case 0
             If oTrans.SearchTransaction(txtSearch(Index).Text, False) Then
                ClearFields
                LoadMaster
             End If
          Case 1
             If oTrans.SearchTransaction(txtSearch(Index).Text) Then
                ClearFields
               LoadMaster
            End If
         End Select
      End If
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With oApp
      .xLogError Err.Number, Err.Description, "frmMPProductInquiry", lsProcName, Erl
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

VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCustomer 
   BorderStyle     =   0  'None
   Caption         =   "Client Information"
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame1 
      Height          =   3555
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6271
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   5
         Left            =   1620
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2295
         Width           =   5220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   11
         Left            =   4305
         MaxLength       =   30
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   3060
         Width           =   2535
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   645
         Index           =   4
         Left            =   1620
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "frmCustomer.frx":0000
         Top             =   1635
         Width           =   5220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   9
         Left            =   4305
         MaxLength       =   15
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2550
         Width           =   2535
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   1620
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   180
         Width           =   1950
      End
      Begin VB.OptionButton optGender 
         Caption         =   "Female"
         Height          =   210
         Index           =   1
         Left            =   2430
         TabIndex        =   19
         Tag             =   "wt0;fb0"
         Top             =   3120
         Width           =   840
      End
      Begin VB.OptionButton optGender 
         Caption         =   "Male"
         Height          =   210
         Index           =   0
         Left            =   1620
         TabIndex        =   18
         Tag             =   "wt0;fb0"
         Top             =   3120
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   10
         Left            =   4305
         MaxLength       =   15
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   2805
         Width           =   2535
      End
      Begin VB.ComboBox cmbField 
         Height          =   315
         ItemData        =   "frmCustomer.frx":0006
         Left            =   1620
         List            =   "frmCustomer.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2790
         Width           =   1620
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   6
         Left            =   1620
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   2550
         Width           =   1620
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   960
         Left            =   1620
         Tag             =   "wt0;fb0"
         Top             =   630
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   1693
         BackColor       =   12632256
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   3
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   600
            Width           =   3870
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   2
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   345
            Width           =   3870
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   1
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   90
            Width           =   3870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Middle Name"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   5
            Top             =   345
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   3
            Top             =   90
            Width           =   765
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Town/City"
         Height          =   195
         Index           =   9
         Left            =   585
         TabIndex        =   11
         Top             =   2295
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "e-Mail Add."
         Height          =   195
         Index           =   16
         Left            =   3435
         TabIndex        =   24
         Top             =   3090
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   195
         Index           =   12
         Left            =   225
         TabIndex        =   2
         Top             =   690
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Address"
         Height          =   195
         Index           =   11
         Left            =   225
         TabIndex        =   9
         Top             =   1635
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1665
         Tag             =   "et0;ht2"
         Top             =   210
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. No."
         Height          =   195
         Index           =   5
         Left            =   3435
         TabIndex        =   20
         Top             =   2565
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   0
         Top             =   195
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         Height          =   195
         Index           =   22
         Left            =   570
         TabIndex        =   17
         Top             =   3090
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No."
         Height          =   195
         Index           =   1
         Left            =   3435
         TabIndex        =   22
         Top             =   2820
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         Height          =   195
         Index           =   6
         Left            =   585
         TabIndex        =   13
         Top             =   2550
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status"
         Height          =   195
         Index           =   4
         Left            =   585
         TabIndex        =   15
         Top             =   2820
         Width           =   780
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   7410
      TabIndex        =   28
      Top             =   2055
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      SizeCW          =   1
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
      Picture         =   "frmCustomer.frx":003F
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   7410
      TabIndex        =   26
      Top             =   1215
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
      Picture         =   "frmCustomer.frx":07B9
      CaptionAlign    =   0
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7410
      TabIndex        =   27
      Top             =   1635
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      SizeCW          =   2
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
      Picture         =   "frmCustomer.frx":0F33
      CaptionAlign    =   0
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As FormSkin
Private bLoaded As Boolean
Dim pnindex As Integer

Dim psClient As String
Dim psForm As String

Property Let oForm(oForm As String)
   psForm = oForm
End Property

Property Let Client(Client As String)
   psClient = Client
End Property


Private Sub cmdButton_Click(Index As Integer)
Dim lnctr As Integer
   Select Case Index
   Case 0
      If psForm = "Register" Then
         frmPOS_Register.txtField(3) = txtField(1).Text & ", " & txtField(2).Text
         frmPOS_Register.ClientID = oDriver.FieldValue(0)
      ElseIf psForm = "" Then
         frmCP_POS.txtField(4) = txtField(1).Text & ", " & txtField(2).Text
         frmCP_POS.ClientID = oDriver.FieldValue(0)
      End If
      oDriver.RecordSave
   Case 1
      If psForm = "Register" Then
         frmPOS_Register.txtField(4) = ""
         frmPOS_Register.ClientID = ""
      ElseIf psForm = "" Then
         frmCP_POS.txtField(4) = ""
         frmCP_POS.ClientID = ""
      End If
      psForm = ""
      Unload Me
   Case 2
      If pnindex = 5 Then oDriver.RecordSearch txtField(5).Text
   End Select
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   If bLoaded = False Then
      oDriver.RecordNew
      oDriver.DisableTextbox 0
      bLoaded = True
   End If
End Sub

Private Sub Form_Deactivate()
   psForm = ""
End Sub

Private Sub Form_Load()
   CenterChildForm mdiMain, Me
   
   bLoaded = False
   
   Set oDriver = New FormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me
   
   Set oSkin = New FormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransDetail
      
   oDriver.RecQuery = "SELECT" _
                        & " sClientID, " _
                        & " sLastName, " _
                        & " sFrstName, " _
                        & " sMiddName, " _
                        & " sAddressx, " _
                        & " sTownIDxx, " _
                        & " dBirthDte, " _
                        & " cCvilStat, " _
                        & " cGenderCd, " _
                        & " sPhoneNox, " _
                        & " sMobileNo, " _
                        & " sEmailAdd, "
   oDriver.RecQuery = oDriver.RecQuery _
                        & " sCitizenx, " _
                        & " sBirthPlc, " _
                        & " sHouseNox, " _
                        & " sTaxIDNox, " _
                        & " sAddlInfo, " _
                        & " sCompnyNm, " _
                        & " sClientNo, " _
                        & " sSpouseID, " _
                        & " cRecdStat, " _
                        & " sModified, " _
                        & " dModified, " _
                        & " vTimeStmp " _
                     & " FROM Client_Master "
                        
   oDriver.InitRecForm
   
   oDriver.LookupQuery(5) = "SELECT" _
                           & "  a.sTownIDxx" _
                           & ", a.sTownName" _
                           & ", b.sProvName" _
                           & ", a.sZippCode" _
                        & " FROM TownCity a" _
                           & " LEFT JOIN Province b" _
                              & " ON a.sProvIDxx = b.sProvIDxx" _
                        & " WHERE a.cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY a.sTownName,b.sProvName"

   oDriver.LookupReference(5) = "a.sTownIDxx»a.sTownName»b.sProvName»a.sZippCode"
   oDriver.LookupColumn(5) = "sTownName»sProvName»sZippCode"
   oDriver.LookupTitle(5) = "Town»Province»ZippCode"
      
   oDriver.FieldFormat(6) = "MMM dd, yyyy"
   oDriver.FieldStart = 1
   
   ClearFields
End Sub
Private Sub ClearFields()
   optGender(0).Value = False
   optGender(1).Value = False
   
   cmbField.ListIndex = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
End Sub

Private Sub oDriver_InitValue()
   If oDriver.SetValue(0, getNextCode("Client_Master", "sClientID", True, oApp.Connection, True, oApp.BranchCode)) = False Then Exit Sub
   oDriver.FieldReference(0) = True
   oDriver.FieldValue(20) = xeRecStateActive
   txtField(1).Text = psClient
End Sub

Private Sub oDriver_LoadOtherData()
   If Not IsNull(oDriver.FieldValue(8)) Then
      If Trim(oDriver.FieldValue(8)) <> "" Then
         optGender(oDriver.FieldValue(8)).Value = True
      Else
         optGender(0).Value = False
         optGender(1).Value = False
      End If
   Else
      optGender(0).Value = False
      optGender(1).Value = False
   End If
      
   If Not IsNull(oDriver.FieldValue(7)) Then
      If Trim(oDriver.FieldValue(7)) <> "" Then
         cmbField.ListIndex = oDriver.FieldValue(7)
      Else
         cmbField.ListIndex = -1
      End If
   Else
      cmbField.ListIndex = -1
   End If

End Sub

Private Sub oDriver_SaveComplete()
   MsgBox "Customer Info Added!!!", vbInformation, "Information"
   Unload Me
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
Dim lnctr As Integer
   If optGender(0).Value Or optGender(1).Value Then
      oDriver.FieldValue(8) = IIf(optGender(0).Value, 0, 1)
   End If
   If cmbField.ListIndex > -1 Then oDriver.FieldValue(7) = cmbField.ListIndex
   
   For lnctr = 12 To 19
      oDriver.FieldValue(lnctr) = ""
   Next
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   oDriver.ColumnIndex = Index
   txtField(Index).BackColor = &HE1FEFF
   pnindex = Index
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Or KeyCode = vbKeyF3 Then
      If Index = 5 Then
         oDriver.RecordSearch txtField(Index).Text
         If txtField(Index).Text <> "" Then SetNextFocus
      End If
   End If
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   txtField(Index).BackColor = &H80000005
End Sub

Private Sub txtfield_Validate(Index As Integer, Cancel As Boolean)
   If Index = 6 Then
      If Not IsDate(txtField(Index).Text) Then
         txtField(Index).Text = Format(oApp.ServerDate, "MMM dd, yyyy")
      Else
         txtField(Index).Text = Format(txtField(Index).Text, "MMM dd, yyyy")
      End If
   End If
   txtField(Index).Text = TitleCase(txtField(Index).Text)
   Cancel = Not oDriver.ValidateField(Index)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub




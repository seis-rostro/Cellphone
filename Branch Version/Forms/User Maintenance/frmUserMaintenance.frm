VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmUserMaintenance 
   BorderStyle     =   0  'None
   Caption         =   "User Maintenance"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmUserMaintenance.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   5340
      TabIndex        =   9
      Top             =   3870
      Width           =   1740
   End
   Begin VB.TextBox txtConfirm 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5340
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   3525
      Width           =   3255
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   5340
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3180
      Width           =   3255
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   5340
      TabIndex        =   3
      Top             =   2835
      Width           =   3255
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   5340
      TabIndex        =   1
      Top             =   2490
      Width           =   1665
   End
   Begin VB.PictureBox xrFrame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2025
      Left            =   3825
      ScaleHeight     =   2025
      ScaleWidth      =   4770
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "wt0;fb0"
      Top             =   4245
      Width           =   4770
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Active"
         Height          =   285
         Left            =   1065
         TabIndex        =   18
         Tag             =   "wt0;fb0"
         Top             =   1575
         Width           =   2385
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1065
         TabIndex        =   13
         Top             =   525
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmUserMaintenance.frx":79C2
         Left            =   1065
         List            =   "frmUserMaintenance.frx":79C4
         TabIndex        =   11
         Top             =   180
         Width           =   1815
      End
      Begin VB.PictureBox xrFrame3 
         BackColor       =   &H00C0C0C0&
         Height          =   675
         Left            =   3135
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "wt0;fb0"
         Top             =   135
         Width           =   1380
         Begin VB.CheckBox Check3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Allow View"
            Height          =   315
            Left            =   150
            TabIndex        =   21
            Tag             =   "wt0;fb0"
            Top             =   315
            Width           =   1170
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Allow Lock"
            Height          =   330
            Left            =   150
            TabIndex        =   20
            Tag             =   "wt0;fb0"
            Top             =   15
            Width           =   1170
         End
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   7
         Left            =   1065
         TabIndex        =   15
         Top             =   870
         Width           =   3465
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   15
         Left            =   1065
         TabIndex        =   17
         Top             =   1215
         Width           =   3465
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Type"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   12
         Top             =   585
         Width           =   735
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Member of"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   14
         Top             =   930
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Skin"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   1252
         Width           =   885
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Level"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   5340
      TabIndex        =   39
      Top             =   2490
      Width           =   1665
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   7845
      TabIndex        =   28
      Top             =   6405
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmUserMaintenance.frx":79C6
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   7080
      TabIndex        =   27
      Top             =   6405
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmUserMaintenance.frx":8140
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   6315
      TabIndex        =   25
      Top             =   6405
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmUserMaintenance.frx":88BA
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   5550
      TabIndex        =   24
      Top             =   6405
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmUserMaintenance.frx":9034
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   4785
      TabIndex        =   23
      Top             =   6405
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmUserMaintenance.frx":97AE
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   7845
      TabIndex        =   29
      Top             =   6405
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
      Picture         =   "frmUserMaintenance.frx":9F28
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   6315
      TabIndex        =   26
      Top             =   6405
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Delete"
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
      Picture         =   "frmUserMaintenance.frx":A6A2
      PicturePos      =   1
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  'Transparent
      Caption         =   "Guanzon Merchandising Corporation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   135
      TabIndex        =   30
      Top             =   1155
      Width           =   4260
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Guanzon Bldg., Perez Blvd., Dagupan City"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   135
      TabIndex        =   33
      Top             =   1395
      Width           =   4260
   End
   Begin VB.Label lblTelNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel No: (075) 522 1085; 522 1097"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   135
      TabIndex        =   40
      Top             =   1650
      Width           =   4260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   135
      X2              =   3735
      Y1              =   5820
      Y2              =   5820
   End
   Begin VB.Label lblOwner 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guanzon Group of Companies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000378ED&
      Height          =   195
      Index           =   0
      Left            =   660
      TabIndex        =   37
      Top             =   5955
      Width           =   2565
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmUserMaintenance.frx":AE1C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000378ED&
      Height          =   1260
      Index           =   1
      Left            =   210
      TabIndex        =   36
      Top             =   3885
      Width           =   3450
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © - GMC-SEG - 2004"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   35
      Top             =   6825
      Width           =   2205
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000A&
      BorderColor     =   &H00FFFFFF&
      Height          =   1875
      Left            =   105
      Top             =   3570
      Width           =   3645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warning"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000378ED&
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   34
      Top             =   3660
      Width           =   3450
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Maintenance"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000378ED&
      Height          =   420
      Index           =   1
      Left            =   4215
      TabIndex        =   31
      Tag             =   "wb0"
      Top             =   2025
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee No"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   3825
      TabIndex        =   8
      Top             =   3930
      Width           =   945
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   3825
      TabIndex        =   6
      Top             =   2895
      Width           =   795
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   3825
      TabIndex        =   4
      Top             =   3585
      Width           =   1260
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   3825
      TabIndex        =   2
      Top             =   3240
      Width           =   690
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   3825
      TabIndex        =   0
      Top             =   2550
      Width           =   855
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Maintenance"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   630
      Index           =   0
      Left            =   4395
      TabIndex        =   32
      Tag             =   "eb0"
      Top             =   2040
      Width           =   3900
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "This product is owned by:"
      ForeColor       =   &H000378ED&
      Height          =   675
      Index           =   1
      Left            =   105
      TabIndex        =   38
      Tag             =   "wt0;fb0"
      Top             =   5580
      Width           =   3645
   End
End
Attribute VB_Name = "frmUserMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents oDriver As FormDriver
Attribute oDriver.VB_VarHelpID = -1
Private p_oCrypt As Crypto
Private p_anUser(6) As Integer
Private pbLoaded As Boolean

Private Sub InitEntry()
   
   ' populate the user lever combo
10       Combo1.AddItem "Encoder"
20       Combo1.AddItem "Supervisor"
30       Combo1.AddItem "Manager"
40       Combo1.AddItem "Auditor"
50       Combo1.AddItem "SysAdmin"
60       Combo1.AddItem "Owner"
70       Combo1.AddItem "Engineer"
   
   ' populate the User Type combo
80       Combo2.AddItem "Local"
90       Combo2.AddItem "Global"
End Sub

Private Sub cmdButton_Click(Index As Integer)
Dim lsslq As String
10       Select Case Index
         Case 0
20          oDriver.RecordCancelUpdate
30       Case 1
40          BrowseRecord
50       Case 2
60          oDriver.RecordSave
70       Case 3
80          oDriver.RecordUpdate
90       Case 4
100         oDriver.RecordNew
110      Case 5
120         Unload Me
130      Case 6
            'Delete ELoad Ledger
            lsSQL = "DELETE xxxSysUser " _
                        & " WHERE sUserIDxx = '" & oDriver.FieldValue(0) & "'"
            oApp.Connection.Execute lsSQL, lnrow, adCmdText
            If lnrow <> 0 Then
               oApp.RegisDelete lsSQL
            Else
               MsgBox "Deleted"
            End If
150      Case 7
160         oDriver.RecordSearch
170      End Select
End Sub

Private Sub Command1_Click()
MsgBox oDriver.FieldValue(0)
End Sub

Private Sub Form_Activate()
10       If pbLoaded = False Then
20          oDriver.RecordNew
30          oDriver.DisableTextbox 0
40          pbLoaded = True
50       End If
End Sub

Private Sub Form_DblClick()
   MsgBox txtfield(3).Text
End Sub

Private Sub Form_Load()
         CenterChildForm mdiMain, Me
10       pbLoaded = False
   
20       InitEntry
   
21       p_anUser(0) = 1
22       p_anUser(1) = 2
23       p_anUser(2) = 4
24       p_anUser(3) = 8
25       p_anUser(4) = 16
26       p_anUser(5) = 32
27       p_anUser(6) = 64

30       Set oDriver = New FormDriver
40       Set oDriver.AppDriver = oApp
50       Set oDriver.MainForm = Me
         
         Set oSkin = New FormSkin
         Set oSkin.AppDriver = oApp
         Set oSkin.Form = Me
         oSkin.ApplySkin

   
60       Set p_oCrypt = New Crypto
70       p_oCrypt.Signature = oApp.Machinex
   
80       oDriver.RecQuery = "SELECT" & _
                           "  sUserIDxx" & _
                           ", sLogNamex" & _
                           ", sUserName" & _
                           ", sPassword" & _
                           ", sEmployNo" & _
                           ", nUserLevl" & _
                           ", cUserType" & _
                           ", sProdctID" & _
                           ", cUserStat" & _
                           ", nSysError" & _
                           ", cLogStatx" & _
                           ", cLockStat" & _
                           ", cAllwLock" & _
                           ", cAllwView" & _
                           ", sCompName" & _
                           ", sSkinCode" & _
                           ", sModified" & _
                           ", dModified" & _
                           ", vTimeStmp" & _
                        " FROM xxxSysUser" & _
                        " WHERE nUserLevl <> " & xeSysMaster

90       oDriver.InitRecForm
   
100      oDriver.LookupQuery(7) = "SELECT * FROM xxxAppObject" & _
                              " ORDER BY sApplName"

110      oDriver.LookupReference(7) = "sProdctID»sApplName"
120      oDriver.LookupColumn(7) = "sApplName"
130      oDriver.LookupTitle(7) = "Application"

140      oDriver.LookupQuery(15) = "SELECT sSkinCode, sSkinName FROM xxxSkin" & _
                              " ORDER BY sSkinName"

150      oDriver.LookupReference(15) = "sSkinCode»sSkinName"
160      oDriver.LookupColumn(15) = "sSkinName"
170      oDriver.LookupTitle(15) = "Skin"
180      oDriver.FieldReference(0) = True

190      oDriver.FieldStart = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
10       Set oDriver = Nothing
20       Set oSkin = Nothing
30       Set oApp = Nothing
End Sub

Private Sub oDriver_InitValue()
10       oDriver.FieldValue(0) = getNextCode()
20       oDriver.FieldValue(1) = ""
30       oDriver.FieldValue(2) = ""
40       oDriver.FieldValue(3) = ""
50       oDriver.FieldValue(4) = ""
60       oDriver.FieldValue(5) = 0
70       oDriver.FieldValue(6) = ""
80       oDriver.FieldValue(7) = ""
90       oDriver.FieldValue(8) = 0
100      oDriver.FieldValue(9) = 0
110      oDriver.FieldValue(10) = 0
120      oDriver.FieldValue(11) = 0
130      oDriver.FieldValue(12) = 1
140      oDriver.FieldValue(13) = 1
150      oDriver.FieldValue(14) = ""
160      oDriver.FieldValue(15) = ""
   
170      txtConfirm = ""
180      Check1.Value = 1
190      Check2.Value = 1
200      Check3.Value = 1
   
210      Combo1.ListIndex = -1
220      Combo2.ListIndex = -1
End Sub

Private Sub oDriver_LoadOtherData()
10       p_oCrypt.InBuffer = oDriver.FieldValue(1)
20       p_oCrypt.Encrypt
30       txtfield(1) = p_oCrypt.OutBuffer
   
40       p_oCrypt.InBuffer = oDriver.FieldValue(2)
50       p_oCrypt.Encrypt
60       txtfield(2) = p_oCrypt.OutBuffer
   
70       p_oCrypt.InBuffer = oDriver.FieldValue(3)
80       p_oCrypt.Encrypt
90       txtfield(3) = p_oCrypt.OutBuffer
95       txtConfirm = p_oCrypt.OutBuffer
110      Combo1.ListIndex = getUserLevel(oDriver.FieldValue(5))
120      Combo2.ListIndex = oDriver.FieldValue(6)
130      Check1.Value = oDriver.FieldValue(8)
140      Check2.Value = oDriver.FieldValue(12)
150      Check3.Value = oDriver.FieldValue(13)
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
10       If txtfield(3) <> txtConfirm Then
20          MsgBox "Password do not match!!!" & vbCrLf & "Verify your password!!!"
30          txtfield(3).SetFocus
40          Cancel = True
50          Exit Sub
60       End If
   
70       If oDriver.FieldValue(1) = "" Then
80          MsgBox "Invalid Login Name Detected!!!", vbCritical, "Warning"
90          txtfield(1).SetFocus
100         Cancel = True
110         Exit Sub
120      ElseIf oDriver.FieldValue(2) = "" Then
130         MsgBox "Invalid Name Detected!!!", vbCritical, "Warning"
140         txtfield(2).SetFocus
150         Cancel = True
160         Exit Sub
170      End If
   
180      p_oCrypt.InBuffer = LCase(txtfield(1))
190      p_oCrypt.Encrypt
200      oDriver.FieldValue(1) = p_oCrypt.OutBuffer
   
210      p_oCrypt.InBuffer = txtfield(2)
220      p_oCrypt.Encrypt
230      oDriver.FieldValue(2) = p_oCrypt.OutBuffer
   
240      p_oCrypt.InBuffer = txtfield(3)
250      p_oCrypt.Encrypt
260      oDriver.FieldValue(3) = p_oCrypt.OutBuffer

270      oDriver.FieldValue(5) = p_anUser(Combo1.ListIndex)
280      oDriver.FieldValue(6) = Combo2.ListIndex
290      oDriver.FieldValue(8) = Check1.Value
300      oDriver.FieldValue(12) = Check2.Value
310      oDriver.FieldValue(13) = Check3.Value
End Sub

Private Sub txtField_GotFocus(Index As Integer)
10       oDriver.ColumnIndex = Index
   
20       Select Case Index
   Case 3
30          txtfield(Index).Text = oDriver.FieldValue(Index)
40          txtfield(Index).SelStart = 0
50          txtfield(Index).SelLength = Len(txtfield(Index))
60       End Select
End Sub

Private Sub txtField_LostFocus(Index As Integer)
10       Dim lbCancel As Boolean

20       lbCancel = Not oDriver.ValidateField(Index)
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
10       If KeyCode = vbKeyF3 Then
20          oDriver.RecordSearch
30          KeyCode = 0
40       End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10       Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
20          Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
30             SetNextFocus
40          Case vbKeyUp
50             SetPreviousFocus
60          End Select
70       End Select
End Sub

Private Function getNextCode()
10       Dim lors As Recordset
20       Dim lsSQL As String
30       Dim lnCode As Long
40       Dim lnLen As Long
   
50       lsSQL = "SELECT TOP 1 sUserIdxx FROM xxxSysUser WHERE sUserIDxx LIKE " & _
               strParm(oApp.BranchCode & Format(oApp.ServerDate, "YY") & "%") & _
            " ORDER BY sUserIDxx DESC"
            
60       Set lors = New Recordset
70       lors.Open lsSQL, oApp.Connection, , , adCmdText

80       lsSQL = IIf(lors.EOF, Empty, lors(0))
90       lnLen = lors(0).DefinedSize
100      lnCode = 1
110      If lsSQL <> Empty Then
120         lnCode = CLng(Mid(lsSQL, 5)) + 1
130      Else
140         lnCode = 1
150      End If
160      getNextCode = oApp.BranchCode & Format(oApp.ServerDate, "YY") & Format(lnCode, "0000")
End Function

Private Sub BrowseRecord()
10       Dim lors As Recordset
20       Dim lsSelect As String
30       Dim lasSelect() As String
   
40       Set lors = New Recordset
45       If oApp.UserLevel <> xeSysMaster Then
50          lors.Open "SELECT * FROM xxxSysUser WHERE cAllwView = 1", _
                  oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
55       Else
56          lors.Open "SELECT * FROM xxxSysUser WHERE nUserLevl <> " & xeSysMaster, _
                  oApp.Connection, adOpenKeyset, adLockOptimistic, adCmdText
58       End If

60       Set lors.ActiveConnection = Nothing
70       If lors.EOF Then
80          MsgBox "No Record is Available for Selection!!!", vbCritical, "Warning"
90          GoTo endProc
100      End If
   
110      Do
120         p_oCrypt.InBuffer = lors("sLogNamex")
130         p_oCrypt.Decrypt
140         lors("sLogNamex") = p_oCrypt.OutBuffer
      
150         p_oCrypt.InBuffer = lors("sUserName")
160         p_oCrypt.Decrypt
170         lors("sUserName") = p_oCrypt.OutBuffer
      
180         lors.MoveNext
190      Loop Until lors.EOF
   
200      lsSelect = KwikBrowse(oApp, lors, "sLogNamex»sUserName»sProdctID", "Log Name»User Name»Product")
210      If lsSelect = "" Then
220         MsgBox "No Selection was Made", vbCritical, "Warning"
230         Exit Sub
240      End If
   
250      lasSelect = Split(lsSelect, "»")
260      oDriver.LookupValue(0) = lasSelect(0)
270      oDriver.LoadRecord

endProc:
280      Set lors = Nothing
290      Exit Sub
End Sub

Private Function getUserLevel(ByVal nUserLevel As Integer) As Integer
10       Dim lnctr As Integer

20       For lnctr = 0 To 6
30          If nUserLevel = p_anUser(lnctr) Then
40             getUserLevel = lnctr
50             Exit Function
60          End If
70       Next
End Function

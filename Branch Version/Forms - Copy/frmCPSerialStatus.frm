VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPSerialStatus 
   BorderStyle     =   0  'None
   Caption         =   "CP Serial Status"
   ClientHeight    =   6990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4680
      Index           =   0
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   1350
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   8255
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1050
         TabIndex        =   36
         Top             =   3375
         Width           =   2595
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   15
         Left            =   4470
         TabIndex        =   26
         Top             =   1725
         Width           =   1785
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   1050
         TabIndex        =   17
         Top             =   1725
         Width           =   2595
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   4470
         TabIndex        =   50
         Top             =   4275
         Width           =   1785
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   1050
         TabIndex        =   42
         Top             =   4275
         Width           =   2595
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   430
         Index           =   7
         Left            =   1050
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   2325
         Width           =   5205
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   4470
         TabIndex        =   48
         Top             =   3975
         Width           =   1785
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1050
         TabIndex        =   40
         Top             =   3975
         Width           =   2595
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   13
         Left            =   4470
         TabIndex        =   44
         Top             =   3375
         Width           =   1785
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   14
         Left            =   4470
         TabIndex        =   46
         Top             =   3675
         Width           =   1785
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1050
         TabIndex        =   38
         Top             =   3675
         Width           =   2595
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmCPSerialStatus.frx":0000
         Left            =   4470
         List            =   "frmCPSerialStatus.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1395
         Width           =   1785
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1050
         TabIndex        =   15
         Top             =   1425
         Width           =   2595
      End
      Begin VB.CheckBox Check1 
         Caption         =   "PNP Clearance"
         Height          =   270
         Index           =   3
         Left            =   4800
         TabIndex        =   21
         Tag             =   "et0;fb0"
         Top             =   1020
         Width           =   1410
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CSR Validation"
         Height          =   270
         Index           =   2
         Left            =   4800
         TabIndex        =   22
         Tag             =   "et0;fb0"
         Top             =   735
         Width           =   1365
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Registered"
         Height          =   270
         Index           =   1
         Left            =   3720
         TabIndex        =   20
         Tag             =   "et0;fb0"
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sold"
         Height          =   270
         Index           =   0
         Left            =   3720
         TabIndex        =   19
         Tag             =   "et0;fb0"
         Top             =   750
         Width           =   750
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1050
         TabIndex        =   32
         Top             =   2775
         Width           =   5205
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1050
         TabIndex        =   28
         Top             =   2025
         Width           =   5205
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   1050
         TabIndex        =   34
         Top             =   3075
         Width           =   5205
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1050
         TabIndex        =   13
         Top             =   1125
         Width           =   2595
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1050
         TabIndex        =   11
         Top             =   825
         Width           =   2595
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1050
         TabIndex        =   9
         Top             =   525
         Width           =   2595
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1050
         TabIndex        =   7
         Top             =   60
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Wart. No"
         Height          =   195
         Index           =   23
         Left            =   3690
         TabIndex        =   25
         Top             =   1740
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         Height          =   195
         Index           =   22
         Left            =   3690
         TabIndex        =   23
         Top             =   1410
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Type"
         Height          =   195
         Index           =   21
         Left            =   3690
         TabIndex        =   49
         Top             =   4290
         Width           =   870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Insur. Type"
         Height          =   195
         Index           =   20
         Left            =   135
         TabIndex        =   41
         Top             =   4290
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "D.R No."
         Height          =   195
         Index           =   19
         Left            =   3690
         TabIndex        =   47
         Top             =   4020
         Width           =   870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Purc."
         Height          =   195
         Index           =   18
         Left            =   135
         TabIndex        =   39
         Top             =   3990
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   16
         Left            =   135
         TabIndex        =   29
         Top             =   2340
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Plt No.(H)"
         Height          =   195
         Index           =   15
         Left            =   3690
         TabIndex        =   45
         Top             =   3690
         Width           =   870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Plt No.(P)"
         Height          =   195
         Index           =   14
         Left            =   3690
         TabIndex        =   43
         Top             =   3390
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Refer. Date"
         Height          =   195
         Index           =   13
         Left            =   135
         TabIndex        =   37
         Top             =   3690
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Refer. No."
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   35
         Top             =   3390
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   16
         Top             =   1740
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "MC Status"
         Height          =   195
         Index           =   6
         Left            =   3600
         TabIndex        =   18
         Tag             =   "et0;fb0"
         Top             =   510
         Width           =   885
      End
      Begin VB.Shape Shape2 
         Height          =   795
         Index           =   0
         Left            =   3675
         Top             =   585
         Width           =   2565
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   33
         Top             =   3090
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   31
         Top             =   2790
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   27
         Top             =   2040
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Frame No"
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   10
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   14
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Engine No"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   540
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   1140
         Width           =   705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial ID"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   135
         Width           =   765
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   1155
         Tag             =   "et0;ht2"
         Top             =   165
         Width           =   1920
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   5775
      TabIndex        =   58
      Top             =   6210
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
      Picture         =   "frmCPSerialStatus.frx":0004
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   4995
      TabIndex        =   57
      Top             =   6210
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
      Picture         =   "frmCPSerialStatus.frx":077E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   2655
      TabIndex        =   52
      Top             =   6210
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
      Picture         =   "frmCPSerialStatus.frx":0EF8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   1875
      TabIndex        =   51
      Top             =   6210
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
      Picture         =   "frmCPSerialStatus.frx":1672
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   5775
      TabIndex        =   59
      Top             =   6210
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
      Picture         =   "frmCPSerialStatus.frx":1DEC
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   705
      Index           =   7
      Left            =   3435
      TabIndex        =   53
      Top             =   6210
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
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
      Picture         =   "frmCPSerialStatus.frx":2566
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   2655
      TabIndex        =   54
      Top             =   6210
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
      Picture         =   "frmCPSerialStatus.frx":2CE0
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   3435
      TabIndex        =   55
      Top             =   6210
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
      Picture         =   "frmCPSerialStatus.frx":345A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   8
      Left            =   4215
      TabIndex        =   56
      Top             =   6210
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Ledger"
      AccessKey       =   "L"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPSerialStatus.frx":3BD4
      PicturePos      =   1
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   810
      Index           =   1
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   510
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1429
      BackColor       =   12632256
      Begin VB.TextBox txtOther 
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
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   3
         Top             =   105
         Width           =   2835
      End
      Begin VB.TextBox txtOther 
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
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   5
         Top             =   405
         Width           =   5235
      End
      Begin VB.TextBox txtOther 
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
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   105
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Custo&mer"
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
         Index           =   17
         Left            =   165
         TabIndex        =   4
         Top             =   450
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&ENo."
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
         Index           =   8
         Left            =   3060
         TabIndex        =   2
         Top             =   150
         Width           =   420
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Se&rial ID"
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
         Index           =   7
         Left            =   165
         TabIndex        =   0
         Top             =   150
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmCPSerialStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmMCSerialStatus"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oForm As frmMCSerialLedger
Private oSkin As clsFormSkin
Private bLoaded As Boolean
Private oFormReg As Object

Dim pbLoadRecord As Boolean
Dim pbOtherField As Boolean
Dim pnCtr As Integer, pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lsSearch As String
   Dim lsSelected() As String
   Dim lsSQL As String
   Dim lsOldProc As String
   
   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc
   
   If pbOtherField Then
      txtOther_LostFocus pnIndex
   Else
      txtField_LostFocus pnIndex
   End If
   
   Select Case Index
   Case 0
      oDriver.RecordCancelUpdate
   Case 1
      searchEngine "", True
   Case 2
      oDriver.RecordSave
   Case 3
      If Trim(oDriver.FieldValue(0)) <> "" Then oDriver.RecordUpdate
   Case 5
      Unload Me
   Case 6
      MsgBox "Unable to Delete Record" & vbCrLf & _
             "Deleting Record is prohibited!!!", vbCritical, "Warning"
   Case 7
      oDriver.RecordSearch
      txtField(pnIndex).SetFocus
   Case 8
      If pbLoadRecord Then
         oForm.SerialID = oDriver.FieldValue(0)
         oForm.Caption = "MC Serial Ledger"
         oForm.txtField(0).Text = txtField(0).Text
         oForm.txtField(1).Text = txtField(3).Text
         oForm.txtField(2).Text = txtField(4).Text
         oForm.txtField(3).Text = txtField(2).Text
         oForm.txtField(4).Text = txtField(1).Text
         oForm.Show 1
         
         If oForm.SystemID <> "" And oForm.TransactionNo <> "" Then
            Select Case oForm.SystemID
            Case "MCDv"
               'Set oFormReg = New frmMCDeliveryMaintenance
            Case "MCDA"
               'Set oFormReg = New frmDAMaintenance
            Case "MCPR"
               'Set oFormReg = New frmMCPOReturnReg
            Case "MCBT"
               'Set oFormReg = New frmBDMaintenance
            Case Else
               Exit Sub
            End Select
         
            oFormReg.TransactionNo = oForm.TransactionNo
            oFormReg.Show
         End If
      Else
         MsgBox "Unable to Load Serial Ledger!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Combo1_GotFocus()
   With Combo1
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      SetNextFocus
      KeyCode = 0
   End If
End Sub

Private Sub Combo1_LostFocus()
   With Combo1
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
   
   If bLoaded Then
      oDriver.RecordCancelUpdate
      bLoaded = False
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oDriver = New clsFormDriver
   Set oDriver.AppDriver = oApp
   Set oDriver.MainForm = Me

   Set oForm = New frmMCSerialLedger
      
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin

   oDriver.RecQuery = "SELECT" _
                        & "  sSerialID" _
                        & ", sEngineNo" _
                        & ", sFrameNox" _
                        & ", sModelIDx" _
                        & ", sColorIDx" _
                        & ", sClientID" _
                        & ", sCompnyID" _
                        & ", sBranchCd" _
                        & ", cLocation" _
                        & ", cSoldStat" _
                        & ", cRegister" _
                        & ", cCSRValid" _
                        & ", cPNPClear" _
                        & ", sPlateNoP" _
                        & ", sPlateNoH" _
                        & ", sWarrntNo" _
                        & ", sModified" _
                        & ", dModified" _
                     & " FROM MC_Serial"
   
   oDriver.BrowseQuery = "SELECT Distinct" _
                        & "  a.sSerialID" _
                        & ", a.sEngineNo" _
                        & ", a.sFrameNox" _
                        & ", d.sReferNox" _
                        & ", d.dReferDte" _
                        & ", a.sPlateNoP" _
                        & ", a.sPlateNoH" _
                        & ", b.sModelNme" _
                        & ", c.sColorNme" _
                        & ", a.sBranchCd" _
                        & ", f.sBrandNme" _
                     & " FROM MC_Serial a" _
                        & " Left Join MC_PO_Receiving_Serial e" _
                           & " On a.sSerialID = e.sSerialID" _
                        & " Left Join MC_PO_Receiving_Master d" _
                              & " On d.sTransNox = e.sTransNox" _
                        & " Left Join MC_Model b" _
                           & " On a.sModelIDx = b.sModelIDx" _
                        & " Left Join Brand f" _
                           & " On b.sBrandIDx = f.sBrandIDx" _
                        & " Left Join Color c" _
                           & " On a.sColorIDx = c.sColorIDx" _
                     & " ORDER BY a.sSerialID"

   oDriver.InitRecForm
   
   oDriver.BrowseFReference(0) = True
   oDriver.BrowseFTitle(0) = "Serial ID"
   oDriver.BrowseFTitle(1) = "Engine No"
   oDriver.BrowseFTitle(2) = "Frame No"
   oDriver.BrowseFTitle(3) = "Refer No"
   oDriver.BrowseFTitle(4) = "Refer Date"
   oDriver.BrowseFTitle(5) = "Plate No(P)"
   oDriver.BrowseFTitle(6) = "Plate No(H)"
   oDriver.BrowseFTitle(7) = "Model"
   oDriver.BrowseFTitle(8) = "Color"
   
   oDriver.LookupQuery(3) = "SELECT" _
                              & "  sModelIDx" _
                              & ", sModelNme" _
                           & " FROM MC_Model" _
                           & " ORDER BY sModelNme"
   
   oDriver.LookupReference(3) = "sModelIDx»sModelNme"
   oDriver.LookupColumn(3) = "sModelNme"
   oDriver.LookupTitle(3) = "Model"

   oDriver.LookupQuery(4) = "SELECT" _
                              & "  sColorIDx" _
                              & ", sColorNme" _
                           & " FROM Color" _
                           & " ORDER BY sColorNme"
                        
   oDriver.LookupReference(4) = "sColorIDx»sColorNme"
   oDriver.LookupColumn(4) = "sColorNme"
   oDriver.LookupTitle(4) = "Color"
   
   oDriver.LookupQuery(5) = "SELECT" _
                              & "  sClientID" _
                              & ", CONCAT(sLastName, ', ', sFrstName, ' ', sMiddName) xFullName" _
                           & " FROM Client_Master" _
                           & " ORDER BY xFullName"
   
   oDriver.LookupReference(5) = "sClientID»xFullName"
   oDriver.LookupColumn(5) = "xFullName"
   oDriver.LookupTitle(5) = "Customer"
   
   oDriver.LookupQuery(6) = "SELECT" _
                              & "  sCompnyID" _
                              & ", sCompnyNm" _
                           & " FROM Company" _
                           & " ORDER BY sCompnyNm"
                           
   oDriver.LookupReference(6) = "sCompnyID»sCompnyNm"
   oDriver.LookupColumn(6) = "sCompnyNm"
   oDriver.LookupTitle(6) = "Company"
   
   oDriver.LookupQuery(7) = "SELECT" _
                              & "  sBranchCd" _
                              & ", sBranchNm" _
                           & " FROM Branch" _
                           & " ORDER BY sBranchNm"
                           
   oDriver.LookupReference(7) = "sBranchCd»sBranchNm"
   oDriver.LookupColumn(7) = "sBranchNm"
   oDriver.LookupTitle(7) = "Branch"
   
   oDriver.FieldFormat(0) = "@@@@-@@@@@@"
   oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
   
   oDriver.FieldStart = 1

   Combo1.ListIndex = -1
   Combo1.List(0) = "Warehouse"
   Combo1.List(1) = "Branch"
   Combo1.List(2) = "Supplier"
   Combo1.List(3) = "Customer"
   Combo1.List(4) = "Unknown"
   
   bLoaded = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oDriver = Nothing
   Set oSkin = Nothing
   Set oForm = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
   oDriver.HideButton 3
   oDriver.HideButton 4
   oDriver.HideButton 6
   Combo1.Enabled = False
   
   Check1(0).Enabled = False
   Check1(1).Enabled = False
   Check1(2).Enabled = False
   Check1(3).Enabled = False
   
   txtOther(0).Enabled = True
   txtOther(1).Enabled = True
   txtOther(4).Enabled = True
End Sub

Private Sub oDriver_EnableOtherControl()
   oDriver.DisableTextbox 0
   oDriver.DisableTextbox 5
   oDriver.DisableTextbox 6
   oDriver.DisableTextbox 7
   oDriver.DisableTextbox 13
   oDriver.DisableTextbox 14
   
   oDriver.ShowButton 1
   Combo1.Enabled = True
   
   Check1(0).Enabled = True
   Check1(1).Enabled = True
   Check1(2).Enabled = True
   Check1(3).Enabled = True
   
   txtOther(0).Enabled = False
   txtOther(1).Enabled = False
   txtOther(4).Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub oDriver_InitValue()
   pbLoadRecord = False
   
   For pnCtr = 0 To txtOther.Count - 1
      txtOther(pnCtr).Text = ""
      txtOther(pnCtr).Tag = ""
   Next
End Sub

Private Sub oDriver_LoadOtherData()
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_LoadOtherData"
   On Error GoTo errProc

   Dim lrs As ADODB.Recordset
   
   For pnCtr = 0 To Check1.Count - 1
      Check1(pnCtr).Value = IIf(IsNull(oDriver.FieldValue(pnCtr + 9)), 0, IIf(Trim(oDriver.FieldValue(pnCtr + 9)) = "", 0, oDriver.FieldValue(pnCtr + 9)))
   Next
   
   For pnCtr = 2 To 9
      txtOther(pnCtr).Text = ""
   Next
   
   Combo1.ListIndex = IIf(IsNull(oDriver.FieldValue(8)), -1, IIf(Trim(oDriver.FieldValue(8)) = "", -1, oDriver.FieldValue(8)))
   txtOther(0).Text = Format(oDriver.FieldValue(0), "@@@@-@@@@@@")
   txtOther(1).Text = oDriver.FieldValue(1)
   txtOther(4).Text = txtField(5).Text
   txtOther(0).Tag = txtOther(0).Text
   txtOther(1).Tag = txtOther(1).Text
   txtOther(4).Tag = txtOther(4).Text
   
   Set lrs = New ADODB.Recordset
   lrs.Open "Select" _
               & "  a.sReferNox" _
               & ", a.dReferDte" _
            & " From MC_PO_Receiving_Master a" _
               & ", MC_PO_Receiving_Serial b" _
            & " Where a.sTransNox = b.sTransNox" _
               & " And b.sSerialID = " & strParm(oDriver.FieldValue(0)) _
            , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If Not lrs.EOF Then
      txtOther(2).Text = lrs("sReferNox")
      txtOther(3).Text = Format(lrs("dReferDte"), "MMMM DD, YYYY")
   End If
   lrs.Close
   
   lrs.Open "Select" _
               & "  a.sDRNoxxxx" _
               & ", a.dTransact" _
               & ", CONCAT(c.sAddressx, ', ', d.sTownName, ', ', e.sProvName, ' ', d.sZippCode) xAddressx" _
               & ", f.sInsTypNm" _
               & ", b.cMotorNew" _
            & " From MC_SO_Master a" _
               & ", Client_Master c" _
                  & " Left Join TownCity d" _
                     & " On c.sTownIDxx = d.sTownIDxx" _
                  & " Left Join Province e" _
                     & " On d.sProvIDxx = e.sProvIDxx" _
               & ", MC_SO_Detail b" _
                  & " Left Join Insurance_Type f" _
                     & " On b.sInsTypID = f.sInsTypID" _
            & " Where a.sTransNox = b.sTransNox" _
               & " And a.sClientID = c.sClientID" _
               & " And b.sSerialID = " & strParm(oDriver.FieldValue(0)) _
               & " And a.sClientID = " & strParm(IIf(IsNull(oDriver.FieldValue(5)), "", oDriver.FieldValue(5))) _
            , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If Not lrs.EOF Then
      txtOther(5).Text = IIf(IsNull(lrs("dTransact")), "", Format(lrs("dTransact"), "MMMM DD, YYYY"))
      txtOther(6).Text = IIf(IsNull(lrs("sDRNoxxxx")), "", lrs("sDRNoxxxx"))
      txtOther(7).Text = IIf(IsNull(lrs("xAddressx")), "", lrs("xAddressx"))
      txtOther(8).Text = IIf(IsNull(lrs("sInsTypNm")), "", lrs("sInsTypNm"))
      txtOther(9).Text = IIf(IsNull(lrs("cMotorNew")), "", IIf(Trim(lrs("cMotorNew")) = "", "", IIf(lrs("cMotorNew") = 1, "New Unit", "Repo Unit")))
   End If
   lrs.Close
   
   pbLoadRecord = True

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )"
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
   Dim lnCtr As Integer
   Dim lsOldProc As String
   
   lsOldProc = "oDriver_WillSave"
   On Error GoTo errProc
   
   If Combo1.ListIndex <> -1 Then oDriver.FieldValue(8) = CStr(Combo1.ListIndex)

   For lnCtr = 0 To Check1.Count - 1
      oDriver.FieldValue(lnCtr + 9) = CStr(Check1(lnCtr).Value)
   Next
   
   If oDriver.isModify Then
      oApp.Execute "Update JobOrder_Master Set" _
                                 & " sCouponNo = " & strParm(oDriver.FieldValue(15)) _
                              & " Where sSerialID = " & strParm(oDriver.FieldValue(0)), "JobOrder_Master"
                                    
'      oApp.Execute "Update MC_SO_Detail Set" _
'                                 & " sWarrntNo = " & strParm(oDriver.FieldValue(15)) _
'                              & " Where sSerialID = " & strParm(oDriver.FieldValue(0)), "JobOrder_Detail"
   End If
   
endProc:
   Exit Sub
errProc:
   Cancel = True
   ShowError lsOldProc & "( " & Cancel & " )"

End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   pnIndex = Index
   pbOtherField = False
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   On Error GoTo errProc
   
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
      With txtField(Index)
         If KeyCode = vbKeyF3 Then
            oDriver.RecordSearch .Text
            If .Text <> "" Then SetNextFocus
         Else
            If .Text <> "" Then oDriver.RecordSearch .Text
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
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   On Error GoTo errProc
   
   With txtField(Index)
      If Trim(.Text) <> "" Then .Text = UCase(Left(.Text, 1)) & Right(.Text, Len(.Text) - 1)
      
      Select Case Index
      Case 5
      Case Else
         If Index = 1 Or Index = 2 Then .Text = UCase(.Text)
         Cancel = Not oDriver.ValidateField(Index)
      End Select
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub

Private Sub txtOther_GotFocus(Index As Integer)
   With txtOther(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With
   pnIndex = Index
   pbOtherField = True
End Sub

Private Sub txtOther_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtOther_KeyDown"
   On Error GoTo errProc

   If KeyCode = vbKeyReturn Or KeyCode = vbKeyF3 Then
      With txtOther(Index)
         If Trim(.Text) = "" Then
            oDriver_InitValue
            Exit Sub
         End If
         
         If .Text <> .Tag Then
            Select Case Index
            Case 0
               oDriver.LookupValue(0) = .Text
               oDriver.LoadRecord
            Case 1
               searchEngine .Text, False
            Case 4
               SearchCustomer .Text, False
            End Select
        End If
      
        .Tag = .Text
        .SelStart = 0
        .SelLength = Len(.Text)
        
        If KeyCode = vbKeyF3 Then
            If .Text <> "" Then SetNextFocus
        End If
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift & " )", True
End Sub

Private Sub searchEngine(lsEngineNo As String, lbSearch As Boolean)
   Dim lrs As ADODB.Recordset
   Dim lsSQL As String, lsOldProc As String
   Dim lsCondition As String, lsSearch As String, lsSelected() As String
   
   lsOldProc = "searchEngine"
   On Error GoTo errProc
   
   lsCondition = " sEngineNO Like " & strParm("%" & lsEngineNo)
   lsSQL = AddCondition(oDriver.BrowseQuery, lsCondition)
   
   Set lrs = New ADODB.Recordset
   lrs.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   With txtOther(10)
      .Text = ""
      If lrs.RecordCount = 1 Then
         oDriver.LookupValue(0) = lrs("sSerialID")
         .Text = IIf(IsNull(lrs("sBrandNme")), "", lrs("sBrandNme"))
         oDriver.LoadRecord
      ElseIf lrs.RecordCount > 1 Then
         lsSearch = KwikSearch(oApp, lsSQL _
                              , "sSerialID»sEngineNo»sFrameNox»sReferNox»dReferDte»sPlateNoP»sPlateNoH»sModelNme»sColorNme" _
                              , "Code»EngineNo»FrameNo»ReferNo»ReferDate»PlateNo(P)»PlateNo(H)»Model»Color" _
                              , "@»@»@»@»MM/DD/YYYY»@»@»@»@")
         If lsSearch <> "" Then
            lsSelected = Split(lsSearch, "»")
            oDriver.LookupValue(0) = lsSelected(0)
            .Text = lsSelected(10)
            oDriver.LoadRecord
         End If
      End If
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & lsEngineNo _
                       & ", " & lbSearch & " )"
End Sub

Private Sub SearchCustomer(lsCustName As String, lbSearch As Boolean)
   Dim lrs As ADODB.Recordset
   Dim lsSQL As String, lsCondition As String
   Dim lsSearch As String, lsSelected() As String, lsOldProc As String
   
   lsOldProc = "SearchCustomer"
   On Error GoTo errProc

   lsSelected = GetSplitedName(lsCustName)
   lsCondition = lsCondition & _
            " b.sLastName LIKE " & strParm(lsSelected(0) & "%") & _
               " AND (b.sFrstName LIKE " & strParm(lsSelected(1) & "%") & _
                  " OR b.sFrstName LIKE " & strParm(lsSelected(1) & lsSelected(2) & "%") & _
                  IIf(lsSelected(2) = Empty, " )", _
                     " OR b.sMiddName LIKE " & strParm(lsSelected(2) & "%") & ")")
   
   lsSQL = "Select" _
               & "  a.sSerialID" _
               & ", a.sEngineNo" _
               & ", a.sFrameNox" _
               & ", CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName) as xFullName" _
               & ", d.sBrandNme" _
            & " From MC_Serial a" _
                  & " Left Join MC_Model c" _
                     & " Left Join Brand d" _
                        & " On c.sBrandIDx = d.sBrandIDx" _
                     & " On a.sModelIDx = c.sModelIDx" _
               & ", Client_Master b" _
            & " Where a.sClientID = b.sClientID"
   
   lsSQL = AddCondition(lsSQL, lsCondition)
   
   Set lrs = New ADODB.Recordset
   lrs.Open lsSQL, oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   With txtOther(10)
      .Text = ""
      If lrs.RecordCount = 1 Then
         oDriver.LookupValue(0) = lrs("sSerialID")
         .Text = IIf(IsNull(lrs("sBrandNme")), "", lrs("sBrandNme"))
         oDriver.LoadRecord
      ElseIf lrs.RecordCount > 1 Then
         lsSearch = KwikBrowse(oApp, lrs _
                              , "sSerialID»sEngineNo»sFrameNox»xFullName" _
                              , "Code»EngineNo»FrameNo»Name" _
                              , "@»@»@»@" _
                              , "a.sSerialID»a.sEngineNo»a.sFrameNox»CONCAT(b.sLastName, ', ', b.sFrstName, ' ', b.sMiddName)")
         If lsSearch <> "" Then
            lsSelected = Split(lsSearch, "»")
            oDriver.LookupValue(0) = lsSelected(0)
            .Text = lsSelected(4)
            oDriver.LoadRecord
         End If
      End If
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & lsCustName _
                       & ", " & lbSearch & " )"
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

Private Sub txtOther_LostFocus(Index As Integer)
   With txtOther(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

Private Sub txtOther_Validate(Index As Integer, Cancel As Boolean)
   With txtOther(Index)
      If Trim(.Text) <> "" Then .Text = UCase(Left(.Text, 1)) & Right(.Text, Len(.Text) - 1)
   End With
End Sub

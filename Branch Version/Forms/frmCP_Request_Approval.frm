VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCP_Request_Approval 
   BorderStyle     =   0  'None
   Caption         =   "CP Unit Stock Request Approval"
   ClientHeight    =   9270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   14205
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   4080
      Left            =   8190
      TabIndex        =   4
      Top             =   1080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7197
      _Version        =   393216
      Cols            =   3
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4080
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   1080
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   7197
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   9
         Left            =   1155
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2100
         Width           =   2895
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1155
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   420
         Width           =   3495
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   14
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1125
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   13
         Left            =   3142
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1125
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   1155
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1125
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   25
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   3585
         Width           =   750
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   4710
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   90
         Width           =   1545
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   3135
         Width           =   750
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   3135
         Width           =   750
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   3135
         Width           =   750
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1155
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   720
         Width           =   5100
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   24
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3585
         Width           =   750
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   22
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   3585
         Width           =   750
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1155
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1770
         Width           =   2895
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   16
         Left            =   5535
         TabIndex        =   27
         Tag             =   "et0;hb0"
         Top             =   1935
         Width           =   705
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   13
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2835
         Width           =   750
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   2535
         Width           =   750
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2835
         Width           =   750
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2835
         Width           =   750
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   15
         Left            =   3750
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2535
         Width           =   750
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   14
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2535
         Width           =   750
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   10
         Left            =   1155
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   90
         Width           =   3495
      End
      Begin VB.TextBox txtOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   5535
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "24"
         Top             =   1515
         Width           =   705
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1155
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Notes"
         Height          =   255
         Index           =   28
         Left            =   645
         TabIndex        =   22
         Tag             =   "wt0;fb0"
         Top             =   2130
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Order No"
         Height          =   195
         Index           =   27
         Left            =   465
         TabIndex        =   8
         Tag             =   "wt0;fb0"
         Top             =   420
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Capacity"
         Height          =   195
         Index           =   26
         Left            =   4485
         TabIndex        =   16
         Tag             =   "wt0;fb0"
         Top             =   1065
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estimate"
         Height          =   195
         Index           =   24
         Left            =   2475
         TabIndex        =   14
         Tag             =   "wt0;fb0"
         Top             =   1065
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "On Hand"
         Height          =   195
         Index           =   23
         Left            =   465
         TabIndex        =   12
         Tag             =   "wt0;fb0"
         Top             =   1065
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Back Order"
         Height          =   255
         Index           =   17
         Left            =   4590
         TabIndex        =   61
         Tag             =   "wt0;fb0"
         Top             =   3600
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Back Order"
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   59
         Tag             =   "wt0;fb0"
         Top             =   3165
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reserve"
         Height          =   255
         Index           =   5
         Left            =   1215
         TabIndex        =   58
         Tag             =   "wt0;fb0"
         Top             =   3150
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Branch"
         Height          =   255
         Index           =   1
         Left            =   555
         TabIndex        =   5
         Tag             =   "wt0;fb0"
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Color"
         Height          =   255
         Index           =   14
         Left            =   555
         TabIndex        =   20
         Tag             =   "wt0;fb0"
         Top             =   1800
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Remarks"
         Height          =   195
         Index           =   3
         Left            =   495
         TabIndex        =   10
         Tag             =   "wt0;fb0"
         Top             =   750
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Issuing Branch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   13
         Left            =   210
         TabIndex        =   54
         Tag             =   "wt0;fb0"
         Top             =   3525
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reserve"
         Height          =   255
         Index           =   18
         Left            =   3120
         TabIndex        =   53
         Tag             =   "wt0;fb0"
         Top             =   3600
         Width           =   585
      End
      Begin VB.Shape Shape1 
         Height          =   405
         Index           =   2
         Left            =   1155
         Top             =   3525
         Width           =   5265
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "On Hand"
         Height          =   255
         Index           =   25
         Left            =   1245
         TabIndex        =   52
         Tag             =   "wt0;fb0"
         Top             =   3600
         Width           =   690
      End
      Begin VB.Shape Shape1 
         Height          =   15
         Index           =   1
         Left            =   75
         Top             =   1365
         Width           =   6345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   45
         TabIndex        =   34
         Tag             =   "wt0;fb0"
         Top             =   2985
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   45
         TabIndex        =   33
         Tag             =   "wt0;fb0"
         Top             =   2745
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Approved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   4005
         TabIndex        =   26
         Tag             =   "ht0;fb0"
         Top             =   2010
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "ROQ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   4815
         TabIndex        =   43
         Tag             =   "wt0;fb0"
         Top             =   3165
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Request"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   4515
         TabIndex        =   24
         Tag             =   "wt0;fb0"
         Top             =   1590
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Requesting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   45
         TabIndex        =   32
         Tag             =   "wt0;fb0"
         Top             =   2505
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Model"
         Height          =   255
         Index           =   2
         Left            =   660
         TabIndex        =   18
         Tag             =   "wt0;fb0"
         Top             =   1470
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Max Level"
         Height          =   255
         Index           =   19
         Left            =   2640
         TabIndex        =   39
         Tag             =   "wt0;fb0"
         Top             =   2580
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "On Transit"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   41
         Tag             =   "wt0;fb0"
         Top             =   2865
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         Height          =   1020
         Index           =   0
         Left            =   1155
         Top             =   2475
         Width           =   5265
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "AMC"
         Height          =   255
         Index           =   7
         Left            =   4545
         TabIndex        =   46
         Tag             =   "wt0;fb0"
         Top             =   2865
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Min Level"
         Height          =   255
         Index           =   4
         Left            =   1155
         TabIndex        =   35
         Tag             =   "wt0;fb0"
         Top             =   2565
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Classification"
         Height          =   255
         Index           =   9
         Left            =   4545
         TabIndex        =   44
         Tag             =   "wt0;fb0"
         Top             =   2550
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "On Hand"
         Height          =   255
         Index           =   8
         Left            =   1155
         TabIndex        =   37
         Tag             =   "wt0;fb0"
         Top             =   2850
         Width           =   750
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3960
      Left            =   1575
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5190
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   6985
      _Version        =   393216
      Cols            =   3
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   31
      Top             =   1785
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
      Picture         =   "frmCP_Request_Approval.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   29
      Top             =   525
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCP_Request_Approval.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   48
      Top             =   525
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
      Picture         =   "frmCP_Request_Approval.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   30
      Top             =   1155
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
      Picture         =   "frmCP_Request_Approval.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   49
      Top             =   1155
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
      Picture         =   "frmCP_Request_Approval.frx":1DE8
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   540
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   953
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   9705
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   75
         Width           =   2670
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1170
         TabIndex        =   1
         Top             =   75
         Width           =   5250
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delivery Schedule"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   8070
         TabIndex        =   2
         Tag             =   "wt0;fb0"
         Top             =   135
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cluster"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   465
         TabIndex        =   0
         Tag             =   "wt0;fb0"
         Top             =   135
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCP_Request_Approval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmCP_Request_Approval"

Private oSkin As clsFormSkin
Private oTrans As clsCPReqApproval

Private pnActiveRec As Integer
Private pnIndex As Integer

Private pnPrintRow As Integer
'Private poPrinter As clsPrintDirect

Private pbLoaded As Boolean
Private pbCtrlPress As Boolean
Private pbHasFocus As Boolean
Private pbModified As Boolean
Private pnLastRow As Integer

Private psOrderT As String

Private Const pxeMaxLine As Integer = 65

Private Sub cmdButton_Click(Index As Integer)
   Dim lsProcName As String

   lsProcName = "cmdButton_Click"
   'On Error GoTo errProc

   Select Case Index
   Case 0 ' save
      If Not pbModified Then
         MsgBox "Entry was not Modified!" & vbCrLf & _
                  "No Info will be updated!!!", vbInformation, "Notice"
         GoTo endProc
      End If

      If oTrans.ApproveRequest Then
         MsgBox "Branch Stock Request was Successfully Processed!", vbInformation, "Success"

         pnLastRow = MSFlexGrid2.Row
         Call initButton(xeModeReady)

         With MSFlexGrid2
            If .TextMatrix(1, 2) <> "" Then
               .Col = 1
               .ColSel = .Cols - 1

               If Not .RowIsVisible(.Row) Then
                  If .Rows > 14 Then .TopRow = .Row
               End If

               Call loadSelTrans

               With MSFlexGrid1
                  .Row = 1
                  .Col = 1
                  .ColSel = .Cols - 1

                  Call setFieldInfo
               End With
            Else
               Call InitFields
            End If
         End With

         pbModified = False
      End If
   Case 1 ' cancel
      If pbModified Then
         If MsgBox("Record was Modified!" & vbCrLf & _
                  "Closing entry will disregard any changes made!" & vbCrLf & vbCrLf & _
                  "Do you want to continue?", vbCritical + vbYesNo, "Confirm") <> vbYes Then
            Exit Sub
         End If
      End If
      Call initButton(xeModeReady)
      loadCluster
      txtSearch(0).SetFocus
      pbModified = False
   Case 2 ' Update
      ' check if valid record was loaded
      If txtField(1) <> "" Then
         Call initButton(xeModeUpdate)
         txtOthers(16).SetFocus
      Else
         MsgBox "No Record was loaded for modification!", vbCritical, "Warning"
         txtSearch(0).SetFocus
      End If
   Case 3 ' close
      Unload Me
   Case 4
      If txtField(0).Text = "" Then Exit Sub

      If MsgBox("Print Transaction!!!", vbYesNo + vbQuestion, "Confirm") = vbYes Then
         If Not PrintTrans Then MsgBox "Unable to Print Transaction!!!", vbCritical, "Warning"
      End If
   Case 5
      If pbModified Then
         MsgBox "Transaction is in update mode. Request not granted.", vbCritical, "Warning"
      Else
         Call ClearDetail
         Call InitFields
         Call loadCluster
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Index & " )"
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String

   lsOldProc = "Form_Activate"
   'On Error GoTo errProc

   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If Not pbLoaded Then pbLoaded = True
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Deactivate()
   'pnPrintRow = pxeMaxLine
End Sub

Private Sub Form_Load()
   CenterChildForm mdiMain, Me

   Set oTrans = New clsCPReqApproval
   Set oTrans.AppDriver = oApp
   If Not oTrans.InitTransaction Then Unload Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualLeft

   Call InitGrid
   Call InitFields
   Call clearCluster

   Call initButton(xeModeReady)

'   pnPrintRow = pxeMaxLine
   pnPrintRow = 0
   pnLastRow = 1

End Sub

Private Sub InitGrid()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Cols = 10
      .Rows = 2
      .Font = "MS Sans Serif"
      .RowHeight(0) = 350

      'column title
      .TextMatrix(0, 1) = "Model"
      .TextMatrix(0, 2) = "Color"
      .TextMatrix(0, 3) = "QTY"
      .TextMatrix(0, 4) = "Aprv"
      .TextMatrix(0, 5) = "Class"
      .TextMatrix(0, 6) = "ROQ"
      .TextMatrix(0, 7) = "QOH"
      .TextMatrix(0, 8) = "Rsv"
      .TextMatrix(0, 9) = "BO"
      
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 4
      Next
      
      .Row = 1

      .ColWidth(0) = 400
      .ColWidth(1) = 4500
      .ColWidth(2) = 1500
      .ColWidth(3) = 850
      .ColWidth(4) = 850
      .ColWidth(5) = 850
      .ColWidth(6) = 850
      .ColWidth(7) = 850
      .ColWidth(8) = 850
      .ColWidth(9) = 850

      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
      .ColAlignment(6) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
      .ColAlignment(6) = flexAlignRightCenter
      .ColAlignment(7) = flexAlignRightCenter
      .ColAlignment(8) = flexAlignRightCenter
      .ColAlignment(9) = flexAlignRightCenter
      
      .Col = 1
      .ColSel = .Cols - 1
   End With

   With MSFlexGrid2
      .Cols = 5
      .Rows = 2
      .Font = "MS Sans Serif"
      .RowHeight(0) = 350

      'column title
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "Transact #"
      .TextMatrix(0, 3) = "Req"
      .TextMatrix(0, 4) = "App"
      .Row = 0

      'column alignment
      For lnCtr = 0 To .Cols - 1
         .Col = lnCtr
         .CellFontBold = True
         .CellAlignment = 4
      Next

      .ColWidth(0) = 400
      .ColWidth(1) = 2750
      .ColWidth(2) = 1500
      .ColWidth(3) = 600
      .ColWidth(4) = 600
      
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter

      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   xrFrame1.Enabled = lbShow
   xrFrame3.Enabled = Not lbShow
   
   txtField(1).Enabled = Not lbShow
   txtField(2).Enabled = lbShow
   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = Not lbShow
   cmdButton(3).Visible = Not lbShow
   cmdButton(4).Visible = Not lbShow

   MSFlexGrid2.Enabled = Not lbShow
End Sub

Private Sub InitFields()
   Dim loText As TextBox
   Dim lnCtr As Integer

   Call clearMaster
   Call ClearDetail

'   For Each loText In txtField
'      Select Case loText.Index
'      Case 1
'         loText.Text = Format(oTrans.Master(loText.Index), "Mmm dd, yyyy")
'      Case Else
'         loText.Text = oTrans.Master(loText.Index)
'      End Select
'   Next
'
'   For Each loText In txtOthers
'      loText.Text = oTrans.Detail(0, loText.Index)
'   Next

   pbModified = False
End Sub

Private Sub LoadMaster()
   Dim loText As TextBox

   For Each loText In txtField
      Select Case loText.Index
      Case 1
         loText.Text = Format(oTrans.Master(loText.Index), "Mmm dd, yyyy")
      Case Else
         loText.Text = oTrans.Master(loText.Index)
      End Select
   Next
   'Check1.Value = oTrans.Master("cMotorNew")
End Sub

Private Sub LoadDetail()
   Dim lsProcName As String
   Dim lnCtr As Integer

   lsProcName = "loadDetail"
   'On Error GoTo errProc
   'Debug.Print lsProcName

   With MSFlexGrid1
      .Rows = oTrans.ItemCount + 1

      If .Rows > 15 Then
         .ColWidth(1) = 4370
      Else
         .ColWidth(1) = 4500
      End If

      For lnCtr = 0 To oTrans.ItemCount - 1
         .TextMatrix(lnCtr + 1, 0) = lnCtr + 1
         .TextMatrix(lnCtr + 1, 1) = oTrans.Detail(lnCtr, "sModelNme")
         .TextMatrix(lnCtr + 1, 2) = IFNull(oTrans.Detail(lnCtr, "sColorNme"))
         .TextMatrix(lnCtr + 1, 3) = oTrans.Detail(lnCtr, "nQuantity")
         .TextMatrix(lnCtr + 1, 4) = oTrans.Detail(lnCtr, "nApproved")
         .TextMatrix(lnCtr + 1, 5) = IFNull(oTrans.Detail(lnCtr, "cClassify"), "F")
         .TextMatrix(lnCtr + 1, 6) = oTrans.Detail(lnCtr, "nRecOrder")
         .TextMatrix(lnCtr + 1, 7) = oTrans.Detail(lnCtr, "nQtyOnHnd")
         .TextMatrix(lnCtr + 1, 8) = oTrans.Detail(lnCtr, "nResvOrdr")
         .TextMatrix(lnCtr + 1, 9) = oTrans.Detail(lnCtr, "nBackOrdr")
      Next

      .Row = 1
      .RowSel = 1

      .Col = 1
      .ColSel = .Cols - 1
   End With

   ' assign value of the first row to the receiving fields
   pnActiveRec = 1
   Call setFieldInfo

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Sub clearMaster()
   Dim loText As TextBox

   For Each loText In txtField
      loText.Text = ""
   Next
   
   For Each loText In txtOthers
      loText.Text = ""
   Next
End Sub

Private Sub ClearDetail()
   With MSFlexGrid1
      .Rows = 2
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = 0
      .TextMatrix(1, 4) = 0
      .TextMatrix(1, 5) = ""
      .TextMatrix(1, 6) = 0
      .TextMatrix(1, 7) = 0
      .TextMatrix(1, 8) = 0
      .TextMatrix(1, 9) = 0
   End With
End Sub

Private Sub clearCluster()
   txtSearch(0) = ""
   txtSearch(1) = Format(Now(), "Mmmm dd, yyyy")

   With MSFlexGrid2
      .Rows = 2
      .TextMatrix(1, 0) = ""
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = 0
      .TextMatrix(1, 4) = 0
   End With
End Sub

Private Sub setFieldInfo()
   Dim lsProcName As String
   Dim loText As TextBox
   
   lsProcName = "setFieldInfo"
   'On Error GoTo errProc
   'Debug.Print lsProcName

   With oTrans
      If oTrans.Detail(pnActiveRec - 1, "sStockIDx") = "" Then GoTo endProc
      
      txtOthers(2) = oTrans.Detail(pnActiveRec - 1, "sModelNme")
      txtOthers(3) = IFNull(oTrans.Detail(pnActiveRec - 1, "sColorNme"))
      txtOthers(9) = IFNull(oTrans.Detail(pnActiveRec - 1, "sNotesxxx"))
      txtOthers(5) = oTrans.Detail(pnActiveRec - 1, "nQuantity")
      txtOthers(16) = oTrans.Detail(pnActiveRec - 1, "nApproved")
      txtOthers(14) = oTrans.Detail(pnActiveRec - 1, "nMinLevel")
      txtOthers(15) = oTrans.Detail(pnActiveRec - 1, "nMaxLevel")
      txtOthers(6) = oTrans.Detail(pnActiveRec - 1, "cClassify")
      txtOthers(8) = oTrans.Detail(pnActiveRec - 1, "nQtyOnHnd")
      txtOthers(11) = IFNull(oTrans.Detail(pnActiveRec - 1, "nOnTrnsit"), 0)
      txtOthers(13) = oTrans.Detail(pnActiveRec - 1, "nAveMonSl")
      txtOthers(10) = oTrans.Detail(pnActiveRec - 1, "nResvOrdr")
      txtOthers(12) = oTrans.Detail(pnActiveRec - 1, "nBackOrdr")
      txtOthers(7) = oTrans.Detail(pnActiveRec - 1, "nRecOrder")
      
      If xrFrame1.Enabled Then txtOthers(16).SetFocus
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )", True
End Sub

Private Function loadCluster() As Boolean
   Dim lors As Recordset
   Dim lsProcName As String
   Dim lnCtr As Integer

   lsProcName = "loadCluster"
   'On Error GoTo errProc

   Set lors = New Recordset
   Set lors = oTrans.LoadClusterRequest(oTrans.ClusterID)

   If lors Is Nothing Then
      Call InitFields
      txtSearch(0).SetFocus
      GoTo endProc
   End If

   With MSFlexGrid2
      If lors.RecordCount = 0 Then
         .Rows = 2
      Else
         .Rows = lors.RecordCount + 1
      End If
      lnCtr = 1
      Do Until lors.EOF
         .TextMatrix(lnCtr, 0) = lnCtr
         .TextMatrix(lnCtr, 1) = lors("sBranchNm")
         .TextMatrix(lnCtr, 2) = Format(lors("sTransNox"), "@@@@-@@-@@@@@@")
         .TextMatrix(lnCtr, 3) = Format(IFNull(lors("nQuantity"), 0), "#,##0")
         .TextMatrix(lnCtr, 4) = Format(IFNull(lors("nApproved"), 0), "#,##0")
         
         lnCtr = lnCtr + 1
         lors.MoveNext
      Loop

      If .Rows > 16 Then
         .ColWidth(1) = 2500
      Else
         .ColWidth(1) = 2750
      End If

      .Row = IIf(pnLastRow >= .Rows, .Rows - 1, pnLastRow)
      .Col = 1
      .ColSel = .Cols - 1
   End With

   loadCluster = True
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )", True
End Function

Private Function loadSelTrans() As Boolean
   With MSFlexGrid2
      If .TextMatrix(.Row, 2) = "" Then
         Call oTrans.InitTransaction
         
         Call ClearDetail
         Call clearMaster
      ElseIf oTrans.LoadRequest(Replace(.TextMatrix(.Row, 2), "-", "")) Then
         Call ClearDetail
         Call LoadDetail
         Call LoadMaster
      Else
         Call ClearDetail
         Call clearMaster
      End If
   End With
End Function

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With oApp
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
'      If bEnd Then
'         .xShowError
'         End
'      Else
'         With Err
'            '.Raise .Number, .Source, .Description
'         End With
'      End If
   End With
End Sub

Private Sub MSFlexGrid1_Click()
   With MSFlexGrid1
      If pnActiveRec <> .Row Then
         If .Row > 0 Then
            pnActiveRec = .Row
            Call setFieldInfo
         End If
      End If
   End With
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'   With MSFlexGrid1
'      Select Case KeyCode
'      Case vbKeyReturn, vbKeyUp, vbKeyDown
'         Select Case KeyCode
'         Case vbKeyReturn
'            Call txtField_Validate(pnIndex, False)
'            If pnActiveRec < .Rows - 1 Then
'               pnActiveRec = pnActiveRec + 1
'
'               .Row = pnActiveRec
'               .Col = 1
'               .ColSel = .Cols - 1
'               If .Row > 18 Then .TopRow = .TopRow + 1
'            End If
'
'            Call setFieldInfo
'            With txtField(pnIndex)
'               .SelStart = 0
'               .SelLength = Len(.Text)
'               .SetFocus
'            End With

End Sub

Private Sub MSFlexGrid1_RowColChange()
   If pbLoaded Then
      With MSFlexGrid1
         If pnActiveRec <> .Row Then
            If .Row > 0 Then
               pnActiveRec = .Row
               Call setFieldInfo
            End If
         End If
         
         .Col = 1
         .ColSel = .Cols - 1
      End With
   End If
End Sub

Private Sub MSFlexGrid2_DblClick()
   With MSFlexGrid2
      If .Row = 0 Then
         Exit Sub
      Else
         If .TextMatrix(.Row, 2) = "" Then Exit Sub
         If pbModified Then
            If txtField(0).Text <> .TextMatrix(.Row, 2) Then
               If MsgBox("Record was Modified!" & vbCrLf & _
                        "Selecting another transaction will disregard any changes made!" & vbCrLf & vbCrLf & _
                        "Do you want to continue?", vbCritical + vbYesNo, "Confirm") <> vbYes Then
                  Exit Sub
               End If
            End If
         End If
         
         Call loadSelTrans
         If txtField(0) <> "" Then
            If MsgBox("Do you want to print order cart?", vbQuestion + vbYesNo) = vbYes Then PrintOrders
         End If
      End If
   End With
End Sub

Private Sub MSFlexGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      Call MSFlexGrid2_DblClick
   End If
End Sub

Private Sub txtOthers_GotFocus(Index As Integer)
   With txtOthers(Index)
      Select Case Index
      Case 9, 16
         .SelStart = 0
         .SelLength = Len(.Text)
         .BackColor = oApp.getColor("HT1")
      End Select

      pnIndex = Index
   End With
End Sub

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lbCancel As Boolean
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      With MSFlexGrid1
         Select Case KeyCode
         Case vbKeyReturn
            If pnIndex = 16 Then
               Call txtOthers_Validate(pnIndex, lbCancel)
               
               If lbCancel Then Exit Sub
               If pnActiveRec < .Rows - 1 Then
                  pnActiveRec = pnActiveRec + 1
                  .Row = pnActiveRec
                  Call setFieldInfo
                  
                  If .Row > 16 Then .TopRow = .TopRow + 1
                  
                  With txtOthers(pnIndex)
                     .SelStart = 0
                     .SelLength = Len(.Text)
                     .SetFocus
                  End With
               End If
            Else
               SetNextFocus
            End If
         Case vbKeyDown
            If pbCtrlPress Then
               If pnActiveRec < .Rows - 1 Then
                  ' this does not trigger lost focus or validate
                  If pnIndex = 16 Then
                     Call txtOthers_Validate(pnIndex, False)
                  End If
                  pnActiveRec = pnActiveRec + 1

                  .Row = pnActiveRec
                  .Col = 1
                  .ColSel = .Cols - 1
                  If .Row > 16 Then .TopRow = .TopRow + 1

                  Call setFieldInfo
                  With txtOthers(pnIndex)
                     .SelStart = 0
                     .SelLength = Len(.Text)
                     .SetFocus
                  End With
               End If
            Else
               SetNextFocus
            End If
         Case vbKeyUp
            If pbCtrlPress Then
               If .Row > 1 Then
                  ' this does not trigger lost focus or validate
                  If pnIndex = 11 Then
                     Call txtOthers_Validate(pnIndex, False)
                  End If

                  If .Row = .TopRow Then .TopRow = .TopRow - 1

                  pnActiveRec = pnActiveRec - 1

                  .Row = pnActiveRec
                  .ColSel = .Cols - 1

                  Call setFieldInfo
                  With txtOthers(pnIndex)
                     .SelStart = 0
                     .SelLength = Len(.Text)
                     .SetFocus
                  End With
               End If
            Else
               SetPreviousFocus
            End If
         End Select
      End With
   Case vbKeyControl
      pbCtrlPress = True
      KeyCode = 0
   End Select
End Sub

Private Sub txtOthers_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If pbCtrlPress Then
      If KeyCode = vbKeyControl Then pbCtrlPress = False
   End If
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   With txtOthers(Index)
      .BackColor = oApp.getColor("EB")

      If Index = 4 Then
         Call txtOthers_Validate(Index, False)
      End If
   End With
   pnIndex = 0
End Sub

Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
   Dim lsTransNox As String
   Dim lnCtr As Integer

   'Debug.Print "txtField_Validate"

   With txtOthers(Index)
      Select Case Index
      Case 9
         oTrans.Detail(pnActiveRec - 1, "sIssNotes") = .Text
      Case 16
         If Not IsNumeric(.Text) Then
            .Text = 0
         End If
            
         'Mac 2018.05.31
         'approved order can be greater than the request
         'as requested by ate she
'         If CLng(.Text) > oTrans.Detail(pnActiveRec - 1, "nQuantity") Then
'            Beep
'            .Text = oTrans.Detail(pnActiveRec - 1, "nQuantity")
'            pbModified = True
'         Else
            If CLng(.Text) > 0 Then
               pbModified = True
            End If
'         End If

         ' assign value to detail and to the object
         MSFlexGrid1.TextMatrix(pnActiveRec, 4) = txtOthers(Index)
         oTrans.Detail(pnActiveRec - 1, "nApproved") = .Text
      End Select
   End With
End Sub

Public Function PrintTrans() As Boolean
   Dim lrs As New ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lor As ADODB.Recordset
   
   lsOldProc = "PrintTrans"
   ''On Error GoTo errProc

   PrintTrans = True
   
   Set lrs = New ADODB.Recordset
   
   lrs.Fields.Append "sField01", adVarChar, 60
   lrs.Fields.Append "sField02", adVarChar, 60
   lrs.Fields.Append "sField03", adVarChar, 60
   lrs.Fields.Append "sField04", adVarChar, 60
   lrs.Fields.Append "nField01", adInteger, 5
   lrs.Fields.Append "nField02", adInteger, 5
   lrs.Fields.Append "nField03", adInteger, 5
   lrs.Fields.Append "nField04", adInteger, 5
   lrs.Open
      

  With MSFlexGrid1
      For lnCtr = 0 To oTrans.ItemCount - 1
      
         lrs.AddNew
         lrs("sField01").Value = oTrans.Detail(lnCtr, "sBrandNme")
         lrs("sField02").Value = oTrans.Detail(lnCtr, "sModelCde")
         lrs("sField03").Value = oTrans.Detail(lnCtr, "sModelNme")
         lrs("sField04").Value = oTrans.Detail(lnCtr, "sColorNme")
         lrs("nField01").Value = oTrans.Detail(lnCtr, "nQuantity")
         lrs("nField02").Value = oTrans.Detail(lnCtr, "nApproved")
         lrs("nField03").Value = oTrans.Detail(lnCtr, "nCancelld")
         lrs("nField04").Value = oTrans.Detail(lnCtr, "nIssueQty")
      Next
   End With
 
   
   Set lor = New ADODB.Recordset
   If lor.State = adStateOpen Then lor.Close
   
   lor.Open "SELECT" _
               & "  CONCAT(a.sAddressx, ', ', b.sTownName, ', ', c.sProvName, ' ', b.sZippCode) as Address" _
               & ", d.sCompnyNm " _
            & " From Branch a" _
               & ", TownCity b" _
               & ", Province c" _
               & ", Company d" _
            & " WHERE a.sBranchCd = " & strParm(oApp.BranchCode) _
               & " AND a.sTownIDxx = b.sTownIDxx" _
               & " AND b.sProvIDxx = c.sProvIDxx" _
               & " AND a.sCompnyID = d.sCompnyID" _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   ' assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\MCIssuance.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   oReport.Sections("RH").ReportObjects("txtRefNo").SetText (oTrans.Master("sTransNox"))
   oReport.Sections("RH").ReportObjects("txtDate").SetText txtField(1).Text
   oReport.Sections("RH").ReportObjects("txtBranch").SetText lor("sCompnyNm")
   oReport.Sections("RH").ReportObjects("txtAddress").SetText lor("Address")
   oReport.Sections("PH").ReportObjects("txtReqBranch").SetText txtField(10).Text
   'oReport.Sections("RF").ReportObjects("txtRemaks").SetText txtField(5).Text
   oReport.Sections("PF").ReportObjects("txtRptUser").SetText oApp.UserName
   
   oReport.PrintOutEx False, 1
   lrs.Close

endProc:

   Set oReport = Nothing
   Exit Function
errProc:
   PrintTrans = False
   ShowError lsOldProc & "( " & " )"
End Function

Private Sub txtSearch_GotFocus(Index As Integer)
   With txtSearch(Index)
      If Index = 0 Then
         .SelStart = 0
         .SelLength = Len(.Text)
      End If
      .BackColor = oApp.getColor("HT1")
   End With
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   With txtSearch(Index)
      Select Case KeyCode
      Case vbKeyReturn, vbKeyUp, vbKeyDown, vbKeyF3
         If KeyCode = vbKeyDown Then
            SetPreviousFocus
         Else
            SetNextFocus
         End If
         
         If Index = 0 Then
            If oTrans.getCluster(.Text) Then
               txtSearch(Index) = oTrans.Cluster
               txtSearch(1) = Format(oTrans.DeliverySched, "Mmmm dd, yyyy")
               
               Call loadCluster
               Call LoadDetail
            Else
               Call clearCluster
               Call ClearDetail
            End If
         End If
      End Select
   End With
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
   With txtSearch(Index)
      .BackColor = oApp.getColor("EB")
   End With
End Sub

'Private Sub txtSearch_Validate(Index As Integer, Cancel As Boolean)
'   With txtSearch(Index)
'      If Index = 0 Then
'         If oTrans.getcluster(.Text) Then
'            txtSearch(Index) = oTrans.Cluster
'            txtSearch(1) = Format(oTrans.DeliverySched, "Mmmm dd, yyyy")
'
'            Call loadCluster
'            Call LoadDetail
'         Else
'            Call clearCluster
'            Call clearDetail
'         End If
'      End If
'   End With
'End Sub

'
'Private Function PrintTrans() As Boolean
'   Dim loRS As Recordset
'   Dim lsSQL As String
'   Dim lsLineStr As String
'
'   lsSQL = "SELECT c.sTransNox" & _
'               ", a.sBarrCode" & _
'               ", a.sDescript" & _
'               ", e.sModelNme" & _
'               ", IFNULL(i.cClassify, 'F') cClassify" & _
'               ", IFNULL(i.nAveMonSl, 0) nAveMonSl" & _
'               ", a.nSelPrice" & _
'               ", IFNULL(i.nQtyOnHnd, 0) nQtyOnHnd" & _
'               ", b.nQuantity" & _
'               ", b.nRecOrder" & _
'               ", g.sSectnNme" & _
'               ", h.sBinNamex" & _
'               ", c.sRemarksx" & _
'               ", c.dTransact" & _
'               ", z.sBranchNm" & _
'            " FROM Spareparts a" & _
'               " LEFT JOIN SP_Model e" & _
'                  " ON a.sModelIDx = e.sModelIDx" & _
'               " LEFT JOIN SP_Inventory f" & _
'                  " LEFT JOIN Section g" & _
'                     " ON f.sSectnIDx = g.sSectnIDx" & _
'                  " LEFT JOIN Bin h" & _
'                     " ON f.sBinIDxxx = h.sBinIDxxx" & _
'                  " ON a.sPartsIDx = f.sPartsIDx" & _
'                     " AND f.sBranchCd = " & strParm(oApp.BranchCode)
'   lsSQL = lsSQL & _
'               ", SP_Stock_Request_Detail b" & _
'                  " LEFT JOIN SP_Inventory i" & _
'                     " ON b.sPartsIDx = i.sPartsIDx" & _
'                        " AND b.sTransNox LIKE CONCAT(i.sBranchCd, '%')" & _
'               ", SP_Stock_Request_Master c" & _
'               ", Branch z" & _
'            " WHERE a.sPartsIDx = b.sPartsIDx" & _
'               " AND b.sTransNox = c.sTransNox" & _
'               " AND c.sTransNox LIKE " & strParm(oTrans.RequestingBranch & "%") & _
'               " AND c.cTranStat = " & strParm(xeStateOpen) & _
'               " AND LEFT(c.sTransNox, 4) = z.sBranchCd" & _
'               IIf(txtField(1).Text = "", "", _
'                  " AND c.sTransNox = " & strParm(txtPrefix & Replace(txtField(1), "-", ""))) & _
'            " ORDER BY b.sTransNox" & _
'               ", g.sSectnNme" & _
'               ", h.sBinNamex" & _
'               ", a.sBarrCode"
'
'   Debug.Print lsSQL
'   Set loRS = New Recordset
'   loRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
'
'   If loRS.EOF Then GoTo endProc
'
'   If poPrinter Is Nothing Then
'      Set poPrinter = New clsPrintDirect
'      With poPrinter
'         .FontName = "Draft 17cpi"
'         .FontSize = 10
'
'         If Not .BegPrint() Then GoTo endProc
'      End With
'   End If
'
'   With poPrinter
'      lsSQL = ""
'      Do Until loRS.EOF
'         If pnPrintRow = pxeMaxLine Then
'            lsLineStr = padRight("Part Number", 20) & " " & _
'                              padRight("Description", 45) & " " & _
'                              padRight("Model", 25) & " " & _
'                              padRight("Cls", 3) & " " & _
'                              padRight("AMC", 5) & " " & _
'                              padRight("QOH", 5) & " " & _
'                              padRight("Qty", 6) & " " & _
'                              padRight("ROQ", 6) & " " & _
'                              padRight("Sect", 6) & " " & _
'                              padRight("Bin", 5)
'            .PrintText 0, 2, lsLineStr
'            pnPrintRow = 1
'         End If
'
'         If lsSQL <> loRS("sTransNox") Then
'
'            lsLineStr = padRight(Left(loRS("sBranchNm"), 30), 30) & " " & _
'                           padRight(Left(IFNull(loRS("sRemarksx"), ""), 45), 45) & " " & _
'                           Format(loRS("sTransNox"), "@@@@-@@-@@@@@@") & "   " & _
'                           Format(loRS("dTransact"), "MMM DD, YYYY")
'            lsSQL = loRS("sTransNox")
'
''            lsLineStr = padRight(Left(txtField(0), 30), 30) & " " & _
''                           padRight(Left(IFNull(loRS("sRemarksx"), ""), 45), 45) & " " & _
''                           Format(loRS("sTransNox"), "@@@@-@@-@@@@@@") & "   " & _
''                           Format(loRS("dTransact"), "MMM DD, YYYY")
''            lsSQL = loRS("sTransNox")
'
'            .PrintText pnPrintRow, 2, lsLineStr
'            pnPrintRow = pnPrintRow + 1
'         End If
'         lsLineStr = padRight(Trim(loRS("sBarrCode")), 20) & " " & _
'                        padRight(Trim(loRS("sDescript")), 45) & " " & _
'                        padRight(Trim(IFNull(loRS("sModelNme"), "")), 25) & " " & _
'                        padRight(loRS("cClassify"), 3) & " " & _
'                        padRight(Format(loRS("nAveMonSl"), "#0"), 5) & " " & _
'                        padRight(Format(loRS("nQtyOnHnd"), "#0"), 5) & " " & _
'                        padRight(Format(loRS("nQuantity"), "#0"), 6) & " " & _
'                        padRight(Format(loRS("nRecOrder"), "#0"), 6) & " " & _
'                        padRight(Trim(Left(IFNull(loRS("sSectnNme"), ""), 6)), 6) & " " & _
'                        padRight(Trim(Left(IFNull(loRS("sBinNamex"), ""), 5)), 5)
'
'         .PrintText pnPrintRow, 2, lsLineStr
'         pnPrintRow = pnPrintRow + 1
'
'         loRS.MoveNext
'      Loop
'
'      .EndPrint
'   End With
'
'   PrintTrans = True
'
'endProc:
'   Exit Function
'End Function
'
Private Function PrintOrders() As Boolean
   Dim lors As Recordset
   Dim lsSQL As String
   Dim lsLineStr As String
   Dim lsTransNox As String
  
   lsSQL = "SELECT g.sBranchNm" & _
               ", a.sTransNox" & _
               ", a.dTransact" & _
               ", a.sRemarksX" & _
               ", e.sBrandNme" & _
               ", d.sModelNme" & _
               ", d.sModelCde" & _
               ", f.sColorNme" & _
               ", b.nQtyOnHnd" & _
               ", b.nQuantity" & _
            " FROM CP_Stock_Request_Master a" & _
            ", CP_Stock_Request_Detail b" & _
            ", CP_Inventory c" & _
                  " LEFT JOIN CP_Model d" & _
                     " ON c.sModelIDx = d.sModelIDx" & _
                  " LEFT JOIN CP_Brand e" & _
                     " ON c.sBrandIDx = e.sBrandIDx" & _
                  " LEFT JOIN Color f" & _
                     " ON c.sColorIDx = f.sColorIDx" & _
            ", Branch g" & _
            " WHERE a.sTransNox = b.sTransNox" & _
            " AND b.sStockIDx = c.sStockIDx" & _
            " AND LEFT(a.sTransNox,4) = g.sBranchCd" & _
            " AND a.sTransNox = " & strParm(txtField(0)) & _
            " ORDER BY e.sBrandNme, d.sModelCde,d.sModelNme,f.sColorNme"
   
   Set lors = New Recordset
   lors.Open lsSQL, oApp.Connection, , , adCmdText
   
   If lors.EOF Then GoTo endProc
   
   Dim loPrinter As clsPrintDirect
   Set loPrinter = New clsPrintDirect
   With loPrinter
      .FontName = "Draft 20cpi"
      .FontSize = 10
      
      If Not .BegPrint() Then GoTo endProc
      
      If pnPrintRow = 65 Or pnPrintRow = 0 Then
         If lsTransNox <> lors("sTransNox") Then
            lsLineStr = padRight("Brand", 15) & " " & _
                                 padRight("Code", 25) & " " & _
                                 padRight("Model", 25) & " " & _
                                 padRight("Color", 15) & " " & _
                                 padRight("QOH", 7) & " " & _
                                 padRight("Order", 7)
   
               .PrintText pnPrintRow, 2, lsLineStr
               pnPrintRow = pnPrintRow + 1
   
               lsLineStr = padRight(lors("sBranchNm"), 20) & " " & _
                              Format(lors("sTransNox"), "@@@@-@@-@@@@@@") & "   " & _
                              Format(lors("dTransact"), "MMM DD, YYYY")
               .PrintText pnPrintRow, 2, lsLineStr
               pnPrintRow = pnPrintRow + 2
         End If
      Else
         If lsTransNox <> lors("sTransNox") Then
               lsLineStr = padRight("Brand", 15) & " " & _
                                 padRight("Code", 25) & " " & _
                                 padRight("Model", 25) & " " & _
                                 padRight("Color", 15) & " " & _
                                 padRight("QOH", 7) & " " & _
                                 padRight("Order", 7)
   
               .PrintText pnPrintRow, 2, lsLineStr
               pnPrintRow = pnPrintRow + 1
   
               lsLineStr = padRight(lors("sBranchNm"), 20) & " " & _
                              Format(lors("sTransNox"), "@@@@-@@-@@@@@@") & "   " & _
                              Format(lors("dTransact"), "MMM DD, YYYY")
               .PrintText pnPrintRow, 2, lsLineStr
               pnPrintRow = pnPrintRow + 2
         End If
      End If
      
      Do Until lors.EOF
          lsLineStr = Left(padRight(Trim(lors("sBrandNme") + "_______________"), 15), 15) & _
                     Left(padRight(Trim(lors("sModelCde") + "____________________"), 25), 25) & _
                     Left(padRight(Trim(lors("sModelNme") + "____________________"), 25), 25) & _
                     Left(padRight(Trim(lors("sColorNme") + "_______________"), 15), 15) & _
                     Left(padRight(Trim(Format(lors("nQtyOnHnd"), "#0") + "_______"), 7), 7) & _
                     padRight(Trim(Format(lors("nQuantity"), "#0")), 7)
            
            .PrintText pnPrintRow, 2, lsLineStr
            pnPrintRow = pnPrintRow + 1
            
            lsTransNox = lors("sTransNox")
      lors.MoveNext
      Loop
      .EndPrint
   End With
   PrintOrders = True
endProc:
   Exit Function
End Function


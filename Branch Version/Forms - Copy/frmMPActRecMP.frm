VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmMPActRecMP 
   BorderStyle     =   0  'None
   Caption         =   "Accounts Receivable MP"
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3840
      Left            =   1650
      TabIndex        =   24
      Tag             =   "wt0;fb0"
      Top             =   3885
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   6773
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Purchase Details"
      TabPicture(0)   =   "frmMPActRecMP.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "xrFrame2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Payment &Details"
      TabPicture(1)   =   "frmMPActRecMP.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "xrFrame2(1)"
      Tab(1).ControlCount=   1
      Begin xrControl.xrFrame xrFrame2 
         Height          =   3390
         Index           =   1
         Left            =   -74925
         Tag             =   "wt0;fb0"
         Top             =   375
         Width           =   9510
         _ExtentX        =   16775
         _ExtentY        =   5980
         BackColor       =   12632256
         BorderStyle     =   4
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   31
            Left            =   4410
            MaxLength       =   50
            TabIndex        =   78
            Text            =   "0.00"
            Top             =   2850
            Width           =   2025
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   27
            Left            =   7845
            MaxLength       =   50
            TabIndex        =   84
            Text            =   "0.00"
            Top             =   735
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   18
            Left            =   4935
            MaxLength       =   50
            TabIndex        =   64
            Text            =   "0.00"
            Top             =   135
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   33
            Left            =   7845
            MaxLength       =   50
            TabIndex        =   80
            Text            =   "0.00"
            Top             =   135
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   14
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   46
            Top             =   135
            Width           =   2025
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   24
            Left            =   4935
            MaxLength       =   50
            TabIndex        =   76
            Text            =   "0.00"
            Top             =   2070
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   23
            Left            =   4935
            MaxLength       =   50
            TabIndex        =   74
            Text            =   "0.00"
            Top             =   1770
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   22
            Left            =   4935
            MaxLength       =   50
            TabIndex        =   72
            Text            =   "0.00"
            Top             =   1470
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   21
            Left            =   4935
            MaxLength       =   50
            TabIndex        =   70
            Text            =   "0.00"
            Top             =   1170
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   20
            Left            =   4935
            MaxLength       =   50
            TabIndex        =   68
            Text            =   "0.00"
            Top             =   735
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   19
            Left            =   4935
            MaxLength       =   50
            TabIndex        =   66
            Text            =   "0.00"
            Top             =   435
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   25
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   60
            Text            =   "0.00"
            Top             =   2730
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   26
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   58
            Top             =   2430
            Width           =   2025
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   28
            Left            =   7845
            MaxLength       =   50
            TabIndex        =   86
            Text            =   "0.00"
            Top             =   1170
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   29
            Left            =   7845
            MaxLength       =   50
            TabIndex        =   88
            Text            =   "0.00"
            Top             =   1470
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   30
            Left            =   7845
            MaxLength       =   50
            TabIndex        =   90
            Text            =   "0.00"
            Top             =   1770
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   17
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   52
            Top             =   1035
            Width           =   2025
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   16
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   50
            Top             =   735
            Width           =   2025
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   15
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   48
            Top             =   435
            Width           =   2025
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   32
            Left            =   7380
            MaxLength       =   50
            TabIndex        =   92
            Text            =   "0.00"
            Top             =   2850
            Width           =   2025
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   34
            Left            =   7845
            MaxLength       =   50
            TabIndex        =   82
            Text            =   "0.00"
            Top             =   435
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   35
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   62
            Text            =   "0.00"
            Top             =   3030
            Width           =   1545
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   36
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   54
            Top             =   1755
            Width           =   2025
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   37
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   56
            Top             =   2070
            Width           =   2025
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Purchase"
            Height          =   195
            Index           =   19
            Left            =   135
            TabIndex        =   45
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rebates"
            Height          =   195
            Index           =   29
            Left            =   3735
            TabIndex        =   75
            Top             =   2085
            Width           =   600
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Penalty"
            Height          =   195
            Index           =   28
            Left            =   3735
            TabIndex        =   73
            Top             =   1815
            Width           =   525
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly Amort."
            Height          =   195
            Index           =   27
            Left            =   3735
            TabIndex        =   71
            Top             =   1515
            Width           =   1050
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PN Value"
            Height          =   195
            Index           =   26
            Left            =   3735
            TabIndex        =   69
            Top             =   1215
            Width           =   675
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Balance"
            Height          =   195
            Index           =   25
            Left            =   3735
            TabIndex        =   67
            Top             =   780
            Width           =   990
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Down Payment"
            Height          =   195
            Index           =   24
            Left            =   3735
            TabIndex        =   65
            Top             =   480
            Width           =   1080
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Price"
            Height          =   195
            Index           =   23
            Left            =   3735
            TabIndex        =   63
            Top             =   180
            Width           =   810
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Payment"
            Height          =   195
            Index           =   30
            Left            =   135
            TabIndex        =   59
            Top             =   2745
            Width           =   960
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Payment Date"
            Height          =   195
            Index           =   31
            Left            =   135
            TabIndex        =   57
            Top             =   2460
            Width           =   1350
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Total"
            Height          =   195
            Index           =   32
            Left            =   6690
            TabIndex        =   83
            Top             =   765
            Width           =   1020
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rebate Total"
            Height          =   195
            Index           =   34
            Left            =   6690
            TabIndex        =   85
            Top             =   1215
            Width           =   930
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Debit Total"
            Height          =   195
            Index           =   35
            Left            =   6690
            TabIndex        =   87
            Top             =   1515
            Width           =   780
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Total"
            Height          =   195
            Index           =   36
            Left            =   6690
            TabIndex        =   89
            Top             =   1785
            Width           =   810
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Due Date"
            Height          =   195
            Index           =   22
            Left            =   135
            TabIndex        =   51
            Top             =   1080
            Width           =   690
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Term"
            Height          =   195
            Index           =   21
            Left            =   135
            TabIndex        =   49
            Top             =   780
            Width           =   1005
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Pay Date"
            Height          =   195
            Index           =   20
            Left            =   135
            TabIndex        =   47
            Top             =   480
            Width           =   990
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount Due"
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
            Index           =   37
            Left            =   5385
            TabIndex        =   77
            Top             =   2655
            Width           =   1050
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Balance"
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
            Index           =   38
            Left            =   7935
            TabIndex        =   91
            Top             =   2655
            Width           =   1470
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Down Total"
            Height          =   195
            Index           =   39
            Left            =   6690
            TabIndex        =   79
            Top             =   180
            Width           =   825
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Total"
            Height          =   195
            Index           =   40
            Left            =   6690
            TabIndex        =   81
            Top             =   480
            Width           =   765
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delay Average"
            Height          =   195
            Index           =   41
            Left            =   135
            TabIndex        =   61
            Top             =   3045
            Width           =   1050
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rating"
            Height          =   195
            Index           =   42
            Left            =   135
            TabIndex        =   53
            Top             =   1815
            Width           =   465
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Former Account #"
            Height          =   195
            Index           =   45
            Left            =   135
            TabIndex        =   55
            Top             =   2085
            Width           =   1275
         End
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   3405
         Index           =   0
         Left            =   75
         Tag             =   "wt0;fb0"
         Top             =   360
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   6006
         BackColor       =   12632256
         BorderStyle     =   4
         Begin VB.ComboBox cmbField 
            Height          =   315
            ItemData        =   "frmMPActRecMP.frx":0038
            Left            =   1155
            List            =   "frmMPActRecMP.frx":004E
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   270
            Width           =   2025
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   11
            Left            =   5790
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   40
            Top             =   600
            Width           =   3630
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   5790
            MaxLength       =   50
            TabIndex        =   44
            Top             =   1200
            Width           =   3630
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   5790
            MaxLength       =   50
            TabIndex        =   42
            Top             =   900
            Width           =   3630
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   10
            Left            =   1155
            MaxLength       =   50
            TabIndex        =   38
            Top             =   2715
            Width           =   3630
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   1155
            MaxLength       =   50
            TabIndex        =   36
            Top             =   2415
            Width           =   3630
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   1155
            MaxLength       =   50
            TabIndex        =   34
            Top             =   1515
            Width           =   3630
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   1155
            MaxLength       =   50
            TabIndex        =   32
            Top             =   1215
            Width           =   3630
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   1155
            MaxLength       =   50
            TabIndex        =   30
            Top             =   915
            Width           =   3630
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   1155
            MaxLength       =   50
            TabIndex        =   28
            Top             =   615
            Width           =   3630
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Loan Type"
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   25
            Top             =   315
            Width           =   765
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Col. Branch"
            Height          =   195
            Index           =   18
            Left            =   4935
            TabIndex        =   43
            Top             =   1245
            Width           =   825
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Manager"
            Height          =   195
            Index           =   16
            Left            =   4935
            TabIndex        =   41
            Top             =   945
            Width           =   630
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Collector"
            Height          =   195
            Index           =   15
            Left            =   4935
            TabIndex        =   39
            Top             =   645
            Width           =   615
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Route"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   37
            Top             =   2760
            Width           =   435
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Count"
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   35
            Top             =   2460
            Width           =   870
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Serial #"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   33
            Top             =   1560
            Width           =   540
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   31
            Top             =   1260
            Width           =   360
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   29
            Top             =   960
            Width           =   435
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Brand"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   27
            Top             =   660
            Width           =   420
         End
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   6780
      Index           =   0
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   1050
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11959
      BackColor       =   12632256
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   80
         Left            =   7650
         TabIndex        =   20
         Top             =   1260
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   81
         Left            =   7650
         TabIndex        =   22
         Top             =   1560
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   53
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1860
         Width           =   5220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   54
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   15
         Top             =   2160
         Width           =   5220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2460
         Width           =   8640
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   585
         Index           =   3
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1260
         Width           =   5220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   9
         Top             =   960
         Width           =   5220
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   7
         Top             =   660
         Width           =   2070
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1095
         MaxLength       =   50
         TabIndex        =   5
         Top             =   105
         Width           =   2070
      End
      Begin xrControl.xrFrame xrFrame2 
         Height          =   1185
         Index           =   2
         Left            =   8550
         Top             =   60
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   2090
         BackColor       =   12632256
         Begin VB.Image imgField 
            Height          =   1095
            Left            =   30
            Picture         =   "frmMPActRecMP.frx":00A4
            Stretch         =   -1  'True
            Top             =   30
            Width           =   1095
         End
      End
      Begin xrControl.xrButton cmdAddress 
         Height          =   285
         Left            =   8940
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1860
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         Caption         =   "UPDATE"
         AccessKey       =   "UPDATE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Map Coordinates"
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
         Index           =   6
         Left            =   6825
         TabIndex        =   18
         Top             =   1005
         Width           =   1440
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Latitude"
         Height          =   195
         Index           =   12
         Left            =   6900
         TabIndex        =   19
         Top             =   1305
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Longitude"
         Height          =   195
         Index           =   33
         Left            =   6900
         TabIndex        =   21
         Top             =   1605
         Width           =   705
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Co-Borwr #1"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   12
         Top             =   1905
         Width           =   885
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Co-Borwr #2"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   14
         Top             =   2205
         Width           =   885
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   5670
         Top             =   60
         Width           =   2730
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   5700
         Top             =   90
         Width           =   2670
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "UNKNOWN"
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
         Left            =   5730
         TabIndex        =   101
         Tag             =   "eb0;et0"
         Top             =   135
         Width           =   2610
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   7
         Left            =   105
         TabIndex        =   16
         Top             =   2505
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   195
         Index           =   5
         Left            =   105
         TabIndex        =   10
         Top             =   1290
         Width           =   570
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cust Name"
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   8
         Top             =   1035
         Width           =   780
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application #"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   705
         Width           =   930
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1185
         Tag             =   "et0;ht2"
         Top             =   240
         Width           =   2070
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account #"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   4
         Top             =   150
         Width           =   750
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000A&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   5730
         Tag             =   "et0;et0"
         Top             =   120
         Width           =   2625
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   450
      Index           =   1
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   794
      BackColor       =   12632256
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
         Height          =   285
         Index           =   38
         Left            =   1110
         TabIndex        =   1
         Top             =   45
         Width           =   2070
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
         Height          =   285
         Index           =   39
         Left            =   4380
         TabIndex        =   3
         Top             =   60
         Width           =   5355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Full Name"
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
         Left            =   3450
         TabIndex        =   2
         Top             =   105
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Account #"
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
         Index           =   19
         Left            =   105
         TabIndex        =   0
         Top             =   105
         Width           =   900
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   95
      Top             =   555
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
      Picture         =   "frmMPActRecMP.frx":4D38
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   96
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "S&earch"
      AccessKey       =   "e"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMPActRecMP.frx":54B2
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   100
      Top             =   3075
      Width           =   1260
      _ExtentX        =   2223
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
      Picture         =   "frmMPActRecMP.frx":5C2C
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   99
      Top             =   3075
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
      Picture         =   "frmMPActRecMP.frx":63A6
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   93
      Top             =   555
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
      Picture         =   "frmMPActRecMP.frx":6B20
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   94
      Top             =   1185
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
      Picture         =   "frmMPActRecMP.frx":729A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   6
      Left            =   105
      TabIndex        =   98
      Top             =   2445
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmMPActRecMP.frx":7A14
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   7
      Left            =   105
      TabIndex        =   97
      Top             =   1830
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "C&redInfo"
      AccessKey       =   "r"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmMPActRecMP.frx":818E
   End
End
Attribute VB_Name = "frmMPActRecMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmMCActRec"

Private WithEvents oTrans As ggcLoanReceivable.clsLRMasterMP
Attribute oTrans.VB_VarHelpID = -1
Private oFormLedger As frmMCARLedger
Private oFromCredInfom As frmCustCredInfo
Private oSkin As clsFormSkin
Private oRSMaster As ADODB.Recordset
Private oPriceList As clsCPPriceList

Dim pnCtr As Integer, pnTranStatus As Integer, pnIndex As Integer
Dim pbLoaded As Boolean, pbMoveCombo As Boolean
Dim psAcctNmbr As String
Dim psImgePath As String
Dim oForm As Object

Property Set FormMCActRec(Form As Object)
   Set oForm = Form
End Property

Property Let AccountNo(ByVal Value As String)
   psAcctNmbr = Value
End Property

Private Sub cmbField_Click()
   If cmbField.ListIndex = -1 Then oTrans.Master("cLoanType") = cmbField.ListIndex
End Sub

Private Sub cmbField_LostFocus()
   pbMoveCombo = False
End Sub

Private Sub cmdAddress_Click()
   Dim lsSQL As String

   If oTrans.Master("sClientID") = "" Then Exit Sub

   With cmdAddress
      If .Caption = "UPDATE" Then
         .Caption = "SAVE"
         txtField(80).Enabled = True
         txtField(81).Enabled = True
         txtField(80).SetFocus
      Else
         'iMac [09-29-15]
         '  save the coordinates on Client_Coordinates
         lsSQL = "INSERT INTO Client_Coordinates" & _
                  " SET sClientID = " & strParm(oTrans.Master("sClientID")) & _
                     ", nLatitude = " & IIf(txtField(80) = "", 0#, txtField(80)) & _
                     ", nLongitud = " & IIf(txtField(81) = "", 0#, txtField(81)) & _
                     ", sModified = " & strParm(Encrypt(oApp.UserID)) & _
                     ", dModified = " & dateParm(oApp.ServerDate) & _
                  " ON DUPLICATE KEY UPDATE" & _
                     "  nLatitude = " & IIf(txtField(80) = "", 0#, txtField(80)) & _
                     ", nLongitud = " & IIf(txtField(81) = "", 0#, txtField(81)) & _
                     ", sModified = " & strParm(Encrypt(oApp.UserID))

         Debug.Print lsSQL
         If oApp.Execute(lsSQL, "Client_Coordinates") = 0 Then
            MsgBox "Unable to Save Client Coordinates!!!", vbCritical, "Warning"
         End If

         .Caption = "UPDATE"
         txtField(80).Enabled = False
         txtField(81).Enabled = False
      End If
   End With
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
  ''On Error GoTo errProc

   Select Case Index
   Case 0
      If Not AllowRecSave(oApp, xeRecStateActive, xeModeUpdate, mdiMain.Controls(oApp.MenuName).Tag, oTrans.Master("dModified"), "") Then GoTo endProc
      If oTrans.SaveAccount Then
         MsgBox "Transaction Saved Successfully!!!", vbInformation, "Notice"
         initButton xeModeReady
         txtField(38).SetFocus
      Else
         MsgBox "Unable to Save Transaction!!!", vbCritical, "Warning"
      End If
   Case 1
      oTrans.SearchMaster pnIndex
      txtField(pnIndex).SetFocus
   Case 2
      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")

      If lnRep = vbYes Then
         If oTrans.OpenAccount(oTrans.Master(0)) Then
            LoadMaster
         Else
            InitValue
         End If
         initButton xeModeReady
         txtField(38).SetFocus
      Else
         txtField(pnIndex).SetFocus
      End If
   Case 3
      If oTrans.SearchAccount Then LoadMaster
      txtField(pnIndex).SetFocus
   Case 4
      If txtField(0).Text <> "" Then
         If oTrans.UpdateAccount Then
            initButton xeModeUpdate
            txtField(2).SetFocus
            SSTab1.Tab = 0
         Else
            MsgBox "Unable to Update Account!!!", vbCritical, "Warning"
         End If
      Else
         MsgBox "No Account to Update!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   Case 5
      Unload Me
   Case 6
      If txtField(0).Text <> "" Then
         oFormLedger.AccountNo = oTrans.Master("sAcctNmbr")
         Load oFormLedger
         If oFormLedger.browseLedger Then
            oFormLedger.Show 1
         Else
            MsgBox "No Ledger found!!!", vbCritical, "Warning"
         End If
         Unload oFormLedger
      Else
         MsgBox "Unable to Load Serial Ledger!!!" & vbCrLf & _
                "Please Verify your Entry then Try Again!!!", vbCritical, "Warning"
      End If
   Case 7
      If oTrans.Master("sApplicNo") <> "" Then
         oFromCredInfom.TransactionNo = oTrans.Master("sApplicNo")
         oFromCredInfom.Show 1
      End If


   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Property Let TranStatus(lnStatus As Integer)
   pnTranStatus = lnStatus
End Property

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0

   If Not pbLoaded Then
      initButton xeModeReady
      txtField(38).SetFocus
      pbLoaded = True
   End If

   'iMac [2015-12-21]
   '  load the account for coordinates update
   Call txtField_KeyDown(38, vbKeyReturn, 0)
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
  ''On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oTrans = New ggcLoanReceivable.clsLRMasterMP
   Set oTrans.AppDriver = oApp
   Set oFormLedger = New frmMCARLedger
   Set oFromCredInfom = New frmCustCredInfo

   oTrans.TransStatus = pnTranStatus
   oTrans.Active = True 'pnTranStatus = xeActStatActive

   oTrans.InitAccount
   oTrans.NewAccount

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = oForm
   oSkin.ApplySkin xeFormTransEqualLeft
      
   Set oPriceList = New clsCPPriceList
   Set oPriceList.AppDriver = oApp
   oPriceList.DateTransact = oApp.ServerDate
   oPriceList.InitTransaction


   InitValue
   Label2.Caption = Format(AccountStat(pnTranStatus), ">")

   txtField(38) = psAcctNmbr
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oTrans = Nothing
   Set oFormLedger = Nothing
   Set oFromCredInfom = Nothing
End Sub

Private Sub loadImage()
   If FileExists(psImgePath) Then
      imgField.Picture = LoadPicture(psImgePath)
   Else
      imgField.Picture = Nothing
      'Mac 2018-07-19
      '  TODO:
      '     Get the download the image from the server to the local pc.
   End If
End Sub

Function FileExists(ByVal sFileName As String) As Boolean
   Dim intReturn As Integer
   On Error GoTo FileExists_Error
   intReturn = GetAttr(sFileName)
   FileExists = True
Exit Function
FileExists_Error:
    FileExists = False
End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
   With SSTab1
      If .Tab = 0 Then
         If cmdButton(0).Visible Then txtField(10).SetFocus
      ElseIf .Tab = 1 Then
         If cmdButton(0).Visible Then txtField(14).SetFocus
      End If
   End With
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")

      Select Case Index
      Case 10
         SSTab1.Tab = 0
      Case 14, 15, 17, 26
         .Text = Format(.Text, "MM/DD/YYYY")
         If Index = 14 Then SSTab1.Tab = 1
      End Select
   End With
   pnIndex = Index
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

Private Sub oTrans_MasterRetrieved(ByVal Index As Integer)
   With txtField(Index)
      Select Case Index
      Case 14, 15, 17, 26
         .Text = Format(oTrans.Master(Index), "MMMM DD, YYYY")
      Case 18 To 25, 27 To 35
         .Text = Format(oTrans.Master(Index), "#,##0.00")
      Case Else
         .Text = oTrans.Master(Index)
      End Select
   End With
End Sub

Private Sub InitValue()
   Dim loTxt As TextBox

   For Each loTxt In txtField
      pnCtr = loTxt.Index
      Select Case pnCtr
      Case 18 To 25, 27 To 35
         txtField(pnCtr).Text = "0.00"
      Case 36
         txtField(pnCtr).Text = "UNKNOWN"
      Case Else
        txtField(pnCtr).Text = ""
        txtField(pnCtr).Tag = ""
      End Select
   Next
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String

   lsOldProc = "txtField_KeyDown"
  ''On Error GoTo errProc

   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         Select Case Index
         Case 38, 39
            Call txtField_Validate(Index, False)
         Case Else
            If KeyCode = vbKeyF3 Then
               oTrans.SearchMaster Index, .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oTrans.SearchMaster Index, .Text
            End If
         End Select
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

Private Sub LoadMaster()
   Dim loTxt As TextBox

   With oTrans
      For Each loTxt In txtField
         pnCtr = loTxt.Index
         Select Case pnCtr
         Case 0, 38
            txtField(pnCtr).Text = .Master("sAcctNmbr")
            txtField(pnCtr).Tag = txtField(pnCtr).Text
         Case 2, 39
            txtField(pnCtr).Text = .Master("xFullName")
            txtField(pnCtr).Tag = txtField(pnCtr).Text
         Case 1
            txtField(pnCtr).Text = Format(.Master(pnCtr), IIf(Len(oApp.BranchCode) = 2, "@@@@-@@@@@@", "@@@@@@-@@@@@@"))
         Case 5
            txtField(pnCtr).Text = .Master("sBrandNme")
         Case 6
            txtField(pnCtr).Text = .Master("sModelNme")
         Case 7
            txtField(pnCtr).Text = .Master("sColorNme")
         Case 8
            txtField(pnCtr).Text = .Master("sSerialNo")
         Case 14, 15, 17, 26
            txtField(pnCtr).Text = Format(.Master(pnCtr), "MMMM DD, YYYY")
         Case 18 To 25, 27 To 35
            txtField(pnCtr).Text = Format(.Master(pnCtr), "#,##0.00")
         Case 36
            txtField(pnCtr).Text = RatingStat(.Master("cRatingxx"))
         Case 37
            txtField(pnCtr).Text = RatingStat(.Master("sExAcctNo"))
         Case 53, 54
            txtField(pnCtr).Text = IIf(Trim(IFNull(.Master(pnCtr))) = "", "N-O-N-E", .Master(pnCtr))
         Case 3
            txtField(pnCtr).Text = .Master("xAddressx")
         Case 80
            txtField(80).Text = .Master("nLatitude")
         Case 81
            txtField(81).Text = .Master("nLongitud")
         Case Else
            txtField(pnCtr).Text = .Master(pnCtr)
         End Select
      Next

      'Mac 2018-07-19
      '  load the picture after loading the master info
      psImgePath = .GetImagePath
      Call loadImage

      cmbField.ListIndex = IIf(IsNull(.Master("cLoanType")), -1, .Master("cLoanType"))
   End With
   
   Call loadPriceList
End Sub

Private Sub loadPriceList()
   Dim lors As Recordset
   
   Set lors = New Recordset
   lors.Open "SELECT" & _
                  " b.sModelIDx" & _
               " FROM CP_Inventory_Serial a" & _
                  ", CP_Inventory b" & _
               " WHERE a.sStockIDx = b.sStockIDx" & _
                  " AND a.sSerialID = " & strParm(oTrans.Master("sSerialID")) _
   , oApp.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   oPriceList.DateTransact = oTrans.Master("dPurchase")
   oPriceList.ModelID = lors("sModelIDx")
   
   Select Case oTrans.Master("nAcctTerm")
   Case 3
      oPriceList.DownPayment(0) = oTrans.Master("nDownPaym")
   Case 6
      oPriceList.DownPayment(1) = oTrans.Master("nDownPaym")
   Case 9
      oPriceList.DownPayment(2) = oTrans.Master("nDownPaym")
   Case 12
      oPriceList.DownPayment(3) = oTrans.Master("nDownPaym")
   End Select
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(3).Visible = Not lbShow
   
   cmdButton(4).Visible = False 'Not lbShow
   
   If oApp.UserLevel >= xeEngineer Then
      cmdButton(4).Visible = Not lbShow
   End If

   cmdButton(5).Visible = Not lbShow
   xrFrame1(1).Enabled = Not lbShow

   cmdButton(0).Visible = lbShow
   cmdButton(1).Visible = lbShow
   cmdButton(2).Visible = lbShow

   txtField(2).Enabled = lbShow
   txtField(3).Enabled = lbShow
   txtField(4).Enabled = lbShow

   cmbField.Enabled = lbShow

   txtField(10).Enabled = lbShow
   txtField(14).Enabled = lbShow
   txtField(16).Enabled = lbShow
   txtField(19).Enabled = lbShow
'   txtField(20).Enabled = lbShow
'   txtField(21).Enabled = lbShow
   txtField(53).Enabled = lbShow
   txtField(54).Enabled = lbShow
End Sub

Private Sub txtField_KeyPress(Index As Integer, keyascii As Integer)
   Select Case Index
   Case 80, 81
      Select Case keyascii
        Case vbKey0 To vbKey9
        Case vbKeyBack, vbKeyClear, vbKeyDelete
        Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
        Case Else
          keyascii = 0
          Beep
      End Select
   End Select
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
      Case 18, 20 To 25, 27 To 35
         If Not IsNumeric(.Text) Then
            .Text = "0.00"
         Else
            .Text = Format(.Text, "#,##0.00")
         End If
         oTrans.Master(Index) = CDbl(.Text)
      Case 14
         If Not IsDate(.Text) Then .Text = oApp.ServerDate
         .Text = Format(.Text, "MMMM DD, YYYY")
         oTrans.Master(Index) = .Text
         oTrans.Master("dLastPaym") = oTrans.Master(Index)
         oPriceList.DateTransact = oTrans.Master(Index)
         
         Select Case Day(oTrans.Master(Index))
            Case 1 To 8
               If Month(DateAdd("m", 1, oTrans.Master(Index))) = 1 Then
                  oTrans.Master("dfirstpay") = Month(DateAdd("m", 1, oTrans.Master(Index))) & " 8 " & Year(DateAdd("yyyy", 1, oTrans.Master(Index)))
               Else
                  oTrans.Master("dfirstpay") = Month(DateAdd("m", 1, oTrans.Master(Index))) & " 8 " & Year(oTrans.Master(Index))
               End If
            Case 9 To 18
               If Month(DateAdd("m", 1, oTrans.Master(Index))) = 1 Then
                  oTrans.Master("dfirstpay") = Month(DateAdd("m", 1, oTrans.Master(Index))) & " 18 " & Year(DateAdd("yyyy", 1, oTrans.Master(Index)))
               Else
                  oTrans.Master("dfirstpay") = Month(DateAdd("m", 1, oTrans.Master(Index))) & " 18 " & Year(oTrans.Master(Index))
               End If
            Case 19 To 28
               If Month(DateAdd("m", 1, oTrans.Master(Index))) = 1 Then
                  oTrans.Master("dfirstpay") = Month(DateAdd("m", 1, oTrans.Master(Index))) & " 28 " & Year(DateAdd("yyyy", 1, oTrans.Master(Index)))
               Else
                  oTrans.Master("dfirstpay") = Month(DateAdd("m", 1, oTrans.Master(Index))) & " 28 " & Year(oTrans.Master(Index))
               End If
            Case 29 To 31
               If Month(oTrans.Master(Index)) = 11 Or Month(oTrans.Master(Index)) = 12 Then
                  oTrans.Master("dfirstpay") = Month(DateAdd("m", 2, oTrans.Master(Index))) & " 5 " & Year(DateAdd("yyyy", 1, oTrans.Master(Index)))
               Else
                  oTrans.Master("dfirstpay") = Month(DateAdd("m", 2, oTrans.Master(Index))) & " 5 " & Year(oTrans.Master(Index))
               End If
            End Select
         oTrans.Master("dDueDatex") = DateAdd("m", oTrans.Master("nAcctTerm") - 1, oTrans.Master("dFirstPay"))
         txtField(15) = Format(oTrans.Master("dFirstPay"), "MMMM DD, YYYY")
         txtField(17) = Format(oTrans.Master("dDueDatex"), "MMMM DD, YYYY")
         txtField(26) = Format(oTrans.Master("dLastPaym"), "MMMM DD, YYYY")
      Case 38, 39
         If .Text = "" Then
            InitValue
            Exit Sub
         End If

         If .Text <> .Tag Then
            If oTrans.SearchAccount(.Text, IIf(Index = 38, True, False)) Then
               LoadMaster

            End If
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
         End If
      Case 16, 19
         If Index = 16 Then
            If Not IsNumeric(.Text) Then .Text = "3"
            oTrans.Master("nAcctTerm") = CDbl(.Text)
         Else
            If Not IsNumeric(.Text) Then
               Select Case oTrans.Master("nAcctTerm")
               Case 3
                  .Text = Format(oPriceList.MinimumDown(0), "#,##0.00")
               Case 6
                  .Text = Format(oPriceList.MinimumDown(1), "#,##0.00")
               Case 9
                  .Text = Format(oPriceList.MinimumDown(2), "#,##0.00")
               Case 12
                  .Text = Format(oPriceList.MinimumDown(3), "#,##0.00")
               End Select
            End If
            
            oTrans.Master("nDownPaym") = CDbl(.Text)
            txtField(Index) = Format(.Text, "#,##0.00")
         End If
         
         oTrans.Master("nMonAmort") = oPriceList.getMonthly(oTrans.Master("nDownPaym"), oTrans.Master("nAcctTerm"), 0, 0, 0)
         oTrans.Master("dDueDatex") = DateAdd("m", oTrans.Master("nAcctTerm") - 1, oTrans.Master("dFirstPay"))
         oTrans.Master("nPNValuex") = CDbl(txtField(22)) * oTrans.Master("nAcctTerm")
         oTrans.Master("nGrossPrc") = CDbl(txtField(21)) + oTrans.Master("nDownPaym")
         
         txtField(17) = Format(oTrans.Master("dDueDatex"), "MMMM DD, YYYY")
         txtField(22) = Format(oTrans.Master("nMonAmort"), "#,##0.00")
         txtField(21) = Format(oTrans.Master("nPNValuex"), "#,##0.00")
         txtField(18) = Format(oTrans.Master("nGrossPrc"), "#,##0.00")
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

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPClusterDeliveryReg 
   BorderStyle     =   0  'None
   Caption         =   "Cellphone Delivery by Cluster"
   ClientHeight    =   8445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14490
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   14490
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   540
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   12810
      _ExtentX        =   22595
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
         Index           =   2
         Left            =   9825
         TabIndex        =   5
         Top             =   75
         Width           =   2175
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
         Index           =   1
         Left            =   6960
         TabIndex        =   3
         Top             =   75
         Width           =   1485
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
         Left            =   1260
         TabIndex        =   1
         Top             =   75
         Width           =   3975
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Index           =   19
         Left            =   9285
         TabIndex        =   6
         Top             =   135
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate No"
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
         Index           =   18
         Left            =   6075
         TabIndex        =   4
         Top             =   135
         Width           =   750
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   17
         Left            =   495
         TabIndex        =   2
         Top             =   135
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gridTransfer 
      Height          =   3150
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5175
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   5556
      _Version        =   393216
      Cols            =   3
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   13155
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3615
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
      Picture         =   "frmCPClusterDeliveryReg.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   13155
      TabIndex        =   8
      Top             =   3000
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Print"
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
      Picture         =   "frmCPClusterDeliveryReg.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   13155
      TabIndex        =   9
      Top             =   525
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
      Picture         =   "frmCPClusterDeliveryReg.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   13155
      TabIndex        =   10
      Top             =   2385
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
      Picture         =   "frmCPClusterDeliveryReg.frx":166E
   End
   Begin xrControl.xrFrame frmeCluster 
      Height          =   2400
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1095
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   4233
      Enabled         =   0   'False
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1275
         TabIndex        =   21
         Top             =   480
         Width           =   1545
      End
      Begin VB.TextBox txtOthers 
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
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "0W1-16000001"
         Top             =   90
         Width           =   1845
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   19
         Top             =   480
         Width           =   1380
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1275
         TabIndex        =   18
         Top             =   1080
         Width           =   3945
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1275
         TabIndex        =   17
         Top             =   1380
         Width           =   3945
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1275
         TabIndex        =   16
         Top             =   1680
         Width           =   3945
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   780
         Width           =   615
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1275
         TabIndex        =   14
         Top             =   780
         Width           =   3945
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1380
         Width           =   615
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   2  'Center
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
         Index           =   2
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1980
         Width           =   615
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   1275
         TabIndex        =   11
         Top             =   1980
         Width           =   3945
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
         Left            =   3750
         TabIndex        =   44
         Tag             =   "eb0;et0"
         Top             =   120
         Width           =   2235
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   0
         Left            =   3615
         Top             =   45
         Width           =   2505
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   0
         Left            =   3645
         Top             =   75
         Width           =   2445
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate No"
         Height          =   195
         Index           =   13
         Left            =   555
         TabIndex        =   32
         Top             =   525
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transact No"
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
         TabIndex        =   31
         Top             =   135
         Width           =   1065
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   3
         Left            =   540
         TabIndex        =   30
         Top             =   2025
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   2
         Left            =   3420
         TabIndex        =   29
         Top             =   525
         Width           =   345
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver"
         Height          =   195
         Index           =   5
         Left            =   750
         TabIndex        =   28
         Top             =   1125
         Width           =   420
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Helper 1"
         Height          =   195
         Index           =   7
         Left            =   570
         TabIndex        =   27
         Top             =   1425
         Width           =   600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Helper 2"
         Height          =   195
         Index           =   11
         Left            =   570
         TabIndex        =   26
         Top             =   1725
         Width           =   600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity"
         Height          =   195
         Index           =   12
         Left            =   5505
         TabIndex        =   25
         Top             =   570
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Space"
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
         Index           =   15
         Left            =   5535
         TabIndex        =   24
         Top             =   1770
         Width           =   555
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cluster"
         Height          =   195
         Index           =   16
         Left            =   690
         TabIndex        =   23
         Top             =   825
         Width           =   480
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Load"
         Height          =   195
         Index           =   14
         Left            =   5625
         TabIndex        =   22
         Top             =   1185
         Width           =   360
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   1
         Left            =   1335
         Tag             =   "et0;ht2"
         Top             =   195
         Width           =   1845
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   0
         Left            =   3675
         Tag             =   "et0;et0"
         Top             =   105
         Width           =   2400
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gridDelivery 
      Height          =   4050
      Left            =   6495
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1095
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   7144
      _Version        =   393216
      Cols            =   3
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin xrControl.xrFrame frmeMaster 
      Height          =   1635
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   3510
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   2884
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "001-16000001"
         Top             =   480
         Width           =   4875
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   4740
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Apr 26, 2016"
         Top             =   780
         Width           =   1380
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   465
         Index           =   4
         Left            =   1245
         MultiLine       =   -1  'True
         TabIndex        =   36
         Text            =   "frmCPClusterDeliveryReg.frx":1DE8
         Top             =   1080
         Width           =   4875
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
         Index           =   0
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "-"
         Top             =   90
         Width           =   1875
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1245
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Apr 26, 2016"
         Top             =   780
         Width           =   1380
      End
      Begin VB.Shape Shape4 
         Height          =   330
         Index           =   1
         Left            =   3645
         Top             =   60
         Width           =   2445
      End
      Begin VB.Shape Shape3 
         Height          =   390
         Index           =   1
         Left            =   3615
         Top             =   30
         Width           =   2505
      End
      Begin VB.Label Label1 
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
         Left            =   3750
         TabIndex        =   45
         Tag             =   "eb0;et0"
         Top             =   105
         Width           =   2235
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order No"
         Height          =   195
         Index           =   20
         Left            =   480
         TabIndex        =   43
         Top             =   825
         Width           =   645
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   6
         Left            =   4305
         TabIndex        =   42
         Top             =   825
         Width           =   345
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   8
         Left            =   495
         TabIndex        =   41
         Top             =   1125
         Width           =   630
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1350
         Tag             =   "et0;ht2"
         Top             =   180
         Width           =   1875
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery No"
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
         Index           =   4
         Left            =   120
         TabIndex        =   40
         Top             =   135
         Width           =   1005
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   1
         Left            =   615
         TabIndex        =   39
         Top             =   525
         Width           =   510
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Index           =   2
         Left            =   3675
         Tag             =   "et0;et0"
         Top             =   90
         Width           =   2400
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   13155
      TabIndex        =   46
      Top             =   1140
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Confirm"
      AccessKey       =   "Confirm"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPClusterDeliveryReg.frx":1E34
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   13155
      TabIndex        =   47
      Top             =   1755
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Unconfirm"
      AccessKey       =   "Unconfirm"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPClusterDeliveryReg.frx":25AE
   End
End
Attribute VB_Name = "frmCPClusterDeliveryReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Delivery Schedule Object
'
' Copyright 2016 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  iMac [ 06/20/2016 09:00 am ]
'     Start creating this form.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "frmCPClusterDeliveryReg"
Private Const pxeTransNoPict As String = "@@@@-@@-@@@@@@"
Private Const pxeDateLong As String = "MMMM DD, YYYY"
Private Const pxeDateShort As String = "MM/DD/YY"

Private WithEvents oTrans As clsCPClusterDelivery
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Private pbLoaded As Boolean
Private pbTModified As Boolean
Private pnOrder As Integer
Private pnTransfer As Integer
Private pnDelivery As Integer
Private pnCtr As Integer
Private pnIndex As Integer

Private Sub cmdButton_Click(Index As Integer)
   Dim lbConfirmed As Boolean
   Dim loForm As frmDate
   
   With oTrans
      Select Case Index
      Case 0 'browse
         Call txtSearch_KeyDown(pnIndex, vbKeyF3, 0)
      Case 1 'cancel
         If oTrans.CancelTransaction(Replace(txtOthers(0), "-", "")) Then
            Label2.Caption = TransStat(oTrans.Master("cTranStat"))
            
            MsgBox "Transaction was cancelled successfully.", vbInformation, "Notice"
         End If
      Case 2 'print
         If .Master("sSerialID") = "" Or .Master("sClustrID") = "" Then Exit Sub
         
         Call PrintTransaction
      Case 3 'close
         Unload Me
      Case 4 'confirm
         If Replace(txtOthers(0), "-", "") <> "" Then
            'Mac 2018.05.31
            'check if all transfers are printed
            Dim lnCtr As Integer
            For lnCtr = 1 To oTrans.ItemCount
               pnDelivery = lnCtr
               pnOrder = pnDelivery
               
               Call loadTransfer
            Next
         
            If oTrans.Master("cTranStat") < xeStateClosed Then
               If MsgBox("Confirm transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                  If MsgBox("Create employee OB?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                     Set loForm = New frmDate
                     
                     loForm.AppDriver = oApp
                     loForm.DateFrom = CDate(txtOthers(1))
                     loForm.Show vbModal
                     
                     If loForm.Cancelled = False Then
                        lbConfirmed = .ConfirmTransaction(True, loForm.DateEntry)
                     End If
                  Else
                     lbConfirmed = .ConfirmTransaction(False)
                  End If
               End If
               
               If lbConfirmed Then
                  'she 2016-08-05 this will temporary posting of delivery
                  'suppose to be posting from vehicle log
                  If MsgBox("Post transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                     .PostTransaction (.Master("sTransNox"))
                  End If
               
                  Label2.Caption = TransStat(oTrans.Master("cTranStat"))
                  
                  MsgBox "Transaction was confirmed succesfully.", vbInformation, "Success"
               End If
            End If
         End If
      Case 5
         If Replace(txtOthers(0), "-", "") <> "" Then
            If oTrans.Master("cTranStat") = xeStateClosed Then
               If MsgBox("Unconfirm transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                  If .UnConfirmTransaction Then
                     Label2.Caption = TransStat(oTrans.Master("cTranStat"))
                  
                     MsgBox "Transaction was unconfirmed succesfully.", vbInformation, "Success"
                  End If
               End If
            End If
         End If
      End Select
   End With
End Sub

Private Sub Form_Activate()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Activate"
   ''On Error GoTo errProc
   
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
 
   pbLoaded = True
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      Case vbKeyF8
         If txtOthers(0) = "" Then Exit Sub
         If MsgBox("Are you sure to delete transaction?", vbQuestion & vbYesNo, "Confirm") = vbOK Then
            If oTrans.DeleteTransaction Then
               MsgBox "Transaction deleted successfully.", vbInformation, "Notice"
            End If
         End If
      End Select
   End Select
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   lsOldProc = "Form_Load"

   ''On Error GoTo errProc

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight
   CenterChildForm mdiMain, Me
   
   Set oTrans = New clsCPClusterDelivery
   Set oTrans.AppDriver = oApp
   oTrans.DisplayMessage = True
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
   
   Call initTrans
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub initTrans()
   Call InitEntry
   Call initGridTransfer
   Call initGridDelivery
End Sub

Private Sub initGridDelivery()
   With gridDelivery
      .Clear
      .Cols = 5
      .Rows = 2
      .Font = "MS Sans Serif"
      .RowHeight(0) = 350

      'column title
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "Trans #"
      .TextMatrix(0, 3) = "Source"
      .TextMatrix(0, 4) = "Qty"

      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 4
      Next

      .ColWidth(0) = 350
      .ColWidth(1) = 2800
      .ColWidth(2) = 1400
      .ColWidth(3) = 1400
      .ColWidth(4) = 400

      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignLeftCenter

      .Row = 1
      .Col = 1
      
      .HighLight = flexHighlightNever
   End With
End Sub

Private Sub initGridTransfer()
   With gridTransfer
      .Clear
      .Cols = 7
      .Rows = 1
      .Font = "MS Sans Serif"
      .RowHeight(0) = 350

      'column title
      .TextMatrix(0, 1) = "Engine No"
      .TextMatrix(0, 2) = "Frame No"
      .TextMatrix(0, 3) = "Code"
      .TextMatrix(0, 4) = "Model"
      .TextMatrix(0, 5) = "Color"
      .TextMatrix(0, 6) = "Company"

      .Row = 0

      'column alignment
      For pnCtr = 0 To .Cols - 1
         .Col = pnCtr
         .CellFontBold = True
         .CellAlignment = 4
      Next

      .ColWidth(0) = 350
      .ColWidth(1) = 2000
      .ColWidth(2) = 2000
      .ColWidth(3) = 2000
      .ColWidth(4) = 2500
      .ColWidth(5) = 1900
      .ColWidth(6) = 2000

      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignLeftCenter
      .ColAlignment(5) = flexAlignLeftCenter

      .Rows = 2
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Function InitEntry()
   Dim loText As TextBox
   
   For Each loText In txtOthers
      loText.Text = ""
   Next
   txtOthers(1) = Format(oApp.ServerDate, pxeDateLong)
   
   For Each loText In txtField
      loText.Text = ""
   Next
   
   For Each loText In txtTotal
      loText.Text = "0"
   Next
   
   For Each loText In txtSearch
      loText.Text = ""
   Next
   
   txtField(3) = Format(oApp.ServerDate, pxeDateLong)
End Function

Private Sub LoadMaster()
   Dim loText As TextBox
   
   With oTrans
      For Each loText In txtOthers
         loText.Text = IFNull(.Master(loText.Index), "")
      Next
      
      txtOthers(0).Text = Format(.Master("sTransNox"), pxeTransNoPict)
      txtOthers(1).Text = Format(.Master("dTransact"), pxeDateLong)
      txtTotal(0).Text = .Capacity
      
      txtSearch(0) = .Master("sClustrDs")
      txtSearch(1) = .Master("sPlateNox")
      txtSearch(2) = .Master("dTransact")
      
      Label2.Caption = TransStat(.Master("cTranStat"))
   End With
End Sub

Private Sub loadDelivery()
   Dim lnCtr As Integer
   
   With gridDelivery
      .Rows = oTrans.ItemCount + 1
      
      For lnCtr = 1 To .Rows - 1
         .TextMatrix(lnCtr, 0) = lnCtr
         .TextMatrix(lnCtr, 1) = oTrans.Detail(lnCtr - 1, "sBranchNm")
         .TextMatrix(lnCtr, 2) = Format(oTrans.Detail(lnCtr - 1, "sReferNox"), pxeTransNoPict)
         .TextMatrix(lnCtr, 3) = getSource(oTrans.Detail(lnCtr - 1, "sSourceCd"))
         .TextMatrix(lnCtr, 4) = oTrans.Detail(lnCtr - 1, "nNoItemsx")
      Next
      
      Call compTotal
   End With
End Sub

Private Sub loadTransfer()
   With oTrans
      ' open transfer
      If Trim(.Detail(pnDelivery - 1, "sReferNox")) = "" Then Exit Sub
      
      If Not .Issuance.OpenTransaction(.Detail(pnDelivery - 1, "sReferNox")) Then
         Call setTransferInfo
         Call setTransDetInfo
                  
         Exit Sub
      End If
      
      txtField(0) = Format(oTrans.Issuance.TransNo, pxeTransNoPict)
      txtField(1) = oTrans.Issuance.Destination
      txtField(2) = Format(oTrans.Issuance.OrderNo, pxeTransNoPict)
      txtField(3) = Format(oTrans.Issuance.TransDate, pxeDateLong)
      txtField(4) = oTrans.Issuance.Remarks
      
      Label1.Caption = TransStat(oTrans.Issuance.TransStatus)
      
      Call loadTransferDetail
      
      'print transfer on load
      If oTrans.Issuance.TransNo <> "" Then
         If oTrans.Issuance.TransStatus < xeStateClosed Then
            MsgBox "Transaction no. " & oTrans.Issuance.TransNo & " is not printed." & vbCrLf & _
                     "System will print the transaction. Please prepare the printer.", vbInformation, "Notice"
                     
            Call PrintTrans
            Label1.Caption = TransStat(oTrans.Issuance.TransStatus)
         End If
      End If
   End With
End Sub

Private Sub loadTransferDetail()
   Call initGridTransfer
   With gridTransfer
      .Rows = oTrans.Issuance.ItemCount + 1
      
      For pnTransfer = 1 To .Rows - 1
         Call setTransDetInfo
      Next
      
      .Row = .Rows - 1
      pnTransfer = .Row
      Call setMCInfo
      
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub setTransferInfo()
   oTrans.Issuance.NewTransaction
   oTrans.Issuance.OrderNo = IFNull(oTrans.Request(pnOrder - 1, "sTransNox"), "")
   oTrans.Issuance.Destination = oTrans.Request(pnOrder - 1, "sBranchNm")
   
   txtField(0) = Format(oTrans.Issuance.TransNo, pxeTransNoPict)
   txtField(1) = oTrans.Request(pnOrder - 1, "sBranchNm")
   txtField(2) = Format(oTrans.Request(pnOrder - 1, "sTransNox"), pxeTransNoPict)
   txtField(3) = Format(oTrans.Master("dTransact"), pxeDateLong)
   txtField(4) = oTrans.Issuance.Remarks

   Call initGridTransfer
   pnTransfer = 1
End Sub

Private Sub setMCInfo()
'   txtDetail(1) = gridTransfer.TextMatrix(pnTransfer, 1)
'   txtDetail(2) = gridTransfer.TextMatrix(pnTransfer, 2)
End Sub

Private Sub setTransDetInfo()
   With gridTransfer
      .TextMatrix(pnTransfer, 0) = oTrans.Issuance.Detail(pnTransfer - 1, "nEntryNox")
      .TextMatrix(pnTransfer, 1) = oTrans.Issuance.Detail(pnTransfer - 1, "sSerialNo")
      .TextMatrix(pnTransfer, 2) = oTrans.Issuance.Detail(pnTransfer - 1, "sBarCodex")
      .TextMatrix(pnTransfer, 3) = oTrans.Issuance.Detail(pnTransfer - 1, "sModelCde")
      .TextMatrix(pnTransfer, 4) = oTrans.Issuance.Detail(pnTransfer - 1, "sModelNme")
      .TextMatrix(pnTransfer, 5) = IFNull(oTrans.Issuance.Detail(pnTransfer - 1, "sColorNme"))
      .TextMatrix(pnTransfer, 6) = IFNull(oTrans.Issuance.Detail(pnTransfer - 1, "sBrandNme"), "")
   End With
End Sub

Private Function getSource(ByVal lsSourceCd As String) As String
   Select Case LCase(lsSourceCd)
   Case "mcdv"
      getSource = "MC Transfer"
   Case "spdv"
      getSource = "SP Transfer"
   Case "cpdl"
      getSource = "CP Transfer"
   Case "spwt"
      getSource = "SP Warranty Transfer"
   Case "cpjt"
      getSource = "CP Job Order"
   Case "asdv"
      getSource = "Asset Transfer"
   Case "sudv"
      getSource = "Supplies Transfer"
   Case "rgdv"
      getSource = "MC Reg Transfer"
   Case "cpdv"
      getSource = "CP Unit Transfer"
   Case "ckdv"
      getSource = "Check Transfer"
   Case "dcdv"
      getSource = "Document Transfer"
   Case Else
      getSource = "Oth Transfer"
   End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
   pbLoaded = False
End Sub

Private Sub gridDelivery_Click()
   With gridDelivery
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub gridDelivery_DblClick()
   With gridDelivery
      If .TextMatrix(.Row, 3) <> "CP Unit Transfer" Then Exit Sub
   
      pnDelivery = .Row
      pnOrder = pnDelivery
      
      Call loadTransfer
   End With
End Sub

Private Sub gridTransfer_Click()
   With gridTransfer
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub gridTransfer_DblClick()
   With gridTransfer
      pnTransfer = .Row
      
      Call setMCInfo
   End With
End Sub

Private Sub oTrans_LoadDelivery()
   Call loadDelivery
End Sub

Private Sub oTrans_LoadTransaction()
   Call loadDelivery
   Call LoadMaster
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

Private Sub compTotal()
   Dim lnCtr As Integer
   Dim lnTotal As Integer
   
   With gridDelivery
      For lnCtr = 1 To .Rows - 1
         If IsNumeric(.TextMatrix(lnCtr, 4)) Then
            lnTotal = lnTotal + CInt(.TextMatrix(lnCtr, 4))
         End If
      Next
   End With
   
   txtTotal(1).Text = lnTotal
   txtTotal(2).Text = CInt(txtTotal(0)) - lnTotal
End Sub

Private Function detailOK() As Boolean
   Dim lnCtr As Integer
   Dim lnRow As Integer, lnCol As Integer
   
   With gridTransfer
      lnCtr = 1
      Do Until lnCtr = .Rows - 1
         If .TextMatrix(lnCtr, 1) = "" Then
            If oTrans.Issuance.deleteDetail(lnCtr - 1) = False Then Exit Function
            
            For lnRow = lnCtr To .Rows - 1
               For lnCol = 0 To .Cols - 1
                  .TextMatrix(lnRow, lnCol) = .TextMatrix(lnRow + 1, lnCol)
               Next
            Next
         Else
            lnCtr = lnCtr + 1
         End If
      Loop
   End With

   detailOK = True
End Function

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   If Index = 13 Then
      Label2.Caption = TransStat(oTrans.Master("cTranStat"))
   End If
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
   Call HighlightOn(txtSearch(Index))
   
   pnIndex = Index
End Sub

Private Sub txtSearch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lbSearch As Boolean

   Select Case KeyCode
   Case vbKeyReturn, vbKeyF3
      If Index = 2 Then
         If IsDate(txtSearch(Index)) Then
            txtSearch(Index) = Format(CDate(txtSearch(Index)), "YYYY-MM-DD")
         Else
            txtSearch(Index) = ""
         End If
      End If
   
      If KeyCode = vbKeyReturn Then
         Select Case Index
         Case 0
            lbSearch = oTrans.SearchTransaction("sClustrDs", txtSearch(Index), False)
         Case 1
            lbSearch = oTrans.SearchTransaction("sPlateNox", txtSearch(Index), False)
         Case 2
            lbSearch = oTrans.SearchTransaction("dTransact", txtSearch(Index), False)
         End Select
      Else
         Select Case Index
         Case 0
            lbSearch = oTrans.SearchTransaction("sClustrDs", txtSearch(Index) & "%", True)
         Case 1
            lbSearch = oTrans.SearchTransaction("sPlateNox", txtSearch(Index) & "%", True)
         Case 2
            lbSearch = oTrans.SearchTransaction("dTransact", txtSearch(Index) & "%", True)
         End Select
      End If
      
      If lbSearch Then
         loadTrans
      Else
         initSearch
      End If
   End Select
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
   Call HighlightOff(txtSearch(Index))
End Sub

Private Sub initSearch()
   txtSearch(0) = ""
   txtSearch(1) = ""
   txtSearch(2) = ""
End Sub

Private Sub loadTrans()
   Call LoadMaster
   Call loadDelivery
End Sub

Private Function PrintTransaction() As Boolean
   Dim lrs As ADODB.Recordset
   Dim lnCtr As Integer
   Dim lsMCInvID As String
   Dim lasMCInv() As String
   Dim lanMCInv() As Integer
   Dim lbFirst As Boolean
   Dim lsOldProc As String
   
   lsOldProc = "PrintTransaction"
   ''On Error GoTo errProc

   PrintTransaction = True
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 5
   lrs.Fields.Append "sField01", adVarChar, 120
   lrs.Fields.Append "sField02", adVarChar, 50
   lrs.Fields.Append "sField03", adVarChar, 50
   lrs.Open

   With oTrans
      For lnCtr = 0 To .ItemCount - 1
'         If .Detail(lnCtr, "nNoItemsx") > 0 Then
            lrs.AddNew
            
            lrs("sField01") = .Detail(lnCtr, "sBranchNm")
            lrs("sField02") = .Detail(lnCtr, "sReferNox")
            lrs("sField03") = getSource(oTrans.Detail(lnCtr - 1, "sSourceCd"))
'            lrs("nField01") = .Detail(lnCtr, "nNoItemsx")
'         End If
      Next
   End With

   'assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CPClusterStockDelivery.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs
   
   oReport.Sections("RH").ReportObjects("txtHeadTitle").SetText oApp.BranchName
   oReport.Sections("RH").ReportObjects("txtCluster").SetText oTrans.Master("sClustrDs")
   oReport.Sections("RH").ReportObjects("txtHeadDescription").SetText oApp.Address & ", " & oApp.TownCity & " " & oApp.Province & " " & oApp.ZippCode
   oReport.Sections("RH").ReportObjects("sTransNox").SetText "CP" & "-" & Right(oTrans.Master("sTransNox"), 11)
   oReport.Sections("RH").ReportObjects("dTransact").SetText Format(oTrans.Master("dTransact"), "Mmm dd, yyyy")
   oReport.Sections("PH").ReportObjects("sRemarksx").SetText IFNull(oTrans.Master("sRemarksx"), "")
   oReport.Sections("PH").ReportObjects("sDriverxx").SetText oTrans.Master("sDriverxx")
   oReport.Sections("PH").ReportObjects("sHelper01").SetText oTrans.Master("sHelper01")
   oReport.Sections("PH").ReportObjects("sHelper02").SetText IFNull(oTrans.Master("sHelper02"))
   oReport.Sections("PF").ReportObjects("sPrepared").SetText oApp.UserName
   oReport.Sections("PF").ReportObjects("PlateNo").SetText IFNull(oTrans.Master("sPlateNox"))
   
   PrintTransaction = True
   oReport.PrintOutEx False, 1

endPoc:
   Set oReport = Nothing
   Set lrs = Nothing
   Exit Function
errProc:
   PrintTransaction = False
   ShowError lsOldProc & "( " & " )"
End Function

Private Function PrintTrans() As Boolean
   Dim loreport As frmRepViewer
   Dim lrs As Recordset
   Dim lors As Recordset
   Dim lnCtr As Integer
   Dim lsOldProc As String
   Dim lnTotlWSerial As Double
   Dim lnTotlWOSerial As Double
   Dim lsSourceNo As String
   
   lsOldProc = "PrinTrans"
   ''On Error GoTo errProc

   PrintTrans = True
   
   Set lrs = New ADODB.Recordset

   lrs.Fields.Append "nField01", adInteger, 3
   lrs.Fields.Append "nField02", adChar, 1
   lrs.Fields.Append "sField01", adVarChar, 20
   lrs.Fields.Append "sField02", adVarChar, 128
   lrs.Fields.Append "sField03", adVarChar, 30
   lrs.Fields.Append "sField04", adVarChar, 12
   lrs.Fields.Append "sField05", adVarChar, 100
   lrs.Fields.Append "lField01", adCurrency
   lrs.Open
   
   With oTrans.Issuance
      lsSourceNo = IFNull(.OrderNo, "")
      lnTotlWOSerial = 0
      lnTotlWSerial = 0
      
      For lnCtr = 0 To .ItemCount - 1
         lrs.AddNew
         lrs.Fields("nField01") = .Detail(lnCtr, "nQuantity")
         lrs.Fields("nField02") = .Detail(lnCtr, "cHsSerial")
         lrs.Fields("sField01") = .Detail(lnCtr, "sBarrCode")
         lrs.Fields("sField02") = .Detail(lnCtr, "sDescript")
         lrs.Fields("sField03") = .Detail(lnCtr, "sSerialNo")
         lrs.Fields("sField04") = .Detail(lnCtr, "sStockIDx")
         lrs.Fields("sField05") = .Detail(lnCtr, "sBrandNme") & " " & .Detail(lnCtr, "sModelNme") & " " & .Detail(lnCtr, "sColorNme")
         lrs.Fields("lField01") = .Detail(lnCtr, "nSelPrice")
      Next
      lrs.Sort = "nField02 DESC,sField05,sField05,sField03"
   End With

   'assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\CP_Transfer.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs

   Set lors = New ADODB.Recordset
   If lors.State = adStateOpen Then lors.Close

   lors.Open "SELECT" _
               & "  a.sAddressx" _
               & ", CONCAT(b.sTownName, ', ' , c.sProvName, ' ' , b.sZippCode) xTownName" _
               & ", a.sBranchNm" _
            & " FROM Branch a" _
               & " LEFT JOIN TownCity b" _
                  & " LEFT JOIN Province c" _
                     & " ON b.sProvIDxx = c.sProvIDxx" _
                  & " ON a.sTownIDxx = b.sTownIDxx" _
            & " WHERE a.sBranchCd = " & strParm(oTrans.Issuance.DestinatCode) _
            , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText

   oReport.Sections("PHa").ReportObjects("txtRefNo").SetText "CP" & "-" & Right(oTrans.Issuance.TransNo, 10)
   oReport.Sections("PHa").ReportObjects("txtDate").SetText txtField(3).Text
   oReport.Sections("PHb").ReportObjects("txtTo").SetText lors("sBranchNm")
   oReport.Sections("PHb").ReportObjects("txtToAddress").SetText lors("sAddressx") & IFNull(lors("xTownName"), "")
   oReport.Sections("PHb").ReportObjects("txtFrom").SetText oApp.ClientName
   oReport.Sections("PHb").ReportObjects("txtFromAddress").SetText oApp.Address & ", " & oApp.TownCity & ", " & oApp.Province & " " & oApp.ZippCode
   oReport.Sections("RFb").ReportObjects("txtRemarks").SetText lsSourceNo & " " & txtField(4).Text
   oReport.Sections("RFb").ReportObjects("txtWithSerial").SetText oTrans.Issuance.ItemCount
   oReport.Sections("RFb").ReportObjects("txtWOutSerial").SetText "0"
   oReport.Sections("PF").ReportObjects("txtRptUser").SetText oApp.UserName
   
   Set loreport = New frmRepViewer
   Set loreport.ReportSource = oReport
   loreport.Show
   
   PrintTrans = True
endPoc:
   If oTrans.Issuance.TransStatus <> xeStateClosed Then
      If BranchAutomate(oTrans.Issuance.DestinatCode) Then
         oTrans.Issuance.ConfirmTransaction (oTrans.Issuance.TransNo)
      End If
   End If
   Set loreport = Nothing
   Set oReport = Nothing
   Set lrs = Nothing
   Set lors = Nothing
   Exit Function
errProc:
   PrintTrans = False
   ShowError lsOldProc & "( " & " )"
End Function

Private Function BranchAutomate(ByVal sBranchCd As String) As Boolean
   Dim lrs As Recordset
   
   Set lrs = New Recordset
   lrs.Open "SELECT * FROM Branch" & _
               " WHERE sBranchCd = " & strParm(sBranchCd) & _
                  " AND cAutomate = " & strParm(xeYes) _
   , oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   If Not lrs.EOF Then BranchAutomate = True
   Set lrs = Nothing
End Function

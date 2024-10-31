VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCPClusterDelivery 
   BorderStyle     =   0  'None
   Caption         =   "Cellphone Delivery by Cluster"
   ClientHeight    =   8745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   14490
   ShowInTaskbar   =   0   'False
   Tag             =   "Approved Orders"
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   44
      Top             =   525
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
      Picture         =   "frmCPClusterDelivery.frx":0000
   End
   Begin xrControl.xrFrame frmeDetail 
      Height          =   450
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   4545
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   794
      ClipControls    =   0   'False
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   3870
         TabIndex        =   46
         Top             =   60
         Width           =   2370
      End
      Begin VB.TextBox txtDetail 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   525
         TabIndex        =   45
         Top             =   60
         Width           =   2520
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BCODE"
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
         Left            =   3180
         TabIndex        =   47
         Top             =   120
         Width           =   645
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMEI"
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
         Index           =   10
         Left            =   60
         TabIndex        =   33
         Top             =   120
         Width           =   405
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1755
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
      Picture         =   "frmCPClusterDelivery.frx":077A
   End
   Begin xrControl.xrFrame frmeCluster 
      Height          =   2400
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   525
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   4233
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   1275
         TabIndex        =   15
         Top             =   1980
         Width           =   3945
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1980
         Width           =   615
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1380
         Width           =   615
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1275
         TabIndex        =   7
         Top             =   780
         Width           =   3945
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   5505
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   780
         Width           =   615
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1275
         TabIndex        =   13
         Top             =   1680
         Width           =   3945
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1275
         TabIndex        =   11
         Top             =   1380
         Width           =   3945
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1275
         TabIndex        =   9
         Top             =   1080
         Width           =   3945
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   5
         Top             =   480
         Width           =   1380
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
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "0W1-16000001"
         Top             =   90
         Width           =   1845
      End
      Begin VB.TextBox txtOthers 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1275
         TabIndex        =   3
         Top             =   480
         Width           =   1545
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1365
         Tag             =   "et0;ht2"
         Top             =   180
         Width           =   1845
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Load"
         Height          =   195
         Index           =   14
         Left            =   5625
         TabIndex        =   18
         Top             =   1185
         Width           =   360
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cluster"
         Height          =   195
         Index           =   16
         Left            =   690
         TabIndex        =   6
         Top             =   825
         Width           =   480
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
         TabIndex        =   20
         Top             =   1770
         Width           =   555
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity"
         Height          =   195
         Index           =   12
         Left            =   5505
         TabIndex        =   16
         Top             =   570
         Width           =   615
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Helper 2"
         Height          =   195
         Index           =   11
         Left            =   570
         TabIndex        =   12
         Top             =   1725
         Width           =   600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Helper 1"
         Height          =   195
         Index           =   7
         Left            =   570
         TabIndex        =   10
         Top             =   1425
         Width           =   600
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver"
         Height          =   195
         Index           =   5
         Left            =   750
         TabIndex        =   8
         Top             =   1125
         Width           =   420
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   2
         Left            =   3420
         TabIndex        =   4
         Top             =   525
         Width           =   345
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Index           =   3
         Left            =   540
         TabIndex        =   14
         Top             =   2025
         Width           =   630
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
         TabIndex        =   0
         Top             =   135
         Width           =   1065
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate No"
         Height          =   195
         Index           =   13
         Left            =   555
         TabIndex        =   2
         Top             =   525
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gridTransfer 
      Height          =   3540
      Left            =   1575
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5040
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   6244
      _Version        =   393216
      Cols            =   3
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1140
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmCPClusterDelivery.frx":0EF4
   End
   Begin MSFlexGridLib.MSFlexGrid gridOrder 
      Height          =   1995
      Left            =   7980
      TabIndex        =   22
      Top             =   525
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   3519
      _Version        =   393216
      Cols            =   3
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid gridDelivery 
      Height          =   2460
      Left            =   7980
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2535
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   4339
      _Version        =   393216
      Cols            =   3
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin xrControl.xrFrame frmeMaster 
      Height          =   1590
      Left            =   1575
      Tag             =   "wt0;fb0"
      Top             =   2940
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   2805
      ClipControls    =   0   'False
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Apr 26, 2016"
         Top             =   780
         Width           =   1380
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
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "-"
         Top             =   90
         Width           =   1875
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   465
         Index           =   4
         Left            =   1245
         MultiLine       =   -1  'True
         TabIndex        =   32
         Text            =   "frmCPClusterDelivery.frx":166E
         Top             =   1080
         Width           =   3555
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   3420
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Apr 26, 2016"
         Top             =   780
         Width           =   1380
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "001-16000001"
         Top             =   480
         Width           =   3555
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "&Save"
         Height          =   450
         Index           =   0
         Left            =   5010
         TabIndex        =   40
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "&Del. Row"
         Height          =   450
         Index           =   1
         Left            =   4995
         TabIndex        =   41
         Top             =   540
         Width           =   1245
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "&Cancel"
         Height          =   450
         Index           =   2
         Left            =   4995
         TabIndex        =   42
         Top             =   1005
         Width           =   1245
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   195
         Index           =   17
         Left            =   615
         TabIndex        =   25
         Top             =   525
         Width           =   510
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
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   135
         Width           =   1005
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
         Caption         =   "Remarks"
         Height          =   195
         Index           =   8
         Left            =   495
         TabIndex        =   31
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Index           =   6
         Left            =   2985
         TabIndex        =   29
         Top             =   825
         Width           =   345
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order No"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   27
         Top             =   825
         Width           =   645
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   5
      Left            =   90
      TabIndex        =   43
      Top             =   2370
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
      Picture         =   "frmCPClusterDelivery.frx":16BA
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   525
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&OK"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmCPClusterDelivery.frx":1E34
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1140
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
      Picture         =   "frmCPClusterDelivery.frx":25AE
   End
End
Attribute VB_Name = "frmCPClusterDelivery"
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
'  iMac [ 04/25/2016 09:00 am ]
'     Start creating this form.
'  XerSys [ 06/02/2016 04:40 am ]
'     Continue coding this form
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const pxeMODULENAME As String = "frmCPClusterDelivery"
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

Private Sub cmdButton_Click(Index As Integer)
   Dim lbConfirmed As Boolean
   Dim loForm As frmDate
   
   With oTrans
      Select Case Index
      Case 0 'New
         If .NewTransaction Then
            Call initTrans
            Call initButtMn(xeModeAddNew)
            Call LoadMaster
            txtOthers(2).SetFocus
         End If
      Case 2 'OK
         If .Master("sSerialID") = "" Or txtOthers(2) = "" Then
            MsgBox "Invalid vehicle detected.", vbCritical, "Warning"
            Exit Sub
         End If
         
         If .Master("sClustrID") = "" Then
            MsgBox "Invalid cluster detected.", vbCritical, "Warning"
            Exit Sub
         End If
         
         If .Master("sDriverxx") = "" Then
            MsgBox "Invalid driver detected.", vbCritical, "Warning"
            Exit Sub
         End If
         
         If .Master("sHelper01") = "" Then
            MsgBox "Invalid helper detected.", vbCritical, "Warning"
            Exit Sub
         End If
      
         If .SaveTransaction Then
            MsgBox "Transaction saved successfully.", vbInformation, "Success"
            
            If .OpenTransaction(Replace(txtOthers(0), "-", "")) Then
               Call LoadMaster
               Call loadDelivery
               Call LoadOrder

               Call initButtMn(xeModeUnknown)
               Call initButtDt(xeModeUnknown)
            End If
         End If
      Case 1 'Close
         Unload Me
      Case 3 'Cancel
         If MsgBox("Unsaved changes will be disregarded." & vbCrLf & vbCrLf & _
               "Do you want to continue?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            
            .InitTransaction
            Call initTrans
         End If
      Case 4 'Printrans
         If Replace(txtOthers(0), "-", "") <> "" Then
            If MsgBox("Do you want to print transaction?", vbQuestion & vbYesNo, "Confirm") = vbYes Then
               Call PrintTransaction
            End If
         End If
      Case 5 'confirm
         If Replace(txtOthers(0), "-", "") <> "" Then
            If .Master("cTranStat") < xeStateClosed Then
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
                  
                  'she 2016-08-05 this will temporary posting of delivery
                  'suppose to be posting from vehicle log
                  If MsgBox("Post transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                     If .PostTransaction(.Master("sTransNox")) Then
                        MsgBox "Transaction Posted Successfully..", vbInformation, "Success"
                     End If
                  End If
               End If
               
               If lbConfirmed Then
                  MsgBox "Transaction was confirmed succesfully.", vbInformation, "Success"
                  
                  .InitTransaction
                  Call initTrans
               End If
            End If
         End If
      End Select
   End With
End Sub

Private Sub cmdDetail_Click(Index As Integer)
   Select Case Index
   Case 0
      If cmdDetail(Index).Caption = "&Save" Then
         If detailOK Then
            oTrans.Issuance.TransDate = oTrans.Master("dTransact")
            
            If Not oTrans.Issuance.SaveTransaction Then Exit Sub
            If pnOrder < gridOrder.Rows Then
               gridOrder.TextMatrix(pnOrder, 4) = oTrans.Issuance.TotalItems 'ItemCount
               gridOrder.TextMatrix(pnOrder, 5) = "Yes"
            End If
            
            Call compTotal
            Call loadDelivery
            Call loadTransfer
            Call initButtDt(xeModeUpdate)
         Else
            MsgBox "Unable to save transfer. Please verify your entry.", vbInformation, "Notice"
         End If
      Else
         Call initButtDt(xeModeAddNew)
         txtDetail(1).SetFocus
      End If
   Case 1
      If cmdDetail(Index).Caption = "&Del. Row" Then
         ' Delete Detail
         If oTrans.Issuance.Detail(pnTransfer - 1, "sSerialID") <> "" Then
            If oTrans.Issuance.deleteDetail(pnTransfer - 1) Then loadTransferDetail
         End If
      Else
         ' Print
         If oTrans.Issuance.TransNo <> "" Then
            Call PrintTransfer
         End If
      End If
   Case 2 ' cancel
      If pbTModified Then
         If MsgBox("Issuance is modified!" & vbCrLf & _
               "Cancellation will disregard entry!" & vbCrLf & vbCrLf & _
               "Continue Anyway?", vbQuestion + vbYesNo, "Confirm") <> vbYes Then Exit Sub
      End If
      
      If oTrans.Issuance.EditMode = xeModeAddNew Then
         Call setTransferInfo
         Call initButtDt(xeModeUnknown)
      Else
         Call loadTransfer
         Call initButtDt(xeModeUpdate)
      End If
   End Select
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
         If GetFocus = txtDetail(1).hwnd = True Then
            txtDetail(1).SetFocus
         ElseIf GetFocus = txtDetail(2).hwnd = True Then
            txtDetail(2).SetFocus
         Else
            SetNextFocus
         End If
      Case vbKeyUp
         SetPreviousFocus
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
   oSkin.ApplySkin xeFormTransEqualLeft
   CenterChildForm mdiMain, Me
   
   Set oTrans = New clsCPClusterDelivery
   Set oTrans.AppDriver = oApp
   oTrans.DisplayMessage = True
   oTrans.Branch = oApp.BranchCode
   oTrans.InitTransaction
   
   Call initTrans
   Call initButtDt(xeUnknown)
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub initTrans()
   Call InitEntry
   Call initGridOrder
   Call initGridTransfer
   Call initGridDelivery
   
   Call initButtMn(xeModeUnknown)
   Call initButtDt(xeModeUnknown)
End Sub

Private Sub initButtMn(ByVal lnStat As Integer)
   Dim lbShow As Boolean
   
   lbShow = (lnStat = xeModeAddNew Or lnStat = xeModeUpdate)
   
   cmdButton(0).Visible = Not lbShow
   cmdButton(1).Visible = Not lbShow
   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow
   cmdButton(4).Visible = Not lbShow
   cmdButton(5).Visible = False 'Not lbShow
   
   frmeCluster.Enabled = lbShow
   frmeMaster.Enabled = lbShow
   frmeDetail.Enabled = Not lbShow
   gridOrder.Enabled = lbShow
   gridDelivery.Enabled = lbShow
   gridTransfer.Enabled = lbShow
   If pbLoaded Then If Not lbShow Then cmdButton(0).SetFocus
End Sub

Private Sub initButtDt(ByVal lnStat As Integer)
   Dim lbShow As Boolean
   
   lbShow = (lnStat = xeModeAddNew Or lnStat = xeModeUpdate)
   
   frmeMaster.Enabled = lbShow
   frmeDetail.Enabled = lbShow
   
   cmdDetail(0).Visible = lbShow
   cmdDetail(1).Visible = lbShow
   cmdDetail(2).Visible = lbShow
   
   If lbShow Then
      If lnStat = xeModeAddNew Then
         cmdDetail(0).Caption = "&Save"
         cmdDetail(1).Caption = "&Del. Row"
         
         lbShow = True
      ElseIf lnStat = xeModeUpdate Then
         cmdDetail(0).Caption = "&Update"
         cmdDetail(1).Caption = "&Print"
         
         lbShow = False
      End If
      
      frmeDetail.Enabled = lbShow
   End If
   txtField(4).Locked = Not lbShow
End Sub

Private Sub initGridOrder()
   With gridOrder
      .Clear
      .Cols = 6
      .Rows = 2
      .Font = "MS Sans Serif"
      .RowHeight(0) = 350

      'column title
      .TextMatrix(0, 1) = "Branch"
      .TextMatrix(0, 2) = "Trans #"
      .TextMatrix(0, 3) = "Date"
      .TextMatrix(0, 4) = "Qty"
      .TextMatrix(0, 5) = "Prc"

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
      .ColWidth(3) = 900
      .ColWidth(4) = 500
      .ColWidth(5) = 400

      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignLeftCenter
      .ColAlignment(3) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter

      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
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
'      .ColSel = .Cols - 1
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
      .TextMatrix(0, 1) = "IMEI"
      .TextMatrix(0, 2) = "BarCode"
      .TextMatrix(0, 3) = "Code"
      .TextMatrix(0, 4) = "Model"
      .TextMatrix(0, 5) = "Color"
      .TextMatrix(0, 6) = "Brand"

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
   End With
End Sub

Private Sub loadDelivery()
   Dim lnCtr As Integer
   
   With gridDelivery
      .Rows = oTrans.ItemCount + 1
      
      For lnCtr = 1 To .Rows - 1
         .TextMatrix(lnCtr, 0) = lnCtr
         .TextMatrix(lnCtr, 1) = IFNull(oTrans.Detail(lnCtr - 1, "sBranchNm"))
         .TextMatrix(lnCtr, 2) = Format(oTrans.Detail(lnCtr - 1, "sReferNox"), pxeTransNoPict) 'sTransNox
         .TextMatrix(lnCtr, 3) = getSource(oTrans.Detail(lnCtr - 1, "sSourceCd"))
         .TextMatrix(lnCtr, 4) = IFNull(oTrans.Detail(lnCtr - 1, "nNoItemsx"), "0")
      Next
      
      .Row = 1
      pnDelivery = .Row
      pnOrder = pnDelivery
      
      Call compTotal
   End With
End Sub

Private Sub loadTransfer()
   With oTrans
      ' open transfer
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
      
      .Issuance.addDetail
      Call initButtDt(xeModeUpdate)
      
      Call loadTransferDetail
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

Private Sub LoadOrder()
   Dim lnCtr As Integer
   
   With gridOrder
      .Rows = oTrans.RequestCount + 1
      
      For lnCtr = 1 To .Rows - 1
         .TextMatrix(lnCtr, 0) = lnCtr
         .TextMatrix(lnCtr, 1) = oTrans.Request(lnCtr - 1, "sBranchNm")
         .TextMatrix(lnCtr, 2) = Format(oTrans.Request(lnCtr - 1, "sTransNox"), pxeTransNoPict)
         .TextMatrix(lnCtr, 3) = IFNull(oTrans.Request(lnCtr - 1, "dTransact"), "")
         .TextMatrix(lnCtr, 4) = IFNull(oTrans.Request(lnCtr - 1, "nRequestx"), 0) 'nApproved
         .TextMatrix(lnCtr, 5) = IIf(oTrans.Request(lnCtr - 1, "cConfirmd") = 1, "Yes", "No")
      Next
      
      If .Rows > 7 Then
         .ColWidth(1) = 2550
      Else
         .ColWidth(1) = 2800
      End If
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
   Call initButtDt(xeModeAddNew)
   pnTransfer = 1
End Sub

Private Sub setMCInfo()
   txtDetail(1) = gridTransfer.TextMatrix(pnTransfer, 1)
   txtDetail(2) = gridTransfer.TextMatrix(pnTransfer, 2)
End Sub

Private Sub setTransDetInfo()
   With gridTransfer
      .TextMatrix(pnTransfer, 0) = pnTransfer
      .TextMatrix(pnTransfer, 1) = oTrans.Issuance.Detail(pnTransfer - 1, "sSerialNo")
      .TextMatrix(pnTransfer, 2) = oTrans.Issuance.Detail(pnTransfer - 1, "sBarrCode")
      .TextMatrix(pnTransfer, 3) = IFNull(oTrans.Issuance.Detail(pnTransfer - 1, "sModelCde"))
      .TextMatrix(pnTransfer, 4) = IFNull(oTrans.Issuance.Detail(pnTransfer - 1, "sModelNme"))
      .TextMatrix(pnTransfer, 5) = IFNull(oTrans.Issuance.Detail(pnTransfer - 1, "sColorNme"))
      .TextMatrix(pnTransfer, 6) = IFNull(oTrans.Issuance.Detail(pnTransfer - 1, "sBrandNme"))
   End With
End Sub

Private Function getSource(ByVal lsSourceCd As String) As String
   Select Case LCase(lsSourceCd)
   Case "mcdv"
      getSource = "MC Transfer"
   Case "spdl"
      getSource = "SP Transfer"
   Case "cpdl"
      getSource = "CP Transfer"
   Case "spwt"
      getSource = "SP Warranty Transfer"
   Case "cpjt"
      getSource = "CP Job Order"
   Case "asdl"
      getSource = "Asset Transfer"
   Case "sudl"
      getSource = "Supplies Transfer"
   Case "rgdl"
      getSource = "MC Reg Transfer"
   Case "cpdv"
      getSource = "CP Unit Transfer"
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
      If .TextMatrix(.Row, 3) = "MC Transfer" Then
         pnDelivery = .Row
         pnOrder = pnDelivery
         
         Call loadTransfer
      Else
         Call procOthrTransfer
      End If
   End With
End Sub

Private Sub procOthrTransfer()
   Dim lnCtr As Integer
   Dim lors As Recordset
   Dim loForm As frmCPClusterDeliveryDetail
   
   Set loForm = New frmCPClusterDeliveryDetail
         
   With loForm
      .Cluster = oTrans.Master("sClustrID")
      .Delivery = oTrans
      .Branch = gridDelivery.TextMatrix(gridDelivery.Row, 1)
      .Show vbModal
      
      If .Cancelled = False Then
         Set lors = .Transfer
         
         'add transfer
         If Not lors.EOF Then
            lors.MoveFirst
            Do Until lors.EOF
               oTrans.Detail(oTrans.ItemCount - 1, "sTransNox") = oTrans.Master("sTransNox")
               oTrans.Detail(oTrans.ItemCount - 1, "sReferNox") = lors("sReferNox")
               oTrans.Detail(oTrans.ItemCount - 1, "sSourceCd") = lors("sSourceCd")
               oTrans.Detail(oTrans.ItemCount - 1, "sBranchCd") = lors("sBranchCd")
               oTrans.Detail(oTrans.ItemCount - 1, "sBranchNm") = lors("sDestinat")
            
               oTrans.saveOthrTransfer
               oTrans.addDetail
               lors.MoveNext
            Loop
         End If
         
         'remove transfer
         Set lors = .UnTransfer
         
         With gridTransfer
            If Not lors.EOF Then
               lors.MoveFirst
                           
               For lnCtr = 0 To oTrans.ItemCount - 1
                  If oTrans.Detail(lnCtr, "sReferNox") = lors("sReferNox") And _
                     oTrans.Detail(lnCtr, "sSourceCd") = lors("sSourceCd") Then
                     
                     oTrans.deleteDetail (lnCtr)
                     Exit For
                  End If
               Next
               
               oTrans.unsaveOthrTransfer lors("sReferNox"), lors("sSourceCd")
            End If
         End With
      End If
      
      Call compTotal
      Call loadDelivery
      
      Unload loForm
   End With
End Sub

Private Sub gridOrder_Click()
   With gridOrder
      .Col = 1
      .ColSel = .Cols - 1
   End With
End Sub

Private Sub gridOrder_DblClick()
   Dim lnCtr As Integer
   
   With gridOrder
      If txtField(0).Text <> "" Then
         If pnOrder = .Row Then Exit Sub
      End If
      
      If pbTModified Then
         If MsgBox("Issuance Transaction is not yet Saved!" & vbCrLf & _
               "Loading new request will reset all changes made." & vbCrLf & vbCrLf & _
               "Continue loading new request?", vbQuestion + vbYesNo, "Confirm") <> vbYes Then
            Exit Sub
         End If
      End If
   
      pnOrder = .Row
      If .TextMatrix(pnOrder, 1) = "" Then Exit Sub
      
      ' check if order is processed
      If .TextMatrix(pnOrder, 5) = "Yes" Then
         If MsgBox("Order was already processed..." & vbCrLf & _
                     "Do you want to create another transfer for that order?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            With oTrans
               For lnCtr = 1 To .ItemCount
                  If oTrans.Detail(lnCtr - 1, "sReferNox") = oTrans.Request(pnOrder - 1, "sTransNox") Then
                     pnDelivery = lnCtr
                     Exit For
                  End If
               Next
            End With
            Call loadTransfer
            Call initButtDt(xeModeUpdate)
         Else
            Call setTransferInfo
            Call initButtDt(xeModeAddNew)
         End If
      Else
         Call setTransferInfo
         Call initButtDt(xeModeAddNew)
      End If
      
      If frmeDetail.Enabled Then txtDetail(1).SetFocus
   End With
End Sub

Private Sub gridTransfer_Click()
   With gridTransfer
      pnTransfer = .Row
   
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

Private Sub oTrans_IssuanceMaster(ByVal Index As Variant, ByVal Value As Variant)
   Select Case Index
   Case 1
      txtField(3) = Format(Value, "Mmm dd, yyyy")
   End Select
End Sub

Private Sub oTrans_LoadDelivery()
   Call loadDelivery
End Sub

Private Sub oTrans_LoadTransaction()
   Call LoadMaster
   Call loadDelivery
   Call LoadOrder
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal Value As Variant)
   With txtOthers(Index)
      If Index = 1 Then
         .Text = Format(Value, "mmmm dd, yyyy")
         txtField(3) = .Text
      Else
         If Index = 2 Then
            'check the capacity of the truck
            txtTotal(0).Text = oTrans.Capacity
            Call compTotal
         ElseIf Index = 3 Then
            If Value <> "" Then
               Call LoadOrder
            Else
               Call initGridOrder
            End If
            
         End If
         .Text = Value
      End If
   End With
End Sub

Private Sub txtDetail_GotFocus(Index As Integer)
   Call HighlightOn(txtDetail(Index))
End Sub

Private Sub txtDetail_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   With gridTransfer
      Select Case Index
         Case 1, 2
            If KeyCode = vbKeyReturn Then
               oTrans.Issuance.Detail(pnTransfer - 1, Index) = txtDetail(Index)
      
               Call setTransDetInfo
               With gridTransfer
                  If .TextMatrix(.Rows - 1, 1) <> "" Then
                     oTrans.Issuance.addDetail
                     .Rows = .Rows + 1
                  End If
                  
                  .Row = .Rows - 1
                  .Col = 1
                  .ColSel = .Cols - 1
                  pnTransfer = .Row
               End With
               Call setMCInfo
            ElseIf KeyCode = vbKeyF3 Then
               If oTrans.Issuance.searchDetail(pnTransfer - 1, Index, txtDetail(Index)) Then
                  Call setTransDetInfo
                  
                  With gridTransfer
                     If .TextMatrix(.Rows - 1, 1) <> "" Then
                        oTrans.Issuance.addDetail
                        .Rows = .Rows + 1
                     End If
                     
                     .Row = .Rows - 1
                     .Col = 1
                     .ColSel = .Cols - 1
                     pnTransfer = .Row
                  End With
                  Call setMCInfo
               End If
            End If
      End Select
   End With
End Sub

Private Sub txtDetail_LostFocus(Index As Integer)
   Call HighlightOff(txtDetail(Index))
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   Call HighlightOn(txtField(Index))
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   Call HighlightOff(txtField(Index))
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 3
      If CDate(txtField(Index)) Then
         oTrans.Issuance.TransDate = CDate(txtField(Index).Text)
      Else
         oTrans.Issuance.TransDate = oTrans.Master("dTransact")
      End If
   Case 4
      oTrans.Issuance.Remarks = txtField(Index).Text
   End Select
End Sub

Private Sub txtOthers_GotFocus(Index As Integer)
   Call HighlightOn(txtOthers(Index))
End Sub

Private Sub txtOthers_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF3, vbKeyReturn
      Select Case Index
      Case 2 To 6
         Call oTrans.SearchMaster(Index, txtOthers(Index))
      End Select
   End Select
End Sub

Private Sub txtOthers_LostFocus(Index As Integer)
   Call HighlightOff(txtOthers(Index))
End Sub

Private Sub txtOthers_Validate(Index As Integer, Cancel As Boolean)
   With txtOthers(Index)
      Select Case Index
      Case 1
         If Not IsDate(.Text) Then
            .Text = oTrans.Master(Index)
         End If
         
         oTrans.Master(Index) = .Text
      Case Else
         oTrans.Master(Index) = .Text
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
      
      If oTrans.Issuance.ItemCount = 0 Then
         detailOK = False
         Exit Function
      End If
   End With

   detailOK = True
End Function

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
         If .Detail(lnCtr, "nNoItemsx") > 0 Then
            lrs.AddNew
            
            lrs("sField01") = .Detail(lnCtr, "sBranchNm")
            lrs("sField02") = .Detail(lnCtr, "sReferNox")
            lrs("sField03") = getSource(oTrans.Detail(lnCtr - 1, "sSourceCd"))
'            lrs("nField01") = .Detail(lnCtr, "nNoItemsx")
         End If
      Next
   End With

   'assign important info to the report
   Set oReport = oRepApp.OpenReport(oApp.AppPath & "\Reports\MCClusterStockDelivery.rpt")
   oReport.DiscardSavedData
   oReport.FieldMappingType = crAutoFieldMapping
   oReport.Database.SetDataSource lrs
   
   oReport.Sections("RH").ReportObjects("txtHeadTitle").SetText oApp.BranchName
   oReport.Sections("RH").ReportObjects("txtCluster").SetText oTrans.Master("sClustrDs")
   oReport.Sections("RH").ReportObjects("txtHeadDescription").SetText oApp.Address & ", " & oApp.TownCity & " " & oApp.Province & " " & oApp.ZippCode
   oReport.Sections("RH").ReportObjects("sTransNox").SetText "MC" & "-" & Right(oTrans.Master("sTransNox"), 11)
   oReport.Sections("RH").ReportObjects("dTransact").SetText Format(oTrans.Master("dTransact"), "Mmm dd, yyyy")
   oReport.Sections("PH").ReportObjects("sRemarksx").SetText IFNull(oTrans.Master("sRemarksx"))
   oReport.Sections("PH").ReportObjects("sDriverxx").SetText IFNull(oTrans.Master("sDriverxx"))
   oReport.Sections("PH").ReportObjects("sHelper01").SetText IFNull(oTrans.Master("sHelper01"))
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

Private Function PrintTransfer() As Boolean
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

   PrintTransfer = True
   
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
         lrs.Fields("sField05") = .Detail(lnCtr, "sBrandNme")
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
   oReport.Sections("PHa").ReportObjects("txtDate").SetText txtField(1).Text
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
   
   PrintTransfer = True
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
   PrintTransfer = False
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


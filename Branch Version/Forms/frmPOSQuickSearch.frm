VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmPOSQuickSearch 
   BorderStyle     =   0  'None
   Caption         =   "MC Serial"
   ClientHeight    =   9585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9585
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   585
      Left            =   105
      Tag             =   "wt0;fb0"
      Top             =   150
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   1032
      BackColor       =   12632256
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmPOSQuickSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmQuickSearh"

Private oSkin As clsFormSkin

Private Sub Form_Load()
   CenterChildForm mdiMain, Me
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransMaintenance
End Sub

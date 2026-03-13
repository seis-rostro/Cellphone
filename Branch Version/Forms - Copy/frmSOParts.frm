VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmSOParts 
   BorderStyle     =   0  'None
   Caption         =   "CP Serial"
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrFrame xrFrame2 
      Height          =   4275
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   570
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   7541
      BorderStyle     =   1
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   7
         Left            =   3450
         TabIndex        =   19
         Text            =   "0,000.00"
         Top             =   3390
         Width           =   2250
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1095
         TabIndex        =   10
         Top             =   810
         Width           =   4605
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1095
         TabIndex        =   9
         Top             =   1110
         Width           =   4605
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1095
         TabIndex        =   8
         Top             =   1410
         Width           =   4605
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1095
         TabIndex        =   7
         Top             =   1710
         Width           =   4605
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1095
         TabIndex        =   6
         Top             =   2010
         Width           =   4605
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   1095
         TabIndex        =   5
         Top             =   2310
         Width           =   4605
      End
      Begin VB.TextBox txtField 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   8
         Left            =   4275
         TabIndex        =   4
         Text            =   "0,000"
         Top             =   2610
         Width           =   1425
      End
      Begin VB.CheckBox chkHsSerial 
         Caption         =   "w/ Serial"
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
         Left            =   1095
         TabIndex        =   3
         Tag             =   "wt0;fb0"
         Top             =   2610
         Width           =   1095
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1095
         TabIndex        =   2
         Top             =   240
         Width           =   2820
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price"
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
         Index           =   27
         Left            =   1575
         TabIndex        =   20
         Top             =   3495
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "Qty On Hand"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2940
         TabIndex        =   18
         Tag             =   "wt0;fb0"
         Top             =   2700
         Width           =   1320
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Index           =   19
         Left            =   195
         TabIndex        =   17
         Top             =   825
         Width           =   795
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barr Code"
         Height          =   195
         Index           =   18
         Left            =   195
         TabIndex        =   16
         Top             =   270
         Width           =   705
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1200
         Tag             =   "et0;ht2"
         Top             =   360
         Width           =   2820
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   315
         Index           =   1
         Left            =   195
         TabIndex        =   15
         Top             =   1125
         Width           =   420
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   14
         Top             =   1410
         Width           =   435
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Made"
         Height          =   195
         Index           =   6
         Left            =   195
         TabIndex        =   13
         Top             =   1725
         Width           =   405
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Index           =   7
         Left            =   195
         TabIndex        =   12
         Top             =   2025
         Width           =   360
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   195
         Index           =   8
         Left            =   195
         TabIndex        =   11
         Top             =   2325
         Width           =   630
      End
      Begin VB.Shape Shape2 
         Height          =   3960
         Left            =   105
         Top             =   120
         Width           =   5700
      End
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ESC"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   4
      Left            =   6225
      TabIndex        =   1
      Top             =   1050
      Width           =   1275
   End
   Begin VB.Label lblField 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F5-OK"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   0
      Left            =   6225
      TabIndex        =   0
      Top             =   555
      Width           =   1275
   End
End
Attribute VB_Name = "frmSOParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmSOParts"

Private oSkin As clsFormSkin

Dim pbCancelled As Boolean
Dim psStockIDxx As String
Dim pnUnitPrice As Currency
Dim pnLastPrice As Currency

Property Get Cancelled() As Boolean
   Cancelled = pbCancelled
End Property

Property Get UnitPrice() As Currency
   UnitPrice = pnUnitPrice
End Property

Property Let StockID(lsStockID As String)
   psStockIDxx = lsStockID
End Property

Private Sub Form_Activate()
   Call LoadInventory(psStockIDxx)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF5, vbKeyReturn
      If Not isEntryOK Then Exit Sub
      Me.Hide
      pnUnitPrice = txtField(7).Text
      pbCancelled = False
   Case vbKeyEscape
      Me.Hide
      pbCancelled = True
   End Select
End Sub

Private Sub Form_Load()
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransDetail
   oSkin.DisableClose = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
End Sub

Private Sub LoadInventory(ByVal lsStockIDx As String)
   Dim loRS As Recordset
   Dim lsSQL As String
   
   Set loRS = New ADODB.Recordset
   lsSQL = "SELECT" & _
               "  a.sBarrCode" & _
               ", a.sDescript" & _
               ", b.sBrandNme" & _
               ", c.sModelNme" & _
               ", d.sMadeName" & _
               ", e.sColorNme" & _
               ", f.sCategrNm" & _
               ", a.cHsSerial" & _
               ", g.nQtyOnHnd" & _
               ", IF(IFNULL(h.nSelPrice, '') = '', a.nSelPrice,h.nSelPrice) nSelPrice" & _
               ", IF(Ifnull(h.nLastPrce, '') = '', a.nLastPrce, h.nLastPrce) nLastPrce"
               
   lsSQL = lsSQL & _
            " FROM CP_Inventory a" & _
               " LEFT JOIN CP_Brand b" & _
                  " ON a.sBrandIDx = b.sBrandIDx" & _
               " LEFT JOIN CP_Model c" & _
                  " ON a.sModelIDx = c.sModelIDx" & _
               " LEFT JOIN CP_Model_Price h" & _
                  " ON c.sModelIDx = h.sModelIDx" & _
                  " AND a.sCategID1 = 'C001001'" & _
               " LEFT JOIN Made d" & _
                  " ON a.sMadeIDxx = d.sMadeIDxx" & _
               " LEFT JOIN Color e" & _
                  " ON a.sColorIDx = e.sColorIDx" & _
               " LEFT Join Category f" & _
                  " ON a.sCategID1 = f.sCategrID" & _
               ", CP_Inventory_Master g" & _
            " WHERE a.sStockIDx = g.sStockIDx" & _
               " AND g.sBranchCd = " & strParm(oApp.BranchCode) & _
               " AND a.sStockIDx = " & strParm(lsStockIDx)
   Debug.Print lsSQL
   loRS.Open lsSQL, oApp.Connection, adOpenStatic, adLockReadOnly, adCmdText
   
   txtField(0).Text = loRS("sBarrCode")
   txtField(1).Text = loRS("sDescript")
   txtField(2).Text = IFNull(loRS("sBrandNme"), "")
   txtField(3).Text = IFNull(loRS("sModelNme"), "")
   txtField(4).Text = IFNull(loRS("sMadeName"), "")
   txtField(5).Text = IFNull(loRS("sColorNme"), "")
   txtField(6).Text = loRS("sCategrNm")
   txtField(7).Text = Format(loRS("nSelPrice"), "#,##0.00")
   txtField(8).Text = loRS("nQtyOnHnd")
   
   chkHsSerial.Value = loRS("cHsSerial")
   pnUnitPrice = loRS("nSelPrice")
   pnLastPrice = loRS("nLastPrce")
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   With txtField(Index)
      If Not IsNumeric(.Text) Then .Text = pnUnitPrice
      If .Text > 99999999.99 Then .Text = pnUnitPrice
      .Text = Format(.Text, "#,##0.00")
      pnUnitPrice = CDbl(.Text)
   End With
End Sub

Private Function isEntryOK() As Boolean
   Dim lsOldProc As String
   Dim lsUserID As String, lsUserName As String
   Dim lnUserRights As Integer
   
   '2016-01-09
   'disable muna accdg to sir rex
   'branch said "what if nakabreak sila how the associate will entry the sales?"
   
'   If CDbl(txtField(7)) = 0# Then
'      '2016-01-08 need manager approval for accessories giveaways with 0.00 amt
'      If oApp.UserLevel < xeManager Then
'         If GetApproval(oApp, lnUserRights, lsUserID, lsUserName, oApp.MenuName) = False Then
'            MsgBox "Approving Officer Has no Right to Save this transaction!!!" & vbCrLf & _
'                     "Request can not be granted!!!", vbCritical, "Warning"
'            GoTo endProc
'         End If
'      End If
'   ElseIf CDbl(txtField(7)) < pnUnitPrice Then
'      If oApp.UserLevel < xeManager Then
'         If GetApproval(oApp, lnUserRights, lsUserID, lsUserName, oApp.MenuName) = False Then
'            MsgBox "Approving Officer Has no Right to Save this transaction!!!" & vbCrLf & _
'                     "Request can not be granted!!!", vbCritical, "Warning"
'            GoTo endProc
'         End If
'      End If
'   End If

   If CDbl(txtField(7)) < pnLastPrice Then
      If CDbl(txtField(7)) <> 0# Then 'allow price if = 0.00 that mean item is giveaway
         If oApp.UserLevel < xeEngineer Then
            If GetApproval(oApp, lnUserRights, lsUserID, lsUserName, oApp.MenuName) = False Then
               MsgBox "Approving Officer Has no Right to Save this transaction!!!" & vbCrLf & _
                        "Request can not be granted!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      End If
   End If

   isEntryOK = True

endProc:
   Exit Function
End Function

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

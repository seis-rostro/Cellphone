VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmSearchStock 
   BorderStyle     =   0  'None
   Caption         =   "Add Stock"
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1770
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1305
      Width           =   3615
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   915
      Width           =   3615
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   525
      Width           =   3615
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Left            =   5055
      TabIndex        =   2
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
      Picture         =   "frmSearchStock.frx":0000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descript"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   180
      TabIndex        =   6
      Top             =   1365
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BarCode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   585
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brand"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   435
      TabIndex        =   4
      Top             =   975
      Width           =   630
   End
End
Attribute VB_Name = "frmSearchStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmSearchModel"

Private oSkin As clsFormSkin

Private poRS As Recordset
Private poRSNew As Recordset

Private psStockIDx As String

Property Set ROQ(ByVal loROQ As Recordset)
   Set poRS = loROQ
End Property

Private Sub cmdButton_Click()
   If findOnROQ Then
      frmCPInvStockRequest.AddModel = poRSNew

      Set poRSNew = Nothing
   Else
      MsgBox "Unable to add requested item." & vbCrLf & _
               "Selected Item Is Not on For-ordering List.", vbCritical, "Warning"
   End If

   Unload Me
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
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

Private Sub Form_Load()
   Dim lsOldProc As String

   lsOldProc = "Form_Load"
   
   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransEqualRight

   createROQRec
   ClearFields
End Sub

Private Function getStock(ByVal fsBarrCode As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lasDetail() As String

   lsOldProc = pxeMODULENAME & ".getModel"

   lsSQL = "SELECT" & _
                     "  a.sStockIDx" & _
                     ", a.sDescript" & _
                     ", a.sBarrCode" & _
                     ", c.sBrandNme" & _
                  " FROM CP_Inventory a" & _
                     " LEFT JOIN CP_Brand c" & _
                        " ON a.sBrandIDx = c.sBrandIDx" & _
                     ", CP_Inventory_Master b" & _
                  " WHERE a.cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND a.sStockIDx = b.sStockIDx" & _
                     " AND b.sBranchCd = " & strParm(oApp.BranchCode) & _
                     " AND a.sCategID1 <> 'C001001'" & _
                     " AND a.cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND a.sBarrCode LIKE " & strParm(fsBarrCode & "%") & _
                  " ORDER BY a.sBarrCode"


   Set lors = New Recordset
   With lors
      Debug.Print lsSQL
      .Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set .ActiveConnection = Nothing

      If .EOF Then
         GoTo endWithClear
      ElseIf .RecordCount = 1 Then
         psStockIDx = lors("sStockIDx")
         txtField(0) = lors("sBarrCode")
         txtField(1) = lors("sBrandNme")
         txtField(2) = lors("sDescript")
      Else
         lsSQL = KwikBrowse(oApp, lors, "sStockIDx»sBarrCode»sBrandNme»sDescript", "Stock ID»Bar Code»Brand»Description")
         If lsSQL <> "" Then
            lasDetail = Split(lsSQL, "»")

            psStockIDx = lasDetail(0)
            txtField(0) = lasDetail(2)
            txtField(1) = lasDetail(3)
            txtField(2) = lasDetail(1)
         Else
            GoTo endWithClear
         End If
      End If
   End With

   getStock = True

endProc:
   Exit Function
endWithClear:
   psStockIDx = ""
   txtField(0) = ""
   txtField(1) = ""
   txtField(2) = ""
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function

Private Sub ClearFields()
   txtField(0) = ""
   txtField(1) = ""
   txtField(2) = ""
End Sub

Private Sub createROQRec()
   Set poRSNew = New Recordset

   With poRSNew
      .Fields.Append "sStockIDx", adVarChar, 12
      .Fields.Append "sBarrCode", adVarChar, 20
      .Fields.Append "sDescript", adVarChar, 128
      .Fields.Append "sBrandNme", adVarChar, 25
      .Fields.Append "nAveMonSl", adInteger
      .Fields.Append "nAveMonMd", adInteger
      .Fields.Append "nMinLevel", adInteger
      .Fields.Append "nMaxLevel", adInteger
      .Fields.Append "cClassify", adChar, 1
      .Fields.Append "cClassMdl", adChar, 1
      .Fields.Append "cInvTypex", adChar, 1
      .Fields.Append "nQuantity", adInteger
      .Fields.Append "nOnTranst", adInteger
      .Fields.Append "nOnTrnsMd", adInteger
      .Fields.Append "nRecOrder", adInteger
      .Fields.Append "nRecOrdMd", adInteger
      .Fields.Append "nQtyOnHnd", adInteger
      .Fields.Append "nQtyOnHMd", adInteger
      .Fields.Append "sBrandIDx", adVarChar, 9
      .Open
   End With
End Sub

Private Function findOnROQ() As Boolean
   Dim lors As Recordset
   Dim lsOldProc As String
   Dim lnPos As Integer
   Dim lsSQL As String

   lsOldProc = "findOnROQ"

   ''On Error GoTo errProc

   If psStockIDx = "" Then
      MsgBox "No Model Selected.", vbInformation, "Notice"
      findOnROQ = False
      GoTo endProc
   End If

   If TypeName(poRS) = "Nothing" Then
      findOnROQ = False
      GoTo endProc
   End If

   With poRS
      Debug.Print psStockIDx
      Call .Find("sStockIDx = " & strParm(psStockIDx), 0, adSearchForward, adBookmarkFirst)

      If Not .EOF Then
         poRSNew.AddNew
         poRSNew.Fields("sStockIDx") = .Fields("sStockIDx")
         poRSNew.Fields("sBarrCode") = .Fields("sBarrCode")
         poRSNew.Fields("sDescript") = .Fields("sDescript")
         poRSNew.Fields("sBrandNme") = .Fields("sBrandNme")
         poRSNew.Fields("nAveMonSl") = .Fields("nAveMonSl")
         poRSNew.Fields("nAveMonMd") = .Fields("nAveMonMd")
         poRSNew.Fields("nMinLevel") = .Fields("nMinLevel")
         poRSNew.Fields("nMaxLevel") = .Fields("nMaxLevel")
         poRSNew.Fields("cClassify") = .Fields("cClassify")
         poRSNew.Fields("cClassMdl") = .Fields("cClassMdl")
         poRSNew.Fields("nQuantity") = .Fields("nQuantity")
         poRSNew.Fields("nOnTranst") = .Fields("nOnTranst")
         poRSNew.Fields("nOnTrnsMd") = .Fields("nOnTrnsMd")
         poRSNew.Fields("nRecOrder") = .Fields("nRecOrder")
         poRSNew.Fields("nRecOrdMd") = .Fields("nRecOrdMd")
         poRSNew.Fields("nQtyOnHnd") = .Fields("nQtyOnHnd")
         poRSNew.Fields("nQtyOnHMd") = .Fields("nQtyOnHMd")
         poRSNew.Fields("sBrandIDx") = .Fields("sBrandIDx")
      Else
         lsSQL = "SELECT a.sStockIDx" & _
                     ", f.sBarrCode" & _
                     ", f.sDescript" & _
                     ", d.sBrandNme" & _
                     ", a.nAveMonSl" & _
                     ", a.nMinLevel" & _
                     ", a.nMaxLevel" & _
                     ", a.cClassify" & _
                     ", e.nOnTranst" & _
                     ", a.nQtyOnHnd" & _
                     ", f.nSelPrice" & _
                     ", d.sBrandIDx" & _
                     ", f.sDescript" & _
                     ", f.cInvTypex" & _
                  " FROM CP_Inventory_Master a"
         lsSQL = lsSQL & _
                        " LEFT JOIN " & _
                           " ( SELECT c.sStockIDx" & _
                                 ", COUNT(c.sStockIDx) nOnTranst" & _
                              " FROM CP_Transfer_Detail a" & _
                                 ", CP_Transfer_Master b" & _
                                 ", CP_Inventory c" & _
                              " WHERE a.sTransNox = b.sTransNox" & _
                                 " AND a.sStockIDx = c.sStockIDx" & _
                                 " AND b.sDestinat = " & strParm(oApp.BranchCode) & _
                                 " AND NOT ( b.cTranStat = " & strParm(xeStateCancelled) & _
                                    " OR b.cTranStat = " & strParm(xeStatePosted) & ")" & _
                                 " AND c.sCategID1 <> 'C001001'" & _
                              " GROUP BY c.sStockIDx" & _
                              " ORDER BY c.sStockIDx) e" & _
                           " ON a.sStockIDx = e.sStockIDx" & _
                     ", CP_Model b" & _
                     ", CP_Brand d" & _
                     ", CP_Inventory f"
         lsSQL = lsSQL & _
                  " WHERE a.sStockIDx = f.sStockIDx" & _
                     " AND f.sModelIDx = b.sModelIDx" & _
                     " AND b.sBrandIDx = d.sBrandIDx" & _
                     " AND a.sBranchCd = " & strParm(oApp.BranchCode) & _
                     " AND a.cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND f.sCategID1 <> 'C001001'" & _
                     " AND a.sStockIDx = " & strParm(psStockIDx) & _
                  " GROUP BY f.sModelIDx, f.sColorIDx" & _
                  " ORDER BY d.sBrandNme, f.sBarrCode"
                  
         Set lors = New Recordset
         lors.Open lsSQL, oApp.Connection, , adCmdText
         Set lors.ActiveConnection = Nothing
         
         If lors.RecordCount = 0 Then
            findOnROQ = False
            .Cancel
            Exit Function
         End If
         
         poRSNew.AddNew
         poRSNew.Fields("sStockIDx") = lors("sStockIDx")
         poRSNew.Fields("sBarrCode") = lors("sBarrCode")
         poRSNew.Fields("sDescript") = lors("sDescript")
         poRSNew.Fields("sBrandNme") = lors("sBrandNme")
         poRSNew.Fields("nAveMonSl") = 0
         poRSNew.Fields("nAveMonMd") = 0
         poRSNew.Fields("nMinLevel") = lors("nMinLevel")
         poRSNew.Fields("nMaxLevel") = lors("nMaxLevel")
         poRSNew.Fields("cClassify") = "F"
         poRSNew.Fields("cClassMdl") = "F"
         poRSNew.Fields("nQuantity") = 0
         poRSNew.Fields("nOnTranst") = 0
         poRSNew.Fields("nOnTrnsMd") = 0
         poRSNew.Fields("nRecOrder") = 0
         poRSNew.Fields("nRecOrdMd") = 0
         poRSNew.Fields("nQtyOnHnd") = 0
         poRSNew.Fields("nQtyOnHMd") = 0
         poRSNew.Fields("sBrandIDx") = lors("sBrandIDx")
      End If

      .Cancel
   End With

   findOnROQ = (poRSNew.RecordCount <> 0)
endProc:
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
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

Private Sub Form_Unload(Cancel As Integer)
   Set poRS = Nothing
   Set poRSNew = Nothing
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case Index
      Case 0
         If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
            If getStock(txtField(Index)) Then SetNextFocus
         End If
   End Select
End Sub

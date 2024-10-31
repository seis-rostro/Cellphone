VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmSearchUnit 
   BorderStyle     =   0  'None
   Caption         =   "Add Stock"
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1395
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
      Picture         =   "frmSearchUnit.frx":0000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
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
      TabIndex        =   4
      Top             =   585
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
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
      Left            =   225
      TabIndex        =   3
      Top             =   975
      Width           =   570
   End
End
Attribute VB_Name = "frmSearchUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pxeMODULENAME = "frmSearchUnit"

Private oSkin As clsFormSkin

Private poRS As Recordset
Private poRSNew As Recordset

Private psStockIDx As String

Property Set ROQ(ByVal loROQ As Recordset)
   Set poRS = loROQ
End Property

Private Sub cmdButton_Click()
   If findOnROQ Then
      frmCPInvUnitRequest.AddModel = poRSNew

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

Private Function getStock(ByVal fsModelNme As String) As Boolean
   Dim lsOldProc As String
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lasDetail() As String

   lsOldProc = pxeMODULENAME & ".getModel"

   lsSQL = "SELECT" & _
                     "  a.sModelIDx" & _
                     ", a.sModelNme" & _
                     ", a.sBrandIDx" & _
                     ", c.sBrandNme" & _
                     ", e.sColorNme" & _
                     ", d.sStockIDx" & _
                  " FROM CP_Model a" & _
                     ", CP_Brand c" & _
                     ", CP_Inventory d" & _
                     ", Color e" & _
                  " WHERE a.cRecdStat = " & strParm(xeRecStateActive) & _
                     " AND a.sModelIDx = d.sModelIDx" & _
                     " AND a.sBrandIDx = c.sBrandIDx" & _
                     " AND d.sColorIDx = e.sColorIDx" & _
                     " AND d.cHsSerial = " & strParm(xeYes) & _
                     " AND a.sModelNme LIKE " & strParm(fsModelNme & "%") & _
                  " GROUP BY a.sModelIDx, d.sColorIDx" & _
                  " ORDER BY a.sModelNme"
'2016-03-14 she
' remove this table to be able to order any model kahit wala sa inventory nila
'", CP_Inventory_Master b"
'" AND b.sBranchCd = " & strParm(oApp.BranchCode)
'" AND b.sStockIDx = d.sStockIDx"

   Set lors = New Recordset
   With lors
      Debug.Print lsSQL
      .Open lsSQL, oApp.Connection, adOpenStatic, adLockOptimistic, adCmdText
      Set .ActiveConnection = Nothing

      If .EOF Then
         GoTo endWithClear
      ElseIf .RecordCount = 1 Then
         psStockIDx = lors("sStockIDx")
         txtField(0) = lors("sModelNme")
         txtField(1) = lors("sColorNme")
      Else
         lsSQL = KwikBrowse(oApp, lors, "sBrandNme»sModelNme»sColorNme", "Brand»Model»Color")
         If lsSQL <> "" Then
            lasDetail = Split(lsSQL, "»")

            psStockIDx = lasDetail(5)
            txtField(0) = lasDetail(1)
            txtField(1) = lasDetail(4)
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
   Exit Function
errProc:
   ShowError lsOldProc & "( " & " )", True
End Function

Private Sub ClearFields()
   txtField(0) = ""
   txtField(1) = ""
End Sub

Private Sub createROQRec()
   Set poRSNew = New Recordset

   With poRSNew
      .Fields.Append "sStockIDx", adVarChar, 12
      .Fields.Append "sModelIDx", adVarChar, 9
      .Fields.Append "sBrandNme", adVarChar, 25
      .Fields.Append "sModelNme", adVarChar, 50
      .Fields.Append "sColorNme", adVarChar, 15
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
      .Fields.Append "sColorIDx", adVarChar, 7
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
         poRSNew.Fields("sModelIDx") = .Fields("sModelIDx")
         poRSNew.Fields("sBrandNme") = .Fields("sBrandNme")
         poRSNew.Fields("sModelNme") = .Fields("sModelNme")
         poRSNew.Fields("sColorNme") = .Fields("sColorNme")
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
         poRSNew.Fields("sColorIDx") = .Fields("sColorIDx")
      Else
         lsSQL = "SELECT" & _
                    "  a.sStockIDx" & _
                    ", d.sBrandNme" & _
                    ", b.sModelNme" & _
                    ", c.sColorNme" & _
                    ", '0' nAveMonSl" & _
                    ", '0' nMinLevel" & _
                    ", '0' nMaxLevel" & _
                    ", 'F' cClassify" & _
                    ", e.nOnTranst" & _
                    ", '0' nQtyOnHnd" & _
                    ", a.nSelPrice" & _
                    ", b.sModelIDx" & _
                    ", c.sColorIDx" & _
                    ", d.sBrandIDx" & _
                    ", a.cInvTypex" & _
                  " FROM CP_Inventory a"
      lsSQL = lsSQL & _
                      " LEFT JOIN (SELECT" & _
                                       "  c.sStockIDx" & _
                                       ", COUNT(c.sSerialID)    nOnTranst" & _
                                     " FROM CP_Transfer_Detail a," & _
                                       " CP_Transfer_Master b," & _
                                       " CP_Inventory_Serial c" & _
                                     " WHERE a.sTransNox = b.sTransNox" & _
                                       " AND a.sSerialID = c.sSerialID" & _
                                       " AND NOT(b.cTranStat = '3'" & _
                                          " OR b.cTranStat = '2')" & _
                                     " GROUP BY c.sStockIDx" & _
                                     " ORDER BY c.sStockIDx) e" & _
                        " ON a.sStockIDx = e.sStockIDx" & _
                    ", CP_Model b" & _
                    ", Color c" & _
                    ", CP_Brand d" & _
                  " WHERE a.sModelIDx = b.sModelIDx" & _
                     " AND a.sBrandIDx = d.sBrandIDx" & _
                     " AND a.sColorIDx = c.sColorIDx" & _
                     " AND a.cHsSerial = " & strParm(xeYes) & _
                     " AND a.sStockIDx = " & strParm(psStockIDx) & _
                  " GROUP BY a.sModelIDx, a.sColorIDx" & _
                  " ORDER BY d.sBrandNme, b.sModelNme, c.sColorNme"
                     
            'change filter  AND a.sCategID1 = 'C001001' to AND a.cHsSerial = " & strParm(xeYes)
            'to search all serialized item not only the cellphone catgory
            Debug.Print lsSQL
            Set lors = New Recordset
            lors.Open lsSQL, oApp.Connection, , adCmdText
            Set lors.ActiveConnection = Nothing
            
            If Not lors.EOF Then
               poRSNew.AddNew
               poRSNew.Fields("sStockIDx") = lors("sStockIDx")
               poRSNew.Fields("sModelIDx") = lors("sModelIDx")
               poRSNew.Fields("sBrandNme") = lors("sBrandNme")
               poRSNew.Fields("sModelNme") = lors("sModelNme")
               poRSNew.Fields("sColorNme") = lors("sColorNme")
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
               poRSNew.Fields("sColorIDx") = lors("sColorIDx")
            Else
               findOnROQ = False
               .Cancel
               
               Exit Function
            End If
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

VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLookUp 
   BorderStyle     =   0  'None
   Caption         =   "Look Up Table"
   ClientHeight    =   7275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLookUp.frx":0000
   ScaleHeight     =   7275
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00253315&
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Tag             =   "et0;eb0"
      Top             =   1680
      Width           =   3795
   End
   Begin VB.ComboBox cmbSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5790
      TabIndex        =   1
      Tag             =   "et0;eb0"
      Text            =   "Sort Key"
      Top             =   1680
      Width           =   1920
   End
   Begin xrControl.xrButton xrButton1 
      Height          =   330
      Index           =   2
      Left            =   6705
      TabIndex        =   13
      Top             =   1230
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   582
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
   End
   Begin xrControl.xrButton xrButton1 
      Height          =   330
      Index           =   1
      Left            =   6705
      TabIndex        =   12
      Top             =   870
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   582
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
   End
   Begin xrControl.xrButton xrButton1 
      Default         =   -1  'True
      Height          =   330
      Index           =   0
      Left            =   6705
      TabIndex        =   11
      Top             =   510
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   582
      Caption         =   "&Load"
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
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   105
      TabIndex        =   2
      Top             =   6915
      Visible         =   0   'False
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4890
      Left            =   60
      TabIndex        =   10
      Tag             =   "et0;eb0;et0;fb0"
      Top             =   2295
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   8625
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   7783164
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GMC-Software Engineering Group"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   5
      Left            =   1920
      TabIndex        =   9
      Tag             =   "hb1"
      Top             =   1200
      Width           =   2115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Index           =   1
      Left            =   1905
      TabIndex        =   8
      Tag             =   "hb2"
      Top             =   465
      Width           =   1875
   End
   Begin VB.Image Image4 
      Height          =   1455
      Left            =   120
      Picture         =   "frmLookUp.frx":361B
      Top             =   555
      Width           =   1500
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      Height          =   1515
      Left            =   90
      Top             =   525
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   1935
      TabIndex        =   7
      Tag             =   "hb1"
      Top             =   1485
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fields"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   5790
      TabIndex        =   6
      Tag             =   "hb1"
      Top             =   1485
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2003 and beyond"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   4
      Left            =   1920
      TabIndex        =   5
      Tag             =   "hb1"
      Top             =   990
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   1905
      TabIndex        =   4
      Tag             =   "1-1"
      Top             =   495
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 2.00"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   6
      Left            =   1920
      TabIndex        =   3
      Tag             =   "hb1"
      Top             =   780
      Width           =   765
   End
End
Attribute VB_Name = "frmLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Rex S. Adversalo
' XerSys Computing
' Canaoalan, Binmaley, Pangasinan
'
' LookUp (RecordSet) v1.5
'     Display lookup table and allows user to select from a list.
'     Properties:
'        RowSource = Recordset that contains the selection
'        Column    = a string or array of array of field name the will appear on
'                    lookup table
'        ColHead   = a string or array of string of column heading
'        SortKey   = the default column to be use as sort key
'
' Copyright 2002 and beyond
' All Rights Reserved
'
' ººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All rights reserved. No part of this  €€  This Software is Owned by        €
' €  software may be reproduced or trans-  €€                                   €
' €  mitted in any  form or by any means,  €€    GUANZON MERCHANDISING CORP.    €
' €  electronic or mechanical,  including  €€     Guanzon Bldg. Perez Blvd.     €
' €  recording, or by information storage  €€           Dagupan City            €
' €  and retrieval systems, without prior  €€  Tel No. 522-1085 ; 522-0863      €
' €  written permission from the author.   €€                                   €
' ººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ================================================================================
'  01/04/2003 | Rex | Start creating this object.
'  04/10/2003 | Rex | Add another properties, the AutoDisplay and the Column
'  04/22/2003 | Rex | Revise this object, remove the search field and set Auto-
'             |     |   Display to true. Reason: the argument includes the
'             |     |   data source, so it takes only a fraction of a second to
'             |     |   to fill the items to the lookup window.
'  04/24/2003 | Rex | The lookup executes as i wanted, but there is a problem:
'             |     |   ListView fills too slow when recordset exceeds 21000
'             |     |   records, so i decided to use MSFlexGrid.
'  06/21/2004 | Rex | Rewrite this object.
'             |     |   Addt'l/Modified property
'             |     |      * Column Name (variant/array of string)
'             |     |      * Column Head (variant/array of string)
'             |     |      * Column Format (variant/array of string)
'             |     |      * SQL Statement (SQL Source of the recordset) /
'             |     |      * Recordset
'             |     |      * Connection
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const xeColMargin As Integer = 30
Private Const xeCharWidth As Integer = 110
Private Const xeScrollBar As Integer = 240
Private Const xeMaxItem As Integer = 20
Private Const xeMaxRecd As Integer = 32767

Private p_oAppDrivr As AppDriver
Private WithEvents p_oLookup As Recordset
Attribute p_oLookup.VB_VarHelpID = -1
Private p_oSkin As FormSkin
Private p_oMod As MainModules

Private p_sSQLQuery As String
Private p_asFldName() As String
Private p_asColName() As String
Private p_asColHead() As String
Private p_asColPict() As String
Private p_acColType() As String
Private p_anColWdth() As Integer
Private p_sColHead As String
Private p_sColName As Variant
Private p_sFldName As String
Private p_sColPict As Variant
Private p_nSearch As Integer
Private p_bSearch As Boolean

Private p_bSelected As Boolean
Private p_bRowSource As Boolean
Private p_bDisplayd As Boolean

Private pnCtr As Integer
Private pnInterval As Integer
Private pnProgress As Integer
Private pbProgress As Boolean
Private pbFocus As Boolean

Property Set AppDriver(oAppDriver As AppDriver)
10       Set p_oAppDrivr = oAppDriver
End Property

Property Set RowSource(Source As Recordset)
   ' the record source of the Lookup
10       Set p_oLookup = Source
20       p_bRowSource = True
End Property

Property Let SQLSource(Source As String)
10       p_sSQLQuery = Source
End Property

Property Let FldTitle(Title As String)
   ' the column heading of the lookup
   ' the heading item/s must correspond to the order of the column
   '     of the rec source. This will be the only visible column
   '     description that identifies its content
   
10       p_sColHead = Title
End Property

Property Let FldName(Name As String)
   ' added this property to customize the # of column and the order
   '     of column to be displayed
   
10       p_sColName = Name
End Property

Property Let FldCriteria(Value As String)
   ' added this property to implement the runtime filtering of recordset
10       p_sFldName = Value
End Property

Property Let FldFormat(Format As String)
   ' added this property to allow field formating
   
10       p_sColPict = Format
End Property

Property Let showSearch(Value As Boolean)
   ' this will allow the lookup to requery the recordset using the criteria entered
   
10       p_bSearch = Value
End Property

Property Get SelectedItem() As Variant
   ' the selected item
   
10       If p_bSelected Then
20          SelectedItem = getSelectedItem()
30       Else
40          SelectedItem = Empty
50       End If
End Property

Private Sub cmbSearch_GotFocus()
10       pbFocus = True
20       With cmbSearch
30          .Tag = .ListIndex
40       End With
End Sub

Private Sub cmbSearch_LostFocus()
   ' this will allow the user to modify the search key
10       pbFocus = False
20       With cmbSearch
30          If .ListIndex = -1 Or .ListIndex = .Tag Then Exit Sub
40       End With
50       SortList
End Sub

Private Sub Form_Activate()
10       If Not p_bDisplayd Then
20          LoadList
   If p_oLookup.RecordCount = 1 Then
      MSFlexGrid1.RowSel = 1
      xrButton1_Click (0)
      Exit Sub
   End If
30          txtSearch.SetFocus
40       End If
End Sub

Private Sub Form_Initialize()
10       Set p_oSkin = New FormSkin
20       Set p_oMod = New MainModules
   
30       p_bRowSource = False
40       p_bSelected = False
50       p_bDisplayd = False
60       p_bSearch = False
70       p_nSearch = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10       Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
20          If KeyCode <> vbKeyReturn And pbFocus Then Exit Sub
30          Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
40             p_oMod.SetNextFocus
50          Case vbKeyUp
60             p_oMod.SetPreviousFocus
70          End Select
80       End Select
End Sub

Private Sub Form_Load()
10       Set p_oSkin.Form = Me
20       Set p_oSkin.AppDriver = p_oAppDrivr
30       p_oSkin.ApplySkin xeFormQuickSearch
End Sub

Private Sub Form_Unload(Cancel As Integer)
10       If p_bRowSource = False Then Set p_oLookup = Nothing
20       Set p_oMod = Nothing
30       Set p_oSkin = Nothing
End Sub

' assigns the contents of the recordset to the grid
Public Function LoadList() As Boolean
10       Dim lvValue As Variant
20       Dim lnAlignment As Integer
30       Dim lanColWidth() As Long
40       Dim lnTotWidth As Long
50       Dim lsOldProc As String
   
60       lsOldProc = p_oAppDrivr.ProcName("LoadList")
70       LoadList = False
80       On Error Goto errProc
   
90       If p_bRowSource = False And p_sSQLQuery = Empty Then GoTo endProc
100      getFieldInfo
110      showButton
   
   ' assign the column head to the combo box
120      cmbSearch.Clear
130      ReDim lanColWidth(UBound(p_anColWdth))
140      For pnCtr = LBound(p_asColHead) To UBound(p_asColHead)
150         If p_asColHead(pnCtr) = "" Then Exit For
160         cmbSearch.AddItem (p_asColHead(pnCtr))
      
      ' assign the length of the headers as the max width of the columns
170         lanColWidth(pnCtr) = Len(Trim(p_asColHead(pnCtr)))
180      Next
190      cmbSearch.ListIndex = IIf(UBound(p_asColHead) > 1, 1, 0)
200      p_oLookup.Sort = p_asColName(cmbSearch.ListIndex)

   ' reformat the flexgrid
210      With MSFlexGrid1
220         .Cols = UBound(p_asColName) + 1
230         .Rows = 2
240         .Row = 0
      
250         For pnCtr = LBound(p_asColName) To UBound(p_asColName)
         ' get the appropraite alignment of each field
260            If p_acColType(pnCtr) = "s" Then
270               lnAlignment = flexAlignLeftTop
280            Else
290               lnAlignment = flexAlignRightTop
300            End If
         
310            .Col = pnCtr
320            .CellAlignment = lvwColumnLeft
330            .CellFontBold = True
         
340            .TextMatrix(0, pnCtr) = p_asColHead(pnCtr)
350            .ColAlignment(pnCtr) = lnAlignment
360            .ColWidth(pnCtr) = p_anColWdth(pnCtr) * xeCharWidth
370         Next

      ' always move the row to 1 to highlight the record not the header
380         If p_bSearch Then
390            .Row = 1
400            .Col = 0
410            .ColSel = .Cols - 1
420            p_bDisplayd = False
430            LoadList = True
440            GoTo endProc
450         End If
      
      ' check if there's a record to display
460         If p_oLookup.RecordCount = 0 Then GoTo endProc
   
470         If p_oLookup.RecordCount > xeMaxRecd Then
480            MsgBox "Search Record Result Exceeds The Maximum Allowable Record Display!!!" & _
               vbCrLf & "Please Limit Your Selection by Specifying More Detailed Info!!!", vbCritical, "Warning"
490            GoTo endProc
500         End If
510         .Rows = p_oLookup.RecordCount + 1
520         p_oLookup.MoveFirst
      
530         showProgress .Rows + 1
540         .Row = 0
550         Do Until p_oLookup.EOF
560            .Row = .Row + 1
570            For pnCtr = 0 To UBound(p_asColName)
580               lvValue = p_oLookup(p_asColName(pnCtr))
590               If IsNull(p_oLookup(p_asColName(pnCtr))) Then lvValue = Empty
600               .TextMatrix(.Row, pnCtr) = Format(lvValue, p_asColPict(pnCtr))
            
610               If Len(Trim(.TextMatrix(.Row, pnCtr))) > lanColWidth(pnCtr) Then
620                  lanColWidth(pnCtr) = Len(Trim(.TextMatrix(.Row, pnCtr)))
630               End If
640            Next
         
650            p_oLookup.MoveNext
660         Loop

670         .Row = 1
680         .Col = 0
690         .ColSel = .Cols - 1

      ' after fetching all record to the grid, adjust the column width
700         lnTotWidth = 0
710         For pnCtr = 0 To .Cols - 1
720            If lanColWidth(pnCtr) < p_anColWdth(pnCtr) Then p_anColWdth(pnCtr) = lanColWidth(pnCtr)
730            .ColWidth(pnCtr) = p_anColWdth(pnCtr) * xeCharWidth
740            lnTotWidth = lnTotWidth + p_anColWdth(pnCtr)
750         Next
      
760         If .Rows > xeMaxItem Then
770            If (lnTotWidth * xeCharWidth) < .Width - xeScrollBar Then
780               For pnCtr = 0 To .Cols - 1
790                  .ColWidth(pnCtr) = (p_anColWdth(pnCtr) * _
                                    ((.Width - xeScrollBar) / xeCharWidth) / lnTotWidth) * xeCharWidth - xeColMargin
800               Next
810            End If
820         Else
830            If (lnTotWidth * xeCharWidth) < .Width - xeColMargin Then
840               For pnCtr = 0 To .Cols - 1
850                  .ColWidth(pnCtr) = (p_anColWdth(pnCtr) * _
                                    (.Width / xeCharWidth) / lnTotWidth) * xeCharWidth - xeColMargin
860               Next
870            End If
880         End If
890      End With

900      hideProgress
910      p_bDisplayd = True
920      LoadList = True
   
endProc:
930      p_oAppDrivr.ProcName lsOldProc
940      Exit Function
errProc:
950      ShowError lsOldProc
End Function

' retrieves the table and set the field property
Private Sub getFieldInfo()
10       Dim lsSQL As String
20       Dim lsOldProc As String
   
30       lsOldProc = p_oAppDrivr.ProcName("getFieldInfo")
40       On Error Goto errProc
   
   ' if SQL query is passed retrieve the records
50       If Not p_bRowSource Then
60          Set p_oLookup = New Recordset
70          lsSQL = p_sSQLQuery
80          If p_bSearch Then lsSQL = p_oMod.AddCondition(lsSQL, "0 = 1")
90          p_oLookup.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, , adCmdText
100      End If
   
   ' check if client passed a field filter
110      If p_sColName <> "" Then
120         p_asColName = Split(p_sColName, "»", , vbTextCompare)
130      Else
      ' if not include all fields in the lookup
140         ReDim p_asColName(p_oLookup.Fields.Count - 1) As String
150         For pnCtr = 0 To UBound(p_asColName)
160            p_asColName(pnCtr) = p_oLookup.Fields(pnCtr).Name
170         Next
180      End If
   
190      If p_sColHead <> Empty Then
200         p_asColHead = Split(p_sColHead, "»", -1, vbTextCompare)
210      Else
220         ReDim p_asColHead(UBound(p_asColName)) As String
230         For pnCtr = 0 To UBound(p_asColName)
240            p_asColHead(pnCtr) = p_asColName(pnCtr)
250         Next
260      End If
   
   ' after retrieving the field name, create a field criteria
   ' to be used in creating sql statement at runtime
270      If p_sFldName <> Empty Then
280         p_asFldName = Split(p_sFldName, "»", , vbTextCompare)
290      Else
300         ReDim p_asFldName(UBound(p_asColName)) As String
310         For pnCtr = 0 To UBound(p_asColName)
320            p_asFldName(pnCtr) = p_asColName(pnCtr)
330         Next
340      End If

   ' after retrieving the column, set the type and the width
350      ReDim p_acColType(UBound(p_asColName))
360      ReDim p_asColPict(UBound(p_asColName))
370      ReDim p_anColWdth(UBound(p_asColName))
380      For pnCtr = 0 To UBound(p_asColName)
390         p_anColWdth(pnCtr) = p_oLookup(p_asColName(pnCtr)).DefinedSize
400         p_asColPict(pnCtr) = "@"
      
410         If p_anColWdth(pnCtr) < Len(p_asColHead(pnCtr)) Then
420            p_anColWdth(pnCtr) = Len(p_asColHead(pnCtr))
430         End If
      
440         Select Case p_oLookup(p_asColName(pnCtr)).Type
      Case 129, 130, 202, 200    ' string
450            p_acColType(pnCtr) = "s"
460         Case 2, 3, 11, 17, 72      ' numeric without decimal point
470            p_acColType(pnCtr) = "n"
480         Case 4, 5, 6, 131          ' numeric with decimal point
490            p_acColType(pnCtr) = "l"
500         Case 135                   ' datetime
510            p_acColType(pnCtr) = "d"
520         End Select
530      Next
540      If p_sColPict <> Empty Then p_asColPict = Split(p_sColPict, "»", -1, vbTextCompare)
   
endProc:
550      p_oAppDrivr.ProcName lsOldProc
560      Exit Sub
errProc:
570      ShowError lsOldProc
End Sub

Private Function ResultingText(iKeyAscii%) As String
   'Purpose: Works out the text string that results from an original string
   '         comprising the specified elements, following addition of <KeyAscii>
   '         at <iSelStart>
   '
   'Returns: Resulting text string
   
10       Dim sLeft As String             ' string element
20       Dim sSel As String              ' selected string element
30       Dim sRight As String            ' string element
40       Dim sResult As String           ' what well return
   
50       On Error Resume Next
   
60       With txtSearch
70          sLeft = Left$(.Text, .SelStart)         ' SelStart is 0-based
80          sSel = Mid$(.Text, .SelStart + 1, .SelLength)
90          sRight = Mid$(.Text, .SelStart + .SelLength + 1)
100      End With
   
110      Select Case iKeyAscii
      Case vbKeyBack             'Backspace Key
120            If Len(sSel) = 0 Then   'Nothing selected
130               sResult = MinusRightChar(sLeft) & sRight  'Del first char on the left
140            Else                    'Selection exists
150               sResult = sLeft & sRight   'Delete selected text only
160            End If
         
170         Case vbKeyDelete           'Delete key
180            If Len(sSel) = 0 Then   'Nothing selected
190               sResult = sLeft & MinusLeftChar(sRight)    'Del first char on the right
200            Else
210               sResult = sLeft & sRight    'Delete selected text only
220            End If
         
230         Case Else         'an ordinary character
240            sResult = sLeft & Chr$(iKeyAscii) & sRight
250      End Select
260      ResultingText = sResult
End Function

Private Function MinusLeftChar(ByVal sGiven As String) As String

   'Purpose: Returns <sGiven> with the leftmost character removed, or "" if
   '         <sGiven> was empty.
   '
   'Returns: The trimmed string
   '
   'Remarks: Just a safe wrapper for Mid$()
10       On Error Resume Next
   
20       If Len(sGiven) = 0 Then
30          MinusLeftChar = ""
40       Else
50          MinusLeftChar = Mid$(sGiven, 2)
60       End If
End Function

Private Function MinusRightChar(ByVal sGiven As String) As String

   'Purpose: Returns <sGiven> with the rightmost character removed, or "" if
   '         <sGiven> was empty.
   '
   'Returns: The trimmed string
   '
   'Remarks: Just a safe wrapper for Left$()
10       On Error Resume Next
   
20       If Len(sGiven) = 0 Then
30          MinusRightChar = ""
40       Else
50          MinusRightChar = Left$(sGiven, Len(sGiven) - 1)
60       End If
End Function

Private Sub MSFlexGrid1_LostFocus()
10       MSFlexGrid1.BackColorSel = &H800000
End Sub


Private Sub MSFlexGrid1_DblClick()
10       With MSFlexGrid1
20          If .MouseRow = 0 Then
30             If .MouseCol <> (cmbSearch.ListIndex) Then
40                cmbSearch.ListIndex = .MouseCol
50                SortList
60             End If
70          Else
80             xrButton1_Click 0
90          End If
100      End With
End Sub

Private Sub MSFlexGrid1_GotFocus()
10       With MSFlexGrid1
20          .HighLight = flexHighlightAlways
30          .BackColorSel = &HB06F00
40       End With
End Sub

Private Function SearchOn(ByVal lsSeek) As Boolean
10       Dim lnCtr As Long
20       Dim lbFound As Boolean
   
30       lbFound = False
40       With MSFlexGrid1
50          For lnCtr = 1 To .Rows
60             If StrComp(Left(.TextMatrix(lnCtr, cmbSearch.ListIndex), Len(lsSeek)), lsSeek, vbTextCompare) >= 0 Then
70                .TopRow = lnCtr
80                .Row = lnCtr
90                .RowSel = lnCtr
100               .ColSel = MSFlexGrid1.Cols - 1
110               lbFound = True
120               Exit For
130            End If
140         Next
150      End With
160      SearchOn = lbFound
End Function

Private Sub ReLoadList()
10       Dim lvValue As Variant
20       Dim lnCol As Long
30       Dim lsOldProc As String
   
40       lsOldProc = p_oAppDrivr.ProcName("ReLoadList")
50       On Error Goto errProc
   
60       With MSFlexGrid1
70          p_bDisplayd = False
80          If p_oLookup.RecordCount = 0 Then
90             .Rows = 2
100            GoTo endProc
110         End If
      
120         If p_oLookup.RecordCount > xeMaxRecd Then
130            MsgBox "Search Record Result Exceeds The Maximum Allowable Record Display!!!" & _
               vbCrLf & "Please Limit Your Selection by Specifying More Detailed Info!!!", vbCritical, "Warning"
140            GoTo endProc
150         End If

160         p_oLookup.MoveFirst
170         .Rows = p_oLookup.RecordCount + 1
      
180         showProgress .Rows + 1
190         pnCtr = 0
200         p_bDisplayd = True
210         Do Until p_oLookup.EOF
220            pnCtr = pnCtr + 1
230            For lnCol = 0 To UBound(p_asColName)
240               lvValue = p_oLookup(p_asColName(lnCol))
250               If IsNull(p_oLookup(p_asColName(lnCol))) Then lvValue = Empty
260               .TextMatrix(pnCtr, lnCol) = Format(lvValue, p_asColPict(lnCol))
270            Next
         
280            p_oLookup.MoveNext
290         Loop
300         hideProgress
310      End With

endProc:
320      p_oAppDrivr.ProcName lsOldProc
330      Exit Sub
errProc:
340      ShowError lsOldProc
End Sub

Private Sub SortList()
10       If p_bDisplayd = False Then Exit Sub
20       p_oLookup.Sort = p_asColName(cmbSearch.ListIndex)
30       ReLoadList
End Sub

Private Function getSelectedItem() As Variant
10       Dim lvSelected As Variant
20       Dim lsOldProc As String
   
30       lsOldProc = p_oAppDrivr.ProcName("getSelectedItem")
40       On Error Goto errProc
   
50       lvSelected = ""
60       With MSFlexGrid1
70          If .RowSel > 0 Then
80             p_oLookup.MoveFirst
90             p_oLookup.Move .RowSel - 1, adBookmarkFirst
100            For pnCtr = 0 To p_oLookup.Fields.Count - 1
110               Select Case p_oLookup(pnCtr).Type
            Case 2, 3, 11, 17, 72, 4, 5, 6, 131
120                  lvSelected = lvSelected & Format(p_oLookup(pnCtr)) & "»"
130               Case Else
140                  lvSelected = lvSelected & p_oLookup(pnCtr) & "»"
150               End Select
160            Next
170            lvSelected = Left(lvSelected, Len(lvSelected) - 1)

180         End If
190      End With
200      getSelectedItem = lvSelected

endProc:
210      p_oAppDrivr.ProcName lsOldProc
220      Exit Function
errProc:
230      ShowError lsOldProc
End Function

Private Sub p_oLookup_MoveComplete(ByVal adReason As EventReasonEnum, ByVal pError As Error, adStatus As EventStatusEnum, ByVal pRecordset As Recordset)
10       If Not pbProgress Then Exit Sub
20       DoEvents
30       If Not pRecordset.EOF Then MoveProgress
End Sub

Private Sub txtSearch_GotFocus()
10       pbFocus = True
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
   'Remarks: This procedure only exists to trap a delete key, which irritatingly,
   '         does not trigger a KeyPress event
   '
10       Dim lsSearchOn As String          'current string to search on

20       On Error Resume Next
   
30       If p_bDisplayd = False Then Exit Sub
   
   'Check if we're dealing with a Delete key
40       If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or _
          KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp Then
50          MSFlexGrid1.SetFocus
60          Exit Sub
70       ElseIf KeyCode <> vbKeyDelete Then
80          Exit Sub
90       End If
   
   'The delete key was pressed; decide what to search on
100      lsSearchOn = ResultingText(KeyCode)
110      SearchOn lsSearchOn
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
   'Remarks: 1. When the user types something into the text portion, move to the
   '            first list entry which begins with the displayed text
   '         2. Not all keys trigger this event. In particular -
   '              <Delete> - triggers KeyDown by not KeyPress
   '              <BackSpace> - triggers KeyPress by not KeyDown
   '         3. This code was originally in the change() event, but confusing inter
   '            actions kept occurring (list index was being set to -1 by WINDOWS)
   '
10       Dim lsSearchOn As String             'current string to search on

20       On Error Resume Next
   
30       If p_bDisplayd = False Then Exit Sub
40       If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then Exit Sub
   
   'A content-modifying key was entered; decide what to search on
50       lsSearchOn = ResultingText(KeyAscii)
60       If SearchOn(lsSearchOn) = False Then KeyAscii = 0
End Sub

Private Sub txtSearch_LostFocus()
10       pbFocus = False
End Sub

Private Sub xrButton1_Click(Index As Integer)
10       Select Case LCase(xrButton1(Index).Caption)
   Case "&load"
20          If MSFlexGrid1.RowSel < 1 Or p_bDisplayd = False Then
30             MsgBox "Nothing to Load!", vbInformation, "Warning"
40             p_bSelected = False
50             Exit Sub
60          End If
70          p_bSelected = True
80          Me.Hide
90       Case "&close"
100         p_bSelected = False
110         Me.Hide
120      Case "searc&h"
130         getList
140         p_oLookup.Sort = p_asColName(cmbSearch.ListIndex)
150         ReLoadList
160      End Select
End Sub

Private Sub showProgress(ByVal lnMaxLength As Long)
10       pnInterval = 1
20       pnProgress = 1
30       If lnMaxLength > 32767 Then
40          pnInterval = Int(lnMaxLength / 32767)
50          ProgressBar1.Max = 32767
60       Else
70          ProgressBar1.Max = lnMaxLength
80       End If
   
90       pbProgress = True
100      ProgressBar1.Visible = True
End Sub

Private Sub MoveProgress()
10       pnProgress = pnProgress + 1
20       DoEvents
30       ProgressBar1.Value = Int(pnProgress / pnInterval)
40       DoEvents
End Sub

Private Sub hideProgress()
10       pbProgress = False
20       ProgressBar1.Visible = False
End Sub

Private Sub ShowError(ByVal lsProcName As String)
10       With p_oAppDrivr
20          .ShowError "frmLookUp", .ProcName(lsProcName), Err.Number, Err.Description, Erl
30       End With
40       With Err
50          .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
60       End With
End Sub

Private Function showButton()
10       If p_bSearch Then
20          xrButton1(1).Caption = "Searc&h"
30          xrButton1(2).Caption = "&Close"
40          xrButton1(2).Visible = True
50       Else
60          xrButton1(1).Caption = "&Close"
70          xrButton1(2).Visible = False
80       End If
End Function

Private Sub getList()
10       Dim lsOldProc As String
20       Dim lsSQL As String
   
30       lsOldProc = p_oAppDrivr.ProcName("getList")
40       On Error Goto errProc
   
50       If p_sSQLQuery <> Empty Then
60          lsSQL = p_sSQLQuery
70       Else
80          lsSQL = p_oLookup.Source
90       End If
   
100      If txtSearch.Text <> Empty Then
110         lsSQL = p_oMod.AddCondition(lsSQL, p_asFldName(cmbSearch.ListIndex) & " LIKE " & p_oMod.strParm(Trim(txtSearch) & "%"))
120      End If
   
130      If p_oLookup.State = adStateOpen Then p_oLookup.Close
140      p_oLookup.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, , adCmdText
   
endProc:
150      p_oAppDrivr.ProcName lsOldProc
160      Exit Sub
errProc:
170      ShowError lsOldProc
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Rex S. Adversalo
' XerSys Computing
' Canaoalan, Binmaley, Pangasinan
'
' Copyright 2003-2004 and beyond
' All Rights Reserved
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-0863      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  XerSys [ 06/15/2004 11:50 am ]
'     Start revising this object. i adopt the idea of the form driver of my dear friend
'  Marlon Sayson (with his consent). Included in here are new procedure format.
' ------------------------------------------------------------------------------------------
'  XerSys [ 06/17/2004 09:25 am ]
'     Continue revising object.
'        1) Remove the validate event coz i think it is redundant.
'        2) Add another property for fieldproperty data type,
'           Empty = True (Field can contain empty entry]
'                   False (Field can not contain empty entry_
'           Reference = True (Field is one of the reference of the table)
'                     = False (Field is not a reference field)
'        3) SetValue = a function that assigns the entry to the field and at the same
'              time validate the field if it's the corresponse to the data type of the
'              field. Also test if field can contain empty or not.
'  XerSys [ 02/07/2006 10:58 am ]
'     Override the aggreed modification policy...
'        Modification Policy
'           Encoder     = Add Only
'           Supervisor  = Add and modify record keyed a day before up to now
'           Manager     = Add and modify record keyed a month befor up to now
'           Audit       = generate reports
'           SysAdmin    = approved modification of record
'           Sys Owner   = Add, modify and delete record
'           Engineer    = same as owner
'     Modification policy would be manually set depending on the needs of a certain object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

'-----------------------------------------------------------
' Private variables here
'-----------------------------------------------------------
Private p_oAppDrivr As AppDriver

Private p_oRecMain As Recordset
Private p_oMainForm As Object
Private p_oMod As New MainModules

Private p_axTextFld() As FieldProperty
Private p_asAHFName() As String
Private p_asAHTitle() As String
Private p_asAHFPict() As String
Private p_asAHValue() As String
Private p_abAHFRefx() As Boolean

Private p_sQuerySQL As String
Private p_sBrowsSQL As String
Private p_nCurIndex As Integer
Private p_nFldCount As Integer
Private p_nFldStart As Integer
Private p_bFldChnge As Boolean
Private p_bRecExist As Boolean
Private p_nEditMode As xeEditMode

Private pnCtr As Long

'-----------------------------------------------------------
'Class Events Here
'-----------------------------------------------------------
' Triggered before the record is deleted. Can be used in validating record
Public Event WillDelete(Cancel As Boolean)

' Triggered before the record is deleted. Can be used to override the delete
' module of this object...
Public Event Delete(Deleted As Boolean)

' Triggered after the record is deleted. Can be used in performing
' other process after deleting record...
Public Event DeleteComplete()

' Triggered before the record is saved. Can be used in validating record
Public Event WillSave(Cancel As Boolean)

' Triggered before the record is save. Can be used to override the save module
Public Event Save(Saved As Boolean)

' Triggered after the record is saved. Can be used in performing other
' process after saving record...
Public Event SaveComplete()

' Triggered after enabling the TextBox for entry. Can be used to enable other controls
Public Event EnableOtherControl()

' Triggered after disabling the TextBox for entry. Can be used to disable other controls
Public Event DisableOtherControl()

' Triggered after emptying the TextBox for entry. Can be used to set the
' initial value of other controls
Public Event InitValue()

' Triggered after loading the data for the textbox. Can be used to set the value
' of the the other data entry controls
Public Event LoadOtherData()

Property Set AppDriver(oAppDriver As AppDriver)
10       Set p_oAppDrivr = oAppDriver
End Property

Property Set MainForm(oForm As Object)
10       Set p_oMainForm = oForm
End Property

Property Get MainForm() As Object
10       Set MainForm = p_oMainForm
End Property

Property Let RecQuery(Value As String)
10       p_sQuerySQL = Value
End Property

Property Get RecQuery() As String
10       RecQuery = p_sQuerySQL
End Property

Property Let BrowseQuery(Value As String)
10       p_sBrowsSQL = Value
End Property

Property Get BrowseQuery() As String
10       BrowseQuery = p_sBrowsSQL
End Property

Property Let BrowseColumn(ByVal Index As Integer, ByVal Value As String)
10       p_asAHFName(Index) = Value
End Property

Property Get BrowseColumn(ByVal Index As Integer) As String
10       BrowseColumn = p_asAHFName(Index)
End Property

Property Let BrowseFTitle(ByVal Index As Integer, ByVal Value As String)
10       p_asAHTitle(Index) = Value
End Property

Property Get BrowseFTitle(ByVal Index As Integer) As String
10       BrowseFTitle = p_asAHTitle(Index)
End Property

Property Let BrowseFFormat(ByVal Index As Integer, ByVal Value As String)
10       p_asAHFPict(Index) = Value
End Property

Property Get BrowseFFormat(ByVal Index As Integer) As String
10       BrowseFFormat = p_asAHFPict(Index)
End Property

Property Let BrowseFReference(ByVal Index As Integer, ByVal Value As Boolean)
10       p_abAHFRefx(Index) = Value
End Property

Property Get BrowseFReference(ByVal Index As Integer) As Boolean
10       BrowseFReference = p_abAHFRefx(Index)
End Property

Property Get EditMode() As xeEditMode
10       EditMode = p_nEditMode
End Property

Property Let LookupQuery(ByVal Index As Integer, ByVal Value As String)
10       p_axTextFld(Index).Query = Value

20       With p_oMainForm
30          If Value = "" Then
40             .txtField(Index).MaxLength = FieldSize(Index)
50          Else
60             .txtField(Index).MaxLength = 0
70          End If
80       End With
End Property

Property Get LookupQuery(ByVal Index As Integer) As String
10       LookupQuery = p_axTextFld(Index).Query
End Property

Property Get LookupColumn(ByVal Index As Integer) As String
10       LookupColumn = p_axTextFld(Index).QColumn
End Property

Property Let LookupColumn(ByVal Index As Integer, Value As String)
10       p_axTextFld(Index).QColumn = Value
End Property

Property Let LookupTitle(ByVal Index As Integer, Value As Variant)
10       p_axTextFld(Index).QTitle = Value
End Property

Property Get LookupTitle(ByVal Index As Integer) As Variant
10       LookupTitle = p_axTextFld(Index).QTitle
End Property

Property Let LookupFormat(ByVal Index As Integer, ByVal Value As Variant)
10       p_axTextFld(Index).QFormat = Value
End Property

Property Get LookupFormat(ByVal Index As Integer) As Variant
10       LookupFormat = p_axTextFld(Index).QFormat
End Property

Property Let LookupReference(ByVal Index As Integer, ByVal Value As Variant)
10       p_axTextFld(Index).QReference = Value
End Property

Property Get LookupReference(ByVal Index As Integer) As Variant
10       LookupReference = p_axTextFld(Index).QReference
End Property

Property Let LookupValue(ByVal Index As Integer, Value As Variant)
10       p_asAHValue(Index) = Value
End Property

Property Get LookupValue(ByVal Index As Integer) As Variant
10       LookupValue = p_asAHValue(Index)
End Property

Property Let FieldName(ByVal Index As Integer, ByVal Value As String)
10       p_axTextFld(Index).Name = Value
End Property

Property Get FieldName(ByVal Index As Integer) As String
10       FieldName = p_axTextFld(Index).Name
End Property

Property Let FieldValue(ByVal Index As Integer, ByVal Value As Variant)
10       p_axTextFld(Index).Value = Value
End Property

Property Get FieldValue(ByVal Index As Integer) As Variant
10       FieldValue = p_axTextFld(Index).Value
End Property

Property Let FieldType(ByVal Index As Integer, ByVal Value As String)
10       p_axTextFld(Index).Type = Value
End Property

Property Get FieldType(ByVal Index As Integer) As String
10       FieldType = p_axTextFld(Index).Type
End Property

Property Let FieldSize(ByVal Index As Integer, ByVal Value As Integer)
10       p_axTextFld(Index).Size = Value
20       p_oMainForm.txtField(Index).MaxLength = Value
End Property

Property Get FieldSize(ByVal Index As Integer) As Integer
10       FieldSize = p_axTextFld(Index).Size
End Property

Property Let FieldHelp(ByVal Index As Integer, ByVal Value As String)
10       p_axTextFld(Index).Help = Value
End Property

Property Get FieldHelp(ByVal Index As Integer) As String
10       FieldHelp = p_axTextFld(Index).Help
End Property

Property Let FieldReference(ByVal Index As Integer, ByVal Value As Boolean)
10       p_axTextFld(Index).Reference = Value
End Property

Property Get FieldReference(ByVal Index As Integer) As Boolean
10       FieldReference = p_axTextFld(Index).Reference
End Property

Property Let FieldEmpty(ByVal Index As Integer, ByVal Value As Boolean)
10       p_axTextFld(Index).Empty = Value
End Property

Property Get FieldEmpty(ByVal Index As Integer) As Boolean
10       FieldEmpty = p_axTextFld(Index).Empty
End Property

Property Let FieldInit(ByVal Index As Integer, ByVal Value As Boolean)
10       p_axTextFld(Index).Initialize = Value
End Property

Property Get FieldInit(ByVal Index As Integer) As Boolean
10       FieldInit = p_axTextFld(Index).Initialize
End Property

Property Let FieldFormat(ByVal Index As Integer, ByVal Value As String)
10       p_axTextFld(Index).Format = Value
End Property

Property Get FieldFormat(ByVal Index As Integer) As String
10       FieldFormat = p_axTextFld(Index).Format
End Property

Property Let ColumnIndex(ByVal Value As Integer)
10       p_nCurIndex = Value
20       With MainForm
30          If .txtField(Value).Enabled And .txtField(Value).Visible Then
40             .txtField(Value).SelStart = 0
50             .txtField(Value).SelLength = Len(.txtField(Value).Text)
60          End If
70       End With
End Property

Property Get ColumnIndex() As Integer
10       ColumnIndex = p_nCurIndex
End Property

Property Get FieldCount() As Integer
10       FieldCount = p_nFldCount
End Property

Property Let FieldStart(nStart As Integer)
10       p_nFldStart = nStart
End Property

Property Get FieldStart() As Integer
10       FieldStart = p_nFldStart
End Property

Property Get RecStatus() As Integer
10       RecStatus = p_nEditMode
End Property

' Show the command button for record process (cmdButton)
'  Index = -1 all record process button will be enabled; record entry button disabled
'        = >= 0 cmdbutton with that index will be enabled
Sub showButton(ByVal Index As Integer)
10       Dim loButton As Control
20       Dim lsOldProc As String
30       Dim lbShow As Boolean
   
40       lsOldProc = p_oAppDrivr.ProcName("ShowButton")
50       'On Error GoTo errProc
   
60       With MainForm
70          If p_nEditMode = xeModeReady Then lbShow = True
80          If Index < 0 Then
90             For Each loButton In .cmdButton
100               Select Case UCase(loButton.Caption)
            Case "&NEW", "&UPDATE", "&DELETE", "&CLOSE"
110                  loButton.Visible = lbShow
120               Case "&SAVE", "&CANCEL", "SEARC&H"
130                  loButton.Visible = Not lbShow
140               End Select
150            Next
160         Else
170            .cmdButton(Index).Visible = True
180         End If
190      End With
   
endProc:
200      p_oAppDrivr.ProcName lsOldProc
210      Exit Sub
errProc:
220      ShowError lsOldProc
End Sub

' hide the command button for record process (cmdButton)
'  Index = see showbutton
Sub hideButton(ByVal Index As Integer)
10       Dim loButton As Control
20       Dim lsOldProc As String
30       Dim lbShow As Boolean
   
40       lsOldProc = p_oAppDrivr.ProcName("HideButton")
50       'On Error GoTo errProc
   
60       With MainForm
70          If p_nEditMode = xeModeReady Then lbShow = True
80          If Index < 0 Then
90             For Each loButton In .cmdButton
100               Select Case UCase(loButton.Caption)
            Case "&NEW", "&UPDATE", "&DELETE", "&CLOSE"
110                  loButton.Visible = lbShow
120               Case "&SAVE", "&CANCEL", "SEARC&H"
130                  loButton.Visible = Not lbShow
140               End Select
150            Next
160         Else
170            .cmdButton(Index).Visible = False
180         End If
190      End With
   
endProc:
200      p_oAppDrivr.ProcName lsOldProc
210      Exit Sub
errProc:
220      ShowError lsOldProc
End Sub

' Enables the textbox entry withing the MainForm.
'  Index = -1 Enables all textbox entry (txtfield)
'        = any no. Enables the textbox entry of that index
Sub EnableTextbox(ByVal Index As Integer)
10       Dim loText As TextBox
20       Dim lsOldProc As String
   
30       lsOldProc = p_oAppDrivr.ProcName("EnableTextBox")
40       'On Error GoTo errProc
   
50       With MainForm
60          If Index < 0 Then
70             For Each loText In .txtField
80               loText.Enabled = True
90             Next
100            RaiseEvent EnableOtherControl
110         Else
120            .txtField(Index).Enabled = True
130         End If
140      End With
   
endProc:
150      p_oAppDrivr.ProcName lsOldProc
160      Exit Sub
errProc:
170      ShowError lsOldProc
End Sub

' Disables the textboxes within MainForm
'  Index = see EnableTextbox
Sub DisableTextbox(ByVal Index As Integer)
10       Dim loText As TextBox
20       Dim lsOldProc As String
   
30       lsOldProc = p_oAppDrivr.ProcName("DisableTextBox")
40       'On Error GoTo errProc
   
50       With MainForm
60          If Index < 0 Then
70             For Each loText In .txtField
80               loText.Enabled = False
90             Next
100            RaiseEvent DisableOtherControl
110         Else
120            .txtField(Index).Enabled = False
130         End If
140      End With
   
endProc:
150      p_oAppDrivr.ProcName lsOldProc
160      Exit Sub
errProc:
170      ShowError lsOldProc
End Sub

' Triggers Validate Event
' Performs formatting and validation of the record entry
Function ValidateField(ByVal Index As Integer) As Boolean
10       Dim lors As Recordset
20       Dim lsOldProc As String
30       Dim lsSQL As String
40       Dim lasRef() As String
50       Dim vValue As Variant
   
60       lsOldProc = p_oAppDrivr.ProcName("ValidateField")
70       ValidateField = False
80       'On Error GoTo errProc
   
   ' Perform textfields that needs validation here!
90       With MainForm
100         .txtField(Index).Text = Trim(.txtField(Index).Text)

      ' If data encoded is valid/changed then perform the following
110         chkFieldStat Index
120         If p_bFldChnge Then
130            If Index < FieldStart Then
140               .txtField(Index).Text = Format(.txtField(Index), FieldFormat(Index))
150               FieldValue(Index) = .txtField(Index).Text
160            Else
            ' check if field has a lookup table
170               lsSQL = LookupQuery(Index)
180               If lsSQL = Empty Then
190                  ValidateField = SetValue(Index, .txtField(Index).Text)
200                  GoTo endProc
210               End If
220               vValue = .txtField(Index).Text

230               If Trim(vValue) = "" Then
240                  ValidateField = SetValue(Index, .txtField(Index).Text)
250                  GoTo endProc
260               End If

            ' Search for the existence of the TextBox Entry
270               lasRef = Split(LookupReference(Index), "»")
280               lsSQL = lasRef(1) & " = " & p_oMod.strParm(Trim(vValue))
290               lsSQL = p_oMod.AddCondition(LookupQuery(Index), lsSQL)

300               Set lors = New Recordset
310               lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
            
320               If lors.EOF Then
330                  ValidateField = SetValue(Index, Empty)
340                  GoTo endProc
350               End If
            
            'Existing ba?
360               If lors.RecordCount = 1 Then
370                  FieldValue(Index) = lors(0).Value
380                  .txtField(Index).Text = lors(1).Value
390                  ValidateField = True
400                  GoTo endProc
410               End If
420               vValue = p_oMod.KwikBrowse(p_oAppDrivr, lors, LookupColumn(Index), _
                        LookupTitle(Index), LookupFormat(Index))
                                
430               If vValue <> Empty Then
440                  p_asAHValue = Split(vValue, "»")
450                  If p_asAHValue(0) = Empty Then
460                     ValidateField = SetValue(Index, Empty)
470                     GoTo endProc
480                  End If
490               End If
            
500               FieldValue(Index) = Trim(p_asAHValue(0))
510               .txtField(Index).Text = p_asAHValue(1)
520            End If
530            ValidateField = True
540         Else
550            If Not (isFieldEmpty(Index, .txtField(Index)) And FieldEmpty(Index) = False) Then ValidateField = True
560         End If
570      End With
   
endProc:
580      p_oAppDrivr.ProcName lsOldProc
590      Set lors = Nothing
600      Exit Function
errProc:
610      ShowError lsOldProc
End Function

' Empty the contents of  the textboxes within MainForm
' Triggers InitValue
Sub InitTextBoxes()
10       Dim lsOldProc As String
   
20       lsOldProc = p_oAppDrivr.ProcName("InitTextBoxes")
30       'On Error GoTo errProc
   
   ' Initialize first the fields type and format them accordingly
40       For pnCtr = 0 To FieldCount - 1
50          If FieldInit(pnCtr) Then setDefaultValue pnCtr
60       Next
70       Field2Control
80       p_bRecExist = False
   
endProc:
90       p_oAppDrivr.ProcName lsOldProc
100      Exit Sub
errProc:
110      ShowError lsOldProc
End Sub

' Set the value of the field and the textbox. This also validate the textbox
' and return if the textbox has a valid value
Function SetValue(ByVal Index As Integer, Value As Variant) As Boolean
10       Dim lsOldProc As String
   
20       lsOldProc = p_oAppDrivr.ProcName("SetValue")
30       SetValue = False
40       'On Error GoTo errProc
   
50       With MainForm
60          If IsNull(Value) Then Value = Empty
         
70          Select Case FieldType(Index)
      Case "n", "l"
80             If IsNumeric(Value) = False Then
90                Value = IIf(FieldType(Index) = "l", 0#, 0)
100            Else
110               SetValue = True
120            End If
         
         ' after validating data type, check if allowed empty
130            If Value = 0 And FieldEmpty(Index) = False Then
140               SetValue = False
150               GoTo endProc
160            End If
170         Case "d", "t", "p"
180            If IsDate(Value) = False Then
190               Select Case FieldType(Index)
            Case "d"
200                  Value = Now
210               Case "t"
220                  Value = Time
230               Case Else
240                  Value = Date
250               End Select
260            Else
270               SetValue = True
280            End If
290         Case "s"
         ' after validating data type, check if allowed empty
300            If Value = "" And FieldEmpty(Index) = False Then GoTo endProc
310            SetValue = True
320         End Select
330         .txtField(Index) = Format(Value, FieldFormat(Index))
340         FieldValue(Index) = Value
350      End With
   
endProc:
360      p_oAppDrivr.ProcName lsOldProc
370      Exit Function
errProc:
380      ShowError lsOldProc
End Function

' Initialize Textfield sizes using it's field size.
Sub InitFieldSize()
10       Dim lsOldProc As String
20       Dim loText As TextBox
   
30       lsOldProc = p_oAppDrivr.ProcName("InitFieldSize")
40       'On Error GoTo errProc
   
50       With MainForm
60          For Each loText In .txtField
70             If FieldType(loText.Index) = "s" And LookupQuery(loText.Index) = Empty Then
80                loText.MaxLength = FieldSize(loText.Index)
90             End If
100         Next
110      End With

endProc:
120      p_oAppDrivr.ProcName lsOldProc
130      Exit Sub
errProc:
140      ShowError lsOldProc
End Sub

' Initializes the field property base on the query string
Function InitRecForm(Optional Query As Variant, _
                     Optional BQuery As Variant) As Boolean
10       Dim lsOldProc As String
20       Dim lsSQL As String
   
30       lsOldProc = p_oAppDrivr.ProcName("InitRecForm")
40       InitRecForm = False
50       'On Error GoTo errProc
   
60       If IsMissing(Query) = False Then p_sQuerySQL = Query
70       If IsMissing(BQuery) = False Then p_sBrowsSQL = BQuery
   ' if either query string or application driver is not present, abort object
80       If p_sQuerySQL = Empty Then GoTo endProc
90       If isAppDriverOK = False Then GoTo endProc
   
100      If p_sBrowsSQL = Empty Then p_sBrowsSQL = p_sQuerySQL
   
   ' initialize first the browse field
110      lsSQL = p_oMod.getEmptyRec(p_sBrowsSQL)
120      Set p_oRecMain = New Recordset
130      p_oRecMain.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
140      ReDim p_asAHFName(p_oRecMain.Fields.Count - 1) As String
150      ReDim p_asAHTitle(p_oRecMain.Fields.Count - 1) As String
160      ReDim p_asAHFPict(p_oRecMain.Fields.Count - 1) As String
170      ReDim p_abAHFRefx(p_oRecMain.Fields.Count - 1) As Boolean
180      ReDim p_asAHValue(p_oRecMain.Fields.Count - 1) As String
   
190      For pnCtr = 0 To UBound(p_asAHTitle)
200         p_asAHFName(pnCtr) = p_oRecMain(pnCtr).Name
210         p_asAHTitle(pnCtr) = p_oRecMain(pnCtr).Name
220         p_asAHFPict(pnCtr) = "@"
230         p_abAHFRefx(pnCtr) = False
240      Next
250      p_abAHFRefx(0) = True
260      p_oRecMain.Close
   
   ' it must retrieve an empty record coz the default is new record
270      lsSQL = p_oMod.getEmptyRec(p_sQuerySQL)
280      p_oRecMain.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockPessimistic, adCmdText

   'Initialize variables to be use in manipulating the fields
290      p_nFldCount = p_oRecMain.Fields.Count
   
300      ReDim p_axTextFld(p_nFldCount) As FieldProperty
   
310      For pnCtr = 0 To p_nFldCount - 1
320         p_axTextFld(pnCtr).Name = p_oRecMain(pnCtr).Name
330         p_axTextFld(pnCtr).Value = Null
340         p_axTextFld(pnCtr).Size = p_oRecMain(pnCtr).DefinedSize
350         p_axTextFld(pnCtr).Help = Empty
360         p_axTextFld(pnCtr).Format = "@"
370         p_axTextFld(pnCtr).Reference = False
380         p_axTextFld(pnCtr).Initialize = True
390         p_axTextFld(pnCtr).Empty = True
400         p_axTextFld(pnCtr).Query = Empty
410         p_axTextFld(pnCtr).QTitle = Empty
420         p_axTextFld(pnCtr).QFormat = Empty
430         p_axTextFld(pnCtr).QType = Empty
      
440         Select Case p_oRecMain(pnCtr).Type
      Case 129, 130, 202, 200    ' string
450            p_axTextFld(pnCtr).Type = "s"
460         Case 2, 3, 11, 17, 72      ' numeric without decimal point
470            p_axTextFld(pnCtr).Type = "n"
480         Case 4, 5, 6, 131          ' numeric with decimal point
490            p_axTextFld(pnCtr).Type = "l"
500         Case 135                   ' datetime
510            p_axTextFld(pnCtr).Type = "d"
520         End Select
530      Next
   ' always assume that the first rec is the reference field
540      p_axTextFld(0).Reference = True
   
550      p_nEditMode = xeModeReady
560      InitFieldSize
570      InitTextBoxes
580      hideButton -1
590      DisableTextbox -1
600      InitRecForm = True
   
endProc:
610      lsOldProc = p_oAppDrivr.ProcName(lsOldProc)
620      Exit Function
errProc:
630      ShowError lsOldProc
End Function

Sub LoadRecord()
10       Dim lsOldProc As String
20       Dim lsSQL As String
   
30       lsOldProc = p_oAppDrivr.ProcName("LoadRecord")
40       'On Error GoTo errProc
   
   ' Load the data from the selected record
50       GetRecord
60       With MainForm
70          If Not p_oRecMain.EOF Then
80             FieldValue(0) = p_oRecMain(0).Value
90             .txtField(0).Text = FieldValue(0)
100            For pnCtr = 1 To FieldCount - 1
            ' Load the field value to the field recipeint
110               FieldValue(pnCtr) = p_oRecMain(pnCtr)
120            Next
130            Field2Control
140            RaiseEvent LoadOtherData
150         Else
160            InitTextBoxes
170         End If
180      End With

   ' after loading record disable edit button
190      p_nEditMode = xeModeReady
200      p_bRecExist = True
210      hideButton -1
220      DisableTextbox -1
   
endProc:
230      p_oAppDrivr.ProcName lsOldProc
240      Exit Sub
errProc:
250      ShowError lsOldProc
End Sub

Sub BrowseRecord()
10       Dim lsOldProc As String
20       Dim lsAHValue As String
30       Dim lsColumn As String
40       Dim lsTitle As String
50       Dim lsPict As String
   
60       lsOldProc = p_oAppDrivr.ProcName("BrowseRecord")
70       'On Error GoTo errProc
   
80       lsColumn = ""
90       lsTitle = ""
100      lsPict = ""
110      For pnCtr = 0 To UBound(p_asAHTitle)
120         If p_asAHTitle(pnCtr) <> "" Then
130            lsColumn = lsColumn & "»" & p_asAHFName(pnCtr)
140            lsTitle = lsTitle & "»" & p_asAHTitle(pnCtr)
150            lsPict = lsPict & "»" & p_asAHFPict(pnCtr)
160         End If
170      Next
180      lsColumn = Mid(lsColumn, 2, Len(lsColumn) - 1)
190      lsTitle = Mid(lsTitle, 2, Len(lsTitle) - 1)
200      lsPict = Mid(lsPict, 2, Len(lsPict) - 1)
   
210      lsAHValue = p_oMod.KwikSearch(p_oAppDrivr, p_sBrowsSQL, lsColumn, lsTitle, lsPict)
   
220      If lsAHValue = Empty Then
230         MsgBox "No selection was made!!!" & vbCrLf & _
               "No record will be loaded!!!", vbInformation, "Notice"
240         GoTo endProc
250      End If
260      p_asAHValue = Split(lsAHValue, "»")
270      LoadRecord

endProc:
280      p_oAppDrivr.ProcName lsOldProc
290      Exit Sub
errProc:
300      ShowError lsOldProc
End Sub

' Adds new record
Sub RecordNew()
   'Initialize all value of the controls here
10       InitTextBoxes
20       RaiseEvent InitValue
   
30       p_nEditMode = xeModeAddNew
   
   ' enable controls for editing
40       showButton -1
50       EnableTextbox -1
   
60       p_nCurIndex = FieldStart
70       With p_oMainForm
80          .txtField(p_nCurIndex).SetFocus
90       End With
End Sub

' Use this to save the records
Function RecordSave() As Boolean
10       Dim lsOldProc As String
20       Dim Cancel As Boolean
30       Dim Saved As Boolean
   
40       lsOldProc = p_oAppDrivr.ProcName("RecordSave")
50       RecordSave = False
60       'On Error GoTo errProc

70       If MsgBox("Are you sure you want to save this record?", _
         vbYesNo + vbInformation, "Confirm") <> vbYes Then GoTo endProc
         
80       If HasRights2Save = False Then GoTo endProc
   
90       If isRefEmpty = True Then
100         MsgBox "Reference Fields can not Contain Empty Value!!!" & vbCrLf & _
                     " Verify your Entry then Try Again!!!", vbCritical, "Warning"
110         GoTo endProc
120      End If
         
130      p_oAppDrivr.Connection.BeginTrans
140      FieldValue(p_nFldCount - 3) = p_oMod.Encrypt(p_oAppDrivr.UserID)
150      FieldValue(p_nFldCount - 2) = p_oAppDrivr.ServerDate

160      Cancel = False
   ' record validation must be in this event
170      RaiseEvent WillSave(Cancel)
   
180      If Cancel = True Then
190         p_oAppDrivr.Connection.RollbackTrans
200         GoTo endProc
210      End If

220      If isModify = False Then
230         p_oAppDrivr.Connection.CommitTrans
240         RecordSave = True
250         If p_nEditMode = xeModeAddNew Then
260            RecordNew
270         Else
280            p_nEditMode = xeModeReady
290            hideButton -1
300            DisableTextbox -1
310         End If

      ' after saving the main table, trigger the savecomplete event
320         RaiseEvent SaveComplete
330         GoTo endProc
340      End If
   
350      Saved = False
360      RaiseEvent Save(Saved)
   
   ' if record is saved be the client then endproc
370      If Saved Then
380         p_oAppDrivr.Connection.CommitTrans
390         RaiseEvent SaveComplete
400         RecordSave = Saved
410         If p_nEditMode = xeModeAddNew Then
420            RecordNew
430         Else
440            p_nEditMode = xeModeReady
450            hideButton -1
460            DisableTextbox -1
470         End If
480         GoTo endProc
490      End If
   
   ' Save data base on the kind of task
500      If p_nEditMode = xeModeAddNew Then p_oRecMain.AddNew
   
510      For pnCtr = 0 To p_nFldCount - 3
520         p_oRecMain(FieldName(pnCtr)) = FieldValue(pnCtr)
530      Next
540      p_oRecMain("sModified") = p_oMod.Encrypt(p_oAppDrivr.UserID)
550      p_oRecMain("dModified") = p_oAppDrivr.ServerDate
   
560      p_oRecMain.UpdateBatch adAffectAllChapters
570      p_oAppDrivr.Connection.CommitTrans
   
   ' show some message that the record was save successfully
580      If Not (p_oAppDrivr.MDIMain Is Nothing) Then
'590         p_oAppDrivr.MDIMain.StatusBar1.Panels(1).Text = "Record save successfully!!!"
600      End If

610      RecordSave = True
620      If p_nEditMode = xeModeAddNew Then
630         RecordNew
640      Else
650         p_asAHValue(0) = p_oRecMain(0)
660         LoadRecord
      
670         p_nEditMode = xeModeReady
680         hideButton -1
690         DisableTextbox -1
700      End If

   ' after saving the main table, trigger the savecomplete event
710      RaiseEvent SaveComplete
   
endProc:
720      p_oAppDrivr.ProcName lsOldProc
730      Exit Function
errProc:
740      MsgBox Err.Description
750      p_oAppDrivr.Connection.RollbackTrans
760      ShowError lsOldProc
End Function

'Use this to search lookup fields...
Sub RecordSearch(Optional Value As Variant = "")
10       Dim lors As Recordset
20       Dim lsOldProc As String
30       Dim lsAHValue As String
40       Dim lsLookupSQL As String
50       Dim lasRef() As String
   
60       lsOldProc = p_oAppDrivr.ProcName("RecordSearch")
70       'On Error GoTo errProc
   
80       With p_oMainForm
90          If LookupQuery(p_nCurIndex) <> Empty And p_nCurIndex > 0 Then
100            lsLookupSQL = LookupQuery(p_nCurIndex)
110            If Value <> "" Then
120               lasRef = Split(LookupReference(p_nCurIndex), "»")
130               lsLookupSQL = lasRef(1) & " LIKE " & p_oMod.strParm(Trim(Value) & "%")
140               lsLookupSQL = p_oMod.AddCondition(LookupQuery(p_nCurIndex), lsLookupSQL)
150            End If
160            Set lors = New Recordset
170            lors.Open lsLookupSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
         
180            If lors.EOF Then GoTo endProc
         
190            If lors.RecordCount = 1 Then
200               FieldValue(p_nCurIndex) = lors(0)
210               .txtField(p_nCurIndex).Text = lors(1)
220               p_asAHValue(0) = lors(0)
230               p_asAHValue(1) = lors(1)
240               GoTo endProc
250            End If
         
260            lsAHValue = p_oMod.KwikBrowse(p_oAppDrivr, lors, _
                           LookupColumn(p_nCurIndex), LookupTitle(p_nCurIndex), _
                           LookupFormat(p_nCurIndex))

270            If Not (lsAHValue = Empty) Then
280               p_asAHValue = Split(lsAHValue, "»")
290               FieldValue(p_nCurIndex) = Trim(p_asAHValue(0))
300               .txtField(p_nCurIndex).Text = p_asAHValue(1)
310            End If
320            .txtField(p_nCurIndex).SetFocus
330         End If
340      End With

endProc:
350      Set lors = Nothing
360      p_oAppDrivr.ProcName lsOldProc
370      Exit Sub
errProc:
380      ShowError lsOldProc
End Sub

' this enable the record for updating
Sub RecordUpdate()
10       If p_bRecExist = False Then
20          MsgBox "No Record is Loaded to Modify!!!" & vbCrLf & _
            "Verify your Entry then Try Again!!!", vbInformation, "Warning"
30          Exit Sub
40       End If

50       p_nEditMode = xeModeUpdate
   
   ' enable controls for editing
60       showButton -1
70       EnableTextbox -1
   
80       p_nCurIndex = FieldStart
90       With p_oMainForm
100         .txtField(p_nCurIndex).SetFocus
110      End With
End Sub

' actual delete module
Function RecordDelete() As Boolean
10       Dim lsSQL As String
20       Dim lsOldProc As String
30       Dim Cancel As Boolean
40       Dim Deleted As Boolean
50       Dim lnCtr As Integer

60       lsOldProc = p_oAppDrivr.ProcName("RecordDelete")
70       'On Error GoTo errProc
   
80       If p_bRecExist = False Then
90          MsgBox "No Record is Loaded to Modify!!!" & vbCrLf & _
           "Verify your Entry then Try Again!!!", vbInformation, "Warning"
100         GoTo endProc
110      End If

120      If MsgBox("Delete this record?", vbOKCancel + vbCritical, "Confirm") <> vbOK Then
130         MsgBox "Deletion of record has been aborted!", vbInformation, "Notice"
140         GoTo endProc
150      End If
   
160      If HasRights2Delete = False Then GoTo endProc
   
170      p_oAppDrivr.Connection.BeginTrans
   
180      Cancel = False
190      RaiseEvent WillDelete(Cancel)
200      If Cancel Then GoTo endProc
   
210      Deleted = False
220      RaiseEvent Delete(Deleted)
   
230      If Deleted Then
240         p_oAppDrivr.Connection.CommitTrans
250         RaiseEvent DeleteComplete
260         InitTextBoxes
270         p_nEditMode = xeModeReady
280         RecordDelete = True
290         GoTo endProc
300      End If
   
310      lnCtr = InStr(1, p_sQuerySQL, "FROM ", vbTextCompare)
320      If lnCtr = 0 Then
330         lsSQL = p_sQuerySQL
340      Else
350         lsSQL = Trim(Mid(p_sQuerySQL, lnCtr + 5))
360         lnCtr = InStr(1, lsSQL, " ", vbTextCompare)
370         If lnCtr > 0 Then
380            lsSQL = Trim(Mid(lsSQL, 1, lnCtr))
390         End If
400      End If
      
410      lsSQL = "DELETE " & lsSQL
420      For lnCtr = 0 To UBound(p_abAHFRefx)
430         If p_abAHFRefx(lnCtr) Then
440            lsSQL = p_oMod.AddCondition(lsSQL, p_asAHFName(lnCtr) & " = " & _
                  p_oMod.strParm(p_oRecMain(p_asAHFName(lnCtr))))
450         End If
460      Next
   
470      If p_oAppDrivr.RegisDelete(lsSQL) = False Then GoTo endProc
480      p_oRecMain.Delete
   
490      p_oAppDrivr.Connection.CommitTrans
500      RaiseEvent DeleteComplete
510      RecordDelete = True

   ' after deleting record, remove record from display
520      InitTextBoxes
530      p_nEditMode = xeModeReady
   
   ' show some message that the record was save successfully
540      If Not (p_oAppDrivr.MDIMain Is Nothing) Then
'550         p_oAppDrivr.MDIMain.StatusBar1.Panels(1).Text = "Record Deleted successfully!!!"
560      End If
   
endProc:
570      p_oAppDrivr.ProcName lsOldProc
580      Exit Function
errProc:
590      p_oAppDrivr.Connection.RollbackTrans
600      ShowError lsOldProc
End Function

Function RecordCancelUpdate() As Boolean
10       If p_nEditMode = xeModeAddNew Then
20          If p_asAHValue(0) = Empty Then
30             InitTextBoxes
40          Else
50             LoadRecord
60          End If
70       Else
80          LoadRecord
90       End If

100      p_nEditMode = xeModeReady
110      hideButton -1
120      DisableTextbox -1
End Function

Sub ViewUserModify()
10       Dim lors As Recordset
20       Dim lsOldProc As String
30       Dim lsSQL As String
   
40       lsOldProc = p_oAppDrivr.ProcName("isMasterModify")
50       'On Error GoTo errProc
   
60       lsSQL = "SELECT sUserName" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & p_oMod.strParm(p_oMod.Decrypt(p_oRecMain("sModified")))

70       Set lors = New Recordset
80       lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
90       If lors.EOF Then
100         MsgBox "User does not Exist in the User Table!!!", vbCritical, "Warning"
110         GoTo endProc
120      End If
   
130      MsgBox "    ««« LAST MODIFICATION »»»" & vbCrLf & vbCrLf & _
            "User Name      : " & p_oMod.Decrypt(lors("sUserName"), p_oAppDrivr.Machinex) & vbCrLf & _
            "Date Modified : " & Format(p_oRecMain("dModified"), "MMMM DD, YYYY HH:NN:SS") _
            , vbInformation, "Notice"
         
endProc:
140      Set lors = Nothing
150      p_oAppDrivr.ProcName lsOldProc
160      Exit Sub
errProc:
170      ShowError lsOldProc
End Sub

' retrieve the records based by the reference passed
Private Sub GetRecord()
10       Dim lsOldProc As String
20       Dim lsSQL As String
   
30       lsOldProc = p_oAppDrivr.ProcName("getRecord")
40       'On Error GoTo errProc
   
   ' before wasting processor time check if there's a valid value
50       If p_asAHValue(0) = "" Then GoTo endProc
   
60       lsSQL = p_sQuerySQL
70       For pnCtr = 0 To UBound(p_asAHFName)
80          If p_abAHFRefx(pnCtr) Then
90             lsSQL = p_oMod.AddCondition(lsSQL, p_asAHFName(pnCtr) & " = " & p_oMod.strParm(p_asAHValue(pnCtr)))
100         End If
110      Next

120      If p_oRecMain.State = adStateOpen Then
130         If Not p_oRecMain.EOF Then
140            If p_oRecMain.EditMode <> adEditNone Then p_oRecMain.CancelUpdate
150         End If
160         p_oRecMain.Close
170      End If
180      p_oRecMain.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText

endProc:
190      p_oAppDrivr.ProcName lsOldProc
200      Exit Sub
errProc:
210      p_oAppDrivr.Connection.RollbackTrans
220      ShowError lsOldProc
End Sub

Private Sub Class_Initialize()
10       Set p_oRecMain = New Recordset
   
20       p_nEditMode = xeModeAddNew
End Sub

Private Sub Class_Terminate()
10       Set p_oRecMain = Nothing
20       Set p_oMod = Nothing
End Sub

Private Sub chkFieldStat(ByVal Index As Integer)
10       If IsNull(FieldValue(Index)) Then
20          p_bFldChnge = MainForm.txtField(Index) <> Empty
30       Else
40          If LookupQuery(Index) <> Empty Then
50             p_bFldChnge = Trim(MainForm.txtField(Index)) <> p_asAHValue(1)
60          Else
70             p_bFldChnge = FieldValue(Index) <> Trim(MainForm.txtField(Index))
80          End If
90       End If
End Sub

Private Sub setDefaultValue(ByVal lnIndex As Integer)
10       Select Case FieldType(lnIndex)
   Case "s"    ' string
20          FieldValue(lnIndex) = ""
30       Case "n"    ' numeric without decimal value
40          FieldValue(lnIndex) = 0
50       Case "l"    ' numeric with decimal value
60          FieldValue(lnIndex) = 0#
70       Case "b"    ' binary
80       Case "d"    ' datetime
90          FieldValue(lnIndex) = Now
100      Case "p"    ' date (petsa)
110         FieldValue(lnIndex) = Date
120      Case "t"
130         FieldValue(lnIndex) = Time
140      End Select
End Sub

Private Function isAppDriverOK() As Boolean
10       isAppDriverOK = True
20       If p_oAppDrivr Is Nothing Then
30          MsgBox "Application is not yet set!!!" & vbCrLf & _
               "Can not further proceed!!!", vbCritical, "Warning"
40          isAppDriverOK = False
50       End If
End Function

Private Sub Field2Control()
10       Dim lors As Recordset
20       Dim loText As TextBox
30       Dim lsSQL As String
40       Dim lsOldProc As String
50       Dim lasRef() As String
60       Dim lnIndex As Integer
70       Dim lvValue As Variant
   
80       lsOldProc = p_oAppDrivr.ProcName("Field2Control")
90       'On Error GoTo errProc
   
100      Set lors = New Recordset
   
   ' assign the field value to the textbox
110      With MainForm
120         For Each loText In .txtField
130            lnIndex = loText.Index
140            lvValue = IIf(IsNull(FieldValue(lnIndex)), Empty, FieldValue(lnIndex))
         
150            lsSQL = LookupQuery(lnIndex)
160            If lsSQL = Empty Then
170               .txtField(lnIndex) = Format(lvValue, FieldFormat(lnIndex))
180            Else
190               If lvValue = Empty Then
200                  .txtField(lnIndex) = Empty
210               Else
            
220                  lasRef = Split(LookupReference(lnIndex), "»")
230                  If lors.State = adStateOpen Then lors.Close
            
240                  lsSQL = p_oMod.AddCondition(LookupQuery(lnIndex), lasRef(0) & " = " & p_oMod.strParm(lvValue))
250                  lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
                  
               ' always assumed that the 2nd field contains the definition of the code.
260                  If Not lors.EOF Then
270                     .txtField(lnIndex) = lors(1)
280                  Else
290                     .txtField(lnIndex).Text = String(15, "?")
300                  End If
310                  lors.Close
320               End If
330            End If
340         Next
350      End With

endProc:
360      Set lors = Nothing
370      p_oAppDrivr.ProcName lsOldProc
380      Exit Sub
errProc:
390      ShowError lsOldProc
End Sub

Private Function isRefEmpty() As Boolean
10       Dim lsOldProc As String
   
20       lsOldProc = p_oAppDrivr.ProcName("isRefEmpty")
30       isRefEmpty = False
40       'On Error GoTo errProc
   
50       For pnCtr = 0 To FieldCount
60          If FieldReference(pnCtr) = True Then
70             If isFieldEmpty(pnCtr, FieldValue(pnCtr)) Then
80                isRefEmpty = True
90                GoTo endProc
100            End If
110         End If
120      Next
   
endProc:
130      p_oAppDrivr.ProcName lsOldProc
140      Exit Function
errProc:
150      ShowError lsOldProc
End Function

Private Function isFieldEmpty(ByVal Index As Integer, Value As Variant) As Boolean
10       Dim lsOldProc As String
   
20       lsOldProc = p_oAppDrivr.ProcName("isFieldEmpty")
30       isFieldEmpty = True
40       'On Error GoTo errProc
   
50       Select Case FieldType(Index)
   Case "n", "l"
60          If Value = 0 Then GoTo endProc
70       Case "d", "t", "p"
80          If Value = "01/01/1890" Then GoTo endProc
90       Case "s"
100         If Value = "" Then GoTo endProc
110      End Select
120      isFieldEmpty = False
   
endProc:
130      p_oAppDrivr.ProcName lsOldProc
140      Exit Function
errProc:
150      ShowError lsOldProc
End Function

Private Sub ShowError(ByVal lsProcName As String)
10       With p_oAppDrivr
20          .ShowError "FormDriver", .ProcName(lsProcName), Err.Number, Err.Description, Erl
30       End With
40       With Err
50          .Raise .Number, .Source, .Description
60       End With
End Sub

Private Function HasRights2Save() As Boolean
10       Dim lsOldProc As String
20       Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
30       Dim lasRights() As String, lnRights As Integer
40       Dim lnSeek As Integer, lbSeek As Boolean
   
50       lsOldProc = p_oAppDrivr.ProcName("HasRights2Save")
60       'On Error GoTo errProc
70       HasRights2Save = False
   
80       With p_oAppDrivr
90          On Error GoTo errMenu
100         lasRights = Split(.MDIMain.Controls(.MenuName).Tag, "»")
110         'On Error GoTo errProc
      
120         lbSeek = False
130         If p_nEditMode = xeModeAddNew Then
         ' Add Rights
140            lnRights = CInt(lasRights(1))
150            If lnRights And .UserLevel = 0 Then
160               lnSeek = MsgBox("User Doesn't Have the Right to Create New Record!!!" & _
                              vbCrLf & "Seek Authorized User's Approval?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
170               If lnSeek = vbNo Then GoTo endProc
180               lbSeek = True
190            End If
200         ElseIf p_nEditMode = xeModeUpdate Then
         ' update Rights
210            lnRights = CInt(lasRights(2))
220            If (lnRights And .UserLevel) = 0 Then
230               Select Case .UserLevel
            Case xeEncoder
240                  lbSeek = True
250               Case xeSupervisor
260                  If DateDiff("d", p_oRecMain("dModified"), .ServerDate) > 2 Then lbSeek = True
270               Case xeManager
280                  If DateDiff("d", p_oRecMain("dModified"), .ServerDate) > 7 Then lbSeek = True
290               Case xeAudit
300                  lbSeek = True
310               Case xeSysAdmin
320                  lbSeek = True
330               End Select
340            End If
350         End If
      
360         If lbSeek = False Then
370            HasRights2Save = True
380            GoTo endProc
390         End If
      
400         lnSeek = MsgBox("User Doesn't Have the Right to Update Record!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
410         If lnSeek = vbNo Then GoTo endProc
      
420         If p_oMod.GetApproval(p_oAppDrivr, lnUserRght, lsUserIDxx, lsUserName) = False Then GoTo endProc
      
430         If p_nEditMode = xeModeAddNew Then
440            If lnUserRght And .UserLevel = 0 Then
450               MsgBox "Approving User Doesn't Have the Right to Create New Record!!!" & _
                  "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
460               GoTo endProc
470            End If
480         ElseIf p_nEditMode = xeModeUpdate Then
490            If lnUserRght And .UserLevel = 0 Then
500               Select Case lnUserRght
            Case xeSupervisor
510                  If DateDiff("d", p_oRecMain("dModified"), p_oAppDrivr.ServerDate) <= 2 Then lbSeek = False
520               Case xeManager
530                  If DateDiff("d", p_oRecMain("dModified"), p_oAppDrivr.ServerDate) <= 31 Then lbSeek = False
540               Case xeSysAdmin
550                  If .UserLevel = xeSysAdmin Then lbSeek = False
560               End Select
            
570               If lbSeek Then
580                  MsgBox "Approving User Doesn't Have the Right to Update Record!!!" & _
                     "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
590                  GoTo endProc
600               End If
610            End If
620         End If
630      End With
640      HasRights2Save = True

endProc:
650      p_oAppDrivr.ProcName lsOldProc
660      Exit Function
errProc:
670      ShowError lsOldProc
errMenu:
680      ReDim lasRights(4) As String
   
690      lasRights(1) = 192
700      lasRights(2) = 192
710      lasRights(3) = 192
720      Resume Next
End Function

Private Function HasRights2Delete() As Boolean
10       Dim lsOldProc As String
20       Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
30       Dim lasRights() As String, lnRights As Integer
40       Dim lnSeek As Integer
   
50       lsOldProc = p_oAppDrivr.ProcName("HasRights2Delete")
60       'On Error GoTo errProc
70       HasRights2Delete = False
   
80       With p_oAppDrivr
90          lasRights = Split(.MDIMain.Controls(.MenuName).Tag, "»")
      
100         lnRights = CInt(lasRights(4))
110         If lnRights And .UserLevel > 0 Then
120            HasRights2Delete = True
130            GoTo endProc
140         End If
      
150         lnSeek = MsgBox("User Doesn't Have the Right to Delete Record!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
160         If lnSeek = vbNo Then GoTo endProc
      
170         If p_oMod.GetApproval(p_oAppDrivr, lnUserRght, lsUserIDxx, lsUserName) = False Then GoTo endProc
   
180         If (lnUserRght And .UserLevel) > 0 Then
190            HasRights2Delete = True
200            GoTo endProc
210         End If
      
220         MsgBox "Approving User Doesn't Have the Right to Delete Record!!!" & _
            "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
230         GoTo endProc
240      End With
   
endProc:
250      p_oAppDrivr.ProcName lsOldProc
260      Exit Function
errProc:
270      ShowError lsOldProc
End Function

Function isModify() As Boolean
10       Dim lsOldProc As String
   
20       lsOldProc = p_oAppDrivr.ProcName("isModify")
30       'On Error GoTo errProc
40       isModify = True
   
50       If p_nEditMode = xeModeAddNew Then GoTo endProc
   
60       For pnCtr = 0 To p_nFldCount - 3
70          If IsNull(p_oRecMain(FieldName(pnCtr))) And FieldValue(pnCtr) <> Empty Then GoTo endProc
80          If p_oRecMain(FieldName(pnCtr)) <> FieldValue(pnCtr) Then GoTo endProc
90       Next
   
100      isModify = False
   
endProc:
110      p_oAppDrivr.ProcName lsOldProc
120      Exit Function
errProc:
130      ShowError lsOldProc
End Function




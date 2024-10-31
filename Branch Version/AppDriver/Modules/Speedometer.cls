VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Speedometer"
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
' ProgressBar v1.0
'     An object for displaying a progress bar.
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
'  XerSys [ 10/02/2004 11:20 pm ]
'     Start creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -&H1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private p_oProgress As New frmProgress

Private pbDisplayed As Boolean
Private pbMoveProg As Boolean

Property Let ProgressStatus(Value As String)
10       If pbDisplayed = False Then Exit Property
   
20       p_oProgress.ProgressStatus = Value
End Property

Property Let PrimaryRemarks(Value As String)
10       p_oProgress.PrimaryRemarks = Value
End Property

Property Let SecondaryRemarks(Value As String)
10       p_oProgress.SecondaryRemarks = Value
End Property

Property Get Cancelled() As Boolean
10       Cancelled = p_oProgress.Cancelled
End Property

Property Get PriMaxValue() As Long
10       PriMaxValue = p_oProgress.PrimaryMaxValue
End Property

Property Let PriMaxValue(ByVal Value As Long)
10       p_oProgress.PrimaryMaxValue = Value
20       pbMoveProg = Value <> 0
End Property

Property Get SecMaxValue() As Long
10       SecMaxValue = p_oProgress.SecondaryMaxValue
End Property

Property Let SecMaxValue(ByVal Value As Long)
10       p_oProgress.SecondaryMaxValue = Value
20       pbMoveProg = Value <> 0
End Property

Function MoveProgress(Optional SecRemarks As Variant, Optional PriRemarks As Variant) As Boolean
10       MoveProgress = True
20       If pbDisplayed = False Then Exit Function
30       If pbMoveProg = False Then Exit Function
   
40       If Not IsMissing(PriRemarks) Then p_oProgress.PrimaryRemarks = PriRemarks
50       If Not IsMissing(SecRemarks) Then p_oProgress.SecondaryRemarks = SecRemarks
   
60       p_oProgress.MoveProgress
70       MoveProgress = Not p_oProgress.Cancelled
End Function

Sub InitProgress(Optional ProgressStatus As String, Optional PriMaxValue As Variant, Optional SecMaxValue As Variant)

10       Load p_oProgress
20       If ProgressStatus <> Empty Then p_oProgress.ProgressStatus = ProgressStatus
30       p_oProgress.Show

40       SetWindowPos p_oProgress.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, (SWP_NOMOVE Or SWP_NOSIZE)
   
50       pbMoveProg = False
60       If Not IsMissing(SecMaxValue) Then
70          p_oProgress.SecondaryMaxValue = SecMaxValue
80          If Not IsMissing(PriMaxValue) Then
90             p_oProgress.PrimaryMaxValue = PriMaxValue
100         Else
110            p_oProgress.PrimaryMaxValue = 1
120         End If
130         pbMoveProg = True
140      Else
150         If Not IsMissing(PriMaxValue) Then
160            p_oProgress.PrimaryMaxValue = 1
170            p_oProgress.SecondaryMaxValue = PriMaxValue
180            pbMoveProg = True
190         End If
200      End If
   
210      pbDisplayed = True
End Sub

Sub CloseProgress()
10       Unload p_oProgress
20       Set p_oProgress = Nothing
End Sub


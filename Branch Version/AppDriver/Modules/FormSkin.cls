VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const DT_CENTER = &H1

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private WithEvents p_oImage As PictureBox
Attribute p_oImage.VB_VarHelpID = -1
Private WithEvents p_oForm As Form
Attribute p_oForm.VB_VarHelpID = -1
Private p_oAppDrivr As AppDriver
Private p_nFormType As xeFormType
Private p_bDisClose As Boolean

Property Set Form(oForm As Object)
10       Set p_oForm = oForm
End Property

Property Set AppDriver(oAppDriver As AppDriver)
10       Set p_oAppDrivr = oAppDriver
End Property

Property Let DisableClose(ByVal Value As Boolean)
10       p_bDisClose = Value
End Property

Sub ApplySkin(Optional FormType As Variant)
10       Dim loControl As Control
20       Dim lnOldMode As Integer
30       Dim lasTag() As String
40       Dim lsSkin As String
50       Dim lnCtr As Integer

60       If p_oForm Is Nothing Then Exit Sub
70       p_nFormType = xeFormMaintenance
80       If Not IsMissing(FormType) Then p_nFormType = FormType

90       On Error Resume Next
   
100      With p_oForm
110         lnOldMode = .ScaleMode
120         Set p_oImage = .Controls.Add("VB.PictureBox", "PictureBox")
130         .Font = "MS San Serif"
140         .FontSize = 8
150         .FontBold = True
160         .Tag = "wb0;et0"
170         .AutoRedraw = True
180         .ScaleMode = 1
      
190         p_oImage.Picture = LoadPicture(p_oAppDrivr.AppPath & "\Images\x_button.ico")
200         p_oImage.Visible = True
210         p_oImage.TabStop = False
220         p_oImage.CausesValidation = False
230         p_oImage.ScaleMode = vbTwips
240         p_oImage.BorderStyle = 0
250         p_oImage.Appearance = 0
260         p_oImage.Height = 225
270         p_oImage.Height = 225
280         p_oImage.Top = 70
290         p_oImage.Left = .ScaleWidth - 265
      
300         For Each loControl In .Controls
310            lnCtr = lnCtr + 1
320            If loControl.Tag = Empty Then
330               loControl.ForeColor = p_oAppDrivr.getColor("ET0")
340               loControl.BackColor = p_oAppDrivr.getColor("EB0")
350            Else
360               lasTag = Split(loControl.Tag, ";")
370               loControl.ForeColor = p_oAppDrivr.getColor(lasTag(0))
            
380               Select Case UBound(lasTag)
            Case 0
390                  loControl.BackColor = p_oAppDrivr.getColor("EB0")
400               Case 1
410                  loControl.BackColor = p_oAppDrivr.getColor(lasTag(1))
420               Case 2
430                  loControl.BackColor = p_oAppDrivr.getColor(lasTag(1))
440                  loControl.ForeColorFixed = p_oAppDrivr.getColor(lasTag(2))
450                  loControl.BackColorFixed = p_oAppDrivr.getColor("EB0")
460               Case 3
470                  loControl.BackColor = p_oAppDrivr.getColor(lasTag(1))
480                  loControl.ForeColorFixed = p_oAppDrivr.getColor(lasTag(2))
490                  loControl.BackColorFixed = p_oAppDrivr.getColor(lasTag(3))
500               End Select
510            End If

520            If TypeName(loControl) = "xrButton" Then
530               loControl.BackColor = p_oAppDrivr.getColor("HB1")
540               loControl.BackColorDown = p_oAppDrivr.getColor("HB4")
550               loControl.BorderColorFocus = p_oAppDrivr.getColor("BC0")
560               loControl.BorderColorHover = p_oAppDrivr.getColor("BC1")
570               loControl.ForeColor = p_oAppDrivr.getColor("ET0")
580            End If
590         Next
600         If .Tag <> "" Then
610            lasTag = Split(.Tag, ";")
620            .BackColor = p_oAppDrivr.getColor(lasTag(0))
630         End If
'570         lsSkin = p_oAppDrivr.getPicture(xeFormMaintenance)
'580         .Picture = LoadPicture(p_oAppDrivr.AppPath & "\Images\" & lsSkin)
      
640         .ScaleMode = lnOldMode
650      End With
End Sub

Private Sub p_oForm_Activate()
10       Dim lnOldMode As Integer
20       Dim hRgn As Long
30       Dim lbhRPen As Long
40       Dim R As RECT
50       Dim nCount As Integer
60       Dim nCtr As Integer

70       With p_oForm
80          lnOldMode = .ScaleMode
90          .Cls
100         .ScaleMode = 1
110         .Picture = LoadPicture(p_oAppDrivr.AppPath & "\Images\" & p_oAppDrivr.getPicture(xeFormMaintenance))
120         TextOut .hdc, 25, 6, .Caption, Len(.Caption)

      ' draw the region now
130         BeginPath .hdc
      ' this will create the title bar
140         RoundRect .hdc, 0, 0, .ScaleWidth, 26, 25, 25
      
      ' then set scale mode to pixel for, coz it's hard to calculate the
      '  rounded side, accrdng to my two colleagues
150         .ScaleMode = 3
      ' the body of the form varies depending on the type of form
160         Select Case p_nFormType
      Case xeFormMaintenance
170            RoundRect .hdc, 0, 29, .ScaleWidth + 5, .ScaleHeight - 60, 0, 0
180            RoundRect .hdc, 0, .ScaleHeight - 57, .ScaleWidth + 5, .ScaleHeight, 0, 0
190            EndPath .hdc
         
         'convert the path to a region
200            hRgn = PathToRegion(.hdc)
          
         'set the Window-region
210            SetWindowRgn .hWnd, hRgn, True
         
         'Crete a new pen
220            lbhRPen = CreatePen(0, 1, &HE0E0E0)
   
         'Select our pen into the form's device context and delete the old pen
230            DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rectangle
240            RoundRect .hdc, 0, 29, .ScaleWidth, .ScaleHeight - 61, 0, 0
250            RoundRect .hdc, 0, .ScaleHeight - 57, .ScaleWidth, .ScaleHeight - 1, 0, 0
               
260      Case xeFormQuickSearch
270            RoundRect .hdc, 0, 29, 117, 143, 10, 10
280            RoundRect .hdc, 121, 29, .ScaleWidth, 143, 10, 10
290            RoundRect .hdc, 0, 147, .ScaleWidth, .ScaleHeight, 10, 10
300            EndPath .hdc
         
         'convert the path to a region
310            hRgn = PathToRegion(.hdc)
          
         'set the Window-region
320            SetWindowRgn .hWnd, hRgn, True

         'Crete a new pen
330            lbhRPen = CreatePen(0, 2, &HE0E0E0)
         
         'Select our pen into the form's device context and delete the old pen
340            DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rounded rectangle
350            RoundRect .hdc, 0, 29, 117, 143, 10, 10
360            RoundRect .hdc, 121, 29, .ScaleWidth, 143, 10, 10
370            RoundRect .hdc, 0, 147, .ScaleWidth, .ScaleHeight, 10, 10

380         Case xeFormTransaction
390            RoundRect .hdc, 0, 29, 96, .ScaleHeight - 100, 0, 0
400            RoundRect .hdc, 99, .ScaleHeight, .ScaleWidth, 29, 0, 0

         'close the path bracket
410            EndPath .hdc
          
         'convert the path to a region
420            hRgn = PathToRegion(.hdc)
          
         'set the Window-region
430            SetWindowRgn .hWnd, hRgn, True

         'Crete a new pen
440          lbhRPen = CreatePen(0, 1, &HE0E0E0)
   
         'Select our pen into the form's device context and delete the old pen
450            DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rectangle
460            RoundRect .hdc, 0, 29, 95, .ScaleHeight - 101, 0, 0
470            RoundRect .hdc, 99, .ScaleHeight - 1, .ScaleWidth - 1, 29, 0, 0

480         Case xeFormLedger
490            RoundRect .hdc, -1, 29, .ScaleWidth + 20, .ScaleHeight, 0, 0
    
500            nCount = 50
510               For nCtr = 1 To 10
520                  RoundRect .hdc, .ScaleWidth - 20, nCount, .ScaleWidth - 10, nCount + 10, 10, 10
530                  nCount = nCount + 50
540               Next
         'close the path bracket
550            EndPath .hdc
          
         'convert the path to a region
560            hRgn = PathToRegion(.hdc)
          
         'set the Window-region
570            SetWindowRgn .hWnd, hRgn, True

         'Crete a new pen
580            lbhRPen = CreatePen(0, 1, &HE0E0E0)
   
         'Select our pen into the form's device context and delete the old pen
590            DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rectangle
600            RoundRect .hdc, 0, 29, .ScaleWidth, .ScaleHeight - 1, 0, 0

610         Case xeFormTransDetail
620            RoundRect .hdc, .ScaleWidth - 96, 29, .ScaleWidth, 170, 0, 0
630            RoundRect .hdc, 0, .ScaleHeight, .ScaleWidth - 98, 29, 0, 0

         'close the path bracket
640            EndPath .hdc
          
         'convert the path to a region
650            hRgn = PathToRegion(.hdc)
          
         'set the Window-region
660            SetWindowRgn .hWnd, hRgn, True

         'Crete a new pen
670            lbhRPen = CreatePen(0, 1, &HE0E0E0)
   
         'Select our pen into the form's device context and delete the old pen
680            DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rectangle
690            RoundRect .hdc, .ScaleWidth - 96, 29, .ScaleWidth - 1, 169, 0, 0
700            RoundRect .hdc, 0, .ScaleHeight - 1, .ScaleWidth - 99, 29, 0, 0

710        Case xeFormTransMaintenance
720            RoundRect .hdc, .ScaleWidth - 96, 29, .ScaleWidth, .ScaleHeight - 182, 0, 0
730            RoundRect .hdc, 0, .ScaleHeight, .ScaleWidth - 98, 29, 0, 0

         'close the path bracket
740            EndPath .hdc
          
         'convert the path to a region
750            hRgn = PathToRegion(.hdc)
          
         'set the Window-region
760            SetWindowRgn .hWnd, hRgn, True

         'Crete a new pen
770            lbhRPen = CreatePen(0, 1, &HE0E0E0)
   
         'Select our pen into the form's device context and delete the old pen
780            DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rectangle
790            RoundRect .hdc, .ScaleWidth - 96, 29, .ScaleWidth - 1, .ScaleHeight - 183, 0, 0
800            RoundRect .hdc, 0, .ScaleHeight - 1, .ScaleWidth - 99, 29, 0, 0

810         End Select
820         .ScaleMode = lnOldMode
830      End With
End Sub

Private Sub p_oForm_Deactivate()
10       With p_oForm
20          .Picture = LoadPicture(p_oAppDrivr.AppPath & "\Images\tBarGry.jpg")
30          TextOut .hdc, 25, 6, .Caption, Len(.Caption)
40       End With
End Sub

Private Sub p_oForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
10       If Button = vbLeftButton Then
20          If y < 360 Then
30             Call ReleaseCapture
40             Call SendMessage(p_oForm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
50          End If
60       End If
End Sub

Private Sub p_oForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
10       If Screen.MousePointer <> 0 Then Screen.MousePointer = 0
End Sub

Private Sub p_oForm_Unload(Cancel As Integer)
10       Set p_oImage = Nothing
End Sub

Private Sub p_oImage_Click()
10       If Not p_bDisClose Then Unload p_oForm
End Sub

Private Sub p_oImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
10       p_oImage.Picture = LoadPicture(p_oAppDrivr.AppPath & "\Images\x_button1.ico")
End Sub

Private Sub p_oImage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
10       p_oImage.Picture = LoadPicture(p_oAppDrivr.AppPath & "\Images\x_button.ico")
End Sub

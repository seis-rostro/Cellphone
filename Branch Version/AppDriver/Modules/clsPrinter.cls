VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum PrintOrientation
   xePOPortrait = 1
   xePOLandscape = 2
End Enum

Property Let FontName(vValue As String)
   Printer.Font.Name = vValue
End Property

Property Get FontName() As String
   FontName = Printer.Font.Name
End Property

Property Let FontSize(vValue As Integer)
   Printer.Font.Size = vValue
End Property

Property Get FontSize() As Integer
   FontSize = Printer.Font.Size
End Property

Property Let Orientation(vValue As PrintOrientation)
   Printer.Orientation = vValue
End Property

Property Get Orientation() As PrintOrientation
   Orientation = Printer.Orientation
End Property

Property Let FontBold(vValue As Boolean)
   Printer.Font.Bold = vValue
End Property

Property Get FontBold() As Boolean
   FontBold = Printer.Font.Bold
End Property

Property Let FontItalic(vValue As Boolean)
   Printer.Font.Italic = vValue
End Property

Property Get FontItalic() As Boolean
   FontItalic = Printer.Font.Italic
End Property

Public Sub PrintText(ByVal Text As String, ByVal Row As Integer, ByVal Col As Integer)
   Printer.CurrentX = Col
   Printer.CurrentY = Row
   Printer.Print Text
End Sub

Public Sub PrintNow()
   Printer.EndDoc
End Sub

Private Sub Class_Initialize()
   Printer.Font.Name = "Courier New"
   Printer.Font.Size = 10
   Printer.Font.Bold = False
   Printer.Font.Italic = False
   Printer.Orientation = vbPRORPortrait
   Printer.ScaleMode = vbCharacters
End Sub

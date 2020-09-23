Attribute VB_Name = "modRoundText"
'------------------------------------------------------------
' Project Name: Project1
' Module Name: modRoundText
' Date: 07/05/2001
' Time: 12.29
' Revision:
' Author: NDV Software
'------------------------------------------------------------
' ****************************************************************************************************
' Copyright © 1990 - 2001 NDV Software,
' All rights are reserved, ndv@interfree.it
' ****************************************************************************************************
Option Explicit
Global Const PIGRECO = 3.141592654

Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFacename As String * 33
End Type

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'------------------------------------------------------------
' Name: drawCircularText
' Desc: Draw a circle/arc text
' Type: Public
' Parameters:
'    Obj As Object              Destination Object of the printing (Picture o Printer object)
'    Testo As String            Text to print
'    TextStartAngle As Single   Starting Angle of the text
'    Raggio As Single           Radius of the circle/arc on that is printed the text
'    CX As Integer              Center X coord of the circle
'    CY As Integer              Center Y coord of the circle
'    TextSector As Single       Sector of circle to fill with text. Good value are between 0 and 360.
'                               0 -> All the text is printed on the same point
'                               180 -> the text make a semicircular beginning from the starting angle
'                               360 -> the text make a circle beginning from the starting angle
'    All the other graphics parameters like font,color,font propertyes (bold,italic,...) can be
'    setted on the Obj Object before to pass it to the drawCircularText procedure.
'
' Date: lunedì 7 maggio 2001
' Time: 12.25
' Author: NDV Software
' Revision:
'------------------------------------------------------------
Public Sub drawCircularText(Obj As Object, Testo As String, TextStartAngle As Single, Raggio As Single, CX As Integer, CY As Integer, TextSector As Single)
  On Error GoTo Errore
  Dim F As LOGFONT
  Dim hPrevFont As Long
  Dim hFont As Long
  Dim FontName As String
  Dim I As Integer
  Dim Passo As Single
  
  Passo = TextSector / Len(Testo)   'Angular Step
    
  For I = 1 To Len(Testo)
    F.lfEscapement = 10 * TextStartAngle - (10 * Passo * (I - 1)) 'rotation angle, in tenths (x10)
    FontName = Obj.FontName + Chr$(0) 'null terminated
    F.lfFacename = FontName
    F.lfHeight = (Obj.FontSize * -20) / Screen.TwipsPerPixelY
    hFont = CreateFontIndirect(F)
    hPrevFont = SelectObject(Obj.hdc, hFont)
    Obj.CurrentX = CX + Raggio * Sin((-180 + TextStartAngle - (I - 1) * Passo) * PIGRECO / 180)
    Obj.CurrentY = CY + Raggio * Cos((-180 + TextStartAngle - (I - 1) * Passo) * PIGRECO / 180)
    
    'Uncomment next line of code to see the angular sector containing a letter of the text
    'Obj.Line (CX, CY)-(CurrentX, CurrentY)
    
    Obj.Print Mid(Testo, I, 1)
    hFont = SelectObject(Obj.hdc, hPrevFont)
    DeleteObject hFont
  Next I
  
  Exit Sub
Errore:
  Exit Sub
End Sub

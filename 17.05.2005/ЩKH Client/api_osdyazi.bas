Attribute VB_Name = "api_osdyazý"
Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type
Public Type pAttributes
     fontName As String * 25
     fontSize As Integer
     fontBold As Boolean
     fontColor As Long
     textString As String * 60
     textBufferBox As PictureBox
     textBufferWidth As Integer
     textBufferHeight As Integer
     textLocX As Integer
     textLocY As Integer
     scrBufferBox As PictureBox
     LastX As Integer
     LastY As Integer
End Type
 ' Þimdi ise bu API' leride module'de declare ediniz. Bunlarýda program içerisinde nokta koyarken veya nokta üzerindeki rengi öðrenirken v.s. kullanýcaðýz.

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'  Aþaðýdaki kodlarýda module' de tanýmlayýnýz. Bu yordamlar sizin için ekrana yazý yazacak. :)
Public Sub PrintOnScreen(ByRef textAttrib As pAttributes)
      Dim hDcDsk As Long, hWndDsk As Long
      Dim Rec As RECT
      Dim winW As Long, winH As Long
      Dim X As Long, Y As Long, c As Long, orgC As Long
      ' PictureBox için gerekli olan ayarlari yapalim..
      textAttrib.textBufferBox.Font.Name = textAttrib.fontName
      textAttrib.textBufferBox.fontBold = textAttrib.fontBold
      textAttrib.textBufferBox.ForeColor = textAttrib.fontColor
      textAttrib.textBufferBox.fontSize = textAttrib.fontSize
      textAttrib.textBufferBox.Width = textAttrib.textBufferWidth * Screen.TwipsPerPixelX
      textAttrib.textBufferBox.Height = textAttrib.textBufferHeight * Screen.TwipsPerPixelY
      textAttrib.scrBufferBox.Width = Screen.Width
      textAttrib.scrBufferBox.Height = Screen.Height
      textAttrib.scrBufferBox.BackColor = textAttrib.fontColor
      textAttrib.textBufferBox.AutoRedraw = True
      textAttrib.scrBufferBox.AutoRedraw = True
      textAttrib.textBufferBox.Visible = False
      textAttrib.scrBufferBox.Visible = False

      ' Yaziyi pictureBox' a yazdiralim,
      textAttrib.textBufferBox.Cls
      textAttrib.scrBufferBox.Cls
      textAttrib.textBufferBox.Print textAttrib.textString

      GetWindowRect textAttrib.textBufferBox.hWnd, Rec       ' Picture Box' in boyutlarini alalim..
      winW = Rec.Right - Rec.Left       ' Genisligi ve yüksekligi hesaplayalim
      winH = Rec.Bottom - Rec.Top

      hWndDsk = GetDesktopWindow       ' Ekranin handle ini alalim
      hDcDsk = GetWindowDC(hWndDsk)       ' ve bu handle a ait olan hDc(Handle Direct Call) numarasini
                                                                        ' alalim.
      For X = 0 To winW
            For Y = 0 To winH
                  c = GetPixel(textAttrib.textBufferBox.hdc, X, Y) 'PictureBox üzerindeki rengi alalim,
                  If c = textAttrib.fontColor Then 'Eger secilen renk, belirledigimiz renkse..
                        ' Ekran üzerindeki orjinal rengi alalim ve diger picturebox a yazalim.
                        orgC = GetPixel(hDcDsk, textAttrib.textLocX + X, textAttrib.textLocY + Y)
                        ' Ekran üzerine PictureBox dan aldigimiz rengi koyalim.
                        SetPixel hDcDsk, textAttrib.textLocX + X, textAttrib.textLocY + Y, c
                        ' Diger picturebox a ekran üzerinden aldigimiz rengi koyalim.
                        SetPixel textAttrib.scrBufferBox.hdc, textAttrib.textLocX + X, textAttrib.textLocY + Y, orgC
                        DoEvents
                         textAttrib.LastX = textAttrib.textLocX + X ' En son nokta koyulan koordinatlari kaydedelim.
                  End If
            Next Y ' Y yi döndür.
            textAttrib.LastY = textAttrib.textLocY + Y
      Next X ' X i döndür.
End Sub
Public Sub ClearScreen(ByRef textAttrib As pAttributes)
      Dim hDcDsk As Long, hWndDsk As Long
      Dim Rec As RECT
      Dim winW As Long, winH As Long
      Dim X As Long, Y As Long, c As Long, orgC As Long
      hWndDsk = GetDesktopWindow ' Ekranin handle ini alalim
      hDcDsk = GetWindowDC(hWndDsk)       ' ve bu handle a ait olan hDc(Handle Direct Call) numarasini
                                                                        ' alalim.
      For X = 0 To textAttrib.LastX
            For Y = 0 To textAttrib.LastY
                  c = GetPixel(textAttrib.scrBufferBox.hdc, X, Y)
                  If Not c = textAttrib.fontColor Then
                        SetPixel hDcDsk, X, Y, c ' PictureBox tan alýp ekrana yazalim
                        DoEvents
                  End If
            Next Y
      Next X
End Sub


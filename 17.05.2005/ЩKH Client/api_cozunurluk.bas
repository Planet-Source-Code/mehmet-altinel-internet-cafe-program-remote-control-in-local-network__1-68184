Attribute VB_Name = "api_cozunurluk"
Option Explicit

Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long

Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Dim DevM As DEVMODE

Sub EkranDüzenle(Dikey As Integer, Yatay As Integer)
    Dim Yap&
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    DevM.dmPelsWidth = Yatay
    DevM.dmPelsHeight = Dikey
    Yap = ChangeDisplaySettings(DevM, 0)
End Sub

Function VerilenDeðerlerUygun(Dikey As Integer, Yatay As Integer) As Boolean
    VerilenDeðerlerUygun = False
    If Yatay = 640 And Dikey = 480 Then VerilenDeðerlerUygun = True
    If Yatay = 800 And Dikey = 600 Then VerilenDeðerlerUygun = True
    If Yatay = 1024 And Dikey = 768 Then VerilenDeðerlerUygun = True
    If Yatay = 1152 And Dikey = 864 Then VerilenDeðerlerUygun = True
    If Yatay = 1280 And Dikey = 1024 Then VerilenDeðerlerUygun = True
    If Yatay = 1600 And Dikey = 1200 Then VerilenDeðerlerUygun = True

    Dim Durum As Boolean, I&
    I = 0
    Do
        Durum = EnumDisplaySettings(0&, I&, DevM)
        I = I + 1
    Loop Until (Durum = False)
    If Yatay > DevM.dmPelsWidth Then VerilenDeðerlerUygun = False
    If Dikey > DevM.dmPelsHeight Then VerilenDeðerlerUygun = False
End Function


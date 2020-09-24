Attribute VB_Name = "modFunctions"
'=============================================================================================================================
'
' Developed by Walter A. Narvasa
' jawoltze@edsamail.com.ph
'
' Walter A. Narvasa of
' WANCOM SYSTEMS
'
' Hey sir, Kindly rate this code, if you like it.
'
' READ THIS BEFORE USING THE CODE:
'
' You can study and view the source code for creating your
' own apps, but do not reproduce/release Puppeteer Version 1.0
' or partially for any commercial and/or personal purposes. All
' rights of this product is related to it's author. Any violation
' of above conditions will be treated seriously and is punishable.
'
' I do not have full time to add complete explanation, read the help
' file (click Help->Contents) in Puppeteer Version 1.0. Contact me for
' additional help/suggestions
'
' I recently inveted a technology for streaming audio, and is
' now looking promoters/investors to invest in a web-phone network
' project.
'
' VISIT MY WEBSITE : http://jawoltze.gq.nu/
'
'=============================================================================================================================

Option Explicit

Global Moderator As String

'BITBIT FUNCTION & declare SRCCOPY
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020

'READ/WRITE TO INI SETTINGS
#If Win16 Then
    Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer
    Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal filename As String) As Integer
#Else
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If
   
Function ReadINI(Section, KeyName, filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), filename))
 End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function


' GUI Effects Caption & Background
Public Function GUIfx(xCaption As String, xDestination As PictureBox, _
                    nTop As Integer, nLeft As Integer, nWidth As Integer, _
                    nHeight As Integer, xSource As PictureBox, _
                    xTop As Integer, xLeft As Integer, _
                    xCaptionTop As Integer, xCaptionLeft As Integer, _
                    xCaptionFont As String, xCaptionFontSize As Integer, xCaptionFontBold As Boolean)
    Call BitBlt(xDestination.hDC, nTop, nLeft, nWidth, nHeight, xSource.hDC, xTop, xLeft, SRCCOPY)
    xDestination.Refresh
    xDestination.Font = xCaptionFont ' TYPE OF FONT
    xDestination.FontBold = xCaptionFontBold ' FONT BOLD (True or False)
    xDestination.FontSize = xCaptionFontSize 'FONT SIZE
    xDestination.CurrentX = xCaptionLeft 'LEFT CAPTION POSITION
    xDestination.CurrentY = xCaptionTop 'TOP CAPTION POSITION
    xDestination.Print xCaption
End Function



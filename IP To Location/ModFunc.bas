Attribute VB_Name = "ModFunc"
Option Explicit



Public Function GetIP(sHTML As String) As String

    GetIP = Empty
    
    Dim sBuff() As String: sBuff() = Split(sHTML, vbNewLine)

    Dim iStart As Integer, iEnd As Integer: iStart = 0: iEnd = Empty

    iStart = InStr(1, sBuff(2), "<ip>")

    iEnd = InStr(iStart + 1, sBuff(2), "<")

    GetIP = Mid(sBuff(2), iStart + 4, (iEnd - iStart) - 4)

End Function

Public Function DotToLong(DotIP As String) As Double
    
    Dim sIP() As String
    
    sIP = Split(DotIP, ".")
    
    DotToLong = 16777216 * sIP(0) + 65536 * sIP(1) + 256 * sIP(2) + sIP(3)
    
End Function

Public Function LongToLocation(LongIP As Double) As String
    
    Dim FF As Long
    Dim strFrom As String
    Dim strTo As String
    Dim strAbr As String
    Dim strCountry As String
    Dim strContinent As String
    
    FF = FreeFile
    
    Open App.Path & "\ip-to-country.csv" For Input As #FF
        Do
            Input #FF, strFrom, strTo, strAbr, strCountry, strAbr, strContinent
            strFrom = Replace(strFrom, Chr(34), "")
            If LongIP >= CDbl(strFrom) And LongIP <= CDbl(strTo) Then LongToLocation = strContinent & " - " & strCountry: Exit Do
            DoEvents
        Loop Until EOF(FF)
    Close #FF
    
End Function

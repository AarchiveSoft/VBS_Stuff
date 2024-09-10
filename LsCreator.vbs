Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Sub NewLieferschein()
'
' NewLieferschein Makro
'
' Tastenkombination: Strg+n
'
Set wshshell = CreateObject("wscript.shell")

Set shellobj = CreateObject("WScript.Shell")

Set db1 = CreateObject("ADODB.connection")
db1.Open "A3", "", ""

'Adressnummer aus der Aktiven Zelle extrahieren ( nur numerische werte)
For i = 1 To Len(ActiveCell.Value)
   eachletter = Mid(ActiveCell.Value, i, 1)
   If IsNumeric(eachletter) Then
   Adressnummer = Adressnummer & eachletter
   End If
Next

sql1 = "SELECT a.a1_adress_id from org.adresse a where a.a1_adressnummer=" & Adressnummer
Set rs1 = db1.Execute(sql1)

A3execStr = "C:\a3\A3.exe " & Chr(34) & "a3://adress/edit/" & rs1("a1_adress_id") & Chr(34)
shellobj.Run A3execStr
sleeptime = 25

Sleep 3000
'shellobj.AppActivate "A3 Business Software"

shellobj.SendKeys "{F6}"                                                                                                                ' Neuer Auftrag
Sleep 300

Dim init As String

init = ""

s = "j"
'MsgBox Environ("UserName")
Select Case Environ("UserName")
    Case Is = "urs.ziswiler"
        s = "n"
        init = "uz"
    Case Is = "martina.aulestia"
        s = "n"
        init = "ma"
    Case Is = "kurt.stoeckli"
        s = "n"
        init = "ks"
    Case Is = "aaron.hafner"
        s = "n"
        init = "ah"
    Case Is = "jasmin.vetsch"
        s = "n"
        init = "jv"
    Case Is = "jens.roecker"
        s = "n"
        init = "jr"
    Case Is = "zh-rentstation"
        s = "n"
        init = "rs"
    Case Is = "zh-frontdesk"
        s = "n"
        init = "fd"
End Select
 ' MsgBox s
  
shellobj.SendKeys s                                                                                                                      ' Ja, Ittigen / Nein Zürich
Sleep sleeptime

shellobj.SendKeys Format(Now(), "DD.MM.YYYY")         ' Bestell-Datum
Sleep sleeptime


' Das Datum von der 3. Kolonne holen als Lieferdatum. von der Zeile 1 über der Adressnummer (z.b. abgeholt),
' wenn auf der zeile kein datum in der 3. Kolonne steht, so wird das datum bis zu 3 zeilen höher gesucht.
If Not IsDate(Cells(ActiveCell.Row - 1, 3).Value) Then
        If Not IsDate(Cells(ActiveCell.Row - 2, 3).Value) Then
            If Not IsDate(Cells(ActiveCell.Row - 3, 3).Value) Then
              If Not IsDate(Cells(ActiveCell.Row - 4, 3).Value) Then
                 s = Now()
             Else
                 s = Cells(ActiveCell.Row - 4, 3).Value
             End If
        Else
            s = Cells(ActiveCell.Row - 3, 3).Value
        End If
    Else
        s = Cells(ActiveCell.Row - 2, 3).Value
    End If
Else
    s = Cells(ActiveCell.Row - 1, 3).Value
End If

Dim tm As String
Dim currentCell As Range
Dim activeColor As Long

' Get the background color of the active cell
activeColor = ActiveCell.Interior.color

' Check above the active cell
Set currentCell = ActiveCell.Offset(-1, 0)
Do While currentCell.Interior.color = activeColor
    If InStr(1, currentCell.Value, "tm", vbTextCompare) > 0 Then
        tm = currentCell.Value
        Exit Do ' Exit the loop if "tm" is found
    End If
    Set currentCell = currentCell.Offset(-1, 0) ' Move one row up
Loop

' If "tm" wasn't found above, check below the active cell
If tm = "" Then
    Set currentCell = ActiveCell.Offset(1, 0)
    Do While currentCell.Interior.color = activeColor
        If InStr(1, currentCell.Value, "tm", vbTextCompare) > 0 Then
            tm = currentCell.Value
            Exit Do ' Exit the loop if "tm" is found
        End If
        Set currentCell = currentCell.Offset(1, 0) ' Move one row down
    Loop
End If

' If no match is found, assign default value
If tm = "" Then
    tm = "1x TM"
End If

's = Cells(ActiveCell.Row - 1, 4).Value
shellobj.SendKeys Format(s, "DD.MM.YYYY")               ' Liefer-Datum
Sleep sleeptime

shellobj.SendKeys ActiveCell.Offset(-1, 0).Value        ' Versandart
Sleep sleeptime

shellobj.SendKeys "{Enter}{Enter}{Enter}"
Sleep sleeptime

If Mid(ActiveCell.Value, 1, 1) = "M" Then s = "Rent Profoto"      ' Bemerkung
If Mid(ActiveCell.Value, 1, 1) = "R" Then s = "Profoto Reparatur-Ersatz"
If Mid(ActiveCell.Value, 1, 1) = "T" Then s = "Test Profoto"

shellobj.SendKeys s
Sleep sleeptime

shellobj.SendKeys "{Enter}{Enter}"
Sleep sleeptime
shellobj.SendKeys "l{Enter}{Enter}"                  ' Lieferschein
Sleep sleeptime

'Erste TextZeile:
s = "tx" & Mid(ActiveCell.Value, 1, 1) & "LS"           'Miete, Test, Rep.-Ersatz
shellobj.SendKeys s

shellobj.SendKeys "{Enter}^(r)"
shellobj.SendKeys "{Enter}^(r)"

' Suchstring aus der Aktuellen Zelle lesen und damit durch alle offenen Execel-iles tingeln um den strin auf der selben zeile wieder zu finden:
f = Cells(ActiveCell.Row, ActiveCell.Column).Value

Dim datei As Workbook, sheet As Worksheet, text As String

For Each datei In Workbooks
    text = text & "Workbook: " & datei.Name & vbNewLine & "Worksheets: " & vbNewLine

        For Each sheet In datei.Worksheets
            'text = text & sheet.Name & vbNewLine
        For i = 1 To Columns.Count
        ' MsgBox Workbooks(datei.Name).Worksheets(sheet.Name).Cells(ActiveCell.Row, i).Value
            t = Workbooks(datei.Name).Worksheets(sheet.Name).Cells(ActiveCell.Row, i).Value
            If f = t Then
                Workbooks(datei.Name).Worksheets(sheet.Name).Cells(ActiveCell.Row, i).Font.Italic = True
                Workbooks(datei.Name).Worksheets(sheet.Name).Cells(ActiveCell.Row, i).Offset(-1, 0).Font.Italic = True
                Workbooks(datei.Name).Worksheets(sheet.Name).Cells(ActiveCell.Row, i).Offset(1, 0).Font.Italic = True
                Workbooks(datei.Name).Worksheets(sheet.Name).Cells(ActiveCell.Row, i).Offset(-2, 0).Value = "ok " & init
                s = Workbooks(datei.Name).Worksheets(sheet.Name).Cells(8, i).Value
                shellobj.SendKeys s
                Sleep sleeptime
                If ((Mid(ActiveCell.Value, 1, 1) = "t") Or (Mid(ActiveCell.Value, 1, 1) = "T")) Then      ' 1. Detailzeile
                    shellobj.SendKeys "{Enter}1{Enter}{Enter}{Enter}0{Enter}0{Enter}{Enter}"
                Else
                    shellobj.SendKeys "{Enter}1{Enter}{Enter}{Enter}0{Enter}{Enter}{Enter}"
                End If
                
                
            End If
        Next i
        
        Next sheet

        'text = text & vbNewLine
Next datei

' MsgBox text


's = Cells(7, ActiveCell.Column).Value

'shellobj.SendKeys s                               ' Artikelnummer
'Sleep sleeptime
     

'If ((Mid(ActiveCell.Value, 1, 1) = "t") Or (Mid(ActiveCell.Value, 1, 1) = "T")) Then      ' 1. Detailzeile
'    shellobj.SendKeys "{Enter}1{Enter}{Enter}{Enter}0{Enter}0{Enter}{Enter}"
'Else
'    shellobj.SendKeys "{Enter}1{Enter}{Enter}{Enter}0{Enter}{Enter}{Enter}"
'End If


'Letzte TextZeile:
s = "txmie"                                      'Unterschriftsfeld
shellobj.SendKeys "{Enter}^(r)" & s & "{Enter}"
shellobj.SendKeys "{Enter}" & "{Enter}" & "{Enter}" & tm & "{TAB}" & "f" & "{Enter}"
Sleep sleeptime

End Sub


Sub OpenLieferschein()
'
' OpenLieferschein Makro
'
' Tastenkombination: Strg+o
'
Set wshshell = CreateObject("wscript.shell")

Set shellobj = CreateObject("WScript.Shell")

Set db1 = CreateObject("ADODB.connection")
db1.Open "A3", "", ""

'Adressnummer = ActiveCell.Value
For i = 1 To Len(ActiveCell.Value)
   eachletter = Mid(ActiveCell.Value, i, 1)
   If IsNumeric(eachletter) Then
   Adressnummer = Adressnummer & eachletter
   End If
Next


sql1 = "SELECT a.a1_adress_id from org.adresse a where a.a1_adressnummer=" & Adressnummer
Set rs1 = db1.Execute(sql1)

A3execStr = "C:\a3\A3.exe " & Chr(34) & "a3://adress/edit/" & rs1("a1_adress_id") & Chr(34)
shellobj.Run A3execStr
sleeptime = 25

Sleep 3000
'shellobj.AppActivate "A3 Business Software"

shellobj.SendKeys "{F7}"         ' Neuer Auftrag

End Sub



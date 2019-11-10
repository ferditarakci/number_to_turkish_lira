<% @Language= VBScript %>

<%

' @date		: 27/04/2012
' @package	: NumberToTurkishLira
' @author 	: Ferdi Tarakci
' @web		: https://www.ferditarakci.com
' @contact	: bilgi@ferditarakci.com

Response.Charset = "utf-8"
Response.CodePage = 65001


Function NumberToTurkishLira(ByVal Tutar)
	Dim Negatif, Birler, Onlar, Yuzler, Katlar, strTL, Arr, IntBoyut, i

	If Not isNumeric(Tutar) Then Exit Function

	Negatif = ""
	If (Tutar < 0) Then Negatif = "Eksi "

	Birler = Array("", "Bir", "İki", "Üç", "Dört", "Beş", "Altı", "Yedi", "Sekiz", "Dokuz")
	Onlar  = Array("", "On", "Yirmi", "Otuz", "Kırk", "Elli", "Altmış", "Yetmiş", "Seksen", "Doksan")
	Yuzler = Array("", "Yüz", "İki Yüz", "Üç Yüz", "Dört Yüz", "Beş Yüz", "Altı Yüz", "Yedi Yüz", "Sekiz Yüz", "Dokuz Yüz")
	Katlar = Array("Lira", "Bin", "Milyon", "Milyar", "Trilyon", "Katrilyon")

	Tutar = Split(Round(Abs(Tutar), 2), ",")

	Arr = Split(FormatNumber(Tutar(0), 0), ".")
	IntBoyut = Ubound(Arr)

	strTL = ""
	For i = 0 To IntBoyut
		strTL = strTL & Yuzler(Int(Arr(i) / 100) Mod 10) & " "

		strTL = strTL & Onlar(Int(Arr(i) / 10) Mod 10) & " "

		If Not (Tutar(0) >= 1000 And Tutar(0) < 2000) Or Arr(i) > 1 Then _
			strTL = strTL & Birler(Arr(i) Mod 10) & " "

		If (Tutar(0) >= 1000 And Tutar(0) < 2000) Or Not Arr(i) = 0 Then _
			strTL = strTL & Katlar(IntBoyut - i) & " "
	Next

	strTL = Trim(strTL)

	If UBound(Tutar) = 1 Then
		If Len(Tutar(1)) = 1 Then Tutar(1) = Tutar(1) & "0"
		If strTL <> "" Then strTL = strTL & ", "
		strTL = strTL & Onlar(int(Tutar(1) / 10) Mod 10) & " "
		strTL = strTL & Birler(Tutar(1) Mod 10) & " "
		strTL = Negatif & strTL & "Kuruş"
	End If

	NumberToTurkishLira = Trim(strTL)
End Function



'################################################################



Num = -0.5
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)
Response.Write "<br><br>"

Num = -13.11
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)
Response.Write "<br><br>"

Num = -9999999999.99
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)
Response.Write "<br><br>"

Num = 11111111111.11
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)
Response.Write "<br><br>"

Num = 500
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)
Response.Write "<br><br>"

Num = 444444444444.44
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)
Response.Write "<br><br>"

Num = 9.9
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)
Response.Write "<br><br>"

Num = 0.9
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)
Response.Write "<br><br>"

Num = 12.12
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)
Response.Write "<br><br>"

Num = 59421.45
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)
Response.Write "<br><br>"

Num = 1000
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)
Response.Write "<br><br>"

Num = 1985
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)
Response.Write "<br><br>"

Num = 40001.32
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)  
Response.Write "<br><br>"

Num = 9000001
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)  
Response.Write "<br><br>"

Num = 9458761
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)  
Response.Write "<br><br>"

Num = 2147483647.99
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)  
Response.Write "<br><br>"

Num = 2250458761.455
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)  
Response.Write "<br><br>"

Num = 7343457483664.82
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)  
Response.Write "<br><br>"

Num = 650011257345.45
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)  
Response.Write "<br><br>"

Num = 95001125453345.80
Response.Write FormatNumber(Num)
Response.Write "<br>"
Response.Write NumberToTurkishLira(Num)  
Response.Write "<br><br>"


%>

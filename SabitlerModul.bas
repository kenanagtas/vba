Attribute VB_Name = "SabitlerModul"
Public Const Dizin As String = "C:\Users\kenanagtas\Desktop\2022-2023\yýllýk plan\yýllýk plan\yýllýk plan\"
Public Const calismaKitabi As String = "YILLIK PLAN MAKRO SON.xlsm"
Public Const DonemSutunu As Integer = 1
Public Const SinifSutunu As Integer = 2
Public Const GrupSutunu As Integer = 3
Public Const DersSutunu As Integer = 4
Public Const OgretmenSutunu As Integer = 5
Public Const MudurSutunu As Integer = 6
Public Const BaslamaSutunu As Integer = 7
Public Const bitisSutunu As Integer = 8
Public Const YasaSutunu As Integer = 9
Public Const AraTatilSutunu As Integer = 10
Public Const MatYaziliSutunu  As Integer = 11
Public Const MtuYaziliSutunu As Integer = 12
Public Const Teknikler As String = "Soru-cevap eleþtirel düþünme, yaratýcý düþünme, araþtýrma yapma, sorun çözme, sosyal ve kültürel katýlým, giriþimcilik, iletiþim kurma, empati kurma, öz denetim, öz güven, yaratýcýlýk, kararlýlýk, liderlik"
Public Const AracGerec As String = "Etkileþimli tahta sunularý ve EBA materyalleri. MEB Ders Kitabý Multimedya Araçlarý Çalýþma Yapraklarý ve Etkinlikler Gözlem formlarý, Anekdotlar"
Public Const holidayDates = "13.10,29.10,10.11,24.11,13.11,27.12,01.01,18.03,08.04,23.04,01.05,19.05"
Public Const MatexamDates = "01.11,29.12,11.03,20.05"
Public Const MtuexamDates = "03.11,04.01,15.03,24.05"
Public Const siraSutunu As Integer = 3
Public Const haftaSutunu As Integer = 4
Public Const saatSutunu As Integer = 5
Public Const alanSutunu As Integer = 6
Public Const kazanimSutunu As Integer = 7
Public Const konuSutunu As Integer = 8
Public Const teknikSutunu As Integer = 9
Public Const aracSutunu As Integer = 10
Public Const ozelGunSutunu As Integer = 11
Public Const yasaMtu As Integer = 2
Public Const yasaMat As Integer = 3
Public Const yasaSecMat As Integer = 4
Public Const yasaFlMat As Integer = 5


Public Function SortDates(dates() As String) As String()
    Dim i As Integer, j As Integer
    Dim temp As String
    For i = LBound(dates) To UBound(dates) - 1
        For j = i + 1 To UBound(dates)
            If CDate(dates(i) & "." & Year(Now())) > CDate(dates(j) & "." & Year(Now())) Then
                temp = dates(i)
                dates(i) = dates(j)
                dates(j) = temp
            End If
        Next j
    Next i
    SortDates = dates
End Function

Public Function TatilAdi(ozelGun As String, index As Integer) As String

    Dim ad As String
    Dim tatiller As New Collection
    ad = ""
    
    tatiller.Add "13 Ekim Ankara'nýn Baþkent Oluþu", "13.10"
    tatiller.Add "29 Ekim Cumhuriyet Bayramý", "29.10"
    tatiller.Add "10 Kasým Atatürk'ü Anma Haftasý", "10.11"
    tatiller.Add "I. ARA TATÝL", "13.11"
    tatiller.Add "24 Kasým Öðretmenler Günü", "24.11"
    tatiller.Add "27 Aralýk Atatürk'ün Ankara'ya Geliþi Ünlü Matematikçi Cahit Arf'i Anma Haftasý Etkinlikleri", "27.12"
    tatiller.Add "1 Ocak Yýlbaþý Tatili", "01.01"
    tatiller.Add "18 Mart Çanakkale Zaferi ve Çanakkale Þehitlerini Anma Günü", "18.03"
    tatiller.Add "II. ARA TATÝL", "08.04"
    tatiller.Add "23 Nisan Ulusal Egemenlik ve Çocuk Bayramý", "23.04"
    tatiller.Add "19 Mayýs Atatürk'ü Anma ve Gençlik ve Spor Bayramý", "19.05"
    tatiller.Add "1 Mayýs Emek ve Dayanýþma Günü", "01.05"
    
    Dim MatexamArray() As String
    If YillikPlanFormv3Beta.dersCB.value = "MATEMATÝK" Then
        MatexamArray = Split(MatexamDates, ",")
    Else
         MatexamArray = Split(MtuexamDates, ",")
    End If
    tatiller.Add "1. YAZILI", MatexamArray(0)
    tatiller.Add "2. YAZILI", MatexamArray(1)
    tatiller.Add "1. YAZILI", MatexamArray(2)
    tatiller.Add "2. YAZILI", MatexamArray(3)
    On Error Resume Next
    ad = tatiller(ozelGun)
    If Err.Number <> 0 Then
        ad = ""
    End If
    On Error GoTo 0
    
    TatilAdi = ad

End Function
Public Function CheckHoliday(currentDate As Date) As String
    Dim holidays() As String
    Dim i As Integer
    Dim yil As Integer
    Dim hedef As Date
    Dim ozelGunler As String
    Dim tatilgunu As String
    Dim combinedResult As String
    tatilgunu = ""
    
    If YillikPlanFormv3Beta.dersCB.text = "MATEMATÝK" Then
        ozelGunler = MatexamDates & "," & holidayDates
    Else
        ozelGunler = MtuexamDates & "," & holidayDates
    End If
    
    holidays = Split(ozelGunler, ",")
    holidays = SortDates(holidays)
    combinedResult = ""
   
    For i = 0 To UBound(holidays)

        yil = Year(currentDate)
        hedef = CDate(holidays(i) & "." & yil)
        If hedef < currentDate Then
            hedef = CDate(holidays(i) & "." & yil + 1)
        End If
        If hedef >= currentDate And hedef <= currentDate + 6 Then
            If combinedResult <> "" Then combinedResult = combinedResult & vbCrLf
            combinedResult = combinedResult & TatilAdi(holidays(i), i)
        End If
    Next i
    
    CheckHoliday = combinedResult
End Function



Attribute VB_Name = "SabitlerModul"
Public Const Dizin As String = "C:\Users\kenanagtas\Desktop\2022-2023\y�ll�k plan\y�ll�k plan\y�ll�k plan\"
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
Public Const Teknikler As String = "Soru-cevap ele�tirel d���nme, yarat�c� d���nme, ara�t�rma yapma, sorun ��zme, sosyal ve k�lt�rel kat�l�m, giri�imcilik, ileti�im kurma, empati kurma, �z denetim, �z g�ven, yarat�c�l�k, kararl�l�k, liderlik"
Public Const AracGerec As String = "Etkile�imli tahta sunular� ve EBA materyalleri. MEB Ders Kitab� Multimedya Ara�lar� �al��ma Yapraklar� ve Etkinlikler G�zlem formlar�, Anekdotlar"
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
    
    tatiller.Add "13 Ekim Ankara'n�n Ba�kent Olu�u", "13.10"
    tatiller.Add "29 Ekim Cumhuriyet Bayram�", "29.10"
    tatiller.Add "10 Kas�m Atat�rk'� Anma Haftas�", "10.11"
    tatiller.Add "I. ARA TAT�L", "13.11"
    tatiller.Add "24 Kas�m ��retmenler G�n�", "24.11"
    tatiller.Add "27 Aral�k Atat�rk'�n Ankara'ya Geli�i �nl� Matematik�i Cahit Arf'i Anma Haftas� Etkinlikleri", "27.12"
    tatiller.Add "1 Ocak Y�lba�� Tatili", "01.01"
    tatiller.Add "18 Mart �anakkale Zaferi ve �anakkale �ehitlerini Anma G�n�", "18.03"
    tatiller.Add "II. ARA TAT�L", "08.04"
    tatiller.Add "23 Nisan Ulusal Egemenlik ve �ocuk Bayram�", "23.04"
    tatiller.Add "19 May�s Atat�rk'� Anma ve Gen�lik ve Spor Bayram�", "19.05"
    tatiller.Add "1 May�s Emek ve Dayan��ma G�n�", "01.05"
    
    Dim MatexamArray() As String
    If YillikPlanFormv3Beta.dersCB.value = "MATEMAT�K" Then
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
    
    If YillikPlanFormv3Beta.dersCB.text = "MATEMAT�K" Then
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



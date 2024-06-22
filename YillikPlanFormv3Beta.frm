VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} YillikPlanFormv3Beta 
   Caption         =   "UserForm1"
   ClientHeight    =   5980
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7580
   OleObjectBlob   =   "YillikPlanFormv3Beta.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "YillikPlanFormv3Beta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AktifSatir As Integer
Public AktifSutun As Integer
Public KazanimIndex As Integer
Public AlanIndex As Integer
Public Alan As String
Public OncekiAlan As String
Public SayfaAdi As String

Sub FormBileseneniDoldur(ByRef kontrol As Object, sutun As Integer, Optional satir As Integer)
   If TypeOf kontrol Is MSForms.TextBox Then
        kontrol.value = Cells(2, sutun).value
    ElseIf TypeOf kontrol Is MSForms.ComboBox Or TypeOf kontrol Is MSForms.ListBox Then
        sonSatir = Cells(Rows.Count, sutun).End(xlUp).row
        For i = 2 To sonSatir
            kontrol.AddItem Cells(i, sutun).value
        Next i
        
    End If
End Sub
Function SinifDersGrup() As String
    Dim sayfa As String
    Dim snf As String
    Dim grp As String
    Dim drs As String
    snf = Replace(sinifCB.value, ".SINIF", "")
    Select Case grupCB.value
        Case "ANADOLU LÝSESÝ": grp = "and"
        Case "FEN LÝSESÝ": grp = "fl"
    End Select
    Select Case dersCB.value
        Case "MATEMATÝK", "FEN LÝSESÝ MATEMATÝK", "SEÇMELÝ MATEMATÝK"
            drs = "mat"
        Case "MATEMATÝK TARÝHÝ VE UYGULAMALARI"
            drs = "mtu"
    End Select
    sayfa = snf & grp & drs
    SinifDersGrup = sayfa
End Function

Sub SinifDersGrupDongu(i, j, k)
    snf = Replace(sinifCB.List(i), ".SINIF", "")
    
    Select Case grupCB.List(j - 1)
        Case "ANADOLU LÝSESÝ": grp = "and"
        Case "FEN LÝSESÝ": grp = "fl"
    End Select
    
    Select Case dersCB.List(k - 1)
        Case "MATEMATÝK", "FEN LÝSESÝ MATEMATÝK", "SEÇMELÝ MATEMATÝK": drs = "mat"
        Case "MATEMATÝK TARÝHÝ VE UYGULAMALARI": drs = "mtu"
    End Select
End Sub
Sub MetinKutusuDoldur(ByRef kutu As Object, deger As String)
    kutu.value = deger
End Sub

Function SayfaVar(ad As String) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ThisWorkbook
    Dim varMi As Boolean
    varMi = False
    For Each ws In wb.Sheets ' Tüm sayfalarý dolaþ
        If ws.Name = ad Then
            varMi = True ' Sayfa adýný bulursanýz varMi'yi True yap
            Exit For ' Sayfayý bulduðunuzda döngüden çýkýn
        End If
    Next ws
    SayfaVar = varMi
    
End Function




Private Sub CommandButton1_Click()


    Dim ws As Worksheet
    Dim newWorkbook As Workbook
    Dim sourceWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim i As Integer
    
    ' Kaynak çalýþma kitabýný tanýmla
    Set sourceWorkbook = ThisWorkbook
    
    ' Yeni bir çalýþma kitabý oluþtur
    Set newWorkbook = Workbooks.Add
    
    ' Kaynak kitaptaki 2. sayfadan itibaren tüm sayfalarý yeni kitaba kopyala
    For i = 2 To sourceWorkbook.Sheets.Count
        Set ws = sourceWorkbook.Sheets(i)
        ws.Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
        
        ' Yeni kitaptaki son sayfayý seç
        Set targetSheet = newWorkbook.Sheets(newWorkbook.Sheets.Count)
        
        ' Sayfayý etkinleþtir
        targetSheet.Activate
        
        ' Sayfa yönünü ayarla
        ActiveSheet.PageSetup.Orientation = xlLandscape
        
        ' Kaðýt boyutunu ayarla
        ActiveSheet.PageSetup.PaperSize = xlPaperA4
        
        ' Sað kenar boþluðunu ayarla
        ActiveSheet.PageSetup.RightMargin = Application.InchesToPoints(0)
        
        
        ' Yazdýrma alanýný belirlemeye çalýþ
        ActiveSheet.PageSetup.PrintArea = "A1:K76"
        
        ' Yazdýrma alanýný kontrol et ve VBA penceresine yazdýr
        Debug.Print "Þu anki yazdýrma alaný: " & ActiveSheet.PageSetup.PrintArea
    Next i
    
    ' Yeni kitaptaki ilk (boþ) sayfayý sil
    newWorkbook.Sheets(1).Delete
    
 

    



    newWorkbook.SaveAs "C:\Users\kenanagtas\Desktop\c.xlsx"
    Debug.Print ActiveSheet.PageSetup.PrintArea

    MsgBox "ok "


End Sub

Private Sub dersCB_Change()
    If dersCB.text = "MATEMATÝK" Then
        FormBileseneniDoldur Me.yazili11TB, MatYaziliSutunu, 2
        FormBileseneniDoldur Me.yazili12TB, MatYaziliSutunu, 3
        FormBileseneniDoldur Me.yazili21TB, MatYaziliSutunu, 4
        FormBileseneniDoldur Me.yazili22TB, MatYaziliSutunu, 5
    Else
        FormBileseneniDoldur Me.yazili11TB, MtuYaziliSutunu, 2
        FormBileseneniDoldur Me.yazili12TB, MtuYaziliSutunu, 3
        FormBileseneniDoldur Me.yazili21TB, MtuYaziliSutunu, 4
        FormBileseneniDoldur Me.yazili22TB, MtuYaziliSutunu, 5
    End If
    
End Sub

Private Sub DosyaOB_Click()

End Sub

Private Sub kapatBtn_Click()
    Dim wb As Workbook
    Dim aktifWB As Workbook

    ' Aktif çalýþma kitabýný kaydedip referans alalým
    If Not ActiveWorkbook.Name = calismaKitabi Then
        Workbooks(calismaKitabi).Activate
    End If
    Set aktifWB = ActiveWorkbook
    ' Tüm açýk çalýþma kitaplarýný dolaþalým
'    For Each wb In Workbooks
'        ' Aktif çalýþma kitabý deðilse ve kaydedilmiþse kapat
'        If Not wb Is aktifWB And Not wb.ReadOnly Then
'            wb.Close saveChanges:=True
'        End If
'    Next wb
    Unload Me
End Sub
Private Sub BaslikveYasaAyarlari(ws As Worksheet, ByRef icerik As String)

    snf = sinifCB.value
    grp = grupCB.value
    drs = dersCB.value
    icerik = donemTB.value + " ÖÐRETÝM YILI ÖZEL BÝLFEN ÇAYYOLU " + grp + " " + snf + "LAR "
    If drs = "MATEMATÝK" Then
        If grp = "FEN LÝSESÝ" Then
            icerik = icerik + "FEN LÝSESÝ " + drs + " DERSÝ YILLIK PLANI"
        Else
            If snf = "11. SINIF" Or snf = "12.SINIF" Then
                icerik = icerik + "SEÇMELÝ " + drs + " DERSÝ YILLIK PLANI"
            Else
                icerik = icerik + drs + " DERSÝ YILLIK PLANI"
            End If
        End If
    Else
       icerik = icerik + drs + " DERSÝ YILLIK PLANI"
    End If
'    If ((snf = "11" Or snf = "12") And (grp = "and" And drs = "mat")) Then
'        icerik = donemTB.value + " ÖÐRETÝM YILI ÖZEL BÝLFEN ÇAYYOLU " + grp + " " + snf + "LAR " + "SEÇMELÝ " + drs + " DERSÝ YILLIK PLANI"
'
'    ElseIf ((snf = "11" Or snf = "12") And (grp = "and" And drs = "mat")) Then
'        icerik = donemTB.value + " ÖÐRETÝM YILI ÖZEL BÝLFEN ÇAYYOLU " + grp + " " + snf + "LAR " + "FEN LÝSESÝ " + drs + " DERSÝ YILLIK PLANI"
'
'    Else
'        If grp = "FEN LÝSESÝ" Then
'            icerik = donemTB.value + " ÖÐRETÝM YILI ÖZEL BÝLFEN ÇAYYOLU " + grp + " " + snf + "LAR " + "FEN LÝSESÝ " + drs + " DERSÝ YILLIK PLANI"
'        Else
'            icerik = donemTB.value + " ÖÐRETÝM YILI ÖZEL BÝLFEN ÇAYYOLU " + grp + " " + snf + "LAR " + drs + " DERSÝ YILLIK PLANI"
'        End If
'
'    End If
    
    If drs = "MATEMATÝK" Then
        If grp = "FEN LÝSESÝ" Then
            yasaTB.value = ws.Cells(yasaFlMat, 9).value
        Else
            If snf = "11.SINIF" Or snf = "12. SINIF" Then
                yasaTB.value = ws.Cells(yasaSecMat, 9).value
            Else
                yasaTB.value = ws.Cells(yasaMat, 9).value
            End If
        End If
    Else
       yasaTB.value = ws.Cells(yasaMtu, 9).value
    End If
End Sub
Private Sub olusturBTN_Click()
    Dim dosyaAdi As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim icerik As String
    'Oluþturulacak Yýllýk Plan yeni bir dosya olarak oluþturulduðunda bu dosya adi kullanýlacak
    dosyaAdi = donemTB.value + " Yýllýk Planlar.xlsx"
    SayfaAdi = SinifDersGrup
    If sayfaOB.value = True And sinifCB.value <> "TÜM SINIFLAR" Then
        Set wb = ThisWorkbook ' Aktif çalýþma kitabýný al
        Dim yeniSayfa As Worksheet
        If SayfaVar(SayfaAdi) Then
            Application.DisplayAlerts = False ' Uyarýlarý devre dýþý býrak
            wb.Sheets(SayfaAdi).Delete
            Application.DisplayAlerts = True ' Uyarýlarý geri aç
        End If
        Set yeniSayfa = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
     
        yeniSayfa.Name = SayfaAdi
        
        
        'Yeni sayfanýn tüm çizgileri silinir, yazý tipi belirlenir ve yükseklik ve geniþlikler belirlenir
        HucreAyariYap ThisWorkbook, ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ElseIf DosyaOB.value = True Then
        Set wb = Workbooks.Add
        wb.Activate
        
    End If
    AktifSatir = 2
    AktifSutun = 2
    icerik = ""
    If sinifCB.value <> "TÜM SINIFLAR" Then
        BaslikveYasaAyarlari ThisWorkbook.Sheets(1), icerik
        'icerik = donemTB.value + " ÖÐRETÝM YILI ÖZEL BÝLFEN ÇAYYOLU " + grupCB.value + " " + sinifCB.value + "LAR " + dersCB.value + " DERSÝ YILLIK PLANI"
        Baslik icerik, AktifSatir, AktifSutun
        Olustur
        ThisWorkbook.ActiveSheet.Range("K9:K12").merge
        ThisWorkbook.ActiveSheet.Range("K9").value = "15 Temmuz Demokrasi ve Milli Birlik Günü Etkinlikleri "
        TryToFitRows
        
       
    Else
        For sinifIndex = 1 To 4
            For dersIndex = 1 To 2
                For grupIndex = 1 To 2
                    sinifCB.ListIndex = sinifIndex
                    dersCB.ListIndex = dersIndex - 1
                    grupCB.ListIndex = grupIndex - 1
                    SayfaAdi = SinifDersGrup
                    Set wb = ThisWorkbook ' Aktif çalýþma kitabýný al
                      
                       If SayfaVar(SayfaAdi) Then
                           Application.DisplayAlerts = False ' Uyarýlarý devre dýþý býrak
                           wb.Sheets(SayfaAdi).Delete
                           Application.DisplayAlerts = True ' Uyarýlarý geri aç
                       End If
                       Set yeniSayfa = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
                    
                       yeniSayfa.Name = SayfaAdi
                       'Yeni sayfanýn tüm çizgileri silinir, yazý tipi belirlenir ve yükseklik ve geniþlikler belirlenir
                       HucreAyariYap ThisWorkbook, ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                    
                    BaslikveYasaAyarlari ThisWorkbook.Sheets(1), icerik
                    Baslik icerik, AktifSatir, AktifSutun
                    Olustur
                    ThisWorkbook.ActiveSheet.Range("K9:K12").merge
                    ThisWorkbook.ActiveSheet.Range("K9").value = "15 Temmuz Demokrasi ve Milli Birlik Günü Etkinlikleri "
                    TryToFitRows
                Next grupIndex
            Next dersIndex
            
        Next sinifIndex
    End If
    MsgBox "Tamamlandý"
End Sub
Sub TryToFitRows()
    Dim ws As Worksheet
   Set ws = ThisWorkbook.ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).row
    
    Dim i As Long
    For i = 9 To lastRow
        ws.Rows(i).RowHeight = 4 ' Yüksekliði sýkýþtýr
         ' Þimdi otomatik ayarla
        ws.Range("F" & i).VerticalAlignment = xlCenter
        ws.Range("F" & i).HorizontalAlignment = xlCenter
        ws.Rows(i).AutoFit
    Next i
End Sub


Sub YariYil(baslangic As Date, ByRef satir As Integer, ws As Worksheet)
    
    Dim bas As Date, bit As Date
    Dim basay As String, basgun As String, bitay As String, bitgun As String
    Dim rng As Range

    bas = CDate(baslangic)
    bit = bas + 12
    basay = Format(bas, "MMMM")
    basgun = Format(bas, "dd")
    bitay = Format(bit, "MMMM")
    bitgun = Format(bit, "dd")
    
    Set rng = ws.Range(ws.Cells(satir, "B"), ws.Cells(satir + 1, "K"))
    With rng
        .merge
        .value = "YARIYIL TATÝLÝ (" & basgun & " " & basay & "- " & bitgun & " " & bitay & ")"
        .Font.size = 10
        .Font.bold = True
        .BorderAround ColorIndex:=0, Weight:=xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    satir = satir + 3
End Sub
Public Sub Olustur()
    Dim ws As Worksheet
    Dim kazanimlarWs As Worksheet
    Dim baslangicTarihi As Date
    Dim bitisTarihi As Date
    Dim tarih As Date
    Dim satir As Integer
    Dim sutun As Integer
    Dim oncekiAy As String
    Dim oncekiAyHaftaSayisi As Integer
    Dim yariYilTatili As Date
    DegiskenleriOlustur ws, kazanimlarWs, baslangicTarihi, bitisTarihi, tarih, satir, sutun, oncekiAy, oncekiAyHaftaSayisi
    ws.PageSetup.Orientation = xlLandscape
    yariYilTatili = DateSerial(2024, 1, 22)
    Do While tarih <= yariYilTatili - 7
        AyAdlariniOlustur ws, tarih, oncekiAy, oncekiAyHaftaSayisi, satir, sutun
        HaftalikVerileriAl ws, kazanimlarWs, ay, oncekiAy, tarih, oncekiAyHaftaSayisi, satir, sutun
        SonrakiHafta tarih, satir, oncekiAyHaftaSayisi
    Loop

    ' Döngü bittikten sonra son ayý elle yazdýrma
    Dim hucre As Range
  
    Set hucre = ws.Cells(satir - oncekiAyHaftaSayisi, sutun).Resize(oncekiAyHaftaSayisi, 1)
    hucre.merge
    HucreBicimle hucre, TurkceUCase(oncekiAy), True, True, 9, xlCenter, xlContinuous, xlThin, 90, False, False

    YariYil yariYilTatili, satir, ws
    oncekiAy = ""
      tarih = yariYilTatili + 14
      Do While tarih <= bitisTarihi
        AyAdlariniOlustur ws, tarih, oncekiAy, oncekiAyHaftaSayisi, satir, sutun
        HaftalikVerileriAl ws, kazanimlarWs, ay, oncekiAy, tarih, oncekiAyHaftaSayisi, satir, sutun
        SonrakiHafta tarih, satir, oncekiAyHaftaSayisi
    Loop

    Set hucre = ws.Cells(satir - oncekiAyHaftaSayisi, sutun).Resize(oncekiAyHaftaSayisi, 1)
    hucre.merge
    HucreBicimle hucre, TurkceUCase(oncekiAy), True, True, 9, xlCenter, xlContinuous, xlThin, 90, False, False
  
    AlaniBirlestir
    SatirlariBirlestir
    satir = satir + 1
    Set ws = ThisWorkbook.ActiveSheet
    Bilgiler ws, satir

End Sub
Sub SetBorders(targetRange As Range, borderStyle As Long, borderWeight As Long)
    With targetRange
        .Borders(xlEdgeLeft).LineStyle = borderStyle
        .Borders(xlEdgeTop).LineStyle = borderStyle
        .Borders(xlEdgeBottom).LineStyle = borderStyle
        .Borders(xlEdgeRight).LineStyle = borderStyle
        .BorderAround Weight:=borderWeight
    End With
End Sub

Sub SetFontAndAlign(targetCell As Range, Optional fontSize As Variant, Optional isBold As Boolean = False, Optional hAlign As Variant, Optional vAlign As Variant, Optional cellValue As Variant)
    With targetCell
        .Font.size = fontSize
        .Font.bold = isBold
        .HorizontalAlignment = hAlign
        .VerticalAlignment = vAlign
        .value = cellValue
    End With
End Sub
Sub SetBordersAndMerge(targetRange As Range, borderStyle As Long, borderWeight As Long, Optional fontSize As Variant, Optional hAlign As Variant, Optional vAlign As Variant, Optional cellValue As Variant, Optional isBold As Boolean = False)
    With targetRange
        .merge
        .Borders(xlEdgeLeft).LineStyle = borderStyle
        .Borders(xlEdgeTop).LineStyle = borderStyle
        .Borders(xlEdgeBottom).LineStyle = borderStyle
        .Borders(xlEdgeRight).LineStyle = borderStyle
        .BorderAround Weight:=borderWeight
        .Font.size = fontSize
        .HorizontalAlignment = hAlign
        .VerticalAlignment = vAlign
        .value = cellValue
        .Font.bold = isBold
    End With
End Sub
Sub Bilgiler(ws As Worksheet, satir As Integer)

    'YASA
    
    SetBordersAndMerge ws.Range("B" & satir & ":K" & satir + 5), xlContinuous, xlThin, 8, xlLeft, xlTop, yasaTB.value
    satir = satir + 6
    
    'DERS ÖÐRETMENLERÝ
    SetBordersAndMerge ws.Range("B" & satir & ":H" & satir), xlContinuous, xlThin, , xlCenter, xlCenter, "DERS ÖÐRETMENLERÝ", True
    SetBorders ws.Range("B" & satir & ":H" & satir + 10), xlContinuous, xlThin
    SetBorders ws.Range("I" & satir & ":K" & satir + Me.ogretmenLB.ListCount), xlContinuous, xlThin
    

    SetFontAndAlign ws.Cells(satir + 3, 11), 9, True, xlCenter, xlCenter, baslamatarihTB.value
    SetFontAndAlign ws.Cells(satir + 4, 11), 9, True, xlCenter, xlCenter, mudurTB.value

    SetFontAndAlign ws.Cells(satir + 5, 11), 9, True, xlCenter, xlCenter, "Okul Müdürü"
   
    
    ' ÖÐRETMEN ÝSÝMLERÝ
    satir = satir + 2
    For i = 0 To Me.ogretmenLB.ListCount - 1
        If i Mod 2 = 0 Then
            SetFontAndAlign ws.Cells(satir, 3), 9, , , , Me.ogretmenLB.List(i, 0)
        Else
            SetFontAndAlign ws.Cells(satir, 8), 9, , , , Me.ogretmenLB.List(i, 0)
            satir = satir + 2
        End If
    Next i
    
End Sub
Private Sub DegiskenleriOlustur(ByRef ws As Worksheet, _
                          ByRef kazanimlarWs As Worksheet, _
                          ByRef baslangicTarihi As Date, _
                          ByRef bitisTarihi As Date, _
                          ByRef tarih As Date, _
                          ByRef satir As Integer, _
                          ByRef sutun As Integer, _
                          ByRef oncekiAy As String, _
                          ByRef oncekiAyHaftaSayisi As Integer)
    ' Baþlangýç deðiþkenlerini ayarla
    
   
    Set ws = ThisWorkbook.ActiveSheet
    Set kazanimlarWs = Workbooks.Open(Dizin + "kazanimlar.xlsx").Sheets(SayfaAdi)
    lastRow = kazanimlarWs.Cells(kazanimlarWs.Rows.Count, "D").End(xlUp).row
    baslangicTarihi = DateSerial(2023, 9, 4)
    bitisTarihi = DateSerial(2024, 6, 7)
    tarih = baslangicTarihi
    satir = 9
    sutun = 2
    oncekiAy = ""
    KazanimIndex = 1
    AlanIndex = 1
    oncekiAyHaftaSayisi = 0
    Alan = ""
    OncekiAlan = ""
   
    
End Sub


Sub SatirlariBirlestir()

    Dim ws As Worksheet
    Dim currentRow As Long, lastRow As Long, startRow As Long, endRow As Long
    
    Set ws = ThisWorkbook.ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).row
    Application.DisplayAlerts = False
    startRow = 9

    For currentRow = 9 To lastRow + 1
        If ws.Cells(currentRow, 9).value = "" Or InStr(1, ws.Cells(currentRow, 7).value, "ARA TATÝL") > 0 Or currentRow > lastRow Then
            endRow = currentRow - 1
            If Not IsEmpty(ws.Cells(currentRow, 7).value) Then
                If InStr(1, ws.Cells(currentRow, 7).value, "ARA TATÝL") > 0 Then
                    ' "ARA TATÝL" kelimesi içeren satýrlarda ise sadece 7. ve 8. sütunlarý birleþtir
                    ws.Range(ws.Cells(currentRow, 7), ws.Cells(currentRow, 11)).merge
            
                    ' Merge iþlemi sonrasý, sadece ana hücrenin (ilk hücrenin) formatýný deðiþtirin
                    ws.Cells(currentRow, 7).Font.bold = True
                    ws.Cells(currentRow, 7).Font.size = 10
                End If
            End If
            
            ' Eðer baþlangýç satýrý mevcut satýrdan küçükse, bu aralýktaki hücreleri birleþtir
            If startRow <= endRow Then
                ws.Range(ws.Cells(startRow, 9), ws.Cells(endRow, 9)).merge
                ws.Cells(startRow, 9).value = Teknikler
                 ws.Range(ws.Cells(startRow, 10), ws.Cells(endRow, 10)).merge
                ws.Cells(startRow, 10).value = AracGerec
                startRow = currentRow + 1
            Else
                startRow = currentRow
            End If
        End If
    Next currentRow
    Application.DisplayAlerts = True
End Sub



Sub AlaniBirlestir()

    Dim ws As Worksheet
    Dim baslangicSatir As Long, bitisSatir As Long, sutunNo As Long, sonSatir As Long
    Dim kontrolVeri As Variant, hucre As Range

    Set ws = ThisWorkbook.ActiveSheet
    baslangicSatir = 1
    sutunNo = 6 ' F sütunu için

    ' F sütunundaki son veriyi bul
    sonSatir = ws.Cells(ws.Rows.Count, sutunNo).End(xlUp).row
    
    While baslangicSatir <= sonSatir
        ' Eðer hücre boþsa, bir sonraki satýra geç
        If IsEmpty(ws.Cells(baslangicSatir, sutunNo)) Then
            baslangicSatir = baslangicSatir + 1
            GoTo NextIteration
        End If

        kontrolVeri = ws.Cells(baslangicSatir, sutunNo).value
        ' Ayný veriye sahip ardýþýk satýrlarý bul
        bitisSatir = baslangicSatir
        Do
            If ws.Cells(bitisSatir, sutunNo).value = kontrolVeri And Not IsEmpty(ws.Cells(bitisSatir, sutunNo)) Then
                bitisSatir = bitisSatir + 1
            Else
                Exit Do
            End If
        
        Loop

        ' Birden fazla ayný veriye sahip satýr varsa birleþtir
       If bitisSatir > baslangicSatir Then
            Set hucre = ws.Cells(baslangicSatir, sutunNo).Resize(bitisSatir - baslangicSatir, 1)
            Application.DisplayAlerts = False
            hucre.merge
            hucre.HorizontalAlignment = xlCenter
            hucre.VerticalAlignment = xlCenter
            
            Application.DisplayAlerts = True
        End If
        ' Sonraki veriye geç
        baslangicSatir = bitisSatir + 1

NextIteration:
    Wend

End Sub



Private Function TurkceUCase(ByVal metin As String) As String
    Dim harf As String
    Dim sonuc As String
    Dim i As Integer
    
    sonuc = ""
    For i = 1 To Len(metin)
        harf = Mid(metin, i, 1)
        Select Case harf
            Case "ç"
                harf = "Ç"
            Case "ý"
                harf = "I"
            Case "ð"
                harf = "Ð"
            Case "ö"
                harf = "Ö"
            Case "þ"
                harf = "Þ"
            Case "ü"
                harf = "Ü"
            Case "i"
                harf = "Ý"
            Case Else
                harf = UCase(harf)
        End Select
        sonuc = sonuc & harf
    Next i
    TurkceUCase = sonuc
End Function

Private Sub AyAdlariniOlustur(ByRef ws As Worksheet, _
                                ByVal tarih As Date, _
                                ByRef oncekiAy As String, _
                                ByRef oncekiAyHaftaSayisi As Integer, _
                                ByRef satir As Integer, _
                                ByVal sutun As Integer)
    
    
   
    Dim haftaBaslangic As Date
    Dim haftaBitis As Date
    Dim ay As String
    Dim hucre As Range

    haftaBaslangic = tarih
    haftaBitis = tarih + 7  ' Haftanýn bitiþ günü

    ay = DominantAy(haftaBaslangic, haftaBitis)
    
   ' Dim ay As String
    'ay = Format(tarih, "mmmm")

    If ay <> oncekiAy Then
        If oncekiAy <> "" Then  ' Baþlangýç ayý deðilse
            Set hucre = ws.Cells(satir - oncekiAyHaftaSayisi, sutun).Resize(oncekiAyHaftaSayisi, 1)
            hucre.merge
            HucreBicimle hucre, TurkceUCase(oncekiAy), True, True, 9, xlCenter, xlContinuous, xlThin, 90, False, False
            satir = satir + 1 ' Ay deðiþikliðinde satýrý arttýr
        End If
      

        oncekiAy = ay
        oncekiAyHaftaSayisi = 0
    End If
        'oncekiAyHaftaSayisi = oncekiAyHaftaSayisi + 1
End Sub

Private Function AlanDegisti() As Boolean
    If Alan <> OncekiAlan Then
          AlanIndex = 1
        AlanDegisti = True
      
    Else
        AlanIndex = AlanIndex + 1
        AlanDegisti = False
    End If
End Function

Function AraTatilBul(ByVal ws As Worksheet, ByVal baslangicSatir As Integer, ByVal bitisSatir As Integer) As Integer
    Dim i As Integer
    AraTatilBul = 0
    For i = baslangicSatir To bitisSatir
        If ws.Cells(i, 7).value = "ARA TATÝL" Then
            AraTatilBul = i
            Exit Function
        End If
    Next i
End Function

Private Sub HaftalikVerileriAl(ByRef ws As Worksheet, _
                               ByRef kazanimlarWs As Worksheet, _
                               ByVal ay As String, _
                               ByVal oncekiAy As String, _
                               ByVal tarih As Date, _
                               ByVal oncekiAyHaftaSayisi As Integer, _
                               ByRef satir As Integer, _
                               ByVal sutun As Integer)
    ' Haftalýk veriyi doldur
    Dim hucre As Range
    Dim haftaBaslangic As Date
    Dim haftaBitis As Date
    Dim ozelGun As String

    haftaBaslangic = tarih
    haftaBitis = tarih + 4

    kazanim = kazanimlarWs.Range("D" & KazanimIndex).value
    konu = kazanimlarWs.Range("E" & KazanimIndex).value
    saat = kazanimlarWs.Range("B" & KazanimIndex).value
    Alan = kazanimlarWs.Range("C" & KazanimIndex).value
    ozelGun = CheckHoliday(tarih)
    HaftalikHucreBicimle ws.Cells(satir, siraSutunu), oncekiAyHaftaSayisi + 1, True, 12, False
    HaftalikHucreBicimle ws.Cells(satir, haftaSutunu), CStr(Format(haftaBaslangic, "dd")) & vbCrLf & CStr(Format(haftaBitis, "dd")), True, 9, False, 0, False, False
    HaftalikHucreBicimle ws.Cells(satir, saatSutunu), saat, True, 9, False

 
    HaftalikHucreBicimle ws.Cells(satir, alanSutunu), Alan, True, 9, False, 90, False, False
    
    HaftalikHucreBicimle ws.Cells(satir, kazanimSutunu), kazanim, False, 8, False, 0, False, False
    HaftalikHucreBicimle ws.Cells(satir, konuSutunu), konu, False, 8, False
    HaftalikHucreBicimle ws.Cells(satir, teknikSutunu), "T", False, 7, False, 0, False, False
    HaftalikHucreBicimle ws.Cells(satir, aracSutunu), "A", False, 7, False, 0, False, False
    'MsgBox month(tarih)

    If InStr(ozelGun, "YAZILI") > 0 Then
         
          
       
        HaftalikHucreBicimle ws.Cells(satir, ozelGunSutunu), ozelGun, False, 8, False
        cellValue = ws.Cells(satir, ozelGunSutunu).value
         startPos = InStr(1, cellValue, "1. YAZILI") ' "1. YAZILI" metninin baþladýðý pozisyonu bul
         strLength = Len("1. YAZILI") ' "1. YAZILI" metninin uzunluðu
         If startPos > 0 Then
            ws.Cells(satir, ozelGunSutunu).Characters(start:=startPos, Length:=strLength).Font.bold = True
        End If
        startPos = InStr(1, cellValue, "2. YAZILI") ' "2. YAZILI" metninin baþladýðý pozisyonu bul
         strLength = Len("2. YAZILI") ' "2. YAZILI" metninin uzunluðu
         If startPos > 0 Then
            With ws.Cells(satir, ozelGunSutunu).Characters(start:=startPos, Length:=strLength).Font
                .bold = True
            End With
        End If
      
    Else
        HaftalikHucreBicimle ws.Cells(satir, ozelGunSutunu), ozelGun, False, 8, False
    End If
  

End Sub

Private Sub HaftalikHucreBicimle(ByRef hucre As Range, _
                                ByVal value As String, _
                                bold As Boolean, _
                                size As Integer, _
                                birlestir As Boolean, _
                                Optional donme As Integer = 0, _
                                Optional ecol As Boolean = False, _
                                Optional erow As Boolean = True)
    ' Hücreyi biçimlendir ve deðerini ayarla
'    Set hucre = hucre.Resize(1, 1)
'    If birlestir Then
'        hucre.merge
'    End If
    HucreBicimle hucre, value, False, bold, size, xlCenter, xlContinuous, xlThin, donme, ecol, erow
End Sub
'Private Function DominantAy(ByVal baslangic As Date, ByVal bitis As Date) As String
'    Dim ay1 As String
'    Dim ay2 As String
'    Dim gunSayisi1 As Integer
'    Dim gunSayisi2 As Integer
'
'    ay1 = Format(baslangic, "mmmm")
'    ay2 = Format(bitis, "mmmm")
'
'    ' Haftanýn ilk günü ve son günü ayný aydaysa
'    If ay1 = ay2 Then
'        DominantAy = ay1
'        Exit Function
'    End If
'
'    ' Baslangic ve bitis tarihleri arasindaki gun sayisini hesapla
'    gunSayisi1 = day(baslangic) - 1
'    gunSayisi2 = day(bitis)
'
'    If gunSayisi1 >= gunSayisi2 Then
'        DominantAy = ay1
'    Else
'        DominantAy = ay2
'    End If
'End Function
Private Function DominantAy(ByVal baslangic As Date, ByVal bitis As Date) As String
    Dim gunSayisi1 As Integer
    Dim gunSayisi2 As Integer
    
    ' Baslangic ve bitis tarihleri arasindaki gun sayisini hesapla
    gunSayisi1 = DateDiff("d", baslangic, DateSerial(Year(baslangic), month(baslangic) + 1, 1))
    gunSayisi2 = DateDiff("d", DateSerial(Year(bitis), month(bitis), 1), bitis)
    
    If gunSayisi1 >= gunSayisi2 Then
        DominantAy = Format(baslangic, "mmmm")
    Else
        DominantAy = Format(bitis, "mmmm")
    End If
End Function


Private Sub SonrakiHafta(ByRef tarih As Date, ByRef satir As Integer, ByRef oncekiAyHaftaSayisi As Integer)
    ' Sonraki haftaya geç
    tarih = tarih + 7
    
    satir = satir + 1
 
    KazanimIndex = KazanimIndex + 1
    oncekiAyHaftaSayisi = oncekiAyHaftaSayisi + 1
End Sub




Private Sub sinifCB_Change()
  
End Sub

Private Sub UserForm_Initialize()
    'Aktif çalýþma kitabýný Yýýlýk Plan makro dosyasý olarak belirle.ÇalýþmaKitabi SabitlerModul de tanýmlý
    AktifSatir = 2
    AktifSutun = 2
    If Not ActiveWorkbook.Name = calismaKitabi Then
        Workbooks(calismaKitabi).Activate
    End If
    ThisWorkbook.Sheets("Bilgiler").Activate
    'Varsayýlan olarak yeni sayfada yýllýk plan hazýrlanýr
    sayfaOB.value = True
    
    'Form Bilesenlerini doldur
    'Üçüncü parametre seçimlik parametre ve satýr numrasýný belirler
    'Her Sütunda 2. satýrdan baþlanýr
    'Sütun numarasý parametre olarak gönderilir
    
    FormBileseneniDoldur Me.donemTB, DonemSutunu
    FormBileseneniDoldur Me.sinifCB, SinifSutunu
    FormBileseneniDoldur Me.grupCB, GrupSutunu
    FormBileseneniDoldur Me.dersCB, DersSutunu
    FormBileseneniDoldur Me.baslamatarihTB, BaslamaSutunu
    FormBileseneniDoldur Me.bitistarihTB, bitisSutunu
    FormBileseneniDoldur Me.ogretmenLB, OgretmenSutunu
    FormBileseneniDoldur Me.mudurTB, MudurSutunu
    FormBileseneniDoldur Me.yasaTB, YasaSutunu
    FormBileseneniDoldur Me.araTatil1TB, AraTatilSutunu, 2
    FormBileseneniDoldur Me.yariyilTB, AraTatilSutunu, 3
    FormBileseneniDoldur Me.aratatil2TB, AraTatilSutunu, 4
    FormBileseneniDoldur Me.yazili11TB, MatYaziliSutunu, 2
    FormBileseneniDoldur Me.yazili12TB, MatYaziliSutunu, 3
    FormBileseneniDoldur Me.yazili21TB, MatYaziliSutunu, 4
    FormBileseneniDoldur Me.yazili22TB, MatYaziliSutunu, 5
End Sub

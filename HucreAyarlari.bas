Attribute VB_Name = "HucreAyarlari"
Public Sub HucreAyariYap(ByRef wb As Workbook, ByRef ws As Worksheet)
   ActiveWindow.DisplayGridlines = False
    If Not wb Is ActiveWorkbook Then
        wb.Activate
    End If
    If Not ws Is ActiveSheet Then
        ws.Activate
    End If
    With ws
        .Cells.Font.Name = "Calibri"
        .Columns("A").ColumnWidth = 0.89
        .Columns("B").ColumnWidth = 2.36
        .Columns("C").ColumnWidth = 2.36
        .Columns("D").ColumnWidth = 2.36
        .Columns("E").ColumnWidth = 3.82
        .Columns("F").ColumnWidth = 6.18
        .Columns("G").ColumnWidth = 35.82
        .Columns("H").ColumnWidth = 17.55
        .Columns("I").ColumnWidth = 18.91
        .Columns("J").ColumnWidth = 14.91
        .Columns("K").ColumnWidth = 25.36
        .Columns("L").ColumnWidth = 4.91
    End With
End Sub

Public Sub Baslik(icerik As String, ByRef AktifSatir As Integer, ByRef AktifSutun As Integer)
    Dim hucreBilgileri As Collection
    Set hucreBilgileri = New Collection
    Dim hucre As Range
    With hucreBilgileri
        .Add Array(icerik, True, True, 13, xlCenter, xlContinuous, xlThick, 0, 2, 9, 1, -9)
        .Add Array("SÜRE", True, True, 8, xlCenter, xlContinuous, xlThick, 0, 1, 3, -1, 1) ' Ýçerik, merge, bold, size, alignment, borderStyle, borderWidth, rotation, satirAtlama
        .Add Array("Alt Öðrenme Alaný", True, True, 8, xlCenter, xlContinuous, xlThick, 90, 2, 0, -2, 1)
        .Add Array("KAZANIMLAR", True, True, 8, xlCenter, xlContinuous, xlThick, 0, 2, 0, -2, 1)
        .Add Array("KONULAR", True, True, 8, xlCenter, xlContinuous, xlThick, 0, 2, 0, -2, 1)
        .Add Array("ÖÐRENME VE ÖÐRETME TEKNÝKLERÝ", True, True, 8, xlCenter, xlContinuous, xlThick, 0, 2, 0, -2, 1)
        .Add Array("KULLANILAN EÐÝTÝM TEKNOLOJÝLERÝ,ARAÇ ve GEREÇLERÝ", True, True, 8, xlCenter, xlContinuous, xlThick, 0, 2, 0, -2, 1)
        .Add Array("DEÐERLENDÝRME", True, True, 8, xlCenter, xlContinuous, xlThick, 0, 2, 0, 0, -9)
        .Add Array("AY", True, True, 8, xlCenter, xlContinuous, xlThick, 0, 0, 0, 0, 1)
        .Add Array("HAFTA", True, True, 8, xlCenter, xlContinuous, xlThick, 0, 0, 1, 0, 1)
        .Add Array("SAAT", True, True, 8, xlCenter, xlContinuous, xlThick, 0, 0, 0, 0, 1)
         
    End With

    Dim satir As Integer
    Dim sutun As Integer
    Dim hucreBilgisi As Variant

    satir = AktifSatir
    sutun = AktifSutun

    For Each hucreBilgisi In hucreBilgileri
        bsatir = satir + hucreBilgisi(8) ' Satýr atlama sayýsý
        bsutun = sutun + hucreBilgisi(9) ' Sütun atlama sayýsý
        Set hucre = Range(Cells(satir, sutun), Cells(bsatir, bsutun))
        HucreBicimle hucre, hucreBilgisi(0), hucreBilgisi(1), hucreBilgisi(2), hucreBilgisi(3), hucreBilgisi(4), _
                              hucreBilgisi(5), hucreBilgisi(6), hucreBilgisi(7)

        
        satir = bsatir + hucreBilgisi(10)
        sutun = bsutun + hucreBilgisi(11)
        'MsgBox " bilgisinden sonra " + satir + " satir ileri" + sutun + " sütun ileri gidildi "
    Next hucreBilgisi
    
End Sub

Public Sub HucreBicimle(hucre As Range, ByVal icerik As String, ByVal merge As Boolean, _
                          ByVal bold As Boolean, ByVal size As Integer, ByVal yer As XlHAlign, _
                          ByVal tip As XlLineStyle, ByVal kalinlik As XlBorderWeight, _
                          ByVal dondur As Integer, Optional ecol As Boolean = False, Optional erow As Boolean = True)
    With hucre
   
        .Borders(xlEdgeLeft).LineStyle = tip
        .Borders(xlEdgeTop).LineStyle = tip
        .Borders(xlEdgeBottom).LineStyle = tip
        .Borders(xlEdgeRight).LineStyle = tip


     
        If merge Then
            .merge
            .value = icerik
        Else
            If .value <> "" Then
                .value = .value + icerik
            Else
                .value = icerik
            End If
        End If
        
        .BorderAround Weight:=kalinlik

        .Font.bold = bold
        .Font.size = size
        .HorizontalAlignment = yer
        .VerticalAlignment = yer
        .Orientation = dondur
        .wrapText = True
        If erow Then
            .EntireRow.AutoFit
        End If
        If ecol Then
            .EntireColumn.AutoFit
        End If
    End With
End Sub



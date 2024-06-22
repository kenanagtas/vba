VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tablo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public AktifSatir As Integer
Public AktifSutun As Integer

Public BaslangicSatiri As Integer
Public BaslangicSutunu As Integer
Public satirgenisligi As Integer
Public sutungenisligi As Integer
Public bitissatiri As Integer
Public bitisSutunu As Integer


Public hucre As Range
Public HucreIcerigi As String
Public DonmeAcisi As Integer
Public Kalin As Boolean
Public Punto As Integer
Public MetniKaydir As Boolean
Sub HucreAyarla()

        Set hucre = Range(Cells(BaslangicSatiri, BaslangicSutunu), Cells(bitissatiri, bitisSutunu))
        With hucre
            .wrapText = MetniKaydir
            .merge
            .value = HucreIcerigi
            .Font.bold = Kalin
            .Font.size = Punto
            .Orientation = DonmeAcisi
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .BorderAround _
                ColorIndex:=0, Weight:=xlThick
        End With
End Sub

Public Sub SatirCiz(a)

     bitissatiri = BaslangicSatiri + satirgenisligi
     bitisSutunu = BaslangicSutunu + sutungenisligi
     HucreAyarla
     AktifSatir = BaslangicSatiri
     AktifSutun = bitisSutunu + a
  
      
      
End Sub

Public Sub SutunCiz(a)

      bitissatiri = BaslangicSatiri + satirgenisligi
      bitisSutunu = BaslangicSutunu + sutungenisligi
      HucreAyarla
      AktifSatir = bitissatiri + a
      AktifSutun = BaslangicSutunu
    
End Sub

Public Sub CizgileriKaldir(baslama, bitis, satirgenisligi, sutungenisligi)
    
End Sub






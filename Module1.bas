Attribute VB_Name = "Module1"
Dim Zamanlayici As Double

Public Sub VerileriGuncelle()
    Dim DosyaYolu As String
    Dim DosyaNo As Integer
    Dim Satir As String
    Dim Veri() As String
    Dim SatirNumarasi As Integer
    
    ' Dosya yolu (Excel dosyasý ile ayný dizinde)
    DosyaYolu = ThisWorkbook.Path & "\veri.dat"
    
    ' Dosyayý aç ve oku
    On Error GoTo Hata
    DosyaNo = FreeFile
    Open DosyaYolu For Input As #DosyaNo
    
    ' Ýlk satýrý 3. satýrdan baþlat (Baþlýk için yer býrakýldý)
    SatirNumarasi = 3
    
    ' Excel'deki önceki verileri temizle
    Sheets("Sayfa1").Range("A3:C100").ClearContents
    
    ' Dosyadaki tüm satýrlarý oku
    Do While Not EOF(DosyaNo)
        Line Input #DosyaNo, Satir
        Veri = Split(Satir, ",")
        
        ' Veri kontrolü (Sembol, Bid ve Ask deðerleri alýnmalý)
        If UBound(Veri) >= 2 Then
            Sheets("Sayfa1").Cells(SatirNumarasi, 1).Value = Now() ' Zaman damgasý
            Sheets("Sayfa1").Cells(SatirNumarasi, 2).Value = Veri(0) ' Sembol
            Sheets("Sayfa1").Cells(SatirNumarasi, 3).Value = Veri(1) ' Bid fiyatý
            Sheets("Sayfa1").Cells(SatirNumarasi, 4).Value = Veri(2) ' Ask fiyatý
            SatirNumarasi = SatirNumarasi + 1 ' Bir sonraki satýra geç
        End If
    Loop
    
    Close #DosyaNo ' Dosyayý kapat
    
    ' Zamanlayýcýyý tekrar baþlat (1 saniye sonra tekrar çalýþtýr)
    Zamanlayici = Now + TimeValue("00:00:01")
    Application.OnTime Zamanlayici, "VerileriGuncelle"
    
    Exit Sub

Hata:
    MsgBox "Hata oluþtu: " & Err.Description, vbCritical
    Close #DosyaNo
End Sub

Public Sub IzlemeyiBaslat()
    ' Ýlk çalýþtýrma
    VerileriGuncelle
End Sub

Public Sub IzlemeyiDurdur()
    ' Zamanlayýcýyý iptal et
    On Error Resume Next
    Application.OnTime EarliestTime:=Zamanlayici, Procedure:="VerileriGuncelle", Schedule:=False
    MsgBox "Ýzleme durduruldu.", vbInformation
End Sub


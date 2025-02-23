Attribute VB_Name = "Module1"
Dim Zamanlayici As Double

Public Sub VerileriGuncelle()
    Dim DosyaYolu As String
    Dim DosyaNo As Integer
    Dim Satir As String
    Dim Veri() As String
    Dim SatirNumarasi As Integer
    
    ' Dosya yolu (Excel dosyas� ile ayn� dizinde)
    DosyaYolu = ThisWorkbook.Path & "\veri.dat"
    
    ' Dosyay� a� ve oku
    On Error GoTo Hata
    DosyaNo = FreeFile
    Open DosyaYolu For Input As #DosyaNo
    
    ' �lk sat�r� 3. sat�rdan ba�lat (Ba�l�k i�in yer b�rak�ld�)
    SatirNumarasi = 3
    
    ' Excel'deki �nceki verileri temizle
    Sheets("Sayfa1").Range("A3:C100").ClearContents
    
    ' Dosyadaki t�m sat�rlar� oku
    Do While Not EOF(DosyaNo)
        Line Input #DosyaNo, Satir
        Veri = Split(Satir, ",")
        
        ' Veri kontrol� (Sembol, Bid ve Ask de�erleri al�nmal�)
        If UBound(Veri) >= 2 Then
            Sheets("Sayfa1").Cells(SatirNumarasi, 1).Value = Now() ' Zaman damgas�
            Sheets("Sayfa1").Cells(SatirNumarasi, 2).Value = Veri(0) ' Sembol
            Sheets("Sayfa1").Cells(SatirNumarasi, 3).Value = Veri(1) ' Bid fiyat�
            Sheets("Sayfa1").Cells(SatirNumarasi, 4).Value = Veri(2) ' Ask fiyat�
            SatirNumarasi = SatirNumarasi + 1 ' Bir sonraki sat�ra ge�
        End If
    Loop
    
    Close #DosyaNo ' Dosyay� kapat
    
    ' Zamanlay�c�y� tekrar ba�lat (1 saniye sonra tekrar �al��t�r)
    Zamanlayici = Now + TimeValue("00:00:01")
    Application.OnTime Zamanlayici, "VerileriGuncelle"
    
    Exit Sub

Hata:
    MsgBox "Hata olu�tu: " & Err.Description, vbCritical
    Close #DosyaNo
End Sub

Public Sub IzlemeyiBaslat()
    ' �lk �al��t�rma
    VerileriGuncelle
End Sub

Public Sub IzlemeyiDurdur()
    ' Zamanlay�c�y� iptal et
    On Error Resume Next
    Application.OnTime EarliestTime:=Zamanlayici, Procedure:="VerileriGuncelle", Schedule:=False
    MsgBox "�zleme durduruldu.", vbInformation
End Sub


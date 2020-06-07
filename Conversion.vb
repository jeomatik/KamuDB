Imports Kamu.Objects
Public Class Conversion

#Region "Get Procedures V4"

    Public Function GetParsellerCollectionV4(_ParselTable As DataTable, _ProjeGUID As String) As Collection
        Dim MyParseller As New Collection
        Dim MyMalikler As New Collection
        Dim MyParsel As New Parsel
        Dim MyMalik As Kisi
        Dim LastAda As String = "-1"
        Dim LastParsel As String = "-1"
        Try
            For Each MyRow As DataRow In _ParselTable.Rows
                If LastAda IsNot MyRow("ADA").ToString Or LastParsel IsNot MyRow("PARSEL").ToString Then
                    If MyMalikler.Count > 0 Then
                        MyParsel.Malikler = MyMalikler
                        MyParseller.Add(MyParsel)
                        MyMalikler = New Collection
                        MyMalik = New Kisi
                        MyParsel = New Parsel
                    End If

                    MyParsel = ParselOlustur(_ProjeGUID, MyParsel, MyRow)
                    MyParsel = ParselKodla(MyParsel, MyRow)

                    MyMalik = MalikOlustur(MyRow)
                    MyMalik = MalikKodla(MyMalik, MyRow)
                    MyMalikler.Add(MyMalik)
                    MyMalik = New Kisi

                    LastAda = MyParsel.AdaNo
                    LastParsel = MyParsel.ParselNo
                Else
                    MyMalik = MalikOlustur(MyRow)
                    MyMalik = MalikKodla(MyMalik, MyRow)
                    MyMalikler.Add(MyMalik)
                    MyMalik = New Kisi
                End If
            Next
            MyParsel.Malikler = MyMalikler
            MyParseller.Add(MyParsel)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return MyParseller
    End Function

    Private Shared Function ParselOlustur(_ProjeGUID As String, MyParsel As Parsel, MyRow As DataRow) As Parsel
        If Not IsDBNull(MyRow("PROJE_ID")) Then
            If IsNumeric(MyRow("PROJE_ID")) Then
                MyParsel.ProjeGUID = _ProjeGUID
            Else
                MyParsel.ProjeGUID = _ProjeGUID
            End If
        Else
            MyParsel.ProjeGUID = _ProjeGUID
        End If
        MyParsel.GUID = Guid.NewGuid().ToString("N")
        MyParsel.Il = MyRow("IL").ToString.Trim
        MyParsel.Ilce = MyRow("ILCE").ToString.Trim
        MyParsel.Koy = MyRow("KOY").ToString.Trim
        MyParsel.Mahalle = MyRow("MAHALLE").ToString.Trim
        MyParsel.AdaNo = MyRow("ADA").ToString
        MyParsel.ParselNo = MyRow("PARSEL").ToString
        MyParsel.PaftaNo = MyRow("PAFTA").ToString.Trim
        MyParsel.Cinsi = MyRow("CINSI").ToString.Trim
        MyParsel.Mevki = MyRow("MEVKI").ToString.Trim
        MyParsel.Cilt = MyRow("CILT").ToString.Trim
        MyParsel.Sayfa = MyRow("SAYFA").ToString.Trim
        If Not IsDBNull(MyRow("TAPU_ALANI")) Then
            MyParsel.TapuAlani = MyRow("TAPU_ALANI")
        End If
        If Not IsDBNull(MyRow("DAIMI_IRTIFAK_ALAN")) Then
            MyParsel.IrtifakAlan = MyRow("DAIMI_IRTIFAK_ALAN")
        End If
        If Not IsDBNull(MyRow("GECICI_IRTIFAK_ALAN")) Then
            MyParsel.GeciciIrtifakAlan = MyRow("GECICI_IRTIFAK_ALAN")
        End If
        If Not IsDBNull(MyRow("MULKIYET_ALAN")) Then
            MyParsel.MulkiyetAlan = MyRow("MULKIYET_ALAN")
        End If
        If Not IsDBNull(MyRow("DAIMI_IRTIFAK_BEDEL")) Then
            MyParsel.IrtifakBedel = MyRow("DAIMI_IRTIFAK_BEDEL")
        End If
        If Not IsDBNull(MyRow("GECICI_IRTIFAK_BEDEL")) Then
            MyParsel.GeciciIrtifakBedel = MyRow("GECICI_IRTIFAK_BEDEL")
        End If
        If Not IsDBNull(MyRow("MULKIYET_BEDEL")) Then
            MyParsel.MulkiyetBedel = MyRow("MULKIYET_BEDEL")
        End If

        Return MyParsel
    End Function

    Private Shared Function MalikOlustur(MyRow As DataRow) As Kisi
        Dim MyMalik As Kisi

        If Not IsDBNull(MyRow("PARSEL_MALIK_TIPI")) Then
            If MyRow("PARSEL_MALIK_TIPI") = 1 Then
                MyMalik = New Kisi(MyRow("MALIK").ToString.Trim)
            Else
                MyMalik = New Kisi(MyRow("MALIK").ToString.Trim, String.Empty)
            End If
        Else
            MyMalik = New Kisi(MyRow("MALIK").ToString.Trim, String.Empty)
        End If

        MyMalik.GUID = Guid.NewGuid().ToString("N")
        MyMalik.Baba = MyRow("BABA").ToString.Trim
        MyMalik.Dusunceler = MyRow("DUSUNCELER").ToString.Trim
        If Not IsDBNull(MyRow("TC_KIMLIK_NO")) Then
            MyMalik.TCKimlikNo = MyRow("TC_KIMLIK_NO")
        End If
        If Not IsDBNull(MyRow("HISSE")) Then
            If MyRow("HISSE").ToString().Contains("TAM") Then
                MyMalik.HissePay = 1
                MyMalik.HissePayda = 1
            ElseIf MyRow("HISSE").ToString().Contains("VRS") Then
                MyMalik.HissePay = 0
                MyMalik.HissePayda = 1
            Else
                If MyRow("HISSE").ToString().Contains("/") Then
                    Dim RSFRSplit As String() = MyRow("HISSE").ToString().Trim.Split("/")
                    MyMalik.HissePay = Val(RSFRSplit(0))
                    MyMalik.HissePayda = Val(RSFRSplit(1))
                Else
                    MyMalik.HissePay = 0
                    MyMalik.HissePayda = 1
                End If
            End If
        Else
            MyMalik.HissePay = 0
            MyMalik.HissePayda = 1
        End If
        If Not IsDBNull(MyRow("TAPUTARIH")) Then
            MyMalik.TapuTarihi = MyRow("TAPUTARIH")
        End If

        Return MyMalik
    End Function

    Private Shared Function ParselKodla(MyParsel As Parsel, MyRow As DataRow) As Parsel
        Dim MyParselKod As New ParselKod With {
            .Kod = MyRow("KOD").ToString.Trim
        }
        If Not IsDBNull(MyRow("KADASTRAL_DURUM")) Then
            Select Case MyRow("KADASTRAL_DURUM")
                Case 1
                    MyParselKod.KadastralDurum = 1
                Case 2
                    MyParselKod.KadastralDurum = 3
                Case 3
                    MyParselKod.IstimlakDisi = True
                Case 4
                    MyParselKod.KadastralDurum = 1
                Case 5
                    MyParselKod.KadastralDurum = 3
                Case 6
                    MyParselKod.KadastralDurum = 4
            End Select
        End If
        If Not IsDBNull(MyRow("PARSEL_MALIK_TIPI")) Then
            MyParselKod.MalikTipi = MyRow("PARSEL_MALIK_TIPI")
        End If
        If Not IsDBNull(MyRow("ISTIMLAK_TURU")) Then
            MyParselKod.IstimlakTuru = MyRow("ISTIMLAK_TURU")
        End If
        If Not IsDBNull(MyRow("ISTIMLAK_SERHI")) Then
            MyParselKod.IstimlakSerhi = MyRow("ISTIMLAK_SERHI")
        End If
        If Not IsDBNull(MyRow("DAVA_DOSYASI_DURUMU")) Then
            MyParselKod.DavaDurumu10 = MyRow("DAVA_DOSYASI_DURUMU")
        End If
        If Not IsDBNull(MyRow("DAVA_DOSYASI_DURUMU_27")) Then
            MyParselKod.DavaDurumu27 = MyRow("DAVA_DOSYASI_DURUMU_27")
        End If
        If Not IsDBNull(MyRow("DEVIR_DURUMU")) Then
            MyParselKod.DevirDurumu = MyRow("DEVIR_DURUMU")
        End If
        If Not IsDBNull(MyRow("PARSEL_ALINMA_DURUMU")) Then
            MyParselKod.EdinimDurumu = MyRow("PARSEL_ALINMA_DURUMU")
        End If
        MyParsel.Kod = MyParselKod
        Return MyParsel
    End Function

    Private Shared Function MalikKodla(MyMalik As Kisi, MyRow As DataRow) As Kisi
        Dim MyMalikKod As New KisiKod
        If Not IsDBNull(MyRow("DAVETIYE_TEBLIG_DURUMU")) Then
            MyMalikKod.DavetiyeTebligDurumu = MyRow("DAVETIYE_TEBLIG_DURUMU")
        End If
        If Not IsDBNull(MyRow("DAVETIYE_ALINMA_DURUMU")) Then
            MyMalikKod.DavetiyeAlinmaDurumu = MyRow("DAVETIYE_ALINMA_DURUMU")
        End If
        If Not IsDBNull(MyRow("GORUSME_DURUMU")) Then
            MyMalikKod.GorusmeDurumu = MyRow("GORUSME_DURUMU")
        End If
        If Not IsDBNull(MyRow("GORUSME_NO")) Then
            MyMalikKod.GorusmeNo = MyRow("GORUSME_NO")
        End If
        If Not IsDBNull(MyRow("GORUSME_TARIHI")) Then
            MyMalikKod.GorusmeTarihi = MyRow("GORUSME_TARIHI")
        End If
        If Not IsDBNull(MyRow("ANLASMA_DURUMU")) Then
            MyMalikKod.AnlasmaDurumu = MyRow("ANLASMA_DURUMU")
        End If
        If Not IsDBNull(MyRow("ANLASMA_TARIHI")) Then
            MyMalikKod.AnlasmaTarihi = MyRow("ANLASMA_TARIHI")
        End If
        If Not IsDBNull(MyRow("ANLASMA_DUSUNCELER")) Then
            MyMalikKod.AnlasmaDusunceler = MyRow("ANLASMA_DUSUNCELER")
        End If
        If Not IsDBNull(MyRow("TESCIL_DURUMU")) Then
            MyMalikKod.TescilDurumu = MyRow("TESCIL_DURUMU")
        End If
        MyMalik.Kod = MyMalikKod
        Return MyMalik
    End Function

    Public Function GetMustemilatCollectionV4(_MustemilatTable As DataTable) As Collection
        Dim MyMustemilatlar As New Collection
        Dim MyMalik As New Kisi
        For Each MyRow As DataRow In _MustemilatTable.Rows
            Dim MyParsel As New Parsel With {
                .Il = MyRow("IL").ToString,
                .Ilce = MyRow("ILCE").ToString,
                .Koy = MyRow("KOY").ToString,
                .Mahalle = MyRow("MAHALLE").ToString,
                .AdaNo = MyRow("ADA").ToString,
                .ParselNo = MyRow("PARSEL").ToString
            }

            Dim MyMustemilat As New Mustemilat With {
                .ParselGUID = GetParselGUID(MyParsel)
            }

            Dim MyKisi As New Kisi(MyRow("SAHIP").ToString, MyRow("BABA").ToString, 0)
            MyMustemilat.SahipGUID = GetKisiGUID(MyKisi)

            If Not IsDBNull(MyRow("BIRIM")) Then
                MyMustemilat.Adet = MyRow("BIRIM")
            End If
            MyMustemilat.Tanim = MyRow("TANIM").ToString
            If Not IsDBNull(MyRow("FIYAT")) Then
                MyMustemilat.Fiyat = MyRow("FIYAT")
            End If
            Select Case MyRow("K_M").ToString.Trim
                Case "K"
                    MyMustemilat.Malik = False
                Case Else
                    MyMustemilat.Malik = True
            End Select
            MyMustemilat.Pay = 1
            MyMustemilat.Payda = 1
            MyMustemilat.OdemeGUID = 0

            MyMustemilatlar.Add(MyMustemilat)
        Next

        Return MyMustemilatlar
    End Function

    Public Function GetMevsimlikCollectionV4(_DataTable As DataTable) As Collection
        Dim MyMevsimlikler As New Collection
        Dim MyMalik As New Kisi
        Dim LastAda As Long = -1
        Dim LastParsel As Long = -1
        For Each MyRow As DataRow In _DataTable.Rows
            Dim MyParsel As New Parsel With {
                .Il = MyRow("IL").ToString,
                .Ilce = MyRow("ILCE").ToString,
                .Koy = MyRow("KOY").ToString,
                .Mahalle = MyRow("MAHALLE").ToString,
                .AdaNo = MyRow("ADA").ToString,
                .ParselNo = MyRow("PARSEL").ToString
            }

            Dim MyKisi As New Kisi(MyRow("SAHIP").ToString, MyRow("BABA").ToString, 0)

            Dim MyMevsimlik As New Mevsimlik
            '.GUID = Guid.NewGuid().ToString("N"),
            '    .ParselGUID = GetParselID(MyParsel),
            '    .SahipGUID = GetKisiID(MyKisi)

            'Kisi ID için parsel bilgiside dikkate alınmalıdır. bu haliyle yanlış maliklere id bağlanacaktır.
            MyKisi = Nothing
            MyParsel = Nothing

            If Not IsDBNull(MyRow("HASAR_ALAN")) Then
                MyMevsimlik.Alan = MyRow("HASAR_ALAN")
            End If
            MyMevsimlik.Tanim = MyRow("TANIM").ToString
            If Not IsDBNull(MyRow("HASAR_BEDEL")) Then
                MyMevsimlik.Bedel = MyRow("HASAR_BEDEL")
            End If
            Select Case MyRow("MK").ToString.Trim
                Case "K"
                    MyMevsimlik.Malik = False
                Case Else
                    MyMevsimlik.Malik = True
            End Select
            If Not IsDBNull(MyRow("HISSE")) Then
                If MyRow("HISSE").ToString().Contains("TAM") Then
                    MyMevsimlik.Pay = 1
                    MyMevsimlik.Payda = 1
                ElseIf MyRow("HISSE").ToString().Contains("VRS") Then
                    MyMevsimlik.Pay = 0
                    MyMevsimlik.Payda = 1
                Else
                    If MyRow("HISSE").ToString().Contains("/") Then
                        Dim RSFRSplit As String() = MyRow("HISSE").ToString().Trim.Split("/")
                        MyMevsimlik.Pay = Val(RSFRSplit(0))
                        MyMevsimlik.Payda = Val(RSFRSplit(1))
                    Else
                        MyMevsimlik.Pay = 0
                        MyMevsimlik.Payda = 1
                    End If
                End If
            Else
                MyMevsimlik.Pay = 0
                MyMevsimlik.Payda = 1
            End If

            MyMevsimlik.OdemeGUID = 0

            MyMevsimlikler.Add(MyMevsimlik)
        Next

        Return MyMevsimlikler
    End Function

#End Region

End Class
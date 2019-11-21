﻿Imports ADOX
Imports Kamu.Objects
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class Create

    Public Function ConvertKamu5ToKamu6(_ConnectionInfo5 As ConnectionInfo, _ConnectionInfo6 As ConnectionInfo) As Boolean
        Dim MyStatus As Boolean = False
        Dim MyKamuDB5 As New Kamu.DB(_ConnectionInfo5)
        Dim MyKamuDB6 As New Kamu.DB(_ConnectionInfo6)

        MyKamuDB5.MyOle.MyConnectionInfo = _ConnectionInfo5
        MyKamuDB6.MyOle.MyConnectionInfo = _ConnectionInfo6

        Using connection As New OleDbConnection(_ConnectionInfo6.ConnectionString)
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Try
                Dim MyCommand As New OleDb.OleDbCommand("ALTER TABLE [PARSEL] ADD ESKI_ID Long", connection)
                MyCommand.ExecuteNonQuery()
                MyCommand = Nothing

                Dim MyCommand1 As New OleDb.OleDbCommand("ALTER TABLE [KISI] ADD ESKI_ID Long", connection)
                MyCommand1.ExecuteNonQuery()
                MyCommand1 = Nothing
            Catch ex As Exception

            End Try
        End Using

        Dim MyParsellerTable As Data.DataTable = MyKamuDB5.GetDataTable("SELECT * FROM PARSELLER ORDER BY IL, ILCE, KOY, MAHALLE, ADA, PARSEL;")
        Dim MyParsellerCollection As Collection = MyKamuDB5.GetParsellerCollectionV5(MyParsellerTable)

        Dim StatusParsel As Boolean = UpdateData(MyParsellerCollection, _ConnectionInfo6)

        Dim MyMustemilatTable As Data.DataTable = MyKamuDB5.GetDataTable("SELECT * FROM MUSTEMILAT ORDER BY IL, ILCE, KOY, MAHALLE, ADA, PARSEL;")
        Dim MyMustemilatCollection As Collection = MyKamuDB5.GetMustemilatCollectionV5(MyMustemilatTable)

        Dim MyMevsimlikTable As Data.DataTable = MyKamuDB5.GetDataTable("SELECT * FROM MEVSIMLIK ORDER BY IL, ILCE, KOY, MAHALLE, ADA, PARSEL;")
        Dim MyMevsimlikCollection As Collection = MyKamuDB5.GetMevsimlikCollectionV5(MyMevsimlikTable)

        Dim StatusMustemilat As Boolean = UpdateData(MyMustemilatCollection, MyMevsimlikCollection, _ConnectionInfo6)

        If StatusParsel And StatusMustemilat Then
            MyStatus = True
            Using connection As New OleDbConnection(_ConnectionInfo6.ConnectionString)
                If Not connection.State = ConnectionState.Open Then connection.Open()
                Try
                    Dim MyCommand As New OleDb.OleDbCommand("ALTER TABLE [PARSEL] DROP ESKI_ID", connection)
                    MyCommand.ExecuteNonQuery()
                    MyCommand = Nothing

                    Dim MyCommand1 As New OleDb.OleDbCommand("ALTER TABLE [KISI] DROP ESKI_ID", connection)
                    MyCommand1.ExecuteNonQuery()
                    MyCommand1 = Nothing
                Catch ex As Exception

                End Try
            End Using
        End If

        MyParsellerCollection = Nothing
        MyMustemilatCollection = Nothing
        MyMevsimlikCollection = Nothing
        MyParsellerTable = Nothing
        MyMustemilatTable = Nothing
        MyMevsimlikTable = Nothing
        MyKamuDB5 = Nothing
        MyKamuDB6 = Nothing
        Return MyStatus
    End Function

    Public Function ConvertKamu6ToKamu6(_SourceConnectionInfo As ConnectionInfo, _TargetConnectionInfo As ConnectionInfo) As Boolean
        Dim MyStatus As Boolean = False
        Dim MyKamuSourceDB As New Kamu.DB(_SourceConnectionInfo)
        Dim MyKamuTargetDB As New Kamu.DB(_TargetConnectionInfo)

        MyKamuSourceDB.MyOle.MyConnectionInfo = _SourceConnectionInfo
        MyKamuTargetDB.MyOle.MyConnectionInfo = _TargetConnectionInfo

        Using connection As New OleDbConnection(_TargetConnectionInfo.ConnectionString)
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Try
                Dim MyCommand As New OleDb.OleDbCommand("ALTER TABLE [PARSEL] ADD ESKI_ID Long", connection)
                MyCommand.ExecuteNonQuery()
                MyCommand = Nothing

                Dim MyCommand1 As New OleDb.OleDbCommand("ALTER TABLE [KISI] ADD ESKI_ID Long", connection)
                MyCommand1.ExecuteNonQuery()
                MyCommand1 = Nothing
            Catch ex As Exception

            End Try
        End Using

        Dim MyFilter As String = "SELECT PARSEL.ID AS ID, PARSEL.PROJE_ID, PARSEL.KOD, PARSEL.IL, PARSEL.ILCE, PARSEL.KOY, PARSEL.MAHALLE, PARSEL.ADA, PARSEL.PARSEL, PARSEL.PAFTA, PARSEL.MEVKI, PARSEL.CINSI, PARSEL.CILT, PARSEL.SAYFA, PARSEL.TAPU_ALANI, PARSEL.SERH_BEYAN, PARSEL_KOD.ID AS PARSEL_KOD_ID, PARSEL_KOD.PARSEL_ID AS PARSEL_KOD_PARSEL_ID, PARSEL_KOD.BOLGE_ID, PARSEL_KOD.KADASTRAL_DURUM, PARSEL_KOD.MALIK_TIPI, PARSEL_KOD.ISTIMLAK_TURU, PARSEL_KOD.ISTIMLAK_SERHI, PARSEL_KOD.DAVA10_DURUMU, PARSEL_KOD.DAVA27_DURUMU, PARSEL_KOD.EDINIM_DURUMU, PARSEL_KOD.ISTIMLAK_DISI, PARSEL_KOD.DEVIR_DURUMU, PARSEL_KOD.ODEME_DURUMU, KISI.ID AS KISI_ID, KISI.ADI, KISI.SOYADI, KISI.TC_KIMLIK_NO, KISI.BABA, KISI.ADRES, KISI.TELEFON, KISI.CINSIYET, KISI.DURUMU, KISI.DOGUM_YERI, KISI.DOGUM_TARIHI, KISI_KOD.ID AS KISI_KOD_ID, KISI_KOD.KISI_ID, KISI_KOD.DAVETIYE_TEBLIG_DURUMU, KISI_KOD.DAVETIYE_ALINMA_DURUMU, KISI_KOD.GORUSME_DURUMU, KISI_KOD.GORUSME_NO, KISI_KOD.GORUSME_TARIHI, KISI_KOD.ANLASMA_DURUMU, KISI_KOD.ANLASMA_TARIHI, KISI_KOD.ANLASMA_DUSUNCELER, KISI_KOD.TESCIL_DURUMU, KAMULASTIRMA.ID AS KAMULASTIRMA_ID, KAMULASTIRMA.PARSEL_ID AS KAMULASTIRMA_PARSEL_ID, KAMULASTIRMA.MULKIYET_ALAN, KAMULASTIRMA.IRTIFAK_ALAN, KAMULASTIRMA.GECICI_IRTIFAK_ALAN, KAMULASTIRMA.MULKIYET_BEDEL, KAMULASTIRMA.IRTIFAK_BEDEL, KAMULASTIRMA.GECICI_IRTIFAK_BEDEL, KAMULASTIRMA.KAMULASTIRMA_AMACI, KAMULASTIRMA.ARAZI_VASFI, KAMULASTIRMA.YAYGIN_MUNAVEBE_SISTEMI, KAMULASTIRMA.DEGERLEME_RAPORU, KAMULASTIRMA.DEGERLEME_TARIHI, KAMULASTIRMA.YILLIK_ORTALAMA_NET_GELIR, KAMULASTIRMA.KAPITALIZASYON_FAIZI, KAMULASTIRMA.OBJEKTIF_ARTIS, KAMULASTIRMA.ART_KISIM_ARTIS, KAMULASTIRMA.VERIM_KAYBI, KAMULASTIRMA.ODEME_ID, MULKIYET.PAY, MULKIYET.PAYDA, MULKIYET.TAPU_TARIHI, MULKIYET.DUSUNCELER FROM ((KISI LEFT JOIN KISI_KOD ON KISI.[ID] = KISI_KOD.[KISI_ID]) INNER JOIN MULKIYET ON KISI.[ID] = MULKIYET.[KISI_ID]) INNER JOIN (PARSEL_KOD RIGHT JOIN (KAMULASTIRMA RIGHT JOIN PARSEL ON KAMULASTIRMA.[PARSEL_ID] = PARSEL.[ID]) ON PARSEL_KOD.[PARSEL_ID] = PARSEL.[ID]) ON MULKIYET.[PARSEL_ID] = PARSEL.[ID] ORDER BY PARSEL.IL, PARSEL.ILCE, PARSEL.KOY, PARSEL.MAHALLE, PARSEL.ADA, PARSEL.PARSEL;"
        Dim MyParsellerTable As Data.DataTable = MyKamuSourceDB.GetDataTable(MyFilter)

        'Dim MyParsellerTable As Data.DataTable = MyKamuSourceDB.GetDataTable("SELECT PARSEL.ID AS ID, PARSEL.PROJE_ID, PARSEL.KOD, PARSEL.IL, PARSEL.ILCE, PARSEL.KOY, PARSEL.MAHALLE, PARSEL.ADA, PARSEL.PARSEL, PARSEL.PAFTA, PARSEL.MEVKI, PARSEL.CINSI, PARSEL.CILT, PARSEL.SAYFA, PARSEL.TAPU_ALANI, PARSEL_KOD.ID AS PARSEL_KOD_ID, PARSEL_KOD.PARSEL_ID AS PARSEL_KOD_PARSEL_ID, PARSEL_KOD.KADASTRAL_DURUM, PARSEL_KOD.MALIK_TIPI, PARSEL_KOD.ISTIMLAK_TURU, PARSEL_KOD.ISTIMLAK_SERHI, PARSEL_KOD.DAVA10_DURUMU, PARSEL_KOD.DAVA27_DURUMU, PARSEL_KOD.EDINIM_DURUMU, PARSEL_KOD.ISTIMLAK_DISI, PARSEL_KOD.DEVIR_DURUMU, PARSEL_KOD.ODEME_DURUMU, KISI.ID AS KISI_ID, KISI.ADI, KISI.SOYADI, KISI.TC_KIMLIK_NO, KISI.BABA, KISI.ADRES, KISI.TELEFON, KISI.CINSIYET, KISI.DURUMU, KISI.DOGUM_YERI, KISI.DOGUM_TARIHI, KISI_KOD.ID AS KISI_KOD_ID, KISI_KOD.KISI_ID, KISI_KOD.DAVETIYE_TEBLIG_DURUMU, KISI_KOD.DAVETIYE_ALINMA_DURUMU, KISI_KOD.GORUSME_DURUMU, KISI_KOD.GORUSME_NO, KISI_KOD.GORUSME_TARIHI, KISI_KOD.ANLASMA_DURUMU, KISI_KOD.ANLASMA_TARIHI, KISI_KOD.ANLASMA_DUSUNCELER, KISI_KOD.TESCIL_DURUMU, KAMULASTIRMA.ID AS KAMULASTIRMA_ID, KAMULASTIRMA.PARSEL_ID AS KAMULASTIRMA_PARSEL_ID, KAMULASTIRMA.MULKIYET_ALAN, KAMULASTIRMA.IRTIFAK_ALAN, KAMULASTIRMA.GECICI_IRTIFAK_ALAN, KAMULASTIRMA.MULKIYET_BEDEL, KAMULASTIRMA.IRTIFAK_BEDEL, KAMULASTIRMA.GECICI_IRTIFAK_BEDEL, KAMULASTIRMA.KAMULASTIRMA_AMACI, KAMULASTIRMA.ARAZI_VASFI, KAMULASTIRMA.YAYGIN_MUNAVEBE_SISTEMI, KAMULASTIRMA.DEGERLEME_RAPORU, KAMULASTIRMA.DEGERLEME_TARIHI, KAMULASTIRMA.YILLIK_ORTALAMA_NET_GELIR, KAMULASTIRMA.KAPITALIZASYON_FAIZI, KAMULASTIRMA.OBJEKTIF_ARTIS, KAMULASTIRMA.ART_KISIM_ARTIS, KAMULASTIRMA.VERIM_KAYBI, KAMULASTIRMA.ODEME_ID, MULKIYET.PAY, MULKIYET.PAYDA, MULKIYET.TAPU_TARIHI, MULKIYET.DUSUNCELER FROM ((KISI INNER JOIN KISI_KOD ON KISI.[ID] = KISI_KOD.[KISI_ID]) INNER JOIN MULKIYET ON KISI.[ID] = MULKIYET.[KISI_ID]) INNER JOIN (PARSEL_KOD INNER JOIN (KAMULASTIRMA INNER JOIN PARSEL ON KAMULASTIRMA.[PARSEL_ID] = PARSEL.[ID]) ON PARSEL_KOD.[PARSEL_ID] = PARSEL.[ID]) ON MULKIYET.[PARSEL_ID] = PARSEL.[ID] ORDER BY PARSEL.IL, PARSEL.ILCE, PARSEL.KOY, PARSEL.MAHALLE, PARSEL.ADA, PARSEL.PARSEL;")

        Dim MyParsellerCollection As Collection = MyKamuSourceDB.GetParselCollection(MyParsellerTable, False, True)
        MyParsellerTable = Nothing

        MyParsellerCollection = MyKamuSourceDB.DefineVerasetDurumu(MyParsellerCollection)

        Dim StatusParsel As Boolean = UpdateData(MyParsellerCollection, _TargetConnectionInfo)
        MyParsellerCollection = Nothing

        Dim MyParselConversionTable As Data.DataTable = MyKamuTargetDB.GetDataTable("SELECT ID, ESKI_ID FROM PARSEL")
        Dim MyKisiConversionTable As Data.DataTable = MyKamuTargetDB.GetDataTable("SELECT ID, ESKI_ID FROM KISI")

        Dim MyMustemilatTable As Data.DataTable = MyKamuSourceDB.GetDataTable("SELECT * FROM MUSTEMILAT ORDER BY ID;")
        Dim MyMustemilatCollection As Collection = MyKamuSourceDB.GetMustemilatCollection(MyMustemilatTable)
        MyMustemilatTable = Nothing

        Dim MyMevsimlikTable As Data.DataTable = MyKamuSourceDB.GetDataTable("SELECT * FROM MEVSIMLIK ORDER BY ID;")
        Dim MyMevsimlikCollection As Collection = MyKamuSourceDB.GetMevsimlikCollection(MyMevsimlikTable)
        MyMevsimlikTable = Nothing

        MyMustemilatCollection = MyKamuSourceDB.DefineMustemilatOwnerShip(MyMustemilatCollection, MyParselConversionTable, MyKisiConversionTable)
        MyMevsimlikCollection = MyKamuSourceDB.DefineMevsimlikOwnerShip(MyMevsimlikCollection, MyParselConversionTable, MyKisiConversionTable)
        MyParselConversionTable = Nothing
        MyKisiConversionTable = Nothing

        Dim StatusMustemilat As Boolean = UpdateData(MyMustemilatCollection, MyMevsimlikCollection, _TargetConnectionInfo)
        MyMustemilatCollection = Nothing
        MyMevsimlikCollection = Nothing

        MyKamuSourceDB = Nothing
        MyKamuTargetDB = Nothing

        If StatusParsel And StatusMustemilat Then
            MyStatus = True
        End If

        Return MyStatus
    End Function

    Public Function CreateKamuDBFromScratch(_FileName As String, _KamuVeriXMLFileName As String, _TableName As String) As Boolean
        Dim MyStatus As Boolean = False
        Try
            If System.IO.File.Exists(_FileName) = True Then
                System.IO.File.Delete(_FileName)
            End If
            Dim MyCatalog As Catalog = New Catalog()
            MyCatalog.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _FileName + ";Jet OLEDB:Engine Type=5")
            MyCatalog = Nothing

            Dim MyConnection As New OleDb.OleDbConnection("PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source =" & _FileName)
            MyConnection.Open()

            Dim KamuVeri As New DataSet
            KamuVeri.ReadXml(_KamuVeriXMLFileName)
            Dim KamuTable As DataTable = KamuVeri.Tables(_TableName)
            For Each MyRow As DataRow In KamuTable.Rows
                Dim strSQL As String = MyRow("TANIMLAMA")
                CreateAccessTable(MyConnection, strSQL)
            Next
            KamuTable = Nothing
            KamuVeri = Nothing

            MyConnection.Close()
            MyConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MsgBox(ex.Message)
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Private Sub CreateAccessTable(_Connection As OleDb.OleDbConnection, strSQL As String)
        Dim MyCommand1 As New OleDb.OleDbCommand(strSQL, _Connection)
        MyCommand1.ExecuteNonQuery()
        MyCommand1 = Nothing
    End Sub

    Private Function UpdateData(ParselCollection As Collection, Connection As ConnectionInfo) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case Connection.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = UpdateDataOleDb(ParselCollection, Connection.ConnectionString)
                Case Connections.SqlConnection
                    MyObject = UpdateDataSQL(ParselCollection, Connection.ConnectionString)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Private Function UpdateData(MustemilatCollection As Collection, MevsimlikCollection As Collection, Connection As ConnectionInfo) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case Connection.ConnectionType
                Case Connections.SqlConnection
                    MyObject = UpdateDataOleDb(MustemilatCollection, MevsimlikCollection, Connection.ConnectionString)
                Case Connections.SqlConnection
                    MyObject = UpdateDataSQL(MustemilatCollection, MevsimlikCollection, Connection.ConnectionString)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Private Function UpdateDataOleDb(ParselCollection As Collection, _ConnectionString As String) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDb.OleDbConnection(_ConnectionString)
            If Not MyKamuConnection.State = ConnectionState.Open Then MyKamuConnection.Open()

            Dim MyQueryStringParsel As String = "SELECT * FROM PARSEL"
            Dim MyParselDataAdapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(New OleDb.OleDbCommand(MyQueryStringParsel, MyKamuConnection))
            Dim MyParselTable As New DataTable
            MyParselDataAdapter.Fill(MyParselTable)

            Dim MyQueryStringParselKod As String = "SELECT * FROM PARSEL_KOD"
            Dim MyParselKodDataAdapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(New OleDb.OleDbCommand(MyQueryStringParselKod, MyKamuConnection))
            Dim MyParselKodTable As New DataTable
            MyParselKodDataAdapter.Fill(MyParselKodTable)

            Dim MyQueryStringKamu As String = "SELECT * FROM KAMULASTIRMA"
            Dim MyKamuDataAdapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(New OleDb.OleDbCommand(MyQueryStringKamu, MyKamuConnection))
            Dim MyKamuTable As New DataTable
            MyKamuDataAdapter.Fill(MyKamuTable)

            Dim MyQueryStringKisi As String = "SELECT * FROM KISI"
            Dim MyKisiDataAdapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(New OleDb.OleDbCommand(MyQueryStringKisi, MyKamuConnection))
            Dim MyKisiTable As New DataTable
            MyKisiDataAdapter.Fill(MyKisiTable)

            Dim MyQueryStringKisiKod As String = "SELECT * FROM KISI_KOD"
            Dim MyKisiKodDataAdapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(New OleDb.OleDbCommand(MyQueryStringKisiKod, MyKamuConnection))
            Dim MyKisiKodTable As New DataTable
            MyKisiKodDataAdapter.Fill(MyKisiKodTable)

            Dim MyQueryStringMulkiyet As String = "SELECT * FROM MULKIYET"
            Dim MyMulkiyetDataAdapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(New OleDb.OleDbCommand(MyQueryStringMulkiyet, MyKamuConnection))
            Dim MyMulkiyetTable As New DataTable
            MyMulkiyetDataAdapter.Fill(MyMulkiyetTable)

            Dim MyQueryStringMiras As String = "SELECT * FROM MIRAS"
            Dim MyMirasDataAdapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(New OleDb.OleDbCommand(MyQueryStringMiras, MyKamuConnection))
            Dim MyMirasTable As New DataTable
            MyMirasDataAdapter.Fill(MyMirasTable)

            Try
                For Each _Parsel As Parsel In ParselCollection
                    Dim MyParselRow As DataRow = MyParselTable.NewRow()
                    MyParselRow("ESKI_ID") = _Parsel.ID
                    MyParselRow("PROJE_ID") = _Parsel.ProjeID
                    MyParselRow("KOD") = _Parsel.Kod.Kod
                    MyParselRow("IL") = _Parsel.Il
                    MyParselRow("ILCE") = _Parsel.Ilce
                    MyParselRow("KOY") = _Parsel.Koy
                    MyParselRow("MAHALLE") = _Parsel.Mahalle
                    MyParselRow("ADA") = _Parsel.AdaNo
                    MyParselRow("PARSEL") = _Parsel.ParselNo
                    MyParselRow("PAFTA") = _Parsel.PaftaNo
                    MyParselRow("MEVKI") = _Parsel.Mevki
                    MyParselRow("CILT") = _Parsel.Cilt
                    MyParselRow("SAYFA") = _Parsel.Sayfa
                    MyParselRow("CINSI") = _Parsel.Cinsi
                    MyParselRow("TAPU_ALANI") = _Parsel.TapuAlani
                    MyParselTable.Rows.Add(MyParselRow)

                    Dim MyParselInfo As System.Reflection.FieldInfo = MyParselRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
                    Dim NewParselID As Integer = CInt(MyParselInfo.GetValue(MyParselRow))

                    Dim MyParselKodRow As DataRow = MyParselKodTable.NewRow()
                    MyParselKodRow("PARSEL_ID") = NewParselID
                    MyParselKodRow("BOLGE_ID") = _Parsel.Kod.BolgeID
                    MyParselKodRow("KADASTRAL_DURUM") = _Parsel.Kod.KadastralDurum
                    MyParselKodRow("MALIK_TIPI") = _Parsel.Kod.MalikTipi
                    MyParselKodRow("ISTIMLAK_TURU") = _Parsel.Kod.IstimlakTuru
                    MyParselKodRow("ISTIMLAK_SERHI") = _Parsel.Kod.IstimlakSerhi
                    MyParselKodRow("DAVA10_DURUMU") = _Parsel.Kod.DavaDurumu10
                    MyParselKodRow("DAVA27_DURUMU") = _Parsel.Kod.DavaDurumu27
                    MyParselKodRow("EDINIM_DURUMU") = _Parsel.Kod.EdinimDurumu
                    MyParselKodRow("ISTIMLAK_DISI") = _Parsel.Kod.IstimlakDisi
                    MyParselKodRow("DEVIR_DURUMU") = _Parsel.Kod.DevirDurumu
                    MyParselKodRow("ODEME_DURUMU") = _Parsel.Kod.OdemeDurumu
                    MyParselKodTable.Rows.Add(MyParselKodRow)

                    Dim MyKamuRow As DataRow = MyKamuTable.NewRow()
                    MyKamuRow("PARSEL_ID") = NewParselID
                    MyKamuRow("MULKIYET_ALAN") = _Parsel.MulkiyetAlan
                    MyKamuRow("IRTIFAK_ALAN") = _Parsel.IrtifakAlan
                    MyKamuRow("GECICI_IRTIFAK_ALAN") = _Parsel.GeciciIrtifakAlan
                    MyKamuRow("MULKIYET_BEDEL") = _Parsel.MulkiyetBedel
                    MyKamuRow("IRTIFAK_BEDEL") = _Parsel.IrtifakBedel
                    MyKamuRow("GECICI_IRTIFAK_BEDEL") = _Parsel.GeciciIrtifakBedel
                    MyKamuRow("KAMULASTIRMA_AMACI") = _Parsel.KamulastirmaAmaci
                    MyKamuRow("ARAZI_VASFI") = _Parsel.AraziVasfi
                    MyKamuRow("YAYGIN_MUNAVEBE_SISTEMI") = _Parsel.YayginMunavebeSistemi
                    MyKamuRow("DEGERLEME_RAPORU") = _Parsel.DegerlemeRaporu
                    MyKamuRow("DEGERLEME_TARIHI") = _Parsel.DegerlemeTarihi
                    MyKamuRow("YILLIK_ORTALAMA_NET_GELIR") = _Parsel.YillikOrtalamaNetGelir
                    MyKamuRow("KAPITALIZASYON_FAIZI") = _Parsel.KapitalizasyonOrani
                    MyKamuRow("OBJEKTIF_ARTIS") = _Parsel.ObjektifArtis
                    MyKamuRow("ART_KISIM_ARTIS") = _Parsel.ArtanKisimArtis
                    MyKamuRow("VERIM_KAYBI") = _Parsel.VerimKaybi
                    MyKamuTable.Rows.Add(MyKamuRow)

                    For Each _Kisi As Kisi In _Parsel.Malikler
                        Dim KisiID As Integer = GetMalikID(_Kisi, MyKisiTable)
                        If KisiID = 0 Then
                            Dim MyKisiRow As DataRow = MyKisiTable.NewRow()
                            MyKisiRow("ESKI_ID") = _Kisi.ID
                            MyKisiRow("ADI") = _Kisi.Adi
                            MyKisiRow("SOYADI") = _Kisi.Soyadi
                            MyKisiRow("TC_KIMLIK_NO") = _Kisi.TCKimlikNo
                            MyKisiRow("BABA") = _Kisi.Baba
                            MyKisiRow("ADRES") = _Kisi.Adres
                            MyKisiRow("TELEFON") = _Kisi.Telefon
                            MyKisiRow("DURUMU") = _Kisi.Durumu
                            MyKisiRow("CINSIYET") = _Kisi.Cinsiyet
                            MyKisiTable.Rows.Add(MyKisiRow)

                            Dim MyKisiInfo As System.Reflection.FieldInfo = MyKisiRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
                            Dim NewKisiID As Integer = CInt(MyKisiInfo.GetValue(MyKisiRow))
                            KisiID = NewKisiID

                            Dim MyKisiKodRow As DataRow = MyKisiKodTable.NewRow()
                            MyKisiKodRow("KISI_ID") = KisiID
                            MyKisiKodRow("DAVETIYE_TEBLIG_DURUMU") = _Kisi.Kod.DavetiyeTebligDurumu
                            MyKisiKodRow("DAVETIYE_ALINMA_DURUMU") = _Kisi.Kod.DavetiyeAlinmaDurumu
                            MyKisiKodRow("GORUSME_DURUMU") = _Kisi.Kod.GorusmeDurumu
                            MyKisiKodRow("GORUSME_NO") = _Kisi.Kod.GorusmeNo
                            MyKisiKodRow("GORUSME_TARIHI") = _Kisi.Kod.GorusmeTarihi
                            MyKisiKodRow("ANLASMA_DURUMU") = _Kisi.Kod.AnlasmaDurumu
                            MyKisiKodRow("ANLASMA_TARIHI") = _Kisi.Kod.AnlasmaTarihi
                            MyKisiKodRow("ANLASMA_DUSUNCELER") = _Kisi.Kod.AnlasmaDusunceler
                            MyKisiKodRow("TESCIL_DURUMU") = _Kisi.Kod.TescilDurumu
                            MyKisiKodTable.Rows.Add(MyKisiKodRow)

                            If Not IsNothing(_Kisi.Varisler) Then


                                For Each Varis As Kisi In _Kisi.Varisler
                                    Dim VarisID As Integer = GetMalikID(Varis, MyKisiTable)
                                    If VarisID = 0 Then
                                        Dim MyVarisKisiRow As DataRow = MyKisiTable.NewRow()
                                        MyVarisKisiRow("ADI") = Varis.Adi
                                        MyVarisKisiRow("SOYADI") = Varis.Soyadi
                                        MyVarisKisiRow("TC_KIMLIK_NO") = Varis.TCKimlikNo
                                        MyVarisKisiRow("BABA") = Varis.Baba
                                        MyVarisKisiRow("ADRES") = Varis.Adres
                                        MyVarisKisiRow("TELEFON") = Varis.Telefon
                                        MyVarisKisiRow("DURUMU") = Varis.Durumu
                                        MyKisiTable.Rows.Add(MyVarisKisiRow)

                                        Dim MyVarisInfo As System.Reflection.FieldInfo = MyVarisKisiRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
                                        Dim NewVarisID As Integer = CInt(MyVarisInfo.GetValue(MyVarisKisiRow))
                                        VarisID = NewVarisID

                                        Dim MyVarisKisiKodRow As DataRow = MyKisiKodTable.NewRow()
                                        MyVarisKisiKodRow("KISI_ID") = VarisID
                                        MyVarisKisiKodRow("DAVETIYE_TEBLIG_DURUMU") = Varis.Kod.DavetiyeTebligDurumu
                                        MyVarisKisiKodRow("DAVETIYE_ALINMA_DURUMU") = Varis.Kod.DavetiyeAlinmaDurumu
                                        MyVarisKisiKodRow("GORUSME_DURUMU") = Varis.Kod.GorusmeDurumu
                                        MyVarisKisiKodRow("GORUSME_NO") = Varis.Kod.GorusmeNo
                                        MyVarisKisiKodRow("GORUSME_TARIHI") = Varis.Kod.GorusmeTarihi
                                        MyVarisKisiKodRow("ANLASMA_DURUMU") = Varis.Kod.AnlasmaDurumu
                                        MyVarisKisiKodRow("ANLASMA_TARIHI") = Varis.Kod.AnlasmaTarihi
                                        MyVarisKisiKodRow("ANLASMA_DUSUNCELER") = Varis.Kod.AnlasmaDusunceler
                                        MyVarisKisiKodRow("TESCIL_DURUMU") = Varis.Kod.TescilDurumu
                                        MyKisiKodTable.Rows.Add(MyVarisKisiKodRow)

                                    End If

                                    Dim MyVarisRow As DataRow = MyMirasTable.NewRow()
                                    MyVarisRow("MURIS") = KisiID
                                    MyVarisRow("VARIS") = VarisID
                                    MyMirasTable.Rows.Add(MyVarisRow)
                                Next
                            End If

                        End If

                        Dim MyMulkiyetRow As DataRow = MyMulkiyetTable.NewRow()
                        MyMulkiyetRow("PARSEL_ID") = NewParselID
                        MyMulkiyetRow("KISI_ID") = KisiID
                        MyMulkiyetRow("PAY") = _Kisi.HissePay
                        MyMulkiyetRow("PAYDA") = _Kisi.HissePayda
                        If _Kisi.TapuTarihi.Year > 1000 And _Kisi.TapuTarihi.Year < 2050 Then
                            MyMulkiyetRow("TAPU_TARIHI") = _Kisi.TapuTarihi.ToShortDateString
                        End If
                        MyMulkiyetRow("DUSUNCELER") = _Kisi.Dusunceler
                        MyMulkiyetTable.Rows.Add(MyMulkiyetRow)
                    Next
                Next
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyParselCommandBuilder As New OleDb.OleDbCommandBuilder
                MyParselCommandBuilder.DataAdapter = MyParselDataAdapter
                MyParselDataAdapter.Update(MyParselTable)
                MyParselTable = Nothing
                MyParselCommandBuilder = Nothing
                MyParselDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyParselKodCommandBuilder As New OleDb.OleDbCommandBuilder
                MyParselKodCommandBuilder.DataAdapter = MyParselKodDataAdapter
                MyParselKodDataAdapter.Update(MyParselKodTable)
                MyParselKodTable = Nothing
                MyParselKodCommandBuilder = Nothing
                MyParselKodDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyKamuCommandBuilder As New OleDb.OleDbCommandBuilder
                MyKamuCommandBuilder.DataAdapter = MyKamuDataAdapter
                MyKamuDataAdapter.Update(MyKamuTable)
                MyKamuTable = Nothing
                MyKamuCommandBuilder = Nothing
                MyKamuDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyKisiCommandBuilder As New OleDb.OleDbCommandBuilder
                MyKisiCommandBuilder.DataAdapter = MyKisiDataAdapter
                MyKisiDataAdapter.Update(MyKisiTable)
                MyKisiTable = Nothing
                MyKisiCommandBuilder = Nothing
                MyKisiDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyKisiKodCommandBuilder As New OleDb.OleDbCommandBuilder
                MyKisiKodCommandBuilder.DataAdapter = MyKisiKodDataAdapter
                MyKisiKodDataAdapter.Update(MyKisiKodTable)
                MyKisiKodTable = Nothing
                MyKisiKodCommandBuilder = Nothing
                MyKisiKodDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyMulkiyetCommandBuilder As New OleDb.OleDbCommandBuilder
                MyMulkiyetCommandBuilder.DataAdapter = MyMulkiyetDataAdapter
                MyMulkiyetDataAdapter.Update(MyMulkiyetTable)
                MyMulkiyetTable = Nothing
                MyMulkiyetCommandBuilder = Nothing
                MyMulkiyetDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyMirasCommandBuilder As New OleDb.OleDbCommandBuilder
                MyMirasCommandBuilder.DataAdapter = MyMirasDataAdapter
                MyMirasDataAdapter.Update(MyMirasTable)
                MyMirasTable = Nothing
                MyMirasCommandBuilder = Nothing
                MyMirasDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return MyStatus
    End Function

    Private Function UpdateDataSQL(ParselCollection As Collection, _ConnectionString As String) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New SqlConnection(_ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryStringParsel As String = "SELECT * FROM PARSEL"
            Dim MyParselDataAdapter As SqlDataAdapter = New SqlDataAdapter(New SqlCommand(MyQueryStringParsel, MyKamuConnection))
            Dim MyParselTable As New DataTable
            MyParselDataAdapter.Fill(MyParselTable)

            Dim MyQueryStringParselKod As String = "SELECT * FROM PARSEL_KOD"
            Dim MyParselKodDataAdapter As SqlDataAdapter = New SqlDataAdapter(New SqlCommand(MyQueryStringParselKod, MyKamuConnection))
            Dim MyParselKodTable As New DataTable
            MyParselKodDataAdapter.Fill(MyParselKodTable)

            Dim MyQueryStringKamu As String = "SELECT * FROM KAMULASTIRMA"
            Dim MyKamuDataAdapter As SqlDataAdapter = New SqlDataAdapter(New SqlCommand(MyQueryStringKamu, MyKamuConnection))
            Dim MyKamuTable As New DataTable
            MyKamuDataAdapter.Fill(MyKamuTable)

            Dim MyQueryStringKisi As String = "SELECT * FROM KISI"
            Dim MyKisiDataAdapter As SqlDataAdapter = New SqlDataAdapter(New SqlCommand(MyQueryStringKisi, MyKamuConnection))
            Dim MyKisiTable As New DataTable
            MyKisiDataAdapter.Fill(MyKisiTable)

            Dim MyQueryStringKisiKod As String = "SELECT * FROM KISI_KOD"
            Dim MyKisiKodDataAdapter As SqlDataAdapter = New SqlDataAdapter(New SqlCommand(MyQueryStringKisiKod, MyKamuConnection))
            Dim MyKisiKodTable As New DataTable
            MyKisiKodDataAdapter.Fill(MyKisiKodTable)

            Dim MyQueryStringMulkiyet As String = "SELECT * FROM MULKIYET"
            Dim MyMulkiyetDataAdapter As SqlDataAdapter = New SqlDataAdapter(New SqlCommand(MyQueryStringMulkiyet, MyKamuConnection))
            Dim MyMulkiyetTable As New DataTable
            MyMulkiyetDataAdapter.Fill(MyMulkiyetTable)

            Try
                For Each _Parsel As Parsel In ParselCollection
                    Dim MyParselRow As DataRow = MyParselTable.NewRow()
                    MyParselRow("PROJE_ID") = _Parsel.ProjeID
                    MyParselRow("KOD") = _Parsel.Kod.Kod
                    MyParselRow("IL") = _Parsel.Il
                    MyParselRow("ILCE") = _Parsel.Ilce
                    MyParselRow("KOY") = _Parsel.Koy
                    MyParselRow("MAHALLE") = _Parsel.Mahalle
                    MyParselRow("ADA") = _Parsel.AdaNo
                    MyParselRow("PARSEL") = _Parsel.ParselNo
                    MyParselRow("PAFTA") = _Parsel.PaftaNo
                    MyParselRow("MEVKI") = _Parsel.Mevki
                    MyParselRow("CILT") = _Parsel.Cilt
                    MyParselRow("SAYFA") = _Parsel.Sayfa
                    MyParselRow("CINSI") = _Parsel.Cinsi
                    MyParselRow("TAPU_ALANI") = _Parsel.TapuAlani
                    MyParselTable.Rows.Add(MyParselRow)

                    Dim MyParselInfo As System.Reflection.FieldInfo = MyParselRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
                    Dim NewParselID As Integer = CInt(MyParselInfo.GetValue(MyParselRow))

                    Dim MyParselKodRow As DataRow = MyParselKodTable.NewRow()
                    MyParselKodRow("PARSEL_ID") = NewParselID
                    MyParselKodRow("BOLGE_ID") = _Parsel.Kod.BolgeID
                    MyParselKodRow("KADASTRAL_DURUM") = _Parsel.Kod.KadastralDurum
                    MyParselKodRow("MALIK_TIPI") = _Parsel.Kod.MalikTipi
                    MyParselKodRow("ISTIMLAK_TURU") = _Parsel.Kod.IstimlakTuru
                    MyParselKodRow("ISTIMLAK_SERHI") = _Parsel.Kod.IstimlakSerhi
                    MyParselKodRow("DAVA10_DURUMU") = _Parsel.Kod.DavaDurumu10
                    MyParselKodRow("DAVA27_DURUMU") = _Parsel.Kod.DavaDurumu27
                    MyParselKodRow("EDINIM_DURUMU") = _Parsel.Kod.EdinimDurumu
                    MyParselKodRow("ISTIMLAK_DISI") = _Parsel.Kod.IstimlakDisi
                    MyParselKodRow("DEVIR_DURUMU") = _Parsel.Kod.DevirDurumu
                    MyParselKodRow("ODEME_DURUMU") = _Parsel.Kod.OdemeDurumu
                    MyParselKodTable.Rows.Add(MyParselKodRow)

                    Dim MyKamuRow As DataRow = MyKamuTable.NewRow()
                    MyKamuRow("PARSEL_ID") = NewParselID
                    MyKamuRow("MULKIYET_ALAN") = _Parsel.MulkiyetAlan
                    MyKamuRow("IRTIFAK_ALAN") = _Parsel.IrtifakAlan
                    MyKamuRow("GECICI_IRTIFAK_ALAN") = _Parsel.GeciciIrtifakAlan
                    MyKamuRow("MULKIYET_BEDEL") = _Parsel.MulkiyetBedel
                    MyKamuRow("IRTIFAK_BEDEL") = _Parsel.IrtifakBedel
                    MyKamuRow("GECICI_IRTIFAK_BEDEL") = _Parsel.GeciciIrtifakBedel
                    MyKamuTable.Rows.Add(MyKamuRow)

                    For Each _Kisi As Kisi In _Parsel.Malikler
                        Dim KisiID As Integer = GetMalikID(_Kisi, MyKisiTable)
                        If KisiID = 0 Then
                            Dim MyKisiRow As DataRow = MyKisiTable.NewRow()
                            MyKisiRow("ADI") = _Kisi.Adi
                            MyKisiRow("SOYADI") = _Kisi.Soyadi
                            MyKisiRow("TC_KIMLIK_NO") = _Kisi.TCKimlikNo
                            MyKisiRow("BABA") = _Kisi.Baba
                            MyKisiRow("ADRES") = _Kisi.Adres
                            MyKisiRow("TELEFON") = _Kisi.Telefon
                            MyKisiRow("DURUMU") = _Kisi.Durumu
                            MyKisiTable.Rows.Add(MyKisiRow)

                            Dim MyKisiInfo As System.Reflection.FieldInfo = MyKisiRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
                            Dim NewKisiID As Integer = CInt(MyKisiInfo.GetValue(MyKisiRow))
                            KisiID = NewKisiID

                            Dim MyKisiKodRow As DataRow = MyKisiKodTable.NewRow()
                            MyKisiKodRow("KISI_ID") = KisiID
                            'MyKisiKodRow("MALIK_TIPI") = _Kisi.Kod.MalikTipi
                            MyKisiKodRow("DAVETIYE_TEBLIG_DURUMU") = _Kisi.Kod.DavetiyeTebligDurumu
                            MyKisiKodRow("DAVETIYE_ALINMA_DURUMU") = _Kisi.Kod.DavetiyeAlinmaDurumu
                            MyKisiKodRow("GORUSME_DURUMU") = _Kisi.Kod.GorusmeDurumu
                            MyKisiKodRow("GORUSME_NO") = _Kisi.Kod.GorusmeNo
                            MyKisiKodRow("GORUSME_TARIHI") = _Kisi.Kod.GorusmeTarihi
                            MyKisiKodRow("ANLASMA_DURUMU") = _Kisi.Kod.AnlasmaDurumu
                            MyKisiKodRow("ANLASMA_TARIHI") = _Kisi.Kod.AnlasmaTarihi
                            MyKisiKodRow("ANLASMA_DUSUNCELER") = _Kisi.Kod.AnlasmaDusunceler
                            MyKisiKodRow("TESCIL_DURUMU") = _Kisi.Kod.TescilDurumu
                            MyKisiKodTable.Rows.Add(MyKisiKodRow)
                        End If

                        Dim MyMulkiyetRow As DataRow = MyMulkiyetTable.NewRow()
                        MyMulkiyetRow("PARSEL_ID") = NewParselID
                        MyMulkiyetRow("KISI_ID") = KisiID
                        MyMulkiyetRow("PAY") = _Kisi.HissePay
                        MyMulkiyetRow("PAYDA") = _Kisi.HissePayda
                        If _Kisi.TapuTarihi.Year > 1000 And _Kisi.TapuTarihi.Year < 2050 Then
                            MyMulkiyetRow("TAPU_TARIHI") = _Kisi.TapuTarihi.ToShortDateString
                        End If
                        MyMulkiyetRow("DUSUNCELER") = _Kisi.Dusunceler
                        MyMulkiyetTable.Rows.Add(MyMulkiyetRow)
                    Next
                Next
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyParselCommandBuilder As New SqlCommandBuilder
                MyParselCommandBuilder.DataAdapter = MyParselDataAdapter
                MyParselDataAdapter.Update(MyParselTable)
                MyParselTable = Nothing
                MyParselCommandBuilder = Nothing
                MyParselDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyParselKodCommandBuilder As New SqlCommandBuilder
                MyParselKodCommandBuilder.DataAdapter = MyParselKodDataAdapter
                MyParselKodDataAdapter.Update(MyParselKodTable)
                MyParselKodTable = Nothing
                MyParselKodCommandBuilder = Nothing
                MyParselKodDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyKamuCommandBuilder As New SqlCommandBuilder
                MyKamuCommandBuilder.DataAdapter = MyKamuDataAdapter
                MyKamuDataAdapter.Update(MyKamuTable)
                MyKamuTable = Nothing
                MyKamuCommandBuilder = Nothing
                MyKamuDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyKisiCommandBuilder As New SqlCommandBuilder
                MyKisiCommandBuilder.DataAdapter = MyKisiDataAdapter
                MyKisiDataAdapter.Update(MyKisiTable)
                MyKisiTable = Nothing
                MyKisiCommandBuilder = Nothing
                MyKisiDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyKisiKodCommandBuilder As New SqlCommandBuilder
                MyKisiKodCommandBuilder.DataAdapter = MyKisiKodDataAdapter
                MyKisiKodDataAdapter.Update(MyKisiKodTable)
                MyKisiKodTable = Nothing
                MyKisiKodCommandBuilder = Nothing
                MyKisiKodDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyMulkiyetCommandBuilder As New SqlCommandBuilder
                MyMulkiyetCommandBuilder.DataAdapter = MyMulkiyetDataAdapter
                MyMulkiyetDataAdapter.Update(MyMulkiyetTable)
                MyMulkiyetTable = Nothing
                MyMulkiyetCommandBuilder = Nothing
                MyMulkiyetDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return MyStatus
    End Function

    Private Function UpdateDataOleDb(MustemilatCollection As Collection, MevsimlikCollection As Collection, _ConnectionString As String) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDb.OleDbConnection(_ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryStringMustemilat As String = "SELECT * FROM MUSTEMILAT"
            Dim MyMustemilatDataAdapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(New OleDb.OleDbCommand(MyQueryStringMustemilat, MyKamuConnection))
            Dim MyMustemilatTable As New DataTable
            MyMustemilatDataAdapter.Fill(MyMustemilatTable)

            Dim MyQueryStringMevsimlik As String = "SELECT * FROM MEVSIMLIK"
            Dim MyMevsimlikDataAdapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(New OleDb.OleDbCommand(MyQueryStringMevsimlik, MyKamuConnection))
            Dim MyMevsimlikTable As New DataTable
            MyMevsimlikDataAdapter.Fill(MyMevsimlikTable)

            Try
                For Each _Mustemilat As Mustemilat In MustemilatCollection
                    Dim MyMustemilatRow As DataRow = MyMustemilatTable.NewRow()
                    MyMustemilatRow("PARSEL_ID") = _Mustemilat.ParselID
                    MyMustemilatRow("SAHIP_ID") = _Mustemilat.SahipID
                    MyMustemilatRow("TANIM") = _Mustemilat.Tanim
                    MyMustemilatRow("ADET") = _Mustemilat.Adet
                    MyMustemilatRow("FIYAT") = _Mustemilat.Fiyat
                    MyMustemilatRow("MALIK") = _Mustemilat.Malik
                    MyMustemilatRow("PAY") = _Mustemilat.Pay
                    MyMustemilatRow("PAYDA") = _Mustemilat.Payda
                    MyMustemilatRow("ODEME_ID") = _Mustemilat.OdemeID
                    MyMustemilatTable.Rows.Add(MyMustemilatRow)
                Next
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                For Each _Mevsimlik As Mevsimlik In MevsimlikCollection
                    Dim MyMevsimlikRow As DataRow = MyMevsimlikTable.NewRow()
                    MyMevsimlikRow("PARSEL_ID") = _Mevsimlik.ParselID
                    MyMevsimlikRow("SAHIP_ID") = _Mevsimlik.SahipID
                    MyMevsimlikRow("TANIM") = _Mevsimlik.Tanim
                    MyMevsimlikRow("ALAN") = _Mevsimlik.Alan
                    MyMevsimlikRow("BEDEL") = _Mevsimlik.Bedel
                    MyMevsimlikRow("MALIK") = _Mevsimlik.Malik
                    MyMevsimlikRow("PAY") = _Mevsimlik.Pay
                    MyMevsimlikRow("PAYDA") = _Mevsimlik.Payda
                    MyMevsimlikRow("ODEME_ID") = _Mevsimlik.OdemeID
                    MyMevsimlikTable.Rows.Add(MyMevsimlikRow)
                Next
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyMustemilatCommandBuilder As New OleDb.OleDbCommandBuilder
                MyMustemilatCommandBuilder.DataAdapter = MyMustemilatDataAdapter
                MyMustemilatDataAdapter.Update(MyMustemilatTable)
                MyMustemilatTable = Nothing
                MyMustemilatCommandBuilder = Nothing
                MyMustemilatDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyMevsimlikCommandBuilder As New OleDb.OleDbCommandBuilder
                MyMevsimlikCommandBuilder.DataAdapter = MyMevsimlikDataAdapter
                MyMevsimlikDataAdapter.Update(MyMevsimlikTable)
                MyMevsimlikTable = Nothing
                MyMevsimlikCommandBuilder = Nothing
                MyMevsimlikDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return MyStatus
    End Function

    Private Function UpdateDataSQL(MustemilatCollection As Collection, MevsimlikCollection As Collection, _ConnectionString As String) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New SqlConnection(_ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryStringMustemilat As String = "SELECT * FROM MUSTEMILAT"
            Dim MyMustemilatDataAdapter As SqlDataAdapter = New SqlDataAdapter(New SqlCommand(MyQueryStringMustemilat, MyKamuConnection))
            Dim MyMustemilatTable As New DataTable
            MyMustemilatDataAdapter.Fill(MyMustemilatTable)

            Dim MyQueryStringMevsimlik As String = "SELECT * FROM MEVSIMLIK"
            Dim MyMevsimlikDataAdapter As SqlDataAdapter = New SqlDataAdapter(New SqlCommand(MyQueryStringMevsimlik, MyKamuConnection))
            Dim MyMevsimlikTable As New DataTable
            MyMevsimlikDataAdapter.Fill(MyMevsimlikTable)

            Try
                For Each _Mustemilat As Mustemilat In MustemilatCollection
                    Dim MyMustemilatRow As DataRow = MyMustemilatTable.NewRow()
                    MyMustemilatRow("PARSEL_ID") = _Mustemilat.ParselID
                    MyMustemilatRow("SAHIP_ID") = _Mustemilat.SahipID
                    MyMustemilatRow("TANIM") = _Mustemilat.Tanim
                    MyMustemilatRow("ADET") = _Mustemilat.Adet
                    MyMustemilatRow("FIYAT") = _Mustemilat.Fiyat
                    MyMustemilatRow("MALIK") = _Mustemilat.Malik
                    MyMustemilatRow("PAY") = _Mustemilat.Pay
                    MyMustemilatRow("PAYDA") = _Mustemilat.Payda
                    MyMustemilatRow("ODEME_ID") = _Mustemilat.OdemeID
                    MyMustemilatTable.Rows.Add(MyMustemilatRow)
                Next
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                For Each _Mevsimlik As Mevsimlik In MevsimlikCollection
                    Dim MyMevsimlikRow As DataRow = MyMevsimlikTable.NewRow()
                    MyMevsimlikRow("PARSEL_ID") = _Mevsimlik.ParselID
                    MyMevsimlikRow("SAHIP_ID") = _Mevsimlik.SahipID
                    MyMevsimlikRow("TANIM") = _Mevsimlik.Tanim
                    MyMevsimlikRow("ALAN") = _Mevsimlik.Alan
                    MyMevsimlikRow("BEDEL") = _Mevsimlik.Bedel
                    MyMevsimlikRow("MALIK") = _Mevsimlik.Malik
                    MyMevsimlikRow("PAY") = _Mevsimlik.Pay
                    MyMevsimlikRow("PAYDA") = _Mevsimlik.Payda
                    MyMevsimlikRow("ODEME_ID") = _Mevsimlik.OdemeID
                    MyMevsimlikTable.Rows.Add(MyMevsimlikRow)
                Next
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyMustemilatCommandBuilder As New SqlCommandBuilder
                MyMustemilatCommandBuilder.DataAdapter = MyMustemilatDataAdapter
                MyMustemilatDataAdapter.Update(MyMustemilatTable)
                MyMustemilatTable = Nothing
                MyMustemilatCommandBuilder = Nothing
                MyMustemilatDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim MyMevsimlikCommandBuilder As New SqlCommandBuilder
                MyMevsimlikCommandBuilder.DataAdapter = MyMevsimlikDataAdapter
                MyMevsimlikDataAdapter.Update(MyMevsimlikTable)
                MyMevsimlikTable = Nothing
                MyMevsimlikCommandBuilder = Nothing
                MyMevsimlikDataAdapter = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return MyStatus
    End Function

    Private Function GetMalikID(_Kisi As Kisi, Kisiler As DataTable) As Long
        Dim MyKisiID As Long = 0
        For Each MyRow As DataRow In Kisiler.Rows
            Dim SorguKisi As New Kisi(MyRow("ADI").ToString, MyRow("SOYADI").ToString, MyRow("BABA").ToString)
            If Not IsDBNull(MyRow("TC_KIMLIK_NO")) Then
                SorguKisi.TCKimlikNo = MyRow("TC_KIMLIK_NO")
            End If
            If SorguKisi.Equals(_Kisi) Then
                If Not IsDBNull(MyRow("ID")) Then
                    MyKisiID = MyRow("ID")
                Else
                    Dim MyKisiInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
                    MyKisiID = CInt(MyKisiInfo.GetValue(MyRow))
                End If


                'Dim MyKisiInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
                'MyKisiID = CInt(MyKisiInfo.GetValue(MyRow))

                'Dim MyKisiRowID As Long
                'If Not IsDBNull(MyRow("ID")) Then
                '    MyKisiRowID = MyRow("ID")
                'End If
                'If (MyKisiID <> MyKisiRowID) And (MyKisiRowID > 0) Then
                '    MyKisiID = MyKisiRowID
                'End If
                Exit For
            End If
        Next
        Return MyKisiID
    End Function

End Class
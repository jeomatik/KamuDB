﻿Imports Kamu.Objects
Public Class DB
    Public ProjectList As Collection
    Public KamuVeriXMLDosya As String
    Public MyOle As New Ole
    Public MySQL As New SQL
    Public MyPgSQL As New PgSQL
    Private _LogTut As Boolean
    Private _ConnectionInfo As ConnectionInfo

    Public Property LogTut() As Boolean
        Get
            Return _LogTut
        End Get
        Set(ByVal value As Boolean)
            _LogTut = value
        End Set
    End Property

    Public Property ConnectionInfo() As ConnectionInfo
        Get
            Return _ConnectionInfo
        End Get
        Set(ByVal value As ConnectionInfo)
            _ConnectionInfo = value
        End Set
    End Property

    Sub New()

    End Sub

    Sub New(ByVal _ConnectionInfo As ConnectionInfo)
        Me.ConnectionInfo = _ConnectionInfo
    End Sub

    Public Function GetDataTable(ByVal _SQLCommand As String) As DataTable
        Try
            Dim MyDataTable As New DataTable
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyDataTable = MyOle.GetDataTable(_SQLCommand)
                Case Connections.SqlConnection
                    _SQLCommand = _SQLCommand.Replace("&", "+")
                    _SQLCommand = _SQLCommand.Replace("True", "1")
                    MyDataTable = MySQL.GetDataTable(_SQLCommand)
                Case Connections.PgSqlConnection
                    MyDataTable = MyPgSQL.GetDataTable(_SQLCommand)
            End Select
            Return MyDataTable
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            Return Nothing
        End Try
    End Function

    Public Function CreateProjectList() As Collection
        Dim MyObject As New Collection()
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.CreateProjectList()
                Case Connections.SqlConnection
                    MyObject = MySQL.CreateProjectList()
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.CreateProjectList()
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function CreateComboList(strTableName As String, strColumnName As String) As Collection
        Dim MyObject As New Collection()
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.CreateComboList(strTableName, strColumnName)
                Case Connections.SqlConnection
                    MyObject = MySQL.CreateComboList(strTableName, strColumnName)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.CreateComboList(strTableName, strColumnName)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function MergeKisi(_AktifKisiID As Long, _PasifKisiID As Long) As Boolean
        Dim MyStatus As Boolean = False
        If ChangeMalik("MULKIYET", _AktifKisiID, _PasifKisiID) Then
            If ChangeMalik("MUSTEMILAT", _AktifKisiID, _PasifKisiID) Then

            End If
            If ChangeMalik("MEVSIMLIK", _AktifKisiID, _PasifKisiID) Then

            End If
            If DeleteKisi(_PasifKisiID) Then
                MyStatus = True
            End If
        End If
        Return MyStatus
    End Function

    Private Function ChangeMalik(_TableName As String, _AktifKisiID As Long, _PasifKisiID As Long)
        Dim MyStatus As Boolean = False
        Try
            Dim _connection As New OleDb.OleDbConnection(ConnectionInfo.ConnectionString)
            Dim command As OleDb.OleDbCommand = _connection.CreateCommand()
            command.CommandText = "UPDATE KISI_ID=" + _AktifKisiID.ToString + " FROM " + _TableName + " WHERE KISI_ID=" + _PasifKisiID.ToString
            _connection.Open()

            Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter With {
                .SelectCommand = command
            }

            Dim table As New DataTable With {
                .Locale = System.Globalization.CultureInfo.InvariantCulture
            }
            adapter.Fill(table)
            adapter = Nothing
            _connection.Close()
            table = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try

        Return MyStatus
    End Function

    Public Function DefineVerasetDurumu(_Parseller As Collection) As Collection
        For Each _Parsel As Parsel In _Parseller
            For Each _Kisi As Kisi In _Parsel.Malikler
                Dim MyVarisler As New Collection
                If _Kisi.Durumu.ToUpper = "ÖLÜ" Then
                    MyVarisler = GetVarisler(_Kisi.ID)
                End If
                _Kisi.Varisler = MyVarisler
                If MyVarisler.Count > 0 Then
                    _Kisi.HasVaris = True
                Else
                    _Kisi.HasVaris = False
                End If
                Dim MyMurisler As Collection = GetMurisler(_Kisi.ID)
                If MyMurisler.Count > 0 Then
                    _Kisi.IsVaris = True
                End If
            Next
        Next
        Return _Parseller
    End Function

    Public Function DefineMustemilatOwnerShip(_Mustemilatlar As Collection, _ParselConversionTable As DataTable, _KisiConversionTable As DataTable) As Collection
        For Each _Mustemilat As Mustemilat In _Mustemilatlar '!!!
            Dim Eski_Parsel_ID As String = _Mustemilat.ParselGUID
            Dim Eski_Kisi_ID As String = _Mustemilat.SahipGUID
            Dim Yeni_Parsel_ID As String = ""
            Dim Yeni_Kisi_ID As String = ""
            If Eski_Parsel_ID <> "" Then
                Dim foundRows() As DataRow = _ParselConversionTable.Select("ESKI_ID=" & Eski_Parsel_ID.ToString)
                If foundRows.Count > 0 Then

                    For i As Integer = 0 To foundRows.GetUpperBound(0)
                        Yeni_Parsel_ID = foundRows(i)(0)
                    Next i
                End If
            End If
            If Eski_Kisi_ID <> "" Then
                Dim foundRows() As DataRow = _KisiConversionTable.Select("ESKI_ID=" & Eski_Kisi_ID.ToString)
                If foundRows.Count > 0 Then
                    For i As Integer = 0 To foundRows.GetUpperBound(0)
                        Yeni_Kisi_ID = foundRows(i)(0)
                    Next i
                End If

            End If
            _Mustemilat.ParselGUID = Yeni_Parsel_ID
            _Mustemilat.SahipGUID = Yeni_Kisi_ID
        Next
        Return _Mustemilatlar
    End Function

    Public Function DefineMevsimlikOwnerShip(_Mevsimlikler As Collection, _ParselConversionTable As DataTable, _KisiConversionTable As DataTable) As Collection
        For Each _Mevsimlik As Mevsimlik In _Mevsimlikler
            Dim Eski_Parsel_ID As String = _Mevsimlik.ParselGUID
            Dim Eski_Kisi_ID As String = _Mevsimlik.SahipGUID
            Dim Yeni_Parsel_ID As String = ""
            Dim Yeni_Kisi_ID As String = ""
            If Eski_Parsel_ID <> "" Then
                Dim foundRows() As DataRow = _ParselConversionTable.Select("ESKI_ID=" & Eski_Parsel_ID.ToString)
                If foundRows.Count > 0 Then
                    For i As Integer = 0 To foundRows.GetUpperBound(0)
                        Yeni_Parsel_ID = foundRows(i)(0)
                    Next i
                End If
            End If
            If Eski_Kisi_ID <> "" Then
                Dim foundRows() As DataRow = _KisiConversionTable.Select("ESKI_ID=" & Eski_Kisi_ID.ToString)
                If foundRows.Count > 0 Then
                    For i As Integer = 0 To foundRows.GetUpperBound(0)
                        Yeni_Kisi_ID = foundRows(i)(0)
                    Next i
                End If
            End If
            _Mevsimlik.ParselGUID = Yeni_Parsel_ID
            _Mevsimlik.SahipGUID = Yeni_Kisi_ID
        Next
        Return _Mevsimlikler
    End Function

#Region "Get Collections"

    'Public Function GetParselCollection(_DataTable As DataTable, Optional WithOutCode As Boolean = False, Optional WithVaris As Boolean = False) As Collection
    '    Dim MyParseller As New Collection
    '    Dim MyMalikler As New Collection
    '    Dim MyParsel As New Parsel
    '    Dim MyMalik As New Kisi
    '    Dim MyVarisler As New Collection
    '    MyMalik.Varisler = MyVarisler
    '    Dim LastAda As String = "-1"
    '    Dim LastParsel As String = "-1"
    '    For Each MyRow As DataRow In _DataTable.Rows
    '        If (LastAda = MyRow("ADA") And LastParsel = MyRow("PARSEL")) Then
    '            If Not IsDBNull(MyRow("KISI_ID")) Then
    '                MyMalik.ID = MyRow("KISI_ID")
    '            End If
    '            MyMalik.Adi = MyRow("ADI").ToString
    '            MyMalik.Soyadi = MyRow("SOYADI").ToString
    '            MyMalik.Baba = MyRow("BABA").ToString
    '            If Not IsDBNull(MyRow("PAY")) Then
    '                MyMalik.HissePay = MyRow("PAY")
    '            End If
    '            If Not IsDBNull(MyRow("PAYDA")) Then
    '                MyMalik.HissePayda = MyRow("PAYDA")
    '            End If
    '            If Not IsDBNull(MyRow("TAPU_TARIHI")) Then
    '                MyMalik.TapuTarihi = MyRow("TAPU_TARIHI")
    '            End If
    '            MyMalik.Dusunceler = MyRow("DUSUNCELER").ToString
    '            MyMalik.Durumu = MyRow("DURUMU").ToString
    '            MyMalik.Telefon = MyRow("TELEFON").ToString
    '            MyMalik.Adres = MyRow("ADRES").ToString
    '            If Not IsDBNull(MyRow("TC_KIMLIK_NO")) Then
    '                MyMalik.TCKimlikNo = MyRow("TC_KIMLIK_NO")
    '            End If

    '            If WithVaris Then
    '                If Not IsDBNull(MyRow("VARIS")) Then
    '                    MyMalik.Varisler.Add(MyRow("VARIS"))
    '                End If
    '            End If

    '            If Not WithOutCode Then
    '                Dim MyMalikKod As New KisiKod
    '                If Not IsDBNull(MyRow("DAVETIYE_TEBLIG_DURUMU")) Then
    '                    MyMalikKod.DavetiyeTebligDurumu = MyRow("DAVETIYE_TEBLIG_DURUMU")
    '                End If
    '                If Not IsDBNull(MyRow("DAVETIYE_ALINMA_DURUMU")) Then
    '                    MyMalikKod.DavetiyeAlinmaDurumu = MyRow("DAVETIYE_ALINMA_DURUMU")
    '                End If
    '                If Not IsDBNull(MyRow("GORUSME_DURUMU")) Then
    '                    MyMalikKod.GorusmeDurumu = MyRow("GORUSME_DURUMU")
    '                End If
    '                If Not IsDBNull(MyRow("GORUSME_NO")) Then
    '                    MyMalikKod.GorusmeNo = MyRow("GORUSME_NO")
    '                End If
    '                If Not IsDBNull(MyRow("GORUSME_TARIHI")) Then
    '                    MyMalikKod.GorusmeTarihi = MyRow("GORUSME_TARIHI")
    '                End If
    '                If Not IsDBNull(MyRow("ANLASMA_DURUMU")) Then
    '                    MyMalikKod.AnlasmaDurumu = MyRow("ANLASMA_DURUMU")
    '                End If
    '                If Not IsDBNull(MyRow("ANLASMA_TARIHI")) Then
    '                    MyMalikKod.AnlasmaTarihi = MyRow("ANLASMA_TARIHI")
    '                End If
    '                If Not IsDBNull(MyRow("ANLASMA_DUSUNCELER")) Then
    '                    MyMalikKod.AnlasmaDusunceler = MyRow("ANLASMA_DUSUNCELER")
    '                End If
    '                If Not IsDBNull(MyRow("TESCIL_DURUMU")) Then
    '                    MyMalikKod.TescilDurumu = MyRow("TESCIL_DURUMU")
    '                End If
    '                MyMalik.Kod = MyMalikKod
    '            End If

    '            MyMalikler.Add(MyMalik)
    '            MyMalik = New Kisi
    '            MyVarisler = New Collection
    '            MyMalik.Varisler = MyVarisler
    '        Else
    '            If MyMalikler.Count > 0 Then
    '                MyParsel.Malikler = MyMalikler
    '                MyParseller.Add(MyParsel)
    '                MyMalikler = New Collection
    '                MyMalik = New Kisi
    '                MyVarisler = New Collection
    '                MyMalik.Varisler = MyVarisler
    '                MyParsel = New Parsel
    '            End If
    '            If Not IsDBNull(MyRow("ID")) Then
    '                MyParsel.ID = MyRow("ID")
    '            End If
    '            If Not IsDBNull(MyRow("PROJE_ID")) Then
    '                MyParsel.ProjeID = MyRow("PROJE_ID")
    '            End If
    '            MyParsel.Il = MyRow("IL").ToString
    '            MyParsel.Ilce = MyRow("ILCE").ToString
    '            MyParsel.Koy = MyRow("KOY").ToString
    '            MyParsel.Mahalle = MyRow("MAHALLE").ToString
    '            MyParsel.AdaNo = MyRow("ADA").ToString
    '            MyParsel.ParselNo = MyRow("PARSEL").ToString
    '            MyParsel.PaftaNo = MyRow("PAFTA").ToString
    '            MyParsel.Cinsi = MyRow("CINSI").ToString
    '            MyParsel.Mevki = MyRow("MEVKI").ToString
    '            MyParsel.Cilt = MyRow("CILT").ToString
    '            MyParsel.Sayfa = MyRow("SAYFA").ToString
    '            If Not IsDBNull(MyRow("TAPU_ALANI")) Then
    '                MyParsel.TapuAlani = MyRow("TAPU_ALANI")
    '            End If
    '            If Not IsDBNull(MyRow("IRTIFAK_ALAN")) Then
    '                MyParsel.IrtifakAlan = MyRow("IRTIFAK_ALAN")
    '            End If
    '            If Not IsDBNull(MyRow("GECICI_IRTIFAK_ALAN")) Then
    '                MyParsel.GeciciIrtifakAlan = MyRow("GECICI_IRTIFAK_ALAN")
    '            End If
    '            If Not IsDBNull(MyRow("MULKIYET_ALAN")) Then
    '                MyParsel.MulkiyetAlan = MyRow("MULKIYET_ALAN")
    '            End If
    '            If Not IsDBNull(MyRow("IRTIFAK_BEDEL")) Then
    '                MyParsel.IrtifakBedel = MyRow("IRTIFAK_BEDEL")
    '            End If
    '            If Not IsDBNull(MyRow("GECICI_IRTIFAK_BEDEL")) Then
    '                MyParsel.GeciciIrtifakBedel = MyRow("GECICI_IRTIFAK_BEDEL")
    '            End If
    '            If Not IsDBNull(MyRow("MULKIYET_BEDEL")) Then
    '                MyParsel.MulkiyetBedel = MyRow("MULKIYET_BEDEL")
    '            End If

    '            If Not WithOutCode Then
    '                Dim MyParselKod As New ParselKod
    '                If Not IsDBNull(MyRow("KADASTRAL_DURUM")) Then
    '                    MyParselKod.KadastralDurum = MyRow("KADASTRAL_DURUM")
    '                End If
    '                If Not IsDBNull(MyRow("MALIK_TIPI")) Then
    '                    MyParselKod.MalikTipi = MyRow("MALIK_TIPI")
    '                End If
    '                If Not IsDBNull(MyRow("ISTIMLAK_TURU")) Then
    '                    MyParselKod.IstimlakTuru = MyRow("ISTIMLAK_TURU")
    '                End If
    '                If Not IsDBNull(MyRow("ISTIMLAK_SERHI")) Then
    '                    MyParselKod.IstimlakSerhi = MyRow("ISTIMLAK_SERHI")
    '                End If
    '                If Not IsDBNull(MyRow("DAVA10_DURUMU")) Then
    '                    MyParselKod.DavaDurumu10 = MyRow("DAVA10_DURUMU")
    '                End If
    '                If Not IsDBNull(MyRow("DAVA27_DURUMU")) Then
    '                    MyParselKod.DavaDurumu27 = MyRow("DAVA27_DURUMU")
    '                End If
    '                If Not IsDBNull(MyRow("ISTIMLAK_DISI")) Then
    '                    MyParselKod.IstimlakDisi = MyRow("ISTIMLAK_DISI")
    '                End If
    '                If Not IsDBNull(MyRow("DEVIR_DURUMU")) Then
    '                    MyParselKod.DevirDurumu = MyRow("DEVIR_DURUMU")
    '                End If
    '                If Not IsDBNull(MyRow("EDINIM_DURUMU")) Then
    '                    MyParselKod.EdinimDurumu = MyRow("EDINIM_DURUMU")
    '                End If
    '                MyParsel.Kod = MyParselKod
    '            End If

    '            LastAda = MyParsel.AdaNo
    '            LastParsel = MyParsel.ParselNo
    '            If Not IsDBNull(MyRow("KISI_ID")) Then
    '                MyMalik.ID = MyRow("KISI_ID")
    '            End If
    '            MyMalik.Adi = MyRow("ADI").ToString
    '            MyMalik.Soyadi = MyRow("SOYADI").ToString
    '            MyMalik.Baba = MyRow("BABA").ToString
    '            If Not IsDBNull(MyRow("PAY")) Then
    '                MyMalik.HissePay = MyRow("PAY")
    '            End If
    '            If Not IsDBNull(MyRow("PAYDA")) Then
    '                MyMalik.HissePayda = MyRow("PAYDA")
    '            End If
    '            If Not IsDBNull(MyRow("TAPU_TARIHI")) Then
    '                MyMalik.TapuTarihi = MyRow("TAPU_TARIHI")
    '            End If
    '            MyMalik.Dusunceler = MyRow("DUSUNCELER").ToString
    '            MyMalik.Durumu = MyRow("DURUMU").ToString
    '            MyMalik.Telefon = MyRow("TELEFON").ToString
    '            MyMalik.Adres = MyRow("ADRES").ToString
    '            If Not IsDBNull(MyRow("TC_KIMLIK_NO")) Then
    '                MyMalik.TCKimlikNo = MyRow("TC_KIMLIK_NO")
    '            End If

    '            If Not WithOutCode Then
    '                Dim MyMalikKod As New KisiKod
    '                If Not IsDBNull(MyRow("DAVETIYE_TEBLIG_DURUMU")) Then
    '                    MyMalikKod.DavetiyeTebligDurumu = MyRow("DAVETIYE_TEBLIG_DURUMU")
    '                End If
    '                If Not IsDBNull(MyRow("DAVETIYE_ALINMA_DURUMU")) Then
    '                    MyMalikKod.DavetiyeAlinmaDurumu = MyRow("DAVETIYE_ALINMA_DURUMU")
    '                End If
    '                If Not IsDBNull(MyRow("GORUSME_DURUMU")) Then
    '                    MyMalikKod.GorusmeDurumu = MyRow("GORUSME_DURUMU")
    '                End If
    '                If Not IsDBNull(MyRow("GORUSME_NO")) Then
    '                    MyMalikKod.GorusmeNo = MyRow("GORUSME_NO")
    '                End If
    '                If Not IsDBNull(MyRow("GORUSME_TARIHI")) Then
    '                    MyMalikKod.GorusmeTarihi = MyRow("GORUSME_TARIHI")
    '                End If
    '                If Not IsDBNull(MyRow("ANLASMA_DURUMU")) Then
    '                    MyMalikKod.AnlasmaDurumu = MyRow("ANLASMA_DURUMU")
    '                End If
    '                If Not IsDBNull(MyRow("ANLASMA_TARIHI")) Then
    '                    MyMalikKod.AnlasmaTarihi = MyRow("ANLASMA_TARIHI")
    '                End If
    '                If Not IsDBNull(MyRow("ANLASMA_DUSUNCELER")) Then
    '                    MyMalikKod.AnlasmaDusunceler = MyRow("ANLASMA_DUSUNCELER")
    '                End If
    '                If Not IsDBNull(MyRow("TESCIL_DURUMU")) Then
    '                    MyMalikKod.TescilDurumu = MyRow("TESCIL_DURUMU")
    '                End If
    '                MyMalik.Kod = MyMalikKod
    '            End If

    '            MyMalikler.Add(MyMalik)
    '            MyMalik = New Kisi
    '            MyVarisler = New Collection
    '            MyMalik.Varisler = MyVarisler
    '        End If
    '    Next
    '    MyParsel.Malikler = MyMalikler
    '    MyParseller.Add(MyParsel)
    '    Return MyParseller
    'End Function

    Public Function GetParselCollection(_DataTable As DataTable, Optional WithOutCode As Boolean = False, Optional VerasetDurumu As Boolean = False) As Collection
        Dim MyParseller As New Collection
        Dim MyMalikler As New Collection
        Dim MyParsel As New Parsel
        Dim MyMalik As New Kisi
        Dim LastAda As String = "-1"
        Dim LastParsel As String = "-1"
        For Each MyRow As DataRow In _DataTable.Rows
            If (LastAda = MyRow("ADA") And LastParsel = MyRow("PARSEL")) Then
                If Not IsDBNull(MyRow("KISI_ID")) Then
                    MyMalik.ID = MyRow("KISI_ID")
                End If
                MyMalik.Adi = MyRow("ADI").ToString
                MyMalik.Soyadi = MyRow("SOYADI").ToString
                MyMalik.Baba = MyRow("BABA").ToString
                If Not IsDBNull(MyRow("PAY")) Then
                    MyMalik.HissePay = MyRow("PAY")
                End If
                If Not IsDBNull(MyRow("PAYDA")) Then
                    MyMalik.HissePayda = MyRow("PAYDA")
                End If
                If Not IsDBNull(MyRow("TAPU_TARIHI")) Then
                    MyMalik.TapuTarihi = MyRow("TAPU_TARIHI")
                End If
                MyMalik.Dusunceler = MyRow("DUSUNCELER").ToString
                MyMalik.Cinsiyet = MyRow("CINSIYET").ToString
                MyMalik.Durumu = MyRow("DURUMU").ToString
                MyMalik.Telefon = MyRow("TELEFON").ToString
                MyMalik.Adres = MyRow("ADRES").ToString
                If Not IsDBNull(MyRow("TC_KIMLIK_NO")) Then
                    MyMalik.TCKimlikNo = MyRow("TC_KIMLIK_NO")
                End If

                If Not WithOutCode Then
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
                End If

                MyMalikler.Add(MyMalik)
                MyMalik = New Kisi
            Else
                If MyMalikler.Count > 0 Then
                    MyParsel.Malikler = MyMalikler
                    MyParseller.Add(MyParsel)
                    MyMalikler = New Collection
                    MyMalik = New Kisi
                    MyParsel = New Parsel
                End If
                If Not IsDBNull(MyRow("ID")) Then
                    MyParsel.ID = MyRow("ID")
                End If
                If Not IsDBNull(MyRow("PROJE_ID")) Then
                    MyParsel.ProjeID = MyRow("PROJE_ID")
                End If

                MyParsel.Il = MyRow("IL").ToString
                MyParsel.Ilce = MyRow("ILCE").ToString
                MyParsel.Koy = MyRow("KOY").ToString
                MyParsel.Mahalle = MyRow("MAHALLE").ToString
                MyParsel.AdaNo = MyRow("ADA").ToString
                MyParsel.ParselNo = MyRow("PARSEL").ToString
                MyParsel.PaftaNo = MyRow("PAFTA").ToString
                MyParsel.Cinsi = MyRow("CINSI").ToString
                MyParsel.Mevki = MyRow("MEVKI").ToString
                MyParsel.Cilt = MyRow("CILT").ToString
                MyParsel.Sayfa = MyRow("SAYFA").ToString
                If Not IsDBNull(MyRow("TAPU_ALANI")) Then
                    MyParsel.TapuAlani = MyRow("TAPU_ALANI")
                End If
                If Not IsDBNull(MyRow("IRTIFAK_ALAN")) Then
                    MyParsel.IrtifakAlan = MyRow("IRTIFAK_ALAN")
                End If
                If Not IsDBNull(MyRow("GECICI_IRTIFAK_ALAN")) Then
                    MyParsel.GeciciIrtifakAlan = MyRow("GECICI_IRTIFAK_ALAN")
                End If
                If Not IsDBNull(MyRow("MULKIYET_ALAN")) Then
                    MyParsel.MulkiyetAlan = MyRow("MULKIYET_ALAN")
                End If
                If Not IsDBNull(MyRow("IRTIFAK_BEDEL")) Then
                    MyParsel.IrtifakBedel = MyRow("IRTIFAK_BEDEL")
                End If
                If Not IsDBNull(MyRow("GECICI_IRTIFAK_BEDEL")) Then
                    MyParsel.GeciciIrtifakBedel = MyRow("GECICI_IRTIFAK_BEDEL")
                End If
                If Not IsDBNull(MyRow("MULKIYET_BEDEL")) Then
                    MyParsel.MulkiyetBedel = MyRow("MULKIYET_BEDEL")
                End If
                MyParsel.KamulastirmaAmaci = MyRow("KAMULASTIRMA_AMACI").ToString
                MyParsel.AraziVasfi = MyRow("ARAZI_VASFI").ToString
                MyParsel.YayginMunavebeSistemi = MyRow("YAYGIN_MUNAVEBE_SISTEMI").ToString
                MyParsel.DegerlemeRaporu = MyRow("DEGERLEME_RAPORU").ToString
                If Not IsDBNull(MyRow("DEGERLEME_TARIHI")) Then
                    MyParsel.DegerlemeTarihi = MyRow("DEGERLEME_TARIHI")
                End If
                If Not IsDBNull(MyRow("YILLIK_ORTALAMA_NET_GELIR")) Then
                    MyParsel.YillikOrtalamaNetGelir = MyRow("YILLIK_ORTALAMA_NET_GELIR")
                End If
                If Not IsDBNull(MyRow("KAPITALIZASYON_FAIZI")) Then
                    MyParsel.KapitalizasyonOrani = MyRow("KAPITALIZASYON_FAIZI")
                End If
                If Not IsDBNull(MyRow("OBJEKTIF_ARTIS")) Then
                    MyParsel.ObjektifArtis = MyRow("OBJEKTIF_ARTIS")
                End If
                If Not IsDBNull(MyRow("ART_KISIM_ARTIS")) Then
                    MyParsel.ArtanKisimArtis = MyRow("ART_KISIM_ARTIS")
                End If
                If Not IsDBNull(MyRow("VERIM_KAYBI")) Then
                    MyParsel.VerimKaybi = MyRow("VERIM_KAYBI")
                End If
                MyParsel.SerhBeyan = MyRow("SERH_BEYAN").ToString
                'If Not IsDBNull(MyRow("ODEME_ID")) Then
                '    MyParsel.kamuodeme = MyRow("ODEME_ID")
                'End If

                If Not WithOutCode Then
                    Dim MyParselKod As New ParselKod
                    MyParselKod.Kod = MyRow("KOD").ToString
                    If Not IsDBNull(MyRow("BOLGE_ID")) Then
                        MyParselKod.BolgeID = MyRow("BOLGE_ID")
                    End If
                    If Not IsDBNull(MyRow("KADASTRAL_DURUM")) Then
                        MyParselKod.KadastralDurum = MyRow("KADASTRAL_DURUM")
                    End If
                    If Not IsDBNull(MyRow("MALIK_TIPI")) Then
                        MyParselKod.MalikTipi = MyRow("MALIK_TIPI")
                    End If
                    If Not IsDBNull(MyRow("ISTIMLAK_TURU")) Then
                        MyParselKod.IstimlakTuru = MyRow("ISTIMLAK_TURU")
                    End If
                    If Not IsDBNull(MyRow("ISTIMLAK_SERHI")) Then
                        MyParselKod.IstimlakSerhi = MyRow("ISTIMLAK_SERHI")
                    End If
                    If Not IsDBNull(MyRow("DAVA10_DURUMU")) Then
                        MyParselKod.DavaDurumu10 = MyRow("DAVA10_DURUMU")
                    End If
                    If Not IsDBNull(MyRow("DAVA27_DURUMU")) Then
                        MyParselKod.DavaDurumu27 = MyRow("DAVA27_DURUMU")
                    End If
                    If Not IsDBNull(MyRow("ISTIMLAK_DISI")) Then
                        MyParselKod.IstimlakDisi = MyRow("ISTIMLAK_DISI")
                    End If
                    If Not IsDBNull(MyRow("DEVIR_DURUMU")) Then
                        MyParselKod.DevirDurumu = MyRow("DEVIR_DURUMU")
                    End If
                    If Not IsDBNull(MyRow("EDINIM_DURUMU")) Then
                        MyParselKod.EdinimDurumu = MyRow("EDINIM_DURUMU")
                    End If
                    If Not IsDBNull(MyRow("ODEME_DURUMU")) Then
                        MyParselKod.OdemeDurumu = MyRow("ODEME_DURUMU")
                    End If
                    MyParsel.Kod = MyParselKod
                End If

                LastAda = MyParsel.AdaNo
                LastParsel = MyParsel.ParselNo
                If Not IsDBNull(MyRow("KISI_ID")) Then
                    MyMalik.ID = MyRow("KISI_ID")
                End If
                MyMalik.Adi = MyRow("ADI").ToString
                MyMalik.Soyadi = MyRow("SOYADI").ToString
                MyMalik.Baba = MyRow("BABA").ToString
                If Not IsDBNull(MyRow("PAY")) Then
                    MyMalik.HissePay = MyRow("PAY")
                End If
                If Not IsDBNull(MyRow("PAYDA")) Then
                    MyMalik.HissePayda = MyRow("PAYDA")
                End If
                If Not IsDBNull(MyRow("TAPU_TARIHI")) Then
                    MyMalik.TapuTarihi = MyRow("TAPU_TARIHI")
                End If
                MyMalik.Dusunceler = MyRow("DUSUNCELER").ToString
                MyMalik.Cinsiyet = MyRow("CINSIYET").ToString
                MyMalik.Durumu = MyRow("DURUMU").ToString
                MyMalik.Telefon = MyRow("TELEFON").ToString
                MyMalik.Adres = MyRow("ADRES").ToString
                If Not IsDBNull(MyRow("TC_KIMLIK_NO")) Then
                    MyMalik.TCKimlikNo = MyRow("TC_KIMLIK_NO")
                End If

                If Not WithOutCode Then
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
                End If

                MyMalikler.Add(MyMalik)
                MyMalik = New Kisi
            End If
        Next
        MyParsel.Malikler = MyMalikler
        MyParseller.Add(MyParsel)
        Return MyParseller
    End Function

    Public Function GetMustemilatCollection(_DataTable As DataTable) As Collection
        Dim MyMustemilatlar As New Collection
        For Each MyRow As DataRow In _DataTable.Rows
            Dim MyMustemilat As New Mustemilat
            If Not IsDBNull(MyRow("ID")) Then
                MyMustemilat.ID = MyRow("ID")
            End If
            If Not IsDBNull(MyRow("PARSEL_ID")) Then
                MyMustemilat.ParselGUID = MyRow("PARSEL_GLOBALID")
            End If
            If Not IsDBNull(MyRow("SAHIP_ID")) Then
                MyMustemilat.SahipGUID = MyRow("SAHIP_GLOBALID")
            End If
            MyMustemilat.Tanim = MyRow("TANIM").ToString
            If Not IsDBNull(MyRow("ADET")) Then
                MyMustemilat.Adet = MyRow("ADET")
            End If
            If Not IsDBNull(MyRow("FIYAT")) Then
                MyMustemilat.Fiyat = MyRow("FIYAT")
            End If
            If Not IsDBNull(MyRow("MALIK")) Then
                MyMustemilat.Malik = MyRow("MALIK")
            End If
            If Not IsDBNull(MyRow("PAY")) Then
                MyMustemilat.Pay = MyRow("PAY")
            End If
            If Not IsDBNull(MyRow("PAYDA")) Then
                MyMustemilat.Payda = MyRow("PAYDA")
            End If
            If Not IsDBNull(MyRow("ODEME_GLOBALID")) Then
                MyMustemilat.OdemeGUID = MyRow("ODEME_GLOBALID")
            End If
            MyMustemilatlar.Add(MyMustemilat)
        Next
        Return MyMustemilatlar
    End Function

    Public Function GetMevsimlikCollection(_DataTable As DataTable) As Collection
        Dim MyMevsimlikler As New Collection
        For Each MyRow As DataRow In _DataTable.Rows
            Dim MyMevsimlik As New Mevsimlik
            If Not IsDBNull(MyRow("ID")) Then
                MyMevsimlik.ID = MyRow("ID")
            End If
            If Not IsDBNull(MyRow("PARSEL_GLOBALID")) Then
                MyMevsimlik.ParselGUID = MyRow("PARSEL_GLOBALID")
            End If
            If Not IsDBNull(MyRow("SAHIP_GLOBALID")) Then
                MyMevsimlik.SahipGUID = MyRow("SAHIP_GLOBALID")
            End If
            MyMevsimlik.Tanim = MyRow("TANIM").ToString
            If Not IsDBNull(MyRow("ALAN")) Then
                MyMevsimlik.Alan = MyRow("ALAN")
            End If
            If Not IsDBNull(MyRow("BEDEL")) Then
                MyMevsimlik.Bedel = MyRow("BEDEL")
            End If
            If Not IsDBNull(MyRow("MALIK")) Then
                MyMevsimlik.Malik = MyRow("MALIK")
            End If
            If Not IsDBNull(MyRow("PAY")) Then
                MyMevsimlik.Pay = MyRow("PAY")
            End If
            If Not IsDBNull(MyRow("PAYDA")) Then
                MyMevsimlik.Payda = MyRow("PAYDA")
            End If
            If Not IsDBNull(MyRow("ODEME_GLOBALID")) Then
                MyMevsimlik.OdemeGUID = MyRow("ODEME_GLOBALID")
            End If
            MyMevsimlikler.Add(MyMevsimlik)
        Next
        Return MyMevsimlikler
    End Function

    Public Function GetDavaAceleCollection(_DataTable As DataTable) As Collection
        Dim MyAceleDavalar As New Collection
        For Each MyRow As DataRow In _DataTable.Rows
            Dim MyDavaAcele As New DavaAcele
            If Not IsDBNull(MyRow("ID")) Then
                MyDavaAcele.ID = MyRow("ID")
            End If
            If Not IsDBNull(MyRow("GLOBALID")) Then
                MyDavaAcele.GUID = MyRow("GLOBALID")
            End If
            If Not IsDBNull(MyRow("PARSEL_GLOBALID")) Then
                MyDavaAcele.ParselGUID = MyRow("PARSEL_GLOBALID")
            End If
            MyDavaAcele.Mahkeme = MyRow("MAHKEME").ToString
            MyDavaAcele.EsasNo = MyRow("ESAS_NO").ToString
            MyDavaAcele.KararNo = MyRow("KARAR_NO").ToString
            If Not IsDBNull(MyRow("KARAR_TARIHI")) Then
                MyDavaAcele.KararTarihi = MyRow("KARAR_TARIHI")
            End If
            If Not IsDBNull(MyRow("DAVA_ACILAN_HISSE_PAY")) Then
                MyDavaAcele.DavaAcilanHissePay = MyRow("DAVA_ACILAN_HISSE_PAY")
            End If
            If Not IsDBNull(MyRow("DAVA_ACILAN_HISSE_PAYDA")) Then
                MyDavaAcele.DavaAcilanHissePayda = MyRow("DAVA_ACILAN_HISSE_PAYDA")
            End If
            If Not IsDBNull(MyRow("TOPLAM_KAMULASTIRMA_BEDELI")) Then
                MyDavaAcele.ToplamKamulastirmaBedeli = MyRow("TOPLAM_KAMULASTIRMA_BEDELI")
            End If
            If Not IsDBNull(MyRow("DAVA_TARIHI")) Then
                MyDavaAcele.DavaTarihi = MyRow("DAVA_TARIHI")
            End If
            If Not IsDBNull(MyRow("KESIF_TARIHI")) Then
                MyDavaAcele.KesifTarihi = MyRow("KESIF_TARIHI")
            End If
            MyDavaAcele.BlokeOluru = MyRow("BLOKE_OLURU").ToString
            If Not IsDBNull(MyRow("OLUR_TARIHI")) Then
                MyDavaAcele.OlurTarihi = MyRow("OLUR_TARIHI")
            End If
            If Not IsDBNull(MyRow("BLOKE_TARIHI")) Then
                MyDavaAcele.BlokeTarihi = MyRow("BLOKE_TARIHI")
            End If
            MyDavaAcele.Avukat = MyRow("AVUKAT").ToString
            MyDavaAcele.Dusunceler = MyRow("DUSUNCELER").ToString
            MyAceleDavalar.Add(MyDavaAcele)
        Next
        Return MyAceleDavalar
    End Function

    Public Function GetDavaTescilCollection(_DataTable As DataTable) As Collection
        Dim MyTescilDavalar As New Collection
        For Each MyRow As DataRow In _DataTable.Rows
            Dim MyDavaTescil As New DavaTescil
            If Not IsDBNull(MyRow("ID")) Then
                MyDavaTescil.ID = MyRow("ID")
            End If
            If Not IsDBNull(MyRow("GLOBALID")) Then
                MyDavaTescil.GUID = MyRow("GLOBALID")
            End If
            If Not IsDBNull(MyRow("PARSEL_GLOBALID")) Then
                MyDavaTescil.ParselGUID = MyRow("PARSEL_GLOBALID")
            End If
            MyDavaTescil.Mahkeme = MyRow("MAHKEME").ToString
            MyDavaTescil.EsasNo = MyRow("ESAS_NO").ToString
            MyDavaTescil.KararNo = MyRow("KARAR_NO").ToString
            If Not IsDBNull(MyRow("KARAR_TARIHI")) Then
                MyDavaTescil.KararTarihi = MyRow("KARAR_TARIHI")
            End If
            If Not IsDBNull(MyRow("DAVA_ACILAN_HISSE_PAY")) Then
                MyDavaTescil.DavaAcilanHissePay = MyRow("DAVA_ACILAN_HISSE_PAY")
            End If
            If Not IsDBNull(MyRow("DAVA_ACILAN_HISSE_PAYDA")) Then
                MyDavaTescil.DavaAcilanHissePayda = MyRow("DAVA_ACILAN_HISSE_PAYDA")
            End If
            If Not IsDBNull(MyRow("TOPLAM_KAMULASTIRMA_BEDELI")) Then
                MyDavaTescil.ToplamKamulastirmaBedeli = MyRow("TOPLAM_KAMULASTIRMA_BEDELI")
            End If
            If Not IsDBNull(MyRow("DAVA_TARIHI")) Then
                MyDavaTescil.DavaTarihi = MyRow("DAVA_TARIHI")
            End If
            If Not IsDBNull(MyRow("BIRINCI_KESIF_TARIHI")) Then
                MyDavaTescil.KesifTarihi1 = MyRow("BIRINCI_KESIF_TARIHI")
            End If
            MyDavaTescil.BlokeOluru = MyRow("BLOKE_OLURU").ToString
            If Not IsDBNull(MyRow("OLUR_TARIHI")) Then
                MyDavaTescil.OlurTarihi = MyRow("OLUR_TARIHI")
            End If
            If Not IsDBNull(MyRow("BLOKE_TARIHI")) Then
                MyDavaTescil.BlokeTarihi = MyRow("BLOKE_TARIHI")
            End If
            MyDavaTescil.Avukat = MyRow("AVUKAT").ToString
            MyDavaTescil.Dusunceler = MyRow("DUSUNCELER").ToString
            MyTescilDavalar.Add(MyDavaTescil)
        Next
        Return MyTescilDavalar
    End Function

    Public Function GetOdemeCollection(_DataTable As DataTable) As Collection
        Dim MyOdemeler As New Collection
        For Each MyRow As DataRow In _DataTable.Rows
            Dim MyOdeme As New Odeme
            If Not IsDBNull(MyRow("ID")) Then
                MyOdeme.ID = MyRow("ID")
            End If
            If Not IsDBNull(MyRow("PARSEL_GLOBALID")) Then
                MyOdeme.ParselGUID = MyRow("PARSEL_GLOBALID")
            End If
            If Not IsDBNull(MyRow("KISI_GLOBALID")) Then
                MyOdeme.KisiGUID = MyRow("KISI_GLOBALID")
            End If
            If Not IsDBNull(MyRow("ONAY_GLOBALID")) Then
                MyOdeme.OnayGUID = MyRow("ONAY_GLOBALID")
            End If
            If Not IsDBNull(MyRow("ODENEN_BEDEL")) Then
                MyOdeme.Tutar = MyRow("ODENEN_BEDEL")
            End If
            If Not IsDBNull(MyRow("ODEME_TARIHI")) Then
                MyOdeme.Tarih = MyRow("ODEME_TARIHI")
            End If
            MyOdeme.Sekli = MyRow("ODEME_SEKLI").ToString
            MyOdeme.Tipi = MyRow("ODEME_TIPI").ToString
            MyOdeme.Kaynak = MyRow("KAYNAK").ToString
            If Not IsDBNull(MyRow("ODEME_DURUMU")) Then
                MyOdeme.Durumu = MyRow("ODEME_DURUMU")
            End If
            MyOdeme.Aciklama = MyRow("ACIKLAMA").ToString
            MyOdemeler.Add(MyOdeme)
        Next
        Return MyOdemeler
    End Function

    Public Function GetBelgeCollection(_DataTable As DataTable) As Collection
        Dim MyBelgeler As New Collection
        For Each MyRow As DataRow In _DataTable.Rows
            Dim MyBelge As New Belge
            If Not IsDBNull(MyRow("ID")) Then
                MyBelge.ID = MyRow("ID")
            End If
            If Not IsDBNull(MyRow("ODEME_ID")) Then
                MyBelge.OdemeID = MyRow("ODEME_ID")
            End If
            MyBelge.Adi = MyRow("ADI").ToString
            MyBelge.Yol = MyRow("YOL").ToString
            MyBelge.Aciklama = MyRow("ACIKLAMA").ToString
            MyBelgeler.Add(MyBelge)
        Next
        Return MyBelgeler
    End Function

    Public Function GetKisiCollection(_DataTable As DataTable, Optional WithOutCode As Boolean = False, Optional WithVaris As Boolean = False) As Collection
        Dim MyKisiler As New Collection
        For Each MyRow As DataRow In _DataTable.Rows
            Dim MyKisi As New Kisi
            If Not IsDBNull(MyRow("ID")) Then
                MyKisi.ID = MyRow("ID")
            End If
            MyKisi.Adi = MyRow("ADI").ToString
            MyKisi.Soyadi = MyRow("SOYADI").ToString
            If Not IsDBNull(MyRow("TC_KIMLIK_NO")) Then
                MyKisi.TCKimlikNo = MyRow("TC_KIMLIK_NO")
            End If
            MyKisi.Baba = MyRow("BABA").ToString
            MyKisi.Adres = MyRow("ADRES").ToString
            MyKisi.Telefon = MyRow("TELEFON").ToString
            MyKisi.Cinsiyet = MyRow("CINSIYET").ToString
            MyKisi.Durumu = MyRow("DURUMU").ToString
            MyKisi.DogumYeri = MyRow("DOGUM_YERI").ToString
            If Not IsDBNull(MyRow("DOGUM_TARIHI")) Then
                MyKisi.DogumTarihi = MyRow("DOGUM_TARIHI")
            End If
            If Not WithOutCode Then
                Dim MyKisiKod As New KisiKod
                'If Not IsDBNull(MyRow("MALIK_TIPI")) Then
                '    MyKisiKod.MalikTipi = MyRow("MALIK_TIPI")
                'End If
                If Not IsDBNull(MyRow("DAVETIYE_TEBLIG_DURUMU")) Then
                    MyKisiKod.DavetiyeTebligDurumu = MyRow("DAVETIYE_TEBLIG_DURUMU")
                End If
                If Not IsDBNull(MyRow("DAVETIYE_ALINMA_DURUMU")) Then
                    MyKisiKod.DavetiyeAlinmaDurumu = MyRow("DAVETIYE_ALINMA_DURUMU")
                End If
                If Not IsDBNull(MyRow("GORUSME_DURUMU")) Then
                    MyKisiKod.GorusmeDurumu = MyRow("GORUSME_DURUMU")
                End If
                If Not IsDBNull(MyRow("GORUSME_NO")) Then
                    MyKisiKod.GorusmeNo = MyRow("GORUSME_NO")
                End If
                If Not IsDBNull(MyRow("GORUSME_TARIHI")) Then
                    MyKisiKod.GorusmeTarihi = MyRow("GORUSME_TARIHI")
                End If
                If Not IsDBNull(MyRow("ANLASMA_DURUMU")) Then
                    MyKisiKod.AnlasmaDurumu = MyRow("ANLASMA_DURUMU")
                End If
                If Not IsDBNull(MyRow("ANLASMA_TARIHI")) Then
                    MyKisiKod.AnlasmaTarihi = MyRow("ANLASMA_TARIHI")
                End If
                If Not IsDBNull(MyRow("ANLASMA_DUSUNCELER")) Then
                    MyKisiKod.AnlasmaDusunceler = MyRow("ANLASMA_DUSUNCELER")
                End If
                If Not IsDBNull(MyRow("TESCIL_DURUMU")) Then
                    MyKisiKod.TescilDurumu = MyRow("TESCIL_DURUMU")
                End If
                MyKisi.Kod = MyKisiKod
            End If

            If WithVaris Then
                MyKisi.Varisler = GetVarisler(MyKisi.ID)
                'Dim MyVaris As New Kisi
                'If Not IsDBNull(MyRow("VARIS")) Then
                '    MyVaris.ID = MyRow("VARIS")
                'End If
                'MyKisi.Varisler.Add(MyVaris)
            End If

            'If Not IsDBNull(MyRow("TIP")) Then
            '    MyKisi.MalikTipi = MyRow("TIP")
            'End If
            'If Not IsDBNull(MyRow("MURIS")) Then
            '    MyKisi.Muris = MyRow("MURIS")
            'End If
            MyKisiler.Add(MyKisi)
        Next
        Return MyKisiler
    End Function

    Public Function GetTakbisParselCollection(_DataTable As DataTable) As Collection
        Dim MyParseller As New Collection
        Dim MyMalikler As New Collection
        Dim MyParsel As New Parsel
        Dim MyMalik As New Kisi
        Dim LastAda As Long = -1
        Dim LastParsel As Long = -1
        For Each MyRow As DataRow In _DataTable.Rows
            If (LastAda = MyRow("ADA").ToString And LastParsel = MyRow("PARSEL").ToString) Then
                MyMalik.Adi = MyRow("ADI").ToString
                MyMalik.Soyadi = MyRow("SOYADI").ToString
                MyMalik.Baba = MyRow("BABA").ToString
                If Not IsDBNull(MyRow("HISSEPAY")) Then
                    MyMalik.HissePay = MyRow("HISSEPAY")
                End If
                If Not IsDBNull(MyRow("HISSEPAYDA")) Then
                    MyMalik.HissePayda = MyRow("HISSEPAYDA")
                End If
                If Not IsDBNull(MyRow("TARIH")) Then
                    MyMalik.TapuTarihi = MyRow("TARIH")
                End If
                MyMalik.Dusunceler = MyRow("DUSUNCELER").ToString
                MyMalikler.Add(MyMalik)
                MyMalik = New Kisi
            Else
                If MyMalikler.Count > 0 Then
                    MyParsel.Malikler = MyMalikler
                    MyParseller.Add(MyParsel)
                    MyMalikler = New Collection
                    MyMalik = New Kisi
                    MyParsel = New Parsel
                End If
                MyParsel.Il = MyRow("IL").ToString
                MyParsel.Ilce = MyRow("ILCE").ToString
                MyParsel.Koy = MyRow("KOY").ToString
                'MyParsel.Mahalle = MyRow("MAHALLE").ToString
                MyParsel.AdaNo = MyRow("ADA").ToString
                MyParsel.ParselNo = MyRow("PARSEL").ToString
                MyParsel.PaftaNo = MyRow("PAFTA").ToString
                MyParsel.Cinsi = MyRow("CINSI").ToString
                MyParsel.Mevki = MyRow("MEVKI").ToString
                MyParsel.Cilt = MyRow("CILT").ToString
                MyParsel.Sayfa = MyRow("SAYFA").ToString
                If Not IsDBNull(MyRow("TAPU_ALANI")) Then
                    MyParsel.TapuAlani = MyRow("TAPU_ALANI")
                End If
                LastAda = MyParsel.AdaNo
                LastParsel = MyParsel.ParselNo
                MyMalik.Adi = MyRow("ADI").ToString
                MyMalik.Soyadi = MyRow("SOYADI").ToString
                MyMalik.Baba = MyRow("BABA").ToString
                MyMalik.Cinsiyet = MyRow("MALIK_CINSIYET").ToString
                If Not IsDBNull(MyRow("TC_KIMLIK_NO")) Then
                    MyMalik.TCKimlikNo = MyRow("TC_KIMLIK_NO")
                End If
                If Not IsDBNull(MyRow("HISSEPAY")) Then
                    MyMalik.HissePay = MyRow("HISSEPAY")
                End If
                If Not IsDBNull(MyRow("HISSEPAYDA")) Then
                    MyMalik.HissePayda = MyRow("HISSEPAYDA")
                End If
                If Not IsDBNull(MyRow("TARIH")) Then
                    MyMalik.TapuTarihi = MyRow("TARIH")
                End If
                MyMalik.Dusunceler = MyRow("DUSUNCELER").ToString
                MyMalikler.Add(MyMalik)
                MyMalik = New Kisi
            End If
        Next
        MyParsel.Malikler = MyMalikler
        MyParseller.Add(MyParsel)
        Return MyParseller
    End Function

#End Region

#Region "Get Procedures"

    Public Function GetProje() As Proje
        Dim MyObject As New Proje
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetProje()
                Case Connections.SqlConnection
                    MyObject = MySQL.GetProje()
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetProje()
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetProje(ProjeGUID As String) As Proje
        Dim MyObject As New Proje
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetProje(ProjeGUID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetProje(ProjeGUID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetProje(ProjeGUID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetProjeDetay(ProjeGUID As String) As ProjeDetay
        Dim MyObject As New ProjeDetay
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetProjeDetay(ProjeGUID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetProjeDetay(ProjeGUID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetProjeDetay(ProjeGUID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetParsel(ParselID As Long) As Parsel
        Dim MyObject As New Parsel
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetParsel(ParselID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetParsel(ParselID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetParsel(ParselID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetParsel(ParselGUID As String) As Parsel
        Dim MyObject As New Parsel
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetParsel(ParselGUID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetParsel(ParselGUID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetParsel(ParselGUID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetParselKod(ParselID As Long) As ParselKod
        Dim MyObject As New ParselKod
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetParselKod(ParselID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetParselKod(ParselID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetParselKod(ParselID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetParselKod(ParselGUID As String) As ParselKod
        Dim MyObject As New ParselKod
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetParselKod(ParselGUID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetParselKod(ParselGUID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetParselKod(ParselGUID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function


    Public Function GetParselDetay(ParselID As Long) As ParselDetay
        Dim MyObject As New ParselDetay
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetParselDetay(ParselID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetParselDetay(ParselID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetParselDetay(ParselID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetEmsaller(ParselID As Long) As Collection
        Dim MyObject As New Collection
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetEmsaller(ParselID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetEmsaller(ParselID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetEmsaller(ParselID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetKisi(KisiID As Long) As Kisi
        Dim MyObject As New Kisi
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetKisi(KisiID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetKisi(KisiID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetKisi(KisiID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetKisi(KisiGUID As String) As Kisi
        Dim MyObject As New Kisi
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetKisi(KisiGUID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetKisi(KisiGUID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetKisi(KisiGUID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    'Public Function GetKisi(KisiID As Long, MulkiyetID As Long) As Kisi
    '    Dim MyObject As New Kisi
    '    Try
    '        Select Case ConnectionInfo.ConnectionType
    '            Case Connections.OleDbConnection
    '                MyObject = MyOle.GetKisi(KisiID, MulkiyetID)
    '           Case Connections.SqlConnection
    '                MyObject = MySQL.GetKisi(KisiID, MulkiyetID)
    '           Case Connections.PgSqlConnection
    '                MyObject = MyPgSQL.GetKisi(KisiID, MulkiyetID)
    '        End Select
    '    Catch ex As Exception
    '        'MyObject = Nothing
    '    End Try
    '    Return MyObject
    'End Function

    Public Function GetKisi(TCKimlikNo As Double) As Kisi
        Dim MyObject As New Kisi
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetKisi(TCKimlikNo)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetKisi(TCKimlikNo)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetKisi(TCKimlikNo)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetKisiKod(KisiID As Long) As KisiKod
        Dim MyObject As New KisiKod
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetKisiKod(KisiID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetKisiKod(KisiID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetKisiKod(KisiID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetVarisler(KisiID As Long) As Collection
        Dim MyObject As New Collection
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetVarisler(KisiID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetVarisler(KisiID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetVarisler(KisiID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetMurisler(KisiID As Long) As Collection
        Dim MyObject As New Collection
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetMurisler(KisiID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetMurisler(KisiID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetMurisler(KisiID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetKamu(KamuID As Long) As Parsel
        Dim MyObject As New Parsel
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetKamu(KamuID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetKamu(KamuID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetKamu(KamuID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetMustemilat(MustemilatID As Long) As Mustemilat
        Dim MyObject As New Mustemilat
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetMustemilat(MustemilatID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetMustemilat(MustemilatID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetMustemilat(MustemilatID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetMustemilatlar(ParselID As Long, SahipID As Long) As Collection
        Dim MyObject As New Collection
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetMustemilatlar(ParselID, SahipID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetMustemilatlar(ParselID, SahipID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetMustemilatlar(ParselID, SahipID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetMevsimlik(MevsimlikID As Long) As Mevsimlik
        Dim MyObject As New Mevsimlik
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetMevsimlik(MevsimlikID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetMevsimlik(MevsimlikID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetMevsimlik(MevsimlikID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetMevsimlikler(ParselID As Long, SahipID As Long) As Collection
        Dim MyObject As New Collection
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetMevsimlikler(ParselID, SahipID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetMevsimlikler(ParselID, SahipID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetMevsimlikler(ParselID, SahipID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetDavaAcele(DavaAceleID As Long) As DavaAcele
        Dim MyObject As New DavaAcele
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetDavaAcele(DavaAceleID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetDavaAcele(DavaAceleID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetDavaAcele(DavaAceleID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetDavaTescil(DavaTescilID As Long) As DavaTescil
        Dim MyObject As New DavaTescil
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetDavaTescil(DavaTescilID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetDavaTescil(DavaTescilID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetDavaTescil(DavaTescilID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetOdeme(OdemeID As Long) As Odeme
        Dim MyObject As New Odeme
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetOdeme(OdemeID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetOdeme(OdemeID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetOdeme(OdemeID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetParselID(_Parsel As Parsel) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetParselID(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetParselID(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetParselID(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetKisiID(_Kisi As Kisi) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetKisiID(_Kisi)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetKisiID(_Kisi)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetKisiID(_Kisi)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetParselGUID(_Parsel As Parsel) As String
        Dim MyObject As String = String.Empty
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetParselGUID(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetParselID(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetParselID(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetKisiGUID(_Kisi As Kisi) As String
        Dim MyObject As String = String.Empty
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetKisiGUID(_Kisi)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetKisiID(_Kisi)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetKisiID(_Kisi)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetProjeGUID(_Proje As Proje) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetProjeGUID(_Proje)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetProjeGUID(_Proje)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetProjeGUID(_Proje)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetKamuID(_Parsel As Parsel) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetKamuID(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetKamuID(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetKamuID(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetDavaAceleID(_Parsel As Parsel) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetDavaAceleID(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetDavaAceleID(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetDavaAceleID(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetDavaTescilID(_Parsel As Parsel) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetDavaTescilID(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetDavaTescilID(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetDavaTescilID(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetMustemilatOdemeID(_Mustemilat As Mustemilat) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetMustemilatOdemeID(_Mustemilat)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetMustemilatOdemeID(_Mustemilat)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetMustemilatOdemeID(_Mustemilat)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetMevsimlikOdemeID(_Mevsimlik As Mevsimlik) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetMevsimlikOdemeID(_Mevsimlik)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetMevsimlikOdemeID(_Mevsimlik)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetMevsimlikOdemeID(_Mevsimlik)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetUser(_Connection As ConnectionInfo, _User As User) As User
        Dim MyObject As New User
        Try
            Select Case _Connection.ConnectionType
                Case Connections.OleDbConnection
                    'MyObject = MyOle.GetUser(_Connection, _User)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetUser(_User)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetUser(_User)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetMulkiyet(KisiID As Long, MulkiyetID As Long) As Kisi
        Dim MyObject As New Kisi
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetMulkiyet(KisiID, MulkiyetID)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetMulkiyet(KisiID, MulkiyetID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetMulkiyet(KisiID, MulkiyetID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function GetMulkiyet(KisiID As Long, MulkiyetID As Long, Optional ByVal GetOption As Boolean = True) As Kisi
        Dim MyObject As New Kisi
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.GetMulkiyet(KisiID, MulkiyetID, GetOption)
                Case Connections.SqlConnection
                    MyObject = MySQL.GetMulkiyet(KisiID, MulkiyetID, GetOption)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.GetMulkiyet(KisiID, MulkiyetID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

#End Region

#Region "Add Procedures"

    Public Function AddParsel(_Parsel As Parsel) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddParsel(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddParsel(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddParsel(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddParselKod(_Parsel As Parsel) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddParselKod(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddParselKod(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddParselKod(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddParselDetay(_Parsel As Parsel) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddParselDetay(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddParselDetay(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddParselDetay(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddEmsal(_Parsel As Parsel, _Emsal As Parsel) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddEmsal(_Parsel, _Emsal)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddEmsal(_Parsel, _Emsal)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddEmsal(_Parsel, _Emsal)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddKisi(_Kisi As Kisi) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddKisi(_Kisi)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddKisi(_Kisi)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddKisi(_Kisi)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddKisiKod(_Kisi As Kisi) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddKisiKod(_Kisi)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddKisiKod(_Kisi)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddKisiKod(_Kisi)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    'Public Function AddKisiBanka(_Kisi As Kisi) As Long
    '    Dim MyObject As Long
    '    Try
    '        Select Case ConnectionInfo.ConnectionType
    '            Case Connections.OleDbConnection
    '                MyObject = MyOle.AddKisiBanka(_Kisi)
    '           Case Connections.SqlConnection
    '                MyObject = MySQL.AddKisiBanka(_Kisi)
    '           Case Connections.PgSqlConnection
    '                MyObject = MyPgSQL.AddKisiBanka(_Kisi)
    '        End Select
    '    Catch ex As Exception
    '        'MyObject = Nothing
    '    End Try
    '    Return MyObject
    'End Function

    Public Function AddKamu(_Parsel As Parsel) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddKamu(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddKamu(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddKamu(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddMulkiyet(_Parsel As Parsel) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddMulkiyet(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddMulkiyet(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddMulkiyet(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddMulkiyet(_Kisi As Kisi) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddMulkiyet(_Kisi)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddMulkiyet(_Kisi)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddMulkiyet(_Kisi)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddMustemilat(_Mustemilat As Mustemilat) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddMustemilat(_Mustemilat)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddMustemilat(_Mustemilat)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddMustemilat(_Mustemilat)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddMevsimlik(_Mevsimlik As Mevsimlik) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddMevsimlik(_Mevsimlik)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddMevsimlik(_Mevsimlik)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddMevsimlik(_Mevsimlik)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddDavaTescil(_DavaTescil As DavaTescil) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddDavaTescil(_DavaTescil)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddDavaTescil(_DavaTescil)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddDavaTescil(_DavaTescil)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddDavaAcele(_DavaAcele As DavaAcele) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddDavaAcele(_DavaAcele)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddDavaAcele(_DavaAcele)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddDavaAcele(_DavaAcele)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddDavali(_Dava As DavaTescil, _Kisi As Kisi) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    'MyObject = MyOle.AddDavali(_Dava, _Kisi)
                Case Connections.SqlConnection
                    'MyObject = MySQL.AddDavali(_Dava, _Kisi)
                Case Connections.PgSqlConnection
                    'MyObject = MyPgSQL.AddDavali(_Dava, _Kisi)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddVaris(_Muris As Kisi, _Varis As Kisi) As Long
        Dim MyObject As Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddVaris(_Muris, _Varis)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddVaris(_Muris, _Varis)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddVaris(_Muris, _Varis)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddOdeme(_Odeme As Odeme) As Long
        Dim MyObject As New Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddOdeme(_Odeme)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddOdeme(_Odeme)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddOdeme(_Odeme)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddOdemeBelge(_Belge As Belge) As Long
        Dim MyObject As New Long
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.AddOdemeBelge(_Belge)
                Case Connections.SqlConnection
                    MyObject = MySQL.AddOdemeBelge(_Belge)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.AddOdemeBelge(_Belge)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function AddLog(_Log As Log) As Long
        Dim MyObject As New Long
        Try
            If LogTut Then
                'Select Case ConnectionInfo.ConnectionType   'hiç bir şey açmadan seçeneklere girince kayıt anında hata veriyor.
                '    Case Connections.OleDbConnection
                '        'MyObject = MyOle.AddLog(_Log)
                '   Case Connections.SqlConnection
                MyObject = MySQL.AddLog(_Log)
                'End Select
            End If
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject

    End Function

#End Region

#Region "Update Procedures"
    Public Function UpdateKamu(_Parsel As Parsel) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateKamu(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateKamu(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateKamu(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function UpdateProject(_Proje As Proje) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateProject(_Proje)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateProject(_Proje)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateProject(_Proje)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function UpdateKisi(_Kisi As Kisi) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateKisi(_Kisi)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateKisi(_Kisi)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateKisi(_Kisi)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function UpdateKisiKod(_Kisi As Kisi) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateKisiKod(_Kisi)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateKisiKod(_Kisi)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateKisiKod(_Kisi)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    'Public Function UpdateKisiBanka(_Kisi As Kisi, _BankaID As Long) As Boolean
    '    Dim MyObject As Boolean
    '    Try
    '        Select Case ConnectionInfo.ConnectionType
    '            Case Connections.OleDbConnection
    '                MyObject = MyOle.UpdateKisiBanka(_Kisi, _BankaID)
    '           Case Connections.SqlConnection
    '                MyObject = MySQL.UpdateKisiBanka(_Kisi, _BankaID)
    '           Case Connections.PgSqlConnection
    '                MyObject = MyPgSQL.UpdateKisiBanka(_Kisi, _BankaID)
    '        End Select
    '    Catch ex As Exception
    '        'MyObject = Nothing
    '    End Try
    '    Return MyObject
    'End Function

    Public Function UpdateParsel(_Parsel As Parsel) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateParsel(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateParsel(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateParsel(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function UpdateParselKod(_Parsel As Parsel) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateParselKod(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateParselKod(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateParselKod(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function UpdateParselDetay(_Parsel As Parsel) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateParselDetay(_Parsel)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateParselDetay(_Parsel)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateParselDetay(_Parsel)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function UpdateMulkiyet(_Kisi As Kisi) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateMulkiyet(_Kisi)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateMulkiyet(_Kisi)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateMulkiyet(_Kisi)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function UpdateMustemilat(_Mustemilat As Mustemilat) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateMustemilat(_Mustemilat)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateMustemilat(_Mustemilat)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateMustemilat(_Mustemilat)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function UpdateMevsimlik(_Mevsimlik As Mevsimlik) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateMevsimlik(_Mevsimlik)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateMevsimlik(_Mevsimlik)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateMevsimlik(_Mevsimlik)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function UpdateDavaTescil(_DavaTescil As DavaTescil) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateDavaTescil(_DavaTescil)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateDavaTescil(_DavaTescil)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateDavaTescil(_DavaTescil)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function UpdateDavaAcele(_DavaAcele As DavaAcele) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateDavaAcele(_DavaAcele)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateDavaAcele(_DavaAcele)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateDavaAcele(_DavaAcele)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function UpdateOdeme(_Odeme As Odeme) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateOdeme(_Odeme)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateOdeme(_Odeme)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateOdeme(_Odeme)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function UpdateOdeme(_Odeme As Odeme, _OnayID As Integer) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.UpdateOdeme(_Odeme, _OnayID)
                Case Connections.SqlConnection
                    MyObject = MySQL.UpdateOdeme(_Odeme, _OnayID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.UpdateOdeme(_Odeme, _OnayID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

#End Region

#Region "Delete Procedures"
    Public Function DeleteParsel(_ID As Long) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.DeleteParsel(_ID)
                Case Connections.SqlConnection
                    MyObject = MySQL.DeleteParsel(_ID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.DeleteParsel(_ID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function DeleteKisi(_ID As Long) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.DeleteKisi(_ID)
                Case Connections.SqlConnection
                    MyObject = MySQL.DeleteKisi(_ID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.DeleteKisi(_ID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function DeleteMustemilat(_ID As Long) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.DeleteMustemilat(_ID)
                Case Connections.SqlConnection
                    MyObject = MySQL.DeleteMustemilat(_ID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.DeleteMustemilat(_ID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function DeleteMevsimlik(_ID As Long) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.DeleteMevsimlik(_ID)
                Case Connections.SqlConnection
                    MyObject = MySQL.DeleteMevsimlik(_ID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.DeleteMevsimlik(_ID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function DeleteMiras(_MurisID As Long, _VarisID As Long) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.DeleteMiras(_MurisID, _VarisID)
                Case Connections.SqlConnection
                    MyObject = MySQL.DeleteMiras(_MurisID, _VarisID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.DeleteMiras(_MurisID, _VarisID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function DeleteEmsal(_ParselID As Long, _EmsalID As Long) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.DeleteEmsal(_ParselID, _EmsalID)
                Case Connections.SqlConnection
                    MyObject = MySQL.DeleteEmsal(_ParselID, _EmsalID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.DeleteEmsal(_ParselID, _EmsalID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function DeleteMalik(_ParselID As Long, _MalikID As Long) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.DeleteMalik(_ParselID, _MalikID)
                Case Connections.SqlConnection
                    MyObject = MySQL.DeleteMalik(_ParselID, _MalikID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.DeleteMalik(_ParselID, _MalikID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function DeleteMalik(_MulkiyetID As Long) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.DeleteMalik(_MulkiyetID)
                Case Connections.SqlConnection
                    MyObject = MySQL.DeleteMalik(_MulkiyetID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.DeleteMalik(_MulkiyetID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function DeleteOdeme(_ID As Long) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.DeleteOdeme(_ID)
                Case Connections.SqlConnection
                    MyObject = MySQL.DeleteOdeme(_ID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.DeleteOdeme(_ID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function DeleteOdemeBelge(_ID As Long) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.DeleteOdemeBelge(_ID)
                Case Connections.SqlConnection
                    MyObject = MySQL.DeleteOdemeBelge(_ID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.DeleteOdemeBelge(_ID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function DeleteDavaTescil(_ID As Long) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.DeleteDavaTescil(_ID)
                Case Connections.SqlConnection
                    MyObject = MySQL.DeleteDavaTescil(_ID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.DeleteDavaTescil(_ID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

    Public Function DeleteDavaAcele(_ID As Long) As Boolean
        Dim MyObject As Boolean
        Try
            Select Case ConnectionInfo.ConnectionType
                Case Connections.OleDbConnection
                    MyObject = MyOle.DeleteDavaAcele(_ID)
                Case Connections.SqlConnection
                    MyObject = MySQL.DeleteDavaAcele(_ID)
                Case Connections.PgSqlConnection
                    MyObject = MyPgSQL.DeleteDavaAcele(_ID)
            End Select
        Catch ex As Exception
            'MyObject = Nothing
        End Try
        Return MyObject
    End Function

#End Region

End Class
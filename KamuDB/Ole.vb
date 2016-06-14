Imports Kamu.Objects
Imports System.Data.OleDb

Public Class Ole
    Public MyConnectionInfo As New ConnectionInfo

    Sub New()

    End Sub

    Sub New(ByVal _Connection As ConnectionInfo)
        MyConnectionInfo = _Connection
    End Sub

    Public Function GetDataTable(ByVal _SQLCommand As String) As DataTable
        Dim MyTable As New DataTable
        MyTable.Locale = System.Globalization.CultureInfo.InvariantCulture
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDb.OleDbCommand = connection.CreateCommand()
                command.CommandText = _SQLCommand
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As New OleDb.OleDbDataAdapter
                adapter.SelectCommand = command

                adapter.Fill(MyTable)

                adapter = Nothing
                command = Nothing
            Catch ex As Exception

            End Try
        End Using
        Return MyTable
    End Function

    Public Function NewProject(_FileName As String, _Project As Proje, _KamuVeriXMLFileName As String) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyCreate As New Create
            MyCreate.CreateKamuDBFromScratch(_FileName, _KamuVeriXMLFileName, "YENI_VERITABANI")

            Dim MyNewConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyNewConnection.Open()

            MyRowID = FillProjectTable("SELECT * FROM PROJE", MyNewConnection, _Project)
            FillTipTable("SELECT * FROM TIP_ANLASMA_DURUMU", "TIP_ANLASMA_DURUMU", "ANLASMA_DURUMU", MyNewConnection, _KamuVeriXMLFileName)
            FillTipTable("SELECT * FROM TIP_DAVET_ALINMA", "TIP_DAVET_ALINMA", "DAVET_ALINMA", MyNewConnection, _KamuVeriXMLFileName)
            FillTipTable("SELECT * FROM TIP_DAVETIYE_TEBLIG", "TIP_DAVETIYE_TEBLIG", "DAVETIYE_TEBLIG", MyNewConnection, _KamuVeriXMLFileName)
            FillTipTable("SELECT * FROM TIP_EDINIM_DURUMU", "TIP_EDINIM_DURUMU", "EDINIM_DURUMU", MyNewConnection, _KamuVeriXMLFileName)
            FillTipTable("SELECT * FROM TIP_GORUSME_DURUMU", "TIP_GORUSME_DURUMU", "GORUSME_DURUMU", MyNewConnection, _KamuVeriXMLFileName)
            FillTipTable("SELECT * FROM TIP_ISTIMLAK_SERHI", "TIP_ISTIMLAK_SERHI", "ISTIMLAK_SERHI", MyNewConnection, _KamuVeriXMLFileName)
            FillTipTable("SELECT * FROM TIP_ISTIMLAK_TURU", "TIP_ISTIMLAK_TURU", "ISTIMLAK_TURU", MyNewConnection, _KamuVeriXMLFileName)
            FillTipTable("SELECT * FROM TIP_KADASTRAL_DURUM", "TIP_KADASTRAL_DURUM", "KADASTRAL_DURUM", MyNewConnection, _KamuVeriXMLFileName)
            FillTipTable("SELECT * FROM TIP_MALIK", "TIP_MALIK", "MALIK", MyNewConnection, _KamuVeriXMLFileName)
            FillTipTable("SELECT * FROM TIP_TESCIL_DURUMU", "TIP_TESCIL_DURUMU", "TESCIL_DURUMU", MyNewConnection, _KamuVeriXMLFileName)

            MyNewConnection.Close()
            MyNewConnection = Nothing
            MyCreate = Nothing
        Catch ex As Exception
            MyRowID = -1
            'MsgBox(ex.Message)
        End Try
        Return MyRowID
    End Function

    Private Function FillProjectTable(MyQueryString As String, MyKamuConnection As OleDbConnection, _Project As Proje) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))
            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow = MyTable.NewRow()
            MyRow("KOD") = _Project.Kod
            MyRow("AD") = _Project.Ad
            MyRow("PROJE_NOTLARI") = _Project.ProjeNotlari
            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))

            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyCommandBuilder = Nothing
            MyTable = Nothing
            MyDataAdapter = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Private Sub FillTipTable(_QueryString As String, _TableName As String, _ColumnName As String, _Connection As OleDbConnection, _KamuVeriXMLFileName As String)
        Try
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(_QueryString, _Connection))
            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim KamuVeri As New DataSet
            KamuVeri.ReadXml(_KamuVeriXMLFileName)
            Dim KamuTable As DataTable = KamuVeri.Tables(_TableName)
            For Each MyTipRow As DataRow In KamuTable.Rows
                Dim MyRow As DataRow = MyTable.NewRow()
                MyRow("ID") = Val(MyTipRow("ID"))
                MyRow(_ColumnName) = MyTipRow(_ColumnName).ToString
                MyTable.Rows.Add(MyRow)
                MyRow = Nothing
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyCommandBuilder = Nothing
            KamuTable = Nothing
            KamuVeri = Nothing
            MyTable = Nothing
            MyDataAdapter = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function CompactDataBase(_SourceConnectionString As String, _TargetConnectionString As String) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim jro As New JRO.JetEngine
            jro.CompactDatabase(_SourceConnectionString, _TargetConnectionString)
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Public Function CreateProjectList() As Collection
        Dim MyProjectList As New Collection()
        Try
            Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
                Dim command As New OleDbCommand("SELECT ID, AD FROM PROJE ORDER BY ID", connection)
                If Not connection.State = ConnectionState.Open Then connection.Open()
                Dim reader As OleDbDataReader = command.ExecuteReader()
                While reader.Read()
                    MyProjectList.Add(New Proje(CLng(reader("ID")), reader("AD").ToString))
                End While
                connection.Close()
                reader.Close()
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
        Return MyProjectList
    End Function

    Public Function CreateComboList(strTableName As String, strColumnName As String) As Collection
        Dim strSQL As String = "SELECT ID, " + strColumnName + " FROM " + strTableName + " ORDER BY ID"
        Dim MyList As New Collection()
        Try
            Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
                Dim command As New OleDbCommand(strSQL, connection)
                If Not connection.State = ConnectionState.Open Then connection.Open()
                Dim reader As OleDbDataReader = command.ExecuteReader()
                MyList.Add(New LookupObject(0, ""))
                While reader.Read()
                    If reader(strColumnName).ToString.Trim <> "" Then
                        MyList.Add(New LookupObject(CLng(reader("ID")), reader(strColumnName).ToString))
                    End If

                End While
                connection.Close()
                reader.Close()
            End Using
        Catch ex As Exception
            Return Nothing
        End Try
        Return MyList
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
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "UPDATE KISI_ID=" + _AktifKisiID.ToString + " FROM " + _TableName + " WHERE KISI_ID=" + _PasifKisiID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()
                adapter.SelectCommand = command

                Dim table As New DataTable
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                adapter.Fill(table)

                table = Nothing
                adapter = Nothing
                command = Nothing
                MyStatus = True
            Catch ex As Exception
                MyStatus = False
            End Try
        End Using
        Return MyStatus
    End Function

#Region "Get Procedures V5"

    Public Function GetParsellerCollectionV5(_DataTable As DataTable) As Collection
        Dim MyParseller As New Collection
        Dim MyMalikler As New Collection
        Dim MyParsel As New Parsel
        Dim MyMalik As New Kisi
        Dim LastAda As String = "-1"
        Dim LastParsel As String = "-1"
        Try
            For Each MyRow As DataRow In _DataTable.Rows
                If (LastAda = MyRow("ADA").ToString And LastParsel = MyRow("PARSEL").ToString) Then
                    If Not IsDBNull(MyRow("PARSEL_MALIK_TIPI")) Then
                        If MyRow("PARSEL_MALIK_TIPI") = 1 Then
                            MyMalik = New Kisi(MyRow("MALIK").ToString.Trim)
                            'MyMalik.MalikTipi = 1
                        Else
                            MyMalik = New Kisi(MyRow("MALIK").ToString.Trim, String.Empty)
                            'MyMalik.MalikTipi = MyRow("PARSEL_MALIK_TIPI")
                        End If
                    Else
                        MyMalik = New Kisi(MyRow("MALIK").ToString.Trim, String.Empty)
                    End If
                    MyMalik.Baba = MyRow("BABA").ToString.Trim
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
                    MyMalik.Dusunceler = MyRow("DUSUNCELER").ToString.Trim


                    Dim MyMalikKod As New KisiKod
                    'If Not IsDBNull(MyRow("PARSEL_MALIK_TIPI")) Then
                    '    MyMalikKod.MalikTipi = MyRow("PARSEL_MALIK_TIPI")
                    'End If
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
                    If Not IsDBNull(MyRow("PROJE_ID")) Then
                        If IsNumeric(MyRow("PROJE_ID")) Then
                            MyParsel.ProjeID = MyRow("PROJE_ID")
                        Else
                            MyParsel.ProjeID = 1 'vt5 de proje id alfasayısal olursa treeview sorun çıkıyor.
                        End If
                    Else
                        MyParsel.ProjeID = 1
                    End If
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

                    Dim MyParselKod As New ParselKod
                    MyParselKod.Kod = MyRow("KOD").ToString.Trim

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
                    'If Not IsDBNull(MyRow("ISTIMLAK_DISI_DURUMU")) Then
                    '    MyParselKod.IstimlakDisi = MyRow("ISTIMLAK_DISI_DURUMU")
                    'End If
                    If Not IsDBNull(MyRow("DEVIR_DURUMU")) Then
                        MyParselKod.DevirDurumu = MyRow("DEVIR_DURUMU")
                    End If
                    If Not IsDBNull(MyRow("PARSEL_ALINMA_DURUMU")) Then
                        MyParselKod.EdinimDurumu = MyRow("PARSEL_ALINMA_DURUMU")
                    End If
                    MyParsel.Kod = MyParselKod

                    LastAda = MyParsel.AdaNo
                    LastParsel = MyParsel.ParselNo
                    If Not IsDBNull(MyRow("PARSEL_MALIK_TIPI")) Then
                        If MyRow("PARSEL_MALIK_TIPI") = 1 Then
                            MyMalik = New Kisi(MyRow("MALIK").ToString.Trim)
                            'MyMalik.MalikTipi = 1
                        Else
                            MyMalik = New Kisi(MyRow("MALIK").ToString.Trim, String.Empty)
                            'MyMalik.MalikTipi = MyRow("PARSEL_MALIK_TIPI")
                        End If
                    Else
                        MyMalik = New Kisi(MyRow("MALIK").ToString.Trim, String.Empty)
                    End If
                    MyMalik.Baba = MyRow("BABA").ToString.Trim
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
                    MyMalik.Dusunceler = MyRow("DUSUNCELER").ToString.Trim

                    Dim MyMalikKod As New KisiKod
                    'If Not IsDBNull(MyRow("PARSEL_MALIK_TIPI")) Then
                    '    MyMalikKod.MalikTipi = MyRow("PARSEL_MALIK_TIPI")
                    'End If
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

    Public Function GetMustemilatCollectionV5(_DataTable As DataTable) As Collection
        Dim MyMustemilatlar As New Collection
        Dim MyMalik As New Kisi
        For Each MyRow As DataRow In _DataTable.Rows
            Dim MyParsel As New Parsel
            MyParsel.Il = MyRow("IL").ToString
            MyParsel.Ilce = MyRow("ILCE").ToString
            MyParsel.Koy = MyRow("KOY").ToString
            MyParsel.Mahalle = MyRow("MAHALLE").ToString
            MyParsel.AdaNo = MyRow("ADA").ToString
            MyParsel.ParselNo = MyRow("PARSEL").ToString

            Dim MyKisi As New Kisi(MyRow("SAHIP").ToString, MyRow("BABA").ToString, 0)

            Dim MyMustemilat As New Mustemilat
            MyMustemilat.ParselID = GetParselID(MyParsel)
            MyMustemilat.SahipID = GetKisiID(MyKisi)
            'Kisi ID için parsel bilgiside dikkate alınmalıdır. bu haliyle yanlış maliklere id bağlanacaktır.
            MyKisi = Nothing
            MyParsel = Nothing

            If Not IsDBNull(MyRow("BIRIM")) Then
                MyMustemilat.Adet = MyRow("BIRIM").ToString
            End If
            MyMustemilat.Tanim = MyRow("TANIM").ToString
            If Not IsDBNull(MyRow("FIYAT")) Then
                MyMustemilat.Fiyat = MyRow("FIYAT").ToString
            End If
            Select Case MyRow("K_M").ToString
                Case "K"
                    MyMustemilat.Malik = False
                Case Else
                    MyMustemilat.Malik = True
            End Select
            MyMustemilat.Pay = 1
            MyMustemilat.Payda = 1
            MyMustemilat.OdemeID = 0

            MyMustemilatlar.Add(MyMustemilat)
        Next

        Return MyMustemilatlar
    End Function

    Public Function GetMevsimlikCollectionV5(_DataTable As DataTable) As Collection
        Dim MyMevsimlikler As New Collection
        Dim MyMalik As New Kisi
        Dim LastAda As Long = -1
        Dim LastParsel As Long = -1
        For Each MyRow As DataRow In _DataTable.Rows
            Dim MyParsel As New Parsel
            MyParsel.Il = MyRow("IL").ToString
            MyParsel.Ilce = MyRow("ILCE").ToString
            MyParsel.Koy = MyRow("KOY").ToString
            MyParsel.Mahalle = MyRow("MAHALLE").ToString
            MyParsel.AdaNo = MyRow("ADA").ToString
            MyParsel.ParselNo = MyRow("PARSEL").ToString

            Dim MyKisi As New Kisi(MyRow("SAHIP").ToString, MyRow("BABA").ToString, 0)

            Dim MyMevsimlik As New Mevsimlik
            MyMevsimlik.ParselID = GetParselID(MyParsel)
            MyMevsimlik.SahipID = GetKisiID(MyKisi)
            'Kisi ID için parsel bilgiside dikkate alınmalıdır. bu haliyle yanlış maliklere id bağlanacaktır.
            MyKisi = Nothing
            MyParsel = Nothing

            If Not IsDBNull(MyRow("HASAR_ALAN")) Then
                MyMevsimlik.Alan = MyRow("HASAR_ALAN").ToString
            End If
            MyMevsimlik.Tanim = MyRow("TANIM").ToString
            If Not IsDBNull(MyRow("HASAR_BEDEL")) Then
                MyMevsimlik.Bedel = MyRow("HASAR_BEDEL").ToString
            End If
            Select Case MyRow("MK").ToString
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

            MyMevsimlik.OdemeID = 0

            MyMevsimlikler.Add(MyMevsimlik)
        Next

        Return MyMevsimlikler
    End Function

#End Region

#Region "Get Procedures"

    Public Function GetProje() As Proje
        Dim MyProje As New Proje
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM PROJE"
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyProje.ID = dataReader("ID")
                MyProje.Kod = dataReader("KOD").ToString
                MyProje.Ad = dataReader("AD").ToString
                MyProje.ProjeNotlari = dataReader("PROJE_NOTLARI").ToString
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            'MyProje = Nothing
        End Try
        Return MyProje
    End Function

    Public Function GetProje(ProjeID As Long) As Proje
        Dim MyProje As New Proje
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM PROJE WHERE ID=" & ProjeID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyProje.ID = ProjeID
                MyProje.Kod = dataReader("KOD").ToString
                MyProje.Ad = dataReader("AD").ToString
                MyProje.ProjeNotlari = dataReader("PROJE_NOTLARI").ToString
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            'MyProje = Nothing
        End Try
        Return MyProje
    End Function

    Public Function GetProjeDetay(ProjeID As Long) As ProjeDetay
        Dim MyProjeDetay As New ProjeDetay
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        Try
            MyProjeDetay.ID = ProjeID
            command.CommandText = "SELECT COUNT(*) AS PARSEL_SAYISI FROM PARSEL WHERE PROJE_ID=" & ProjeID.ToString
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyProjeDetay.ParselSayisi = dataReader("PARSEL_SAYISI")
            Loop
            dataReader.Close()
            dataReader = Nothing
            'command = Nothing
            connection.Close()

            command.CommandText = "SELECT COUNT(*) AS MALIK_SAYISI FROM KISI"
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader1 As OleDbDataReader = command.ExecuteReader()
            Do While dataReader1.Read()
                MyProjeDetay.MalikSayisi = dataReader1("MALIK_SAYISI")
            Loop
            dataReader1.Close()
            dataReader1 = Nothing
            'command = Nothing
            connection.Close()

            command.CommandText = "SELECT DISTINCTROW IL FROM PARSEL WHERE PROJE_ID=" & ProjeID.ToString & " GROUP BY IL"
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader2 As OleDbDataReader = command.ExecuteReader()
            Do While dataReader2.Read()
                MyProjeDetay.IlSayisi += 1
            Loop
            dataReader2.Close()
            dataReader2 = Nothing
            'command = Nothing
            connection.Close()

            command.CommandText = "SELECT DISTINCTROW IL, ILCE FROM PARSEL WHERE PROJE_ID=" & ProjeID.ToString & " GROUP BY IL, ILCE"
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader3 As OleDbDataReader = command.ExecuteReader()
            Do While dataReader3.Read()
                MyProjeDetay.IlceSayisi += 1
            Loop
            dataReader3.Close()
            dataReader3 = Nothing
            'command = Nothing
            connection.Close()

            command.CommandText = "SELECT IL, ILCE, KOY, MAHALLE FROM PARSEL WHERE PROJE_ID=" & ProjeID.ToString & " GROUP BY IL, ILCE, KOY, MAHALLE"
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader4 As OleDbDataReader = command.ExecuteReader()
            Do While dataReader4.Read()
                MyProjeDetay.KoySayisi += 1
            Loop
            dataReader4.Close()
            dataReader4 = Nothing

            command = Nothing
            connection.Close()

            connection = Nothing
        Catch ex As Exception
            'MyParsel = Nothing
        End Try
        Return MyProjeDetay
    End Function

    Public Function GetParsel(ParselID As Long) As Parsel
        Dim MyParsel As New Parsel
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM PARSEL WHERE ID=" & ParselID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyParsel.ID = ParselID
                MyParsel.ProjeID = dataReader("PROJE_ID")
                MyParsel.Code = dataReader("KOD").ToString
                MyParsel.Il = dataReader("IL").ToString
                MyParsel.Ilce = dataReader("ILCE").ToString
                MyParsel.Koy = dataReader("KOY").ToString
                MyParsel.Mahalle = dataReader("MAHALLE").ToString
                MyParsel.AdaNo = dataReader("ADA").ToString
                MyParsel.ParselNo = dataReader("PARSEL").ToString
                MyParsel.PaftaNo = dataReader("PAFTA").ToString
                MyParsel.Mevki = dataReader("MEVKI").ToString
                MyParsel.Cilt = dataReader("CILT").ToString
                MyParsel.Sayfa = dataReader("SAYFA").ToString
                MyParsel.Cinsi = dataReader("CINSI").ToString
                If Not IsDBNull(dataReader("TAPU_ALANI")) Then
                    MyParsel.TapuAlani = dataReader("TAPU_ALANI")
                End If
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
        Catch ex As Exception
            'MyParsel = Nothing
        End Try

        Dim commandk As OleDbCommand = connection.CreateCommand()
        commandk.CommandText = "SELECT * FROM KAMULASTIRMA WHERE PARSEL_ID=" & ParselID.ToString
        Try
            Dim dataReaderk As OleDbDataReader = commandk.ExecuteReader()
            Do While dataReaderk.Read()
                MyParsel.KamuID = dataReaderk("ID")
                If Not IsDBNull(dataReaderk("MULKIYET_ALAN")) Then
                    MyParsel.MulkiyetAlan = dataReaderk("MULKIYET_ALAN")
                End If
                If Not IsDBNull(dataReaderk("IRTIFAK_ALAN")) Then
                    MyParsel.IrtifakAlan = dataReaderk("IRTIFAK_ALAN")
                End If
                If Not IsDBNull(dataReaderk("GECICI_IRTIFAK_ALAN")) Then
                    MyParsel.GeciciIrtifakAlan = dataReaderk("GECICI_IRTIFAK_ALAN")
                End If
                If Not IsDBNull(dataReaderk("MULKIYET_BEDEL")) Then
                    MyParsel.MulkiyetBedel = dataReaderk("MULKIYET_BEDEL")
                End If
                If Not IsDBNull(dataReaderk("IRTIFAK_BEDEL")) Then
                    MyParsel.IrtifakBedel = dataReaderk("IRTIFAK_BEDEL")
                End If
                If Not IsDBNull(dataReaderk("GECICI_IRTIFAK_BEDEL")) Then
                    MyParsel.GeciciIrtifakBedel = dataReaderk("GECICI_IRTIFAK_BEDEL")
                End If
                MyParsel.AraziVasfi = dataReaderk("ARAZI_VASFI").ToString
                MyParsel.KamulastirmaAmaci = dataReaderk("KAMULASTIRMA_AMACI").ToString
                MyParsel.YayginMunavebeSistemi = dataReaderk("YAYGIN_MUNAVEBE_SISTEMI").ToString
                MyParsel.DegerlemeRaporu = dataReaderk("DEGERLEME_RAPORU").ToString
                If Not IsDBNull(dataReaderk("YILLIK_ORTALAMA_NET_GELIR")) Then
                    MyParsel.YillikOrtalamaNetGelir = dataReaderk("YILLIK_ORTALAMA_NET_GELIR")
                End If
                If Not IsDBNull(dataReaderk("KAPITALIZASYON_FAIZI")) Then
                    MyParsel.KapitalizasyonOrani = dataReaderk("KAPITALIZASYON_FAIZI")
                End If
                If Not IsDBNull(dataReaderk("OBJEKTIF_ARTIS")) Then
                    MyParsel.ObjektifArtis = dataReaderk("OBJEKTIF_ARTIS")
                End If
                If Not IsDBNull(dataReaderk("ART_KISIM_ARTIS")) Then
                    MyParsel.ArtanKisimArtis = dataReaderk("ART_KISIM_ARTIS")
                End If
                If Not IsDBNull(dataReaderk("VERIM_KAYBI")) Then
                    MyParsel.VerimKaybi = dataReaderk("VERIM_KAYBI")
                End If
                If Not IsDBNull(dataReaderk("DEGERLEME_TARIHI")) Then
                    MyParsel.DegerlemeTarihi = dataReaderk("DEGERLEME_TARIHI")
                End If
            Loop
            dataReaderk.Close()
            dataReaderk = Nothing
            commandk = Nothing
        Catch ex As Exception
            'MyParsel = Nothing
        End Try
        connection.Close()
        connection = Nothing
        Return MyParsel
    End Function

    Public Function GetParselKod(ParselID As Long) As ParselKod
        Dim MyParselKod As New ParselKod
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM PARSEL_KOD WHERE PARSEL_ID=" & ParselID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyParselKod.ID = dataReader("ID")
                If Not IsDBNull(dataReader("BOLGE_ID")) Then
                    MyParselKod.BolgeID = dataReader("BOLGE_ID")
                End If
                If Not IsDBNull(dataReader("KADASTRAL_DURUM")) Then
                    MyParselKod.KadastralDurum = dataReader("KADASTRAL_DURUM")
                End If
                If Not IsDBNull(dataReader("MALIK_TIPI")) Then
                    MyParselKod.MalikTipi = dataReader("MALIK_TIPI")
                End If
                If Not IsDBNull(dataReader("ISTIMLAK_TURU")) Then
                    MyParselKod.IstimlakTuru = dataReader("ISTIMLAK_TURU")
                End If
                If Not IsDBNull(dataReader("ISTIMLAK_SERHI")) Then
                    MyParselKod.IstimlakSerhi = dataReader("ISTIMLAK_SERHI")
                End If
                If Not IsDBNull(dataReader("DAVA10_DURUMU")) Then
                    MyParselKod.DavaDurumu10 = dataReader("DAVA10_DURUMU")
                End If
                If Not IsDBNull(dataReader("DAVA27_DURUMU")) Then
                    MyParselKod.DavaDurumu27 = dataReader("DAVA27_DURUMU")
                End If
                If Not IsDBNull(dataReader("EDINIM_DURUMU")) Then
                    MyParselKod.EdinimDurumu = dataReader("EDINIM_DURUMU")
                End If
                If Not IsDBNull(dataReader("ISTIMLAK_DISI")) Then
                    MyParselKod.IstimlakDisi = dataReader("ISTIMLAK_DISI")
                End If
                If Not IsDBNull(dataReader("DEVIR_DURUMU")) Then
                    MyParselKod.DevirDurumu = dataReader("DEVIR_DURUMU")
                End If
                If Not IsDBNull(dataReader("ODEME_DURUMU")) Then
                    MyParselKod.OdemeDurumu = dataReader("ODEME_DURUMU")
                End If
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            'MyParsel = Nothing
        End Try
        Return MyParselKod
    End Function

    Public Function GetParselDetay(ParselID As Long) As ParselDetay
        Dim MyParselDetay As New ParselDetay
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM PARSEL_DETAY WHERE PARSEL_ID=" & ParselID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyParselDetay.ID = dataReader("ID")
                If Not IsDBNull(dataReader("ARSA")) Then
                    MyParselDetay.Arsa = dataReader("ARSA")
                End If
                If Not IsDBNull(dataReader("IMAR_DURUMU")) Then
                    MyParselDetay.ImarDurumu = dataReader("IMAR_DURUMU")
                End If
                If Not IsDBNull(dataReader("DOP_KESINTISI")) Then
                    MyParselDetay.DopKesintisi = dataReader("DOP_KESINTISI")
                End If
                If Not IsDBNull(dataReader("VERGI_DEGERI")) Then
                    MyParselDetay.VergiDegeri = dataReader("VERGI_DEGERI")
                End If
                If Not IsDBNull(dataReader("VERGI_DEGERI_YILI")) Then
                    MyParselDetay.VergiDegeriTarihi = dataReader("VERGI_DEGERI_YILI")
                End If
                If Not IsDBNull(dataReader("KAYIP_ORANI")) Then
                    MyParselDetay.KayipOrani = dataReader("KAYIP_ORANI")
                End If
                If Not IsDBNull(dataReader("FAIZ")) Then
                    MyParselDetay.Faiz = dataReader("FAIZ")
                End If
                If Not IsDBNull(dataReader("YARGITAY_SONUC")) Then
                    MyParselDetay.YargitaySonuc = dataReader("YARGITAY_SONUC")
                End If
                MyParselDetay.DavaAciklama = dataReader("YARGITAY_ACIKLAMA").ToString
                MyParselDetay.DavaEsasNo = dataReader("ESAS_NO").ToString
                MyParselDetay.DavaKararNo = dataReader("KARAR_NO").ToString
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            'MyParsel = Nothing
        End Try
        Return MyParselDetay
    End Function

    Public Function GetEmsaller(ParselID As Long) As Collection
        Dim MyEmsaller As New Collection
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM EMSAL WHERE PARSEL_ID=" & ParselID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                If Not IsDBNull(dataReader("EMSAL_ID")) Then
                    Dim MyParsel As New Parsel
                    MyParsel.ID = dataReader("EMSAL_ID")
                    MyParsel = GetParsel(MyParsel.ID)
                    MyEmsaller.Add(MyParsel)
                End If
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            'MyParsel = Nothing
        End Try
        Return MyEmsaller
    End Function

    Public Function GetKisi(KisiID As Long) As Kisi
        Dim MyKisi As New Kisi
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "SELECT * FROM KISI WHERE ID=" & KisiID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()
                Dim dataReader As OleDbDataReader = command.ExecuteReader()
                Do While dataReader.Read()
                    MyKisi.ID = KisiID
                    If Not IsDBNull(dataReader("TC_KIMLIK_NO")) Then
                        MyKisi.TCKimlikNo = dataReader("TC_KIMLIK_NO")
                    End If
                    MyKisi.Adi = dataReader("ADI").ToString
                    MyKisi.Soyadi = dataReader("SOYADI").ToString
                    MyKisi.Cinsiyet = dataReader("CINSIYET").ToString
                    MyKisi.Baba = dataReader("BABA").ToString
                    MyKisi.Durumu = dataReader("DURUMU").ToString
                    MyKisi.Adres = dataReader("ADRES").ToString
                    MyKisi.Telefon = dataReader("TELEFON").ToString
                    If Not IsDBNull(dataReader("DOGUM_TARIHI")) Then
                        MyKisi.DogumTarihi = dataReader("DOGUM_TARIHI")
                    End If
                    MyKisi.DogumYeri = dataReader("DOGUM_YERI").ToString
                    MyKisi.IBAN = dataReader("IBAN").ToString
                    MyKisi.BankaSubeKodu = dataReader("SUBE_KODU").ToString
                Loop
                dataReader.Close()
                dataReader = Nothing
                command = Nothing
            Catch ex As Exception

            End Try
        End Using
        Return MyKisi
    End Function

    'Public Function GetKisi(KisiID As Long, MulkiyetID As Long) As Kisi
    '    Dim MyKisi As New Kisi
    '    Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
    '        Dim command As OleDbCommand = connection.CreateCommand()
    '        command.CommandText = "SELECT * FROM KISI WHERE ID=" & KisiID.ToString
    '        Try
    '            If Not connection.State = ConnectionState.Open Then connection.Open()
    '            Dim dataReader As OleDbDataReader = command.ExecuteReader()
    '            Do While dataReader.Read()
    '                MyKisi.ID = KisiID
    '                If Not IsDBNull(dataReader("TC_KIMLIK_NO")) Then
    '                    MyKisi.TCKimlikNo = dataReader("TC_KIMLIK_NO")
    '                End If
    '                MyKisi.Adi = dataReader("ADI").ToString
    '                MyKisi.Soyadi = dataReader("SOYADI").ToString
    '                MyKisi.Cinsiyet = dataReader("CINSIYET").ToString
    '                MyKisi.Baba = dataReader("BABA").ToString
    '                MyKisi.Durumu = dataReader("DURUMU").ToString
    '                MyKisi.Adres = dataReader("ADRES").ToString
    '                MyKisi.Telefon = dataReader("TELEFON").ToString
    '                'MyKisi.Dusunceler = "" 'dataReader("DUSUNCELER").ToString
    '                If Not IsDBNull(dataReader("DOGUM_TARIHI")) Then
    '                    MyKisi.DogumTarihi = dataReader("DOGUM_TARIHI")
    '                End If
    '                MyKisi.DogumYeri = dataReader("DOGUM_YERI").ToString
    '            Loop
    '            dataReader.Close()
    '            dataReader = Nothing
    '            'command = Nothing
    '        Catch ex As Exception

    '        End Try
    '        'Dim command As OleDbCommand = connection.CreateCommand()
    '        command.CommandText = "SELECT * FROM MULKIYET WHERE ID=" & MulkiyetID.ToString
    '        Try
    '            If Not connection.State = ConnectionState.Open Then connection.Open()
    '            Dim dataReader1 As OleDbDataReader = command.ExecuteReader()
    '            Do While dataReader1.Read()
    '                MyKisi.MulkiyetID = MulkiyetID
    '                If Not IsDBNull(dataReader1("PARSEL_ID")) Then
    '                    MyKisi.ParselID = dataReader1("PARSEL_ID")
    '                End If
    '                If Not IsDBNull(dataReader1("KISI_ID")) Then
    '                    MyKisi.ID = dataReader1("KISI_ID")
    '                End If
    '                If Not IsDBNull(dataReader1("PAY")) Then
    '                    MyKisi.HissePay = dataReader1("PAY")
    '                End If
    '                If Not IsDBNull(dataReader1("PAYDA")) Then
    '                    MyKisi.HissePayda = dataReader1("PAYDA")
    '                End If
    '                If Not IsDBNull(dataReader1("TAPU_TARIHI")) Then
    '                    MyKisi.TapuTarihi = dataReader1("TAPU_TARIHI")
    '                End If
    '                MyKisi.Dusunceler = dataReader1("HISSE_REHIN").ToString
    '                MyKisi.Dusunceler = dataReader1("HISSE_REHIN_ALACAKLI").ToString
    '                MyKisi.Dusunceler = dataReader1("HISSE_SERH").ToString
    '                MyKisi.Dusunceler = dataReader1("DUSUNCELER").ToString
    '            Loop
    '            dataReader1.Close()
    '            dataReader1 = Nothing
    '            command = Nothing



    '            ' connection.Close()
    '            ' connection = Nothing
    '        Catch ex As Exception

    '        End Try
    '    End Using
    '    Return MyKisi
    'End Function

    Public Function GetKisi(TCKimlikNo As Double) As Kisi
        Dim MyKisi As New Kisi
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM KISI WHERE TC_KIMLIK_NO=" & TCKimlikNo.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                If Not IsDBNull(dataReader("ID")) Then
                    MyKisi.ID = dataReader("ID")
                End If
                MyKisi.TCKimlikNo = TCKimlikNo
                MyKisi.Adi = dataReader("ADI").ToString
                MyKisi.Soyadi = dataReader("SOYADI").ToString
                MyKisi.Cinsiyet = dataReader("CINSIYET").ToString
                MyKisi.Baba = dataReader("BABA").ToString
                MyKisi.Durumu = dataReader("DURUMU").ToString
                MyKisi.Adres = dataReader("ADRES").ToString
                MyKisi.Telefon = dataReader("TELEFON").ToString
                If Not IsDBNull(dataReader("DOGUM_TARIHI")) Then
                    MyKisi.DogumTarihi = dataReader("DOGUM_TARIHI")
                End If
                MyKisi.DogumYeri = dataReader("DOGUM_YERI").ToString
                MyKisi.IBAN = dataReader("IBAN").ToString
                MyKisi.BankaSubeKodu = dataReader("SUBE_KODU").ToString
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception

        End Try
        Return MyKisi
    End Function

    Public Function GetKisiKod(KisiID As Long) As KisiKod
        Dim MyKisiKod As New KisiKod
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM KISI_KOD WHERE KISI_ID=" & KisiID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                'If Not IsDBNull(dataReader("MALIK_TIPI")) Then
                '    MyKisiKod.MalikTipi = dataReader("MALIK_TIPI")
                'End If
                MyKisiKod.ID = dataReader("ID")
                If Not IsDBNull(dataReader("DAVETIYE_TEBLIG_DURUMU")) Then
                    MyKisiKod.DavetiyeTebligDurumu = dataReader("DAVETIYE_TEBLIG_DURUMU")
                End If
                If Not IsDBNull(dataReader("DAVETIYE_ALINMA_DURUMU")) Then
                    MyKisiKod.DavetiyeAlinmaDurumu = dataReader("DAVETIYE_ALINMA_DURUMU")
                End If
                If Not IsDBNull(dataReader("GORUSME_DURUMU")) Then
                    MyKisiKod.GorusmeDurumu = dataReader("GORUSME_DURUMU")
                End If
                If Not IsDBNull(dataReader("GORUSME_NO")) Then
                    MyKisiKod.GorusmeNo = dataReader("GORUSME_NO")
                End If
                If Not IsDBNull(dataReader("GORUSME_TARIHI")) Then
                    MyKisiKod.GorusmeTarihi = dataReader("GORUSME_TARIHI")
                End If
                If Not IsDBNull(dataReader("ANLASMA_DURUMU")) Then
                    MyKisiKod.AnlasmaDurumu = dataReader("ANLASMA_DURUMU")
                End If
                If Not IsDBNull(dataReader("ANLASMA_TARIHI")) Then
                    MyKisiKod.AnlasmaTarihi = dataReader("ANLASMA_TARIHI")
                End If
                MyKisiKod.AnlasmaDusunceler = dataReader("ANLASMA_DUSUNCELER").ToString
                If Not IsDBNull(dataReader("TESCIL_DURUMU")) Then
                    MyKisiKod.TescilDurumu = dataReader("TESCIL_DURUMU")
                End If
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            'MyParsel = Nothing
        End Try
        Return MyKisiKod
    End Function

    Public Function GetVarisler(KisiID As Long) As Collection
        Dim MyVarisler As New Collection
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM MIRAS WHERE MURIS=" & KisiID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                If Not IsDBNull(dataReader("VARIS")) Then
                    Dim MyKisi As New Kisi
                    MyKisi.ID = dataReader("VARIS")
                    If MyKisi.ID > 0 Then
                        MyKisi = GetKisi(MyKisi.ID)
                        If Not IsNothing(MyKisi.Adi) Then
                            Dim MyKisiKod As KisiKod = GetKisiKod(MyKisi.ID)
                            MyKisi.Kod = MyKisiKod
                            MyKisiKod = Nothing
                            MyKisi.IsVaris = True
                            Try
                                MyVarisler.Add(MyKisi, MyKisi.ID.ToString)
                            Catch ex As Exception

                            End Try
                        End If
                        MyKisi = Nothing
                    Else
                        MyKisi.IsVaris = False
                    End If
                End If
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            'MyParsel = Nothing
        End Try
        Return MyVarisler
    End Function

    Public Function GetMurisler(KisiID As Long) As Collection
        Dim MyMurisler As New Collection
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM MIRAS WHERE VARIS=" & KisiID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                If Not IsDBNull(dataReader("MURIS")) Then
                    Dim MyKisi As New Kisi
                    MyKisi.ID = dataReader("MURIS")
                    If MyKisi.ID > 0 Then
                        MyKisi = GetKisi(MyKisi.ID)
                        'MyKisi.IsVaris = True
                        MyMurisler.Add(MyKisi, MyKisi.ID.ToString)
                    Else
                        'MyKisi.IsVaris = False
                    End If
                End If
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            'MyParsel = Nothing
        End Try
        Return MyMurisler
    End Function

    Public Function GetKamu(KamuID As Long) As Parsel
        Dim MyParsel As New Parsel
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM KAMULASTIRMA WHERE ID=" & KamuID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyParsel.ID = dataReader("PARSEL_ID")
                If Not IsDBNull(dataReader("MULKIYET_ALAN")) Then
                    MyParsel.MulkiyetAlan = dataReader("MULKIYET_ALAN")
                End If
                If Not IsDBNull(dataReader("IRTIFAK_ALAN")) Then
                    MyParsel.IrtifakAlan = dataReader("IRTIFAK_ALAN")
                End If
                If Not IsDBNull(dataReader("GECICI_IRTIFAK_ALAN")) Then
                    MyParsel.GeciciIrtifakAlan = dataReader("GECICI_IRTIFAK_ALAN")
                End If
                If Not IsDBNull(dataReader("MULKIYET_BEDEL")) Then
                    MyParsel.MulkiyetBedel = dataReader("MULKIYET_BEDEL")
                End If
                If Not IsDBNull(dataReader("IRTIFAK_BEDEL")) Then
                    MyParsel.IrtifakBedel = dataReader("IRTIFAK_BEDEL")
                End If
                If Not IsDBNull(dataReader("GECICI_IRTIFAK_BEDEL")) Then
                    MyParsel.GeciciIrtifakBedel = dataReader("GECICI_IRTIFAK_BEDEL")
                End If
                MyParsel.AraziVasfi = dataReader("ARAZI_VASFI").ToString
                MyParsel.KamulastirmaAmaci = dataReader("KAMULASTIRMA_AMACI").ToString
                If Not IsDBNull(dataReader("YILLIK_ORTALAMA_NET_GELIR")) Then
                    MyParsel.YillikOrtalamaNetGelir = dataReader("YILLIK_ORTALAMA_NET_GELIR")
                End If
                If Not IsDBNull(dataReader("KAPITALIZASYON_FAIZI")) Then
                    MyParsel.KapitalizasyonOrani = dataReader("KAPITALIZASYON_FAIZI")
                End If
                If Not IsDBNull(dataReader("OBJEKTIF_ARTIS")) Then
                    MyParsel.ObjektifArtis = dataReader("OBJEKTIF_ARTIS")
                End If
                If Not IsDBNull(dataReader("ART_KISIM_ARTIS")) Then
                    MyParsel.ArtanKisimArtis = dataReader("ART_KISIM_ARTIS")
                End If
                If Not IsDBNull(dataReader("VERIM_KAYBI")) Then
                    MyParsel.VerimKaybi = dataReader("VERIM_KAYBI")
                End If
                MyParsel.YayginMunavebeSistemi = dataReader("YAYGIN_MUNAVEBE_SISTEMI").ToString
                MyParsel.DegerlemeRaporu = dataReader("DEGERLEME_RAPORU").ToString
                If Not IsDBNull(dataReader("DEGERLEME_TARIHI")) Then
                    MyParsel.DegerlemeTarihi = dataReader("DEGERLEME_TARIHI")
                End If
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            'MyParsel = Nothing
        End Try
        Return MyParsel
    End Function

    Public Function GetMustemilat(MustemilatID As Long) As Mustemilat
        Dim MyMustemilat As New Mustemilat
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM MUSTEMILAT WHERE ID=" & MustemilatID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyMustemilat.ID = MustemilatID
                If Not IsDBNull(dataReader("PARSEL_ID")) Then
                    MyMustemilat.ParselID = dataReader("PARSEL_ID")
                End If
                If Not IsDBNull(dataReader("SAHIP_ID")) Then
                    MyMustemilat.SahipID = dataReader("SAHIP_ID")
                End If
                MyMustemilat.Tanim = dataReader("TANIM").ToString
                If Not IsDBNull(dataReader("ADET")) Then
                    MyMustemilat.Adet = dataReader("ADET")
                End If
                If Not IsDBNull(dataReader("FIYAT")) Then
                    MyMustemilat.Fiyat = dataReader("FIYAT")
                End If
                If Not IsDBNull(dataReader("MALIK")) Then
                    MyMustemilat.Malik = dataReader("MALIK")
                End If
                If Not IsDBNull(dataReader("PAY")) Then
                    MyMustemilat.Pay = dataReader("PAY")
                End If
                If Not IsDBNull(dataReader("PAYDA")) Then
                    MyMustemilat.Payda = dataReader("PAYDA")
                End If
                If Not IsDBNull(dataReader("ODEME_ID")) Then
                    MyMustemilat.OdemeID = dataReader("ODEME_ID")
                End If
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            'MyMustemilat = Nothing
        End Try
        Return MyMustemilat
    End Function

    Public Function GetMustemilatlar(ParselID As Long, SahipID As Long) As Collection
        Dim MyMustemilatlar As New Collection
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "SELECT * FROM MUSTEMILAT WHERE PARSEL_ID=" & ParselID.ToString & " AND SAHIP_ID=" & SahipID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()
                Dim dataReader As OleDbDataReader = command.ExecuteReader()
                Do While dataReader.Read()
                    Dim MyMustemilat As New Mustemilat
                    If Not IsDBNull(dataReader("ID")) Then
                        MyMustemilat.ID = dataReader("ID")
                    End If
                    MyMustemilat.ParselID = ParselID
                    MyMustemilat.SahipID = SahipID
                    MyMustemilat.Tanim = dataReader("TANIM").ToString
                    If Not IsDBNull(dataReader("ADET")) Then
                        MyMustemilat.Adet = dataReader("ADET")
                    End If
                    If Not IsDBNull(dataReader("FIYAT")) Then
                        MyMustemilat.Fiyat = dataReader("FIYAT")
                    End If
                    If Not IsDBNull(dataReader("MALIK")) Then
                        MyMustemilat.Malik = dataReader("MALIK")
                    End If
                    If Not IsDBNull(dataReader("PAY")) Then
                        MyMustemilat.Pay = dataReader("PAY")
                    End If
                    If Not IsDBNull(dataReader("PAYDA")) Then
                        MyMustemilat.Payda = dataReader("PAYDA")
                    End If
                    If Not IsDBNull(dataReader("ODEME_ID")) Then
                        MyMustemilat.OdemeID = dataReader("ODEME_ID")
                    End If
                    MyMustemilatlar.Add(MyMustemilat)
                Loop
                dataReader.Close()
                dataReader = Nothing
                command = Nothing
            Catch ex As Exception
                'MyMustemilat = Nothing
            End Try
        End Using
        Return MyMustemilatlar
    End Function

    Public Function GetMevsimlik(MevsimlikID As Long) As Mevsimlik
        Dim MyMevsimlik As New Mevsimlik
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM MEVSIMLIK WHERE ID=" & MevsimlikID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyMevsimlik.ID = MevsimlikID
                If Not IsDBNull(dataReader("PARSEL_ID")) Then
                    MyMevsimlik.ParselID = dataReader("PARSEL_ID")
                End If
                If Not IsDBNull(dataReader("SAHIP_ID")) Then
                    MyMevsimlik.SahipID = dataReader("SAHIP_ID")
                End If
                MyMevsimlik.Tanim = dataReader("TANIM").ToString
                If Not IsDBNull(dataReader("ALAN")) Then
                    MyMevsimlik.Alan = dataReader("ALAN")
                End If
                If Not IsDBNull(dataReader("BEDEL")) Then
                    MyMevsimlik.Bedel = dataReader("BEDEL")
                End If
                If Not IsDBNull(dataReader("MALIK")) Then
                    MyMevsimlik.Malik = dataReader("MALIK")
                End If
                If Not IsDBNull(dataReader("PAY")) Then
                    MyMevsimlik.Pay = dataReader("PAY")
                End If
                If Not IsDBNull(dataReader("PAYDA")) Then
                    MyMevsimlik.Payda = dataReader("PAYDA")
                End If
                If Not IsDBNull(dataReader("ODEME_ID")) Then
                    MyMevsimlik.OdemeID = dataReader("ODEME_ID")
                End If
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            'MyMevsimlik = Nothing
        End Try
        Return MyMevsimlik
    End Function

    Public Function GetMevsimlikler(ParselID As Long, SahipID As Long) As Collection
        Dim MyMevsimlikler As New Collection
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "SELECT * FROM MEVSIMLIK WHERE PARSEL_ID=" & ParselID.ToString & " AND SAHIP_ID=" & SahipID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()
                Dim dataReader As OleDbDataReader = command.ExecuteReader()
                Do While dataReader.Read()
                    Dim MyMevsimlik As New Mevsimlik
                    If Not IsDBNull(dataReader("PARSEL_ID")) Then
                        MyMevsimlik.ID = dataReader("ID")
                    End If
                    MyMevsimlik.ParselID = ParselID
                    MyMevsimlik.SahipID = SahipID
                    MyMevsimlik.Tanim = dataReader("TANIM").ToString
                    If Not IsDBNull(dataReader("ALAN")) Then
                        MyMevsimlik.Alan = dataReader("ALAN")
                    End If
                    If Not IsDBNull(dataReader("BEDEL")) Then
                        MyMevsimlik.Bedel = dataReader("BEDEL")
                    End If
                    If Not IsDBNull(dataReader("MALIK")) Then
                        MyMevsimlik.Malik = dataReader("MALIK")
                    End If
                    If Not IsDBNull(dataReader("PAY")) Then
                        MyMevsimlik.Pay = dataReader("PAY")
                    End If
                    If Not IsDBNull(dataReader("PAYDA")) Then
                        MyMevsimlik.Payda = dataReader("PAYDA")
                    End If
                    If Not IsDBNull(dataReader("ODEME_ID")) Then
                        MyMevsimlik.OdemeID = dataReader("ODEME_ID")
                    End If
                    MyMevsimlikler.Add(MyMevsimlik)
                Loop
                dataReader.Close()
                dataReader = Nothing
                command = Nothing
            Catch ex As Exception
                'MyMevsimlik = Nothing
            End Try
        End Using
        Return MyMevsimlikler
    End Function

    Public Function GetDavaAcele(DavaAceleID As Long) As DavaAcele
        Dim MyDavaAcele As New DavaAcele
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM DAVA_27 WHERE ID=" & DavaAceleID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyDavaAcele.ID = DavaAceleID
                If Not IsDBNull(dataReader("PARSEL_ID")) Then
                    MyDavaAcele.ParselID = dataReader("PARSEL_ID")
                End If
                MyDavaAcele.Mahkeme = dataReader("MAHKEME").ToString
                MyDavaAcele.EsasNo = dataReader("ESAS_NO").ToString
                MyDavaAcele.KararNo = dataReader("KARAR_NO").ToString
                If Not IsDBNull(dataReader("KARAR_TARIHI")) Then
                    MyDavaAcele.KararTarihi = dataReader("KARAR_TARIHI")
                End If
                If Not IsDBNull(dataReader("DAVA_ACILAN_HISSE_PAY")) Then
                    MyDavaAcele.DavaAcilanHissePay = dataReader("DAVA_ACILAN_HISSE_PAY")
                End If
                If Not IsDBNull(dataReader("DAVA_ACILAN_HISSE_PAYDA")) Then
                    MyDavaAcele.DavaAcilanHissePayda = dataReader("DAVA_ACILAN_HISSE_PAYDA")
                End If
                If Not IsDBNull(dataReader("TOPLAM_KAMULASTIRMA_BEDELI")) Then
                    MyDavaAcele.ToplamKamulastirmaBedeli = dataReader("TOPLAM_KAMULASTIRMA_BEDELI")
                End If
                If Not IsDBNull(dataReader("DAVA_TARIHI")) Then
                    MyDavaAcele.DavaTarihi = dataReader("DAVA_TARIHI")
                End If
                If Not IsDBNull(dataReader("KESIF_TARIHI")) Then
                    MyDavaAcele.KesifTarihi = dataReader("KESIF_TARIHI")
                End If
                MyDavaAcele.BlokeOluru = dataReader("BLOKE_OLURU").ToString
                If Not IsDBNull(dataReader("OLUR_TARIHI")) Then
                    MyDavaAcele.OlurTarihi = dataReader("OLUR_TARIHI")
                End If
                If Not IsDBNull(dataReader("BLOKE_TARIHI")) Then
                    MyDavaAcele.BlokeTarihi = dataReader("BLOKE_TARIHI")
                End If
                MyDavaAcele.Avukat = dataReader("AVUKAT").ToString
                MyDavaAcele.Dusunceler = dataReader("DUSUNCELER").ToString
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            'MyMustemilat = Nothing
        End Try
        Return MyDavaAcele
    End Function

    Public Function GetDavaTescil(DavaTescilID As Long) As DavaTescil
        Dim MyDavaTescil As New DavaTescil
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT * FROM DAVA_10 WHERE ID=" & DavaTescilID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyDavaTescil.ID = DavaTescilID
                If Not IsDBNull(dataReader("PARSEL_ID")) Then
                    MyDavaTescil.ParselID = dataReader("PARSEL_ID")
                End If
                MyDavaTescil.Mahkeme = dataReader("MAHKEME").ToString
                MyDavaTescil.EsasNo = dataReader("ESAS_NO").ToString
                MyDavaTescil.KararNo = dataReader("KARAR_NO").ToString
                If Not IsDBNull(dataReader("KARAR_TARIHI")) Then
                    MyDavaTescil.KararTarihi = dataReader("KARAR_TARIHI")
                End If
                If Not IsDBNull(dataReader("DAVA_ACILAN_HISSE_PAY")) Then
                    MyDavaTescil.DavaAcilanHissePay = dataReader("DAVA_ACILAN_HISSE_PAY")
                End If
                If Not IsDBNull(dataReader("DAVA_ACILAN_HISSE_PAYDA")) Then
                    MyDavaTescil.DavaAcilanHissePayda = dataReader("DAVA_ACILAN_HISSE_PAYDA")
                End If
                If Not IsDBNull(dataReader("TOPLAM_KAMULASTIRMA_BEDELI")) Then
                    MyDavaTescil.ToplamKamulastirmaBedeli = dataReader("TOPLAM_KAMULASTIRMA_BEDELI")
                End If
                If Not IsDBNull(dataReader("DAVA_TARIHI")) Then
                    MyDavaTescil.DavaTarihi = dataReader("DAVA_TARIHI")
                End If
                If Not IsDBNull(dataReader("BIRINCI_KESIF_TARIHI")) Then
                    MyDavaTescil.KesifTarihi1 = dataReader("BIRINCI_KESIF_TARIHI")
                End If
                If Not IsDBNull(dataReader("BIRINCI_DURUSMA_TARIHI")) Then
                    MyDavaTescil.DurusmaTarihi1 = dataReader("BIRINCI_DURUSMA_TARIHI")
                End If
                If Not IsDBNull(dataReader("IKINCI_KESIF_TARIHI")) Then
                    MyDavaTescil.KesifTarihi2 = dataReader("IKINCI_KESIF_TARIHI")
                End If
                If Not IsDBNull(dataReader("IKINCI_DURUSMA_TARIHI")) Then
                    MyDavaTescil.DurusmaTarihi2 = dataReader("IKINCI_DURUSMA_TARIHI")
                End If
                If Not IsDBNull(dataReader("UCUNCU_DURUSMA_TARIHI")) Then
                    MyDavaTescil.DurusmaTarihi3 = dataReader("UCUNCU_DURUSMA_TARIHI")
                End If
                MyDavaTescil.BlokeOluru = dataReader("BLOKE_OLURU").ToString
                If Not IsDBNull(dataReader("OLUR_TARIHI")) Then
                    MyDavaTescil.OlurTarihi = dataReader("OLUR_TARIHI")
                End If
                If Not IsDBNull(dataReader("BLOKE_TARIHI")) Then
                    MyDavaTescil.BlokeTarihi = dataReader("BLOKE_TARIHI")
                End If
                MyDavaTescil.Avukat = dataReader("AVUKAT").ToString
                MyDavaTescil.Dusunceler = dataReader("DUSUNCELER").ToString
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            'MyMustemilat = Nothing
        End Try
        Return MyDavaTescil
    End Function

    Public Function GetOdeme(OdemeID As Long) As Odeme
        Dim MyOdeme As New Odeme
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "SELECT * FROM ODEME WHERE ID=" & OdemeID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()
                Dim dataReader As OleDbDataReader = command.ExecuteReader()
                Do While dataReader.Read()
                    MyOdeme.ID = OdemeID
                    If Not IsDBNull(dataReader("PARSEL_ID")) Then
                        MyOdeme.ParselID = dataReader("PARSEL_ID")
                    End If
                    If Not IsDBNull(dataReader("KISI_ID")) Then
                        MyOdeme.KisiID = dataReader("KISI_ID")
                    End If
                    If Not IsDBNull(dataReader("ONAY_ID")) Then
                        MyOdeme.OnayID = dataReader("ONAY_ID")
                    End If
                    If Not IsDBNull(dataReader("ODENEN_BEDEL")) Then
                        MyOdeme.Tutar = dataReader("ODENEN_BEDEL")
                    End If
                    If Not IsDBNull(dataReader("ODEME_TARIHI")) Then
                        MyOdeme.Tarih = dataReader("ODEME_TARIHI")
                    End If
                    MyOdeme.Sekli = dataReader("ODEME_SEKLI").ToString
                    MyOdeme.Tipi = dataReader("ODEME_TIPI").ToString
                    MyOdeme.Kaynak = dataReader("KAYNAK").ToString
                    If Not IsDBNull(dataReader("ODEME_DURUMU")) Then
                        MyOdeme.Durumu = dataReader("ODEME_DURUMU")
                    End If
                    MyOdeme.Aciklama = dataReader("ACIKLAMA").ToString
                Loop
                dataReader.Close()
                dataReader = Nothing
                command = Nothing
            Catch ex As Exception
                'MyOdeme = Nothing
            End Try

            Dim Belgeler As New Collection
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "SELECT * FROM ODEME_BELGE WHERE ODEME_ID=" & OdemeID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()
                Dim dataReader As OleDbDataReader = command.ExecuteReader()
                Do While dataReader.Read()
                    Dim MyBelge As New Belge
                    MyBelge.ID = dataReader("ID")
                    MyBelge.OdemeID = OdemeID
                    MyBelge.Adi = dataReader("ADI").ToString
                    MyBelge.Yol = dataReader("YOL").ToString
                    MyBelge.Aciklama = dataReader("ACIKLAMA").ToString
                    Belgeler.Add(MyBelge)
                Loop
                dataReader.Close()
                dataReader = Nothing
                command = Nothing
            Catch ex As Exception
                'MyOdeme = Nothing
            End Try
            MyOdeme.Belgeler = Belgeler
        End Using
        Return MyOdeme
    End Function

    Public Function GetParselID(_Parsel As Parsel) As Long
        Dim MyID As Long = -1
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT ID FROM PARSEL WHERE IL='" + _Parsel.Il.ToString + "' AND ILCE='" + _Parsel.Ilce.ToString + "' AND KOY='" + _Parsel.Koy.ToString + "' AND MAHALLE='" + _Parsel.Mahalle.ToString + "' AND ADA='" + _Parsel.AdaNo.ToString + "' AND PARSEL='" + _Parsel.ParselNo.ToString + "'"
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyID = dataReader("ID")
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            MyID = -1
        End Try
        Return MyID
    End Function

    Public Function GetKisiID(_Kisi As Kisi) As Long
        Dim MyID As Long = -1
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT ID FROM KISI WHERE ADI='" & _Kisi.Adi.ToString & "' AND SOYADI='" & _Kisi.Soyadi.ToString & "' AND BABA='" & _Kisi.Baba.ToString & "'"
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyID = dataReader("ID")
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            MyID = -1
        End Try
        Return MyID
    End Function

    Public Function GetProjeID(_Proje As Proje) As Long
        Dim MyID As Long = -1
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT ID FROM PROJE WHERE AD='" & _Proje.Ad.ToString & "'"
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyID = dataReader("ID")
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            MyID = -1
        End Try
        Return MyID
    End Function

    Public Function GetKamuID(_Parsel As Parsel) As Long
        Dim MyID As Long = -1
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT ID FROM KAMULASTIRMA WHERE PARSEL_ID=" & _Parsel.ID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyID = dataReader("ID")
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            MyID = -1
        End Try
        Return MyID
    End Function

    Public Function GetDavaAceleID(_Parsel As Parsel) As Long
        Dim MyID As Long = -1
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT ID FROM DAVA_27 WHERE PARSEL_ID=" & _Parsel.ID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyID = dataReader("ID")
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            MyID = -1
        End Try
        Return MyID
    End Function

    Public Function GetDavaTescilID(_Parsel As Parsel) As Long
        Dim MyID As Long = -1
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT ID FROM DAVA_10 WHERE PARSEL_ID=" & _Parsel.ID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyID = dataReader("ID")
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            MyID = -1
        End Try
        Return MyID
    End Function

    Public Function GetMustemilatOdemeID(_Mustemilat As Mustemilat) As Long
        Dim MyID As Long = -1
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT ODEME_ID FROM MUSTEMILAT WHERE ID=" & _Mustemilat.ID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyID = dataReader("ODEME_ID")
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            MyID = -1
        End Try
        Return MyID
    End Function

    Public Function GetMevsimlikOdemeID(_Mevsimlik As Mevsimlik) As Long
        Dim MyID As Long = -1
        Dim connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
        Dim command As OleDbCommand = connection.CreateCommand()
        command.CommandText = "SELECT ODEME_ID FROM MEVSIMLIK WHERE ID=" & _Mevsimlik.ID.ToString
        Try
            If Not connection.State = ConnectionState.Open Then connection.Open()
            Dim dataReader As OleDbDataReader = command.ExecuteReader()
            Do While dataReader.Read()
                MyID = dataReader("ODEME_ID")
            Loop
            dataReader.Close()
            dataReader = Nothing
            command = Nothing
            connection.Close()
            connection = Nothing
        Catch ex As Exception
            MyID = -1
        End Try
        Return MyID
    End Function

    Public Function GetMulkiyet(KisiID As Long, MulkiyetID As Long) As Kisi
        Dim MyKisi As New Kisi
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Dim command As OleDbCommand = connection.CreateCommand()
            command.CommandText = "SELECT * FROM KISI WHERE ID=" & KisiID.ToString
            Try
                If Not connection.State = ConnectionState.Open Then connection.Open()
                Dim dataReader As OleDbDataReader = command.ExecuteReader()
                Do While dataReader.Read()
                    MyKisi.ID = KisiID
                    If Not IsDBNull(dataReader("TC_KIMLIK_NO")) Then
                        MyKisi.TCKimlikNo = dataReader("TC_KIMLIK_NO")
                    End If
                    MyKisi.Adi = dataReader("ADI").ToString
                    MyKisi.Soyadi = dataReader("SOYADI").ToString
                    MyKisi.Cinsiyet = dataReader("CINSIYET").ToString
                    MyKisi.Baba = dataReader("BABA").ToString
                    MyKisi.Durumu = dataReader("DURUMU").ToString
                    MyKisi.Adres = dataReader("ADRES").ToString
                    MyKisi.Telefon = dataReader("TELEFON").ToString
                    'MyKisi.Dusunceler = "" 'dataReader("DUSUNCELER").ToString
                    If Not IsDBNull(dataReader("DOGUM_TARIHI")) Then
                        MyKisi.DogumTarihi = dataReader("DOGUM_TARIHI")
                    End If
                    MyKisi.DogumYeri = dataReader("DOGUM_YERI").ToString
                    MyKisi.IBAN = dataReader("IBAN").ToString
                    MyKisi.BankaSubeKodu = dataReader("SUBE_KODU").ToString
                Loop
                dataReader.Close()
                dataReader = Nothing
                'command = Nothing
            Catch ex As Exception

            End Try
            'Dim command As OleDbCommand = connection.CreateCommand()
            command.CommandText = "SELECT * FROM MULKIYET WHERE ID=" & MulkiyetID.ToString
            Try
                If Not connection.State = ConnectionState.Open Then connection.Open()
                Dim dataReader1 As OleDbDataReader = command.ExecuteReader()
                Do While dataReader1.Read()
                    MyKisi.MulkiyetID = MulkiyetID
                    If Not IsDBNull(dataReader1("PARSEL_ID")) Then
                        MyKisi.ParselID = dataReader1("PARSEL_ID")
                    End If
                    If Not IsDBNull(dataReader1("KISI_ID")) Then
                        MyKisi.ID = dataReader1("KISI_ID")
                    End If
                    If Not IsDBNull(dataReader1("PAY")) Then
                        MyKisi.HissePay = dataReader1("PAY")
                    End If
                    If Not IsDBNull(dataReader1("PAYDA")) Then
                        MyKisi.HissePayda = dataReader1("PAYDA")
                    End If
                    If Not IsDBNull(dataReader1("TAPU_TARIHI")) Then
                        MyKisi.TapuTarihi = dataReader1("TAPU_TARIHI")
                    End If
                    MyKisi.Rehin = dataReader1("HISSE_REHIN").ToString
                    MyKisi.RehinAlacakli = dataReader1("HISSE_REHIN_ALACAKLI").ToString
                    MyKisi.SerhBeyan = dataReader1("HISSE_SERH").ToString
                    MyKisi.Dusunceler = dataReader1("DUSUNCELER").ToString
                Loop
                dataReader1.Close()
                dataReader1 = Nothing
                command = Nothing

                ' connection.Close()
                ' connection = Nothing
            Catch ex As Exception

            End Try
        End Using
        Return MyKisi
    End Function

    Public Function GetMulkiyet(KisiID As Long, ParselID As Long, Optional ByRef GetOption As Boolean = True) As Kisi
        Dim MyKisi As New Kisi
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Dim Command As OleDbCommand = connection.CreateCommand()
            Command = connection.CreateCommand()
            Command.CommandText = "SELECT * FROM MULKIYET WHERE PARSEL_ID=" & ParselID.ToString + " AND KISI_ID=" & KisiID.ToString
            Try
                If Not connection.State = ConnectionState.Open Then connection.Open()
                Dim dataReader As OleDbDataReader = Command.ExecuteReader()
                Do While dataReader.Read()
                    MyKisi.MulkiyetID = dataReader("ID")
                    MyKisi.ParselID = ParselID
                    MyKisi.ID = KisiID
                    If Not IsDBNull(dataReader("PAY")) Then
                        MyKisi.HissePay = dataReader("PAY")
                    End If
                    If Not IsDBNull(dataReader("PAYDA")) Then
                        MyKisi.HissePayda = dataReader("PAYDA")
                    End If
                    If Not IsDBNull(dataReader("TAPU_TARIHI")) Then
                        MyKisi.TapuTarihi = dataReader("TAPU_TARIHI")
                    End If
                    MyKisi.Rehin = dataReader("HISSE_REHIN").ToString
                    MyKisi.RehinAlacakli = dataReader("HISSE_REHIN_ALACAKLI").ToString
                    MyKisi.SerhBeyan = dataReader("HISSE_SERH").ToString
                    MyKisi.Dusunceler = dataReader("DUSUNCELER").ToString
                Loop
                dataReader.Close()
                dataReader = Nothing
                Command = Nothing

                ' connection.Close()
                ' connection = Nothing
            Catch ex As Exception

            End Try
        End Using
        Return MyKisi
    End Function

#End Region

#Region "Add Procedures"

    Public Function AddParsel(_Parsel As Parsel) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM PARSEL"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("PROJE_ID") = _Parsel.ProjeID
            MyRow("KOD") = _Parsel.Code
            MyRow("IL") = _Parsel.Il
            MyRow("ILCE") = _Parsel.Ilce
            MyRow("KOY") = _Parsel.Koy
            MyRow("MAHALLE") = _Parsel.Mahalle
            MyRow("ADA") = _Parsel.AdaNo
            MyRow("PARSEL") = _Parsel.ParselNo
            MyRow("PAFTA") = _Parsel.PaftaNo
            MyRow("MEVKI") = _Parsel.Mevki
            MyRow("CILT") = _Parsel.Cilt
            MyRow("SAYFA") = _Parsel.Sayfa
            MyRow("CINSI") = _Parsel.Cinsi
            MyRow("TAPU_ALANI") = _Parsel.TapuAlani

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))
            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddParselKod(_Parsel As Parsel) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM PARSEL_KOD"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("PARSEL_ID") = _Parsel.ID
            MyRow("BOLGE_ID") = _Parsel.Kod.BolgeID
            MyRow("KADASTRAL_DURUM") = _Parsel.Kod.KadastralDurum
            MyRow("MALIK_TIPI") = _Parsel.Kod.MalikTipi
            MyRow("ISTIMLAK_TURU") = _Parsel.Kod.IstimlakTuru
            MyRow("ISTIMLAK_SERHI") = _Parsel.Kod.IstimlakSerhi
            MyRow("DAVA10_DURUMU") = _Parsel.Kod.DavaDurumu10
            MyRow("DAVA27_DURUMU") = _Parsel.Kod.DavaDurumu27
            MyRow("EDINIM_DURUMU") = _Parsel.Kod.EdinimDurumu
            MyRow("ISTIMLAK_DISI") = _Parsel.Kod.IstimlakDisi
            MyRow("DEVIR_DURUMU") = _Parsel.Kod.DevirDurumu
            MyRow("ODEME_DURUMU") = _Parsel.Kod.OdemeDurumu

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyRow = Nothing
            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddParselDetay(_Parsel As Parsel) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM PARSEL_DETAY"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("PARSEL_ID") = _Parsel.ID
            MyRow("ESAS_NO") = _Parsel.Detay.DavaEsasNo
            MyRow("KARAR_NO") = _Parsel.Detay.DavaKararNo
            MyRow("ARSA") = _Parsel.Detay.Arsa
            MyRow("IMAR_DURUMU") = _Parsel.Detay.ImarDurumu
            MyRow("DOP_KESINTISI") = _Parsel.Detay.DopKesintisi
            MyRow("VERGI_DEGERI") = _Parsel.Detay.VergiDegeri
            MyRow("VERGI_DEGERI_YILI") = _Parsel.Detay.VergiDegeriTarihi
            MyRow("KAYIP_ORANI") = _Parsel.Detay.KayipOrani
            MyRow("FAIZ") = _Parsel.Detay.Faiz
            MyRow("YARGITAY_SONUC") = _Parsel.Detay.YargitaySonuc
            MyRow("YARGITAY_ACIKLAMA") = _Parsel.Detay.DavaAciklama

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyRow = Nothing
            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddEmsal(_Parsel As Parsel, _Emsal As Parsel) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM EMSAL"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("PARSEL_ID") = _Parsel.ID
            MyRow("EMSAL_ID") = _Emsal.ID

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))
            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddKisi(_Kisi As Kisi) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim queryString As String = "SELECT * FROM KISI"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(queryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("TC_KIMLIK_NO") = _Kisi.TCKimlikNo
            MyRow("ADI") = _Kisi.Adi
            MyRow("SOYADI") = _Kisi.Soyadi
            MyRow("CINSIYET") = _Kisi.Cinsiyet
            If _Kisi.DogumTarihi.Year > 1752 Then
                MyRow("DOGUM_TARIHI") = _Kisi.DogumTarihi
            End If
            MyRow("DOGUM_YERI") = _Kisi.DogumYeri
            MyRow("BABA") = _Kisi.Baba
            MyRow("DURUMU") = _Kisi.Durumu
            MyRow("ADRES") = _Kisi.Adres
            MyRow("TELEFON") = _Kisi.Telefon
            MyRow("IBAN") = _Kisi.IBAN
            MyRow("SUBE_KODU") = _Kisi.BankaSubeKodu
            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))

            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddKisiKod(_Kisi As Kisi) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim queryString As String = "SELECT * FROM KISI_KOD"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(queryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("KISI_ID") = _Kisi.ID
            'MyRow("MALIK_TIPI") = _Kisi.Kod.MalikTipi
            MyRow("DAVETIYE_TEBLIG_DURUMU") = _Kisi.Kod.DavetiyeTebligDurumu
            MyRow("DAVETIYE_ALINMA_DURUMU") = _Kisi.Kod.DavetiyeAlinmaDurumu
            MyRow("GORUSME_DURUMU") = _Kisi.Kod.GorusmeDurumu
            MyRow("GORUSME_NO") = _Kisi.Kod.GorusmeNo
            MyRow("GORUSME_TARIHI") = _Kisi.Kod.GorusmeTarihi
            MyRow("ANLASMA_DURUMU") = _Kisi.Kod.AnlasmaDurumu
            MyRow("ANLASMA_TARIHI") = _Kisi.Kod.AnlasmaTarihi
            MyRow("ANLASMA_DUSUNCELER") = _Kisi.Kod.AnlasmaDusunceler
            MyRow("TESCIL_DURUMU") = _Kisi.Kod.TescilDurumu

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))
            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    'Public Function AddKisiBanka(_Kisi As Kisi) As Long
    '    Dim MyRowID As Long = -1
    '    Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
    '        Try
    '            If Not connection.State = ConnectionState.Open Then connection.Open()

    '            Dim queryString As String = "SELECT * FROM BANKA"
    '            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(queryString, connection))

    '            Dim MyTable As New DataTable
    '            MyDataAdapter.Fill(MyTable)

    '            Dim MyRow As DataRow

    '            MyRow = MyTable.NewRow()

    '            MyRow("KISI_ID") = _Kisi.ID
    '            MyRow("IBAN") = _Kisi.IBAN
    '            MyRow("SUBE_KODU") = _Kisi.BankaSubeKodu

    '            MyTable.Rows.Add(MyRow)

    '            'Kayıt anında ID alma
    '            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
    '            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))
    '            MyRow = Nothing

    '            Dim MyCommandBuilder As New OleDbCommandBuilder
    '            MyCommandBuilder.DataAdapter = MyDataAdapter
    '            MyDataAdapter.Update(MyTable)

    '            MyTable = Nothing
    '            MyCommandBuilder = Nothing
    '            MyDataAdapter = Nothing
    '        Catch ex As Exception
    '            MyRowID = -1
    '        End Try
    '    End Using
    '    Return MyRowID
    'End Function

    Public Function AddKamu(_Parsel As Parsel) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM KAMULASTIRMA"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("PARSEL_ID") = _Parsel.ID
            MyRow("MULKIYET_ALAN") = _Parsel.MulkiyetAlan
            MyRow("IRTIFAK_ALAN") = _Parsel.IrtifakAlan
            MyRow("GECICI_IRTIFAK_ALAN") = _Parsel.GeciciIrtifakAlan
            MyRow("MULKIYET_BEDEL") = _Parsel.MulkiyetBedel
            MyRow("IRTIFAK_BEDEL") = _Parsel.IrtifakBedel
            MyRow("GECICI_IRTIFAK_BEDEL") = _Parsel.GeciciIrtifakBedel
            MyRow("KAMULASTIRMA_AMACI") = _Parsel.KamulastirmaAmaci
            MyRow("ARAZI_VASFI") = _Parsel.AraziVasfi
            MyRow("YAYGIN_MUNAVEBE_SISTEMI") = _Parsel.YayginMunavebeSistemi
            MyRow("DEGERLEME_RAPORU") = _Parsel.DegerlemeRaporu
            MyRow("DEGERLEME_TARIHI") = _Parsel.DegerlemeTarihi
            MyRow("YILLIK_ORTALAMA_NET_GELIR") = _Parsel.YillikOrtalamaNetGelir
            MyRow("KAPITALIZASYON_FAIZI") = _Parsel.KapitalizasyonOrani
            MyRow("OBJEKTIF_ARTIS") = _Parsel.ObjektifArtis
            MyRow("ART_KISIM_ARTIS") = _Parsel.ArtanKisimArtis
            MyRow("VERIM_KAYBI") = _Parsel.VerimKaybi

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))

            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddMulkiyet(_Parsel As Parsel) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM MULKIYET"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            For Each MyMalik As Kisi In _Parsel.Malikler
                Dim MyRow As DataRow

                MyRow = MyTable.NewRow()
                MyRow("PARSEL_ID") = _Parsel.ID
                MyRow("KISI_ID") = MyMalik.ID
                MyRow("PAY") = MyMalik.HissePay
                MyRow("PAYDA") = MyMalik.HissePayda
                If MyMalik.TapuTarihi.Year > 1752 Then
                    MyRow("TAPU_TARIHI") = MyMalik.TapuTarihi
                End If
                MyRow("HISSE_REHIN") = MyMalik.Rehin
                MyRow("HISSE_REHIN_ALACAKLI") = MyMalik.RehinAlacakli
                MyRow("HISSE_SERH") = MyMalik.SerhBeyan
                MyRow("DUSUNCELER") = MyMalik.Dusunceler

                MyTable.Rows.Add(MyRow)

                'Kayıt anında ID alma
                Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
                MyRowID = CLng(MyFieldInfo.GetValue(MyRow))

                MyRow = Nothing
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddMulkiyet(_Kisi As Kisi) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM MULKIYET"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("PARSEL_ID") = _Kisi.ParselID
            MyRow("KISI_ID") = _Kisi.ID
            MyRow("PAY") = _Kisi.HissePay
            MyRow("PAYDA") = _Kisi.HissePayda
            If _Kisi.TapuTarihi.Year > 1752 Then
                MyRow("TAPU_TARIHI") = _Kisi.TapuTarihi
            End If
            MyRow("HISSE_REHIN") = _Kisi.Rehin
            MyRow("HISSE_REHIN_ALACAKLI") = _Kisi.RehinAlacakli
            MyRow("HISSE_SERH") = _Kisi.SerhBeyan
            MyRow("DUSUNCELER") = _Kisi.Dusunceler

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))

            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddMustemilat(_Mustemilat As Mustemilat) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM MUSTEMILAT"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("PARSEL_ID") = _Mustemilat.ParselID
            MyRow("SAHIP_ID") = _Mustemilat.SahipID
            MyRow("TANIM") = _Mustemilat.Tanim
            MyRow("ADET") = _Mustemilat.Adet
            MyRow("FIYAT") = _Mustemilat.Fiyat
            MyRow("MALIK") = _Mustemilat.Malik
            MyRow("PAY") = _Mustemilat.Pay
            MyRow("PAYDA") = _Mustemilat.Payda
            MyRow("ODEME_ID") = _Mustemilat.OdemeID

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))
            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddMevsimlik(_Mevsimlik As Mevsimlik) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM MEVSIMLIK"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("PARSEL_ID") = _Mevsimlik.ParselID
            MyRow("SAHIP_ID") = _Mevsimlik.SahipID
            MyRow("TANIM") = _Mevsimlik.Tanim
            MyRow("ALAN") = _Mevsimlik.Alan
            MyRow("BEDEL") = _Mevsimlik.Bedel
            MyRow("MALIK") = _Mevsimlik.Malik
            MyRow("PAY") = _Mevsimlik.Pay
            MyRow("PAYDA") = _Mevsimlik.Payda
            MyRow("ODEME_ID") = _Mevsimlik.OdemeID

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))
            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddDavaTescil(_DavaTescil As DavaTescil) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM DAVA_10"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("PARSEL_ID") = _DavaTescil.ParselID
            MyRow("MAHKEME") = _DavaTescil.Mahkeme
            MyRow("ESAS_NO") = _DavaTescil.EsasNo
            MyRow("KARAR_NO") = _DavaTescil.KararNo
            MyRow("KARAR_TARIHI") = _DavaTescil.KararTarihi
            MyRow("DAVA_ACILAN_HISSE_PAY") = _DavaTescil.DavaAcilanHissePay
            MyRow("DAVA_ACILAN_HISSE_PAYDA") = _DavaTescil.DavaAcilanHissePayda
            MyRow("TOPLAM_KAMULASTIRMA_BEDELI") = _DavaTescil.ToplamKamulastirmaBedeli
            MyRow("DAVA_TARIHI") = _DavaTescil.DavaTarihi
            MyRow("BIRINCI_KESIF_TARIHI") = _DavaTescil.KesifTarihi1
            MyRow("BIRINCI_DURUSMA_TARIHI") = _DavaTescil.DurusmaTarihi1
            MyRow("IKINCI_KESIF_TARIHI") = _DavaTescil.KesifTarihi2
            MyRow("IKINCI_DURUSMA_TARIHI") = _DavaTescil.DurusmaTarihi2
            MyRow("UCUNCU_DURUSMA_TARIHI") = _DavaTescil.DurusmaTarihi3
            MyRow("SON_DURUSMA_TARIHI") = _DavaTescil.DurusmaTarihiSon
            MyRow("BLOKE_OLURU") = _DavaTescil.BlokeOluru
            MyRow("OLUR_TARIHI") = _DavaTescil.OlurTarihi
            MyRow("BLOKE_TARIHI") = _DavaTescil.BlokeTarihi
            MyRow("AVUKAT") = _DavaTescil.Avukat
            MyRow("DUSUNCELER") = _DavaTescil.Dusunceler

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))
            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddDavaAcele(_DavaAcele As DavaAcele) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM DAVA_27"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("PARSEL_ID") = _DavaAcele.ParselID
            MyRow("MAHKEME") = _DavaAcele.Mahkeme
            MyRow("ESAS_NO") = _DavaAcele.EsasNo
            MyRow("KARAR_NO") = _DavaAcele.KararNo
            MyRow("KARAR_TARIHI") = _DavaAcele.KararTarihi
            MyRow("DAVA_ACILAN_HISSE_PAY") = _DavaAcele.DavaAcilanHissePay
            MyRow("DAVA_ACILAN_HISSE_PAYDA") = _DavaAcele.DavaAcilanHissePayda
            MyRow("TOPLAM_KAMULASTIRMA_BEDELI") = _DavaAcele.ToplamKamulastirmaBedeli
            MyRow("DAVA_TARIHI") = _DavaAcele.DavaTarihi
            MyRow("KESIF_TARIHI") = _DavaAcele.KesifTarihi
            MyRow("BLOKE_OLURU") = _DavaAcele.BlokeOluru
            MyRow("OLUR_TARIHI") = _DavaAcele.OlurTarihi
            MyRow("BLOKE_TARIHI") = _DavaAcele.BlokeTarihi
            MyRow("AVUKAT") = _DavaAcele.Avukat
            MyRow("DUSUNCELER") = _DavaAcele.Dusunceler

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))
            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddVaris(_Muris As Kisi, _Varis As Kisi) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM MIRAS"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("MURIS") = _Muris.ID
            MyRow("VARIS") = _Varis.ID

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))
            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddOdeme(_Odeme As Odeme) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM ODEME"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            MyRow("ODENEN_BEDEL") = _Odeme.Tutar
            If _Odeme.Tarih.Year > 1752 Then
                MyRow("ODEME_TARIHI") = _Odeme.Tarih
            End If
            MyRow("ODEME_SEKLI") = _Odeme.Sekli
            MyRow("KAYNAK") = _Odeme.Kaynak
            MyRow("ODEME_DURUMU") = _Odeme.Durumu

            MyRow("PARSEL_ID") = _Odeme.ParselID
            MyRow("KISI_ID") = _Odeme.KisiID
            MyRow("ODEME_TIPI") = _Odeme.Tipi
            MyRow("ACIKLAMA") = _Odeme.Aciklama

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))
            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

    Public Function AddOdemeBelge(_Belge As Belge) As Long
        Dim MyRowID As Long = -1
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim MyQueryString As String = "SELECT * FROM ODEME_BELGE"
                Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, connection))

                Dim MyTable As New DataTable
                MyDataAdapter.Fill(MyTable)

                Dim MyRow As DataRow

                MyRow = MyTable.NewRow()

                MyRow("ADI") = _Belge.Adi
                MyRow("ODEME_ID") = _Belge.OdemeID
                MyRow("YOL") = _Belge.Yol
                MyRow("ACIKLAMA") = _Belge.Aciklama

                MyTable.Rows.Add(MyRow)

                'Kayıt anında ID alma
                Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
                MyRowID = CLng(MyFieldInfo.GetValue(MyRow))
                MyRow = Nothing

                Dim MyCommandBuilder As New OleDbCommandBuilder
                MyCommandBuilder.DataAdapter = MyDataAdapter
                MyDataAdapter.Update(MyTable)

                MyTable = Nothing
                MyCommandBuilder = Nothing
                MyDataAdapter = Nothing
            Catch ex As Exception
                MyRowID = -1
            End Try
        End Using
        Return MyRowID
    End Function

    Public Function AddLog(_Log As Log) As Long
        Dim MyRowID As Long = -1
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM Log"
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRow As DataRow

            MyRow = MyTable.NewRow()

            'MyRow("ID") = _Log.ID
            If _Log.ActionDate.Year > 1752 Then
                MyRow("KOMUT_TARIHI") = _Log.ActionDate
            End If
            MyRow("KOMUT_ADI") = _Log.ActionName
            MyRow("KULLANICI") = _Log.User

            MyTable.Rows.Add(MyRow)

            'Kayıt anında ID alma
            Dim MyFieldInfo As System.Reflection.FieldInfo = MyRow.GetType().GetField("_rowID", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            MyRowID = CLng(MyFieldInfo.GetValue(MyRow))
            MyRow = Nothing

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
        Catch ex As Exception
            MyRowID = -1
        End Try
        Return MyRowID
    End Function

#End Region

#Region "Update Procedures"

    Public Function UpdateKamu(_Parsel As Parsel) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM KAMULASTIRMA WHERE ID=" & _Parsel.KamuID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("PARSEL_ID") = _Parsel.ID
                MyRow("MULKIYET_ALAN") = _Parsel.MulkiyetAlan
                MyRow("IRTIFAK_ALAN") = _Parsel.IrtifakAlan
                MyRow("GECICI_IRTIFAK_ALAN") = _Parsel.GeciciIrtifakAlan
                MyRow("MULKIYET_BEDEL") = _Parsel.MulkiyetBedel
                MyRow("IRTIFAK_BEDEL") = _Parsel.IrtifakBedel
                MyRow("GECICI_IRTIFAK_BEDEL") = _Parsel.GeciciIrtifakBedel
                MyRow("KAMULASTIRMA_AMACI") = _Parsel.KamulastirmaAmaci
                MyRow("ARAZI_VASFI") = _Parsel.AraziVasfi
                MyRow("YAYGIN_MUNAVEBE_SISTEMI") = _Parsel.YayginMunavebeSistemi
                MyRow("DEGERLEME_RAPORU") = _Parsel.DegerlemeRaporu
                MyRow("DEGERLEME_TARIHI") = _Parsel.DegerlemeTarihi
                MyRow("YILLIK_ORTALAMA_NET_GELIR") = _Parsel.YillikOrtalamaNetGelir
                MyRow("KAPITALIZASYON_FAIZI") = _Parsel.KapitalizasyonOrani
                MyRow("OBJEKTIF_ARTIS") = _Parsel.ObjektifArtis
                MyRow("ART_KISIM_ARTIS") = _Parsel.ArtanKisimArtis
                MyRow("VERIM_KAYBI") = _Parsel.VerimKaybi
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Public Function UpdateProject(_Proje As Proje) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM PROJE WHERE ID=" & _Proje.ID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("KOD") = _Proje.Kod
                MyRow("AD") = _Proje.Ad
                MyRow("PROJE_NOTLARI") = _Proje.ProjeNotlari
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Public Function UpdateKisi(_Kisi As Kisi) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM KISI WHERE ID=" & _Kisi.ID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("TC_KIMLIK_NO") = _Kisi.TCKimlikNo
                MyRow("ADI") = _Kisi.Adi
                MyRow("SOYADI") = _Kisi.Soyadi
                MyRow("CINSIYET") = _Kisi.Cinsiyet
                If _Kisi.DogumTarihi.Year > 1752 Then
                    MyRow("DOGUM_TARIHI") = _Kisi.DogumTarihi
                End If
                MyRow("DOGUM_YERI") = _Kisi.DogumYeri
                MyRow("BABA") = _Kisi.Baba
                MyRow("DURUMU") = _Kisi.Durumu
                MyRow("ADRES") = _Kisi.Adres
                MyRow("TELEFON") = _Kisi.Telefon
                MyRow("IBAN") = _Kisi.IBAN
                MyRow("SUBE_KODU") = _Kisi.BankaSubeKodu
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Public Function UpdateKisiKod(_Kisi As Kisi) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM KISI_KOD WHERE ID=" & _Kisi.Kod.ID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("KISI_ID") = _Kisi.ID
                MyRow("DAVETIYE_TEBLIG_DURUMU") = _Kisi.Kod.DavetiyeTebligDurumu
                MyRow("DAVETIYE_ALINMA_DURUMU") = _Kisi.Kod.DavetiyeAlinmaDurumu
                MyRow("GORUSME_DURUMU") = _Kisi.Kod.GorusmeDurumu
                MyRow("GORUSME_NO") = _Kisi.Kod.GorusmeNo
                MyRow("GORUSME_TARIHI") = _Kisi.Kod.GorusmeTarihi
                MyRow("ANLASMA_DURUMU") = _Kisi.Kod.AnlasmaDurumu
                MyRow("ANLASMA_TARIHI") = _Kisi.Kod.AnlasmaTarihi
                MyRow("ANLASMA_DUSUNCELER") = _Kisi.Kod.AnlasmaDusunceler
                MyRow("TESCIL_DURUMU") = _Kisi.Kod.TescilDurumu
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    'Public Function UpdateKisiBanka(_Kisi As Kisi, _BankaID As Long) As Boolean
    '    Dim MyStatus As Boolean = False
    '    Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
    '        Try
    '            If Not connection.State = ConnectionState.Open Then connection.Open()

    '            Dim MyQueryString As String = "SELECT * FROM BANKA WHERE ID=" & _BankaID.ToString
    '            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, connection))

    '            Dim MyTable As New DataTable
    '            MyDataAdapter.Fill(MyTable)

    '            Dim MyRows() As DataRow = MyTable.Select()

    '            For Each MyRow As DataRow In MyTable.Select
    '                MyRow("KISI_ID") = _Kisi.ID
    '                MyRow("IBAN") = _Kisi.IBAN
    '                MyRow("SUBE_KODU") = _Kisi.BankaSubeKodu
    '            Next

    '            Dim MyCommandBuilder As New OleDbCommandBuilder
    '            MyCommandBuilder.DataAdapter = MyDataAdapter
    '            MyDataAdapter.Update(MyTable)

    '            MyTable = Nothing
    '            MyCommandBuilder = Nothing
    '            MyDataAdapter = Nothing
    '            MyStatus = True
    '        Catch ex As Exception
    '            MyStatus = False
    '        End Try
    '    End Using
    '    Return MyStatus
    'End Function

    Public Function UpdateParsel(_Parsel As Parsel) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM PARSEL WHERE ID=" & _Parsel.ID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("PROJE_ID") = _Parsel.ProjeID
                MyRow("KOD") = _Parsel.Code
                MyRow("IL") = _Parsel.Il
                MyRow("ILCE") = _Parsel.Ilce
                MyRow("KOY") = _Parsel.Koy
                MyRow("MAHALLE") = _Parsel.Mahalle
                MyRow("ADA") = _Parsel.AdaNo
                MyRow("PARSEL") = _Parsel.ParselNo
                MyRow("PAFTA") = _Parsel.PaftaNo
                MyRow("MEVKI") = _Parsel.Mevki
                MyRow("CILT") = _Parsel.Cilt
                MyRow("SAYFA") = _Parsel.Sayfa
                MyRow("CINSI") = _Parsel.Cinsi
                MyRow("TAPU_ALANI") = _Parsel.TapuAlani
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Public Function UpdateParselKod(_Parsel As Parsel) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM PARSEL_KOD WHERE ID=" & _Parsel.Kod.ID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("BOLGE_ID") = _Parsel.Kod.BolgeID
                MyRow("KADASTRAL_DURUM") = _Parsel.Kod.KadastralDurum
                MyRow("MALIK_TIPI") = _Parsel.Kod.MalikTipi
                MyRow("ISTIMLAK_TURU") = _Parsel.Kod.IstimlakTuru
                MyRow("ISTIMLAK_SERHI") = _Parsel.Kod.IstimlakSerhi
                MyRow("DAVA10_DURUMU") = _Parsel.Kod.DavaDurumu10
                MyRow("DAVA27_DURUMU") = _Parsel.Kod.DavaDurumu27
                MyRow("EDINIM_DURUMU") = _Parsel.Kod.EdinimDurumu
                MyRow("ISTIMLAK_DISI") = _Parsel.Kod.IstimlakDisi
                MyRow("DEVIR_DURUMU") = _Parsel.Kod.DevirDurumu
                MyRow("ODEME_DURUMU") = _Parsel.Kod.OdemeDurumu
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Public Function UpdateParselDetay(_Parsel As Parsel) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM PARSEL_DETAY WHERE ID=" & _Parsel.Detay.ID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("ESAS_NO") = _Parsel.Detay.DavaEsasNo
                MyRow("KARAR_NO") = _Parsel.Detay.DavaKararNo
                MyRow("ARSA") = _Parsel.Detay.Arsa
                MyRow("IMAR_DURUMU") = _Parsel.Detay.ImarDurumu
                MyRow("DOP_KESINTISI") = _Parsel.Detay.DopKesintisi
                MyRow("VERGI_DEGERI") = _Parsel.Detay.VergiDegeri
                MyRow("VERGI_DEGERI_YILI") = _Parsel.Detay.VergiDegeriTarihi
                MyRow("KAYIP_ORANI") = _Parsel.Detay.KayipOrani
                MyRow("FAIZ") = _Parsel.Detay.Faiz
                MyRow("YARGITAY_SONUC") = _Parsel.Detay.YargitaySonuc
                MyRow("YARGITAY_ACIKLAMA") = _Parsel.Detay.DavaAciklama
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Public Function UpdateMulkiyet(_Kisi As Kisi) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM MULKIYET WHERE ID=" & _Kisi.MulkiyetID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("PARSEL_ID") = _Kisi.ParselID
                MyRow("KISI_ID") = _Kisi.ID
                MyRow("PAY") = _Kisi.HissePay
                MyRow("PAYDA") = _Kisi.HissePayda
                If _Kisi.TapuTarihi.Year > 1752 Then
                    MyRow("TAPU_TARIHI") = _Kisi.TapuTarihi
                Else
                    MyRow("TAPU_TARIHI") = Date.Parse("1/1/1000")
                End If
                MyRow("HISSE_REHIN") = _Kisi.Rehin
                MyRow("HISSE_REHIN_ALACAKLI") = _Kisi.RehinAlacakli
                MyRow("HISSE_SERH") = _Kisi.SerhBeyan
                MyRow("DUSUNCELER") = _Kisi.Dusunceler
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Public Function UpdateMustemilat(_Mustemilat As Mustemilat) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM MUSTEMILAT WHERE ID=" & _Mustemilat.ID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("PARSEL_ID") = _Mustemilat.ParselID
                MyRow("SAHIP_ID") = _Mustemilat.SahipID
                MyRow("TANIM") = _Mustemilat.Tanim
                MyRow("ADET") = _Mustemilat.Adet
                MyRow("FIYAT") = _Mustemilat.Fiyat
                MyRow("MALIK") = _Mustemilat.Malik
                MyRow("PAY") = _Mustemilat.Pay
                MyRow("PAYDA") = _Mustemilat.Payda
                MyRow("ODEME_ID") = _Mustemilat.OdemeID
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Public Function UpdateMevsimlik(_Mevsimlik As Mevsimlik) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM MEVSIMLIK WHERE ID=" & _Mevsimlik.ID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("PARSEL_ID") = _Mevsimlik.ParselID
                MyRow("SAHIP_ID") = _Mevsimlik.SahipID
                MyRow("TANIM") = _Mevsimlik.Tanim
                MyRow("ALAN") = _Mevsimlik.Alan
                MyRow("BEDEL") = _Mevsimlik.Bedel
                MyRow("MALIK") = _Mevsimlik.Malik
                MyRow("PAY") = _Mevsimlik.Pay
                MyRow("PAYDA") = _Mevsimlik.Payda
                MyRow("ODEME_ID") = _Mevsimlik.OdemeID
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Public Function UpdateDavaTescil(_DavaTescil As DavaTescil) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM DAVA_10 WHERE ID=" & _DavaTescil.ID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("PARSEL_ID") = _DavaTescil.ParselID
                MyRow("MAHKEME") = _DavaTescil.Mahkeme
                MyRow("ESAS_NO") = _DavaTescil.EsasNo
                MyRow("KARAR_NO") = _DavaTescil.KararNo
                MyRow("KARAR_TARIHI") = _DavaTescil.KararTarihi
                MyRow("DAVA_ACILAN_HISSE_PAY") = _DavaTescil.DavaAcilanHissePay
                MyRow("DAVA_ACILAN_HISSE_PAYDA") = _DavaTescil.DavaAcilanHissePayda
                MyRow("TOPLAM_KAMULASTIRMA_BEDELI") = _DavaTescil.ToplamKamulastirmaBedeli
                MyRow("DAVA_TARIHI") = _DavaTescil.DavaTarihi
                MyRow("BIRINCI_KESIF_TARIHI") = _DavaTescil.KesifTarihi1
                MyRow("BIRINCI_DURUSMA_TARIHI") = _DavaTescil.DurusmaTarihi1
                MyRow("IKINCI_KESIF_TARIHI") = _DavaTescil.KesifTarihi2
                MyRow("IKINCI_DURUSMA_TARIHI") = _DavaTescil.DurusmaTarihi2
                MyRow("UCUNCU_DURUSMA_TARIHI") = _DavaTescil.DurusmaTarihi3
                MyRow("SON_DURUSMA_TARIHI") = _DavaTescil.DurusmaTarihiSon
                MyRow("BLOKE_OLURU") = _DavaTescil.BlokeOluru
                MyRow("OLUR_TARIHI") = _DavaTescil.OlurTarihi
                MyRow("BLOKE_TARIHI") = _DavaTescil.BlokeTarihi
                MyRow("AVUKAT") = _DavaTescil.Avukat
                MyRow("DUSUNCELER") = _DavaTescil.Dusunceler
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Public Function UpdateDavaAcele(_DavaAcele As DavaAcele) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM DAVA_27 WHERE ID=" & _DavaAcele.ID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("PARSEL_ID") = _DavaAcele.ParselID
                MyRow("MAHKEME") = _DavaAcele.Mahkeme
                MyRow("ESAS_NO") = _DavaAcele.EsasNo
                MyRow("KARAR_NO") = _DavaAcele.KararNo
                MyRow("KARAR_TARIHI") = _DavaAcele.KararTarihi
                MyRow("DAVA_ACILAN_HISSE_PAY") = _DavaAcele.DavaAcilanHissePay
                MyRow("DAVA_ACILAN_HISSE_PAYDA") = _DavaAcele.DavaAcilanHissePayda
                MyRow("TOPLAM_KAMULASTIRMA_BEDELI") = _DavaAcele.ToplamKamulastirmaBedeli
                MyRow("DAVA_TARIHI") = _DavaAcele.DavaTarihi
                MyRow("KESIF_TARIHI") = _DavaAcele.KesifTarihi
                MyRow("BLOKE_OLURU") = _DavaAcele.BlokeOluru
                MyRow("OLUR_TARIHI") = _DavaAcele.OlurTarihi
                MyRow("BLOKE_TARIHI") = _DavaAcele.BlokeTarihi
                MyRow("AVUKAT") = _DavaAcele.Avukat
                MyRow("DUSUNCELER") = _DavaAcele.Dusunceler
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Public Function UpdateOdeme(_Odeme As Odeme) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM ODEME WHERE ID=" & _Odeme.ID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("ODENEN_BEDEL") = _Odeme.Tutar
                If _Odeme.Tarih.Year > 1752 Then
                    MyRow("ODEME_TARIHI") = _Odeme.Tarih
                End If
                MyRow("ODEME_SEKLI") = _Odeme.Sekli
                MyRow("KAYNAK") = _Odeme.Kaynak
                MyRow("ODEME_DURUMU") = _Odeme.Durumu

                MyRow("PARSEL_ID") = _Odeme.ParselID
                MyRow("KISI_ID") = _Odeme.KisiID
                MyRow("ODEME_TIPI") = _Odeme.Tipi
                MyRow("ACIKLAMA") = _Odeme.Aciklama

            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

    Public Function UpdateOdeme(_Odeme As Odeme, _OnayID As Integer) As Boolean
        Dim MyStatus As Boolean = False
        Try
            Dim MyKamuConnection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            MyKamuConnection.Open()

            Dim MyQueryString As String = "SELECT * FROM ODEME WHERE ID=" & _Odeme.ID.ToString
            Dim MyDataAdapter As OleDbDataAdapter = New OleDbDataAdapter(New OleDbCommand(MyQueryString, MyKamuConnection))

            Dim MyTable As New DataTable
            MyDataAdapter.Fill(MyTable)

            Dim MyRows() As DataRow = MyTable.Select()

            For Each MyRow As DataRow In MyTable.Select
                MyRow("ONAY_ID") = _OnayID
            Next

            Dim MyCommandBuilder As New OleDbCommandBuilder
            MyCommandBuilder.DataAdapter = MyDataAdapter
            MyDataAdapter.Update(MyTable)

            MyTable = Nothing
            MyCommandBuilder = Nothing
            MyDataAdapter = Nothing
            MyKamuConnection.Close()
            MyKamuConnection = Nothing
            MyStatus = True
        Catch ex As Exception
            MyStatus = False
        End Try
        Return MyStatus
    End Function

#End Region

#Region "Delete Procedures"

    Public Function DeleteParsel(_ID As Long) As Boolean
        Dim MyStatus As Boolean = False
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "DELETE FROM PARSEL WHERE ID=" + _ID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()
                adapter.SelectCommand = command

                Dim table As New DataTable
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                adapter.Fill(table)

                table = Nothing
                adapter = Nothing
                command = Nothing
                MyStatus = True
            Catch ex As Exception
                MyStatus = False
            End Try
        End Using

        Return MyStatus
    End Function

    Public Function DeleteKisi(_ID As Long) As Boolean
        Dim MyStatus As Boolean = False
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "DELETE FROM KISI WHERE ID=" + _ID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()
                adapter.SelectCommand = command

                Dim table As New DataTable
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                adapter.Fill(table)

                table = Nothing
                adapter = Nothing
                command = Nothing
                MyStatus = True
            Catch ex As Exception
                MyStatus = False
            End Try
            Return MyStatus
        End Using

    End Function

    Public Function DeleteMustemilat(_ID As Long) As Boolean
        Dim MyStatus As Boolean = False
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "DELETE FROM MUSTEMILAT WHERE ID=" + _ID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()
                adapter.SelectCommand = command

                Dim table As New DataTable
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                adapter.Fill(table)

                table = Nothing
                adapter = Nothing
                command = Nothing
                MyStatus = True
            Catch ex As Exception
                MyStatus = False
            End Try
            Return MyStatus
        End Using
    End Function

    Public Function DeleteMevsimlik(_ID As Long) As Boolean
        Dim MyStatus As Boolean = False
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "DELETE FROM MEVSIMLIK WHERE ID=" + _ID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()
                adapter.SelectCommand = command

                Dim table As New DataTable
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                adapter.Fill(table)

                table = Nothing
                adapter = Nothing
                command = Nothing
                MyStatus = True
            Catch ex As Exception
                MyStatus = False
            End Try
        End Using
        Return MyStatus
    End Function

    Public Function DeleteMiras(_MurisID As Long, _VarisID As Long) As Boolean
        Dim MyStatus As Boolean = False
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "DELETE FROM MIRAS WHERE MURIS=" + _MurisID.ToString + " AND VARIS=" + _VarisID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()
                adapter.SelectCommand = command

                Dim table As New DataTable
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                adapter.Fill(table)

                table = Nothing
                adapter = Nothing
                command = Nothing
                MyStatus = True
            Catch ex As Exception
                MyStatus = False
            End Try
        End Using
        Return MyStatus
    End Function

    Public Function DeleteEmsal(_ParselID As Long, _EmsalID As Long) As Boolean
        Dim MyStatus As Boolean = False
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "DELETE FROM EMSAL WHERE PARSEL_ID=" + _ParselID.ToString + " AND EMSAL_ID=" + _EmsalID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()
                adapter.SelectCommand = command

                Dim table As New DataTable
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                adapter.Fill(table)

                table = Nothing
                adapter = Nothing
                command = Nothing
                MyStatus = True
            Catch ex As Exception
                MyStatus = False
            End Try
        End Using
        Return MyStatus
    End Function

    Public Function DeleteMalik(_ParselID As Long, _MalikID As Long) As Boolean
        Dim MyStatus As Boolean = False
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "DELETE FROM MULKIYET WHERE PARSEL_ID=" + _ParselID.ToString + " AND KISI_ID=" + _MalikID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()
                adapter.SelectCommand = command

                Dim table As New DataTable
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                adapter.Fill(table)

                table = Nothing
                adapter = Nothing
                command = Nothing
                MyStatus = True
            Catch ex As Exception
                MyStatus = False
            End Try
        End Using
        Return MyStatus
    End Function

    Public Function DeleteMalik(_MulkiyetID As Long) As Boolean
        Dim MyStatus As Boolean = False
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "DELETE FROM MULKIYET WHERE ID=" + _MulkiyetID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()
                adapter.SelectCommand = command

                Dim table As New DataTable
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                adapter.Fill(table)

                table = Nothing
                adapter = Nothing
                command = Nothing
                MyStatus = True
            Catch ex As Exception
                MyStatus = False
            End Try
        End Using
        Return MyStatus
    End Function

    Public Function DeleteOdeme(_ID As Long) As Boolean
        Dim MyStatus As Boolean = False
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "DELETE FROM ODEME WHERE ID=" + _ID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()
                adapter.SelectCommand = command

                Dim table As New DataTable
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                adapter.Fill(table)

                table = Nothing
                adapter = Nothing
                command = Nothing
                MyStatus = True
            Catch ex As Exception
                MyStatus = False
            End Try
        End Using
        Return MyStatus
    End Function

    Public Function DeleteOdemeBelge(_ID As Long) As Boolean
        Dim MyStatus As Boolean = False
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "DELETE FROM ODEME_BELGE WHERE ID=" + _ID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()
                adapter.SelectCommand = command

                Dim table As New DataTable
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                adapter.Fill(table)

                table = Nothing
                adapter = Nothing
                command = Nothing
                MyStatus = True
            Catch ex As Exception
                MyStatus = False
            End Try
        End Using
        Return MyStatus
    End Function


    Public Function DeleteDavaTescil(_ID As Long) As Boolean
        Dim MyStatus As Boolean = False
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "DELETE FROM DAVA_10 WHERE ID=" + _ID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()
                adapter.SelectCommand = command

                Dim table As New DataTable
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                adapter.Fill(table)

                table = Nothing
                adapter = Nothing
                command = Nothing
                MyStatus = True
            Catch ex As Exception
                MyStatus = False
            End Try
        End Using
        Return MyStatus
    End Function

    Public Function DeleteDavaAcele(_ID As Long) As Boolean
        Dim MyStatus As Boolean = False
        Using connection As New OleDbConnection(MyConnectionInfo.ConnectionString)
            Try
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "DELETE FROM DAVA_27 WHERE ID=" + _ID.ToString
                If Not connection.State = ConnectionState.Open Then connection.Open()

                Dim adapter As OleDbDataAdapter = New OleDbDataAdapter()
                adapter.SelectCommand = command

                Dim table As New DataTable
                table.Locale = System.Globalization.CultureInfo.InvariantCulture
                adapter.Fill(table)

                table = Nothing
                adapter = Nothing
                command = Nothing
                MyStatus = True
            Catch ex As Exception
                MyStatus = False
            End Try
        End Using
        Return MyStatus
    End Function
#End Region

End Class

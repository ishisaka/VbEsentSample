Imports Microsoft.Isam.Esent.Interop
Imports System.IO
Imports System.Text

''' <summary>
''' データベース操作クラス
''' </summary>
''' <remarks></remarks>
Public Class Database
    Private _instance As Instance
    Private _instancePath As String
    Private _databasePath As String
    Private Const DatabaseName As String = "Database"

    ''' <summary>
    ''' ESEのインスタンスの作成
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateInstance()
        _instancePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DatabaseName)
        _databasePath = Path.Combine(_instancePath, "database.edb")
        _instance = New Instance(_databasePath)

        'Configure Instance
        _instance.Parameters.CreatePathIfNotExist = True
        _instance.Parameters.TempDirectory = Path.Combine(_instancePath, "temp")
        _instance.Parameters.SystemDirectory = Path.Combine(_instancePath, "system")
        _instance.Parameters.LogFileDirectory = Path.Combine(_instancePath, "logs")
        _instance.Parameters.Recovery = True
        _instance.Parameters.CircularLog = True

        _instance.Init()
    End Sub

    ''' <summary>
    ''' ESEデータベースファイルの作成
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateDatabase()
        Using session = New Session(_instance)
            ' create database file
            Dim database As JET_DBID
            Api.JetCreateDatabase(session, _databasePath, Nothing, database, CreateDatabaseGrbit.OverwriteExisting)

            ' create database schema
            Using transaction = New Transaction(session)
                Dim tableid As JET_TABLEID
                Api.JetCreateTable(session, database, "Events", 1, 100, tableid)

                ' ID
                Dim columnid As JET_COLUMNID
                Api.JetAddColumn(session, tableid, "Id", New JET_COLUMNDEF() With { _
                    .cbMax = 16, _
                    .coltyp = JET_coltyp.Binary, _
                    .grbit = ColumndefGrbit.ColumnFixed Or ColumndefGrbit.ColumnNotNULL _
                }, Nothing, 0, _
                    columnid)
                ' Description
                Api.JetAddColumn(session, tableid, "Description", New JET_COLUMNDEF() With { _
                    .coltyp = JET_coltyp.LongText, _
                    .cp = JET_CP.Unicode, _
                    .grbit = ColumndefGrbit.None _
                }, Nothing, 0, _
                    columnid)
                ' Price
                Api.JetAddColumn(session, tableid, "Price", New JET_COLUMNDEF() With { _
                    .coltyp = JET_coltyp.IEEEDouble, _
                    .grbit = ColumndefGrbit.None _
                }, Nothing, 0, _
                    columnid)
                ' StartTime
                Api.JetAddColumn(session, tableid, "StartTime", New JET_COLUMNDEF() With { _
                    .coltyp = JET_coltyp.Currency, _
                    .grbit = ColumndefGrbit.None _
                }, Nothing, 0, _
                    columnid)

                ' Define table indices
                Dim indexDef = "+Id" & vbNullChar & vbNullChar
                Api.JetCreateIndex(session, tableid, "id_index", CreateIndexGrbit.IndexPrimary, indexDef, indexDef.Length, _
                    100)

                indexDef = "+Price" & vbNullChar & vbNullChar
                Api.JetCreateIndex(session, tableid, "price_index", CreateIndexGrbit.IndexDisallowNull, indexDef, indexDef.Length, _
                    100)

                transaction.Commit(CommitTransactionGrbit.None)
            End Using

            Api.JetCloseDatabase(session, database, CloseDatabaseGrbit.None)
            Api.JetDetachDatabase(session, _databasePath)
        End Using
    End Sub

    ''' <summary>
    ''' データベース操作
    ''' </summary>
    ''' <param name="dataFunc">トランザクション</param>
    ''' <returns>作業結果</returns>
    Private Function ExecuteInTransaction(dataFunc As Func(Of Session, Table, IList(Of Item))) As IList(Of Item)
        Dim results As IList(Of Item)
        Using session = New Session(_instance)
            Dim dbid As JET_DBID
            Api.JetAttachDatabase(session, _databasePath, AttachDatabaseGrbit.None)
            Api.JetOpenDatabase(session, _databasePath, [String].Empty, dbid, OpenDatabaseGrbit.None)
            Using transaction = New Transaction(session)
                Using table = New Table(session, dbid, "Events", OpenTableGrbit.None)
                    results = dataFunc(session, table)
                End Using

                transaction.Commit(CommitTransactionGrbit.None)
            End Using
        End Using

        Return results
    End Function

    ''' <summary>
    ''' Itemをテーブルに追加する
    ''' </summary>
    ''' <param name="ev">Itemオブジェクト</param>
    Public Sub AddEvent(ev As Item)
        ExecuteInTransaction(Function(session, table)
                                 Using updater = New Update(session, table, JET_prep.Insert)
                                     Dim columnId = Api.GetTableColumnid(session, table, "Id")
                                     Api.SetColumn(session, table, columnId, ev.Id)

                                     Dim columnDesc = Api.GetTableColumnid(session, table, "Description")
                                     Api.SetColumn(session, table, columnDesc, ev.Description, Encoding.Unicode)

                                     Dim columnPrice = Api.GetTableColumnid(session, table, "Price")
                                     Api.SetColumn(session, table, columnPrice, ev.Price)

                                     Dim columnStartTime = Api.GetTableColumnid(session, table, "StartTime")
                                     Api.SetColumn(session, table, columnStartTime, DateTime.Now.Ticks)

                                     updater.Save()
                                 End Using
                                 Return Nothing

                             End Function)
    End Sub

    ''' <summary>
    ''' 指定されたIdのデータをテーブルから削除する
    ''' </summary>
    ''' <param name="id">Id</param>
    Public Sub Delete(id As Guid)
        ExecuteInTransaction(Function(session, table)
                                 Api.JetSetCurrentIndex(session, table, Nothing)
                                 Api.MakeKey(session, table, id, MakeKeyGrbit.NewKey)
                                 If Api.TrySeek(session, table, SeekGrbit.SeekEQ) Then
                                     Api.JetDelete(session, table)
                                 End If
                                 Return Nothing

                             End Function)
    End Sub

    ''' <summary>
    ''' 全てのデータを取得する
    ''' </summary>
    ''' <returns>Itemのリスト</returns>
    Public Function GetAllEvents() As IList(Of Item)
        Return ExecuteInTransaction(Function(session, table)
                                        Dim results = New List(Of Item)()
                                        If Api.TryMoveFirst(session, table) Then
                                            Do
                                                results.Add(GetEvent(session, table))
                                            Loop While Api.TryMoveNext(session, table)
                                        End If
                                        Return results

                                    End Function)
    End Function

    ''' <summary>
    ''' Itemを取得する
    ''' </summary>
    ''' <param name="session">Sessionオブジェクト</param>
    ''' <param name="table">Tableオブジェクト</param>
    ''' <returns>Itemオブジェクト</returns>
    Private Function GetEvent(session As Session, table As Table) As Item
        Dim ev = New Item()

        Dim columnId = Api.GetTableColumnid(session, table, "Id")
        ev.Id = If(Api.RetrieveColumnAsGuid(session, table, columnId), Guid.Empty)

        Dim columnDesc = Api.GetTableColumnid(session, table, "Description")
        ev.Description = Api.RetrieveColumnAsString(session, table, columnDesc, Encoding.Unicode)

        Dim columnPrice = Api.GetTableColumnid(session, table, "Price")
        ev.Price = If(Api.RetrieveColumnAsDouble(session, table, columnPrice), 0)

        Dim columnStartTime = Api.GetTableColumnid(session, table, "StartTime")
        Dim ticks = Api.RetrieveColumnAsInt64(session, table, columnStartTime)
        If ticks.HasValue Then
            ev.StartTime = New DateTime(ticks.Value)
        End If

        Return ev
    End Function

    ''' <summary>
    ''' Idを指定してItemを取得する
    ''' </summary>
    ''' <param name="id">Id</param>
    ''' <returns>Itemオブジェクト</returns>
    Public Function GetEventsById(id As Guid) As IList(Of Item)
        Return ExecuteInTransaction(Function(session, table)
                                        Dim results = New List(Of Item)()
                                        Api.JetSetCurrentIndex(session, table, Nothing)
                                        Api.MakeKey(session, table, id, MakeKeyGrbit.NewKey)
                                        If Api.TrySeek(session, table, SeekGrbit.SeekEQ) Then
                                            results.Add(GetEvent(session, table))
                                        End If
                                        Return results

                                    End Function)
    End Function

    ''' <summary>
    ''' Priceの範囲でデータを取得する
    ''' </summary>
    ''' <param name="minPrice">最小値</param>
    ''' <param name="maxPrice">最大値</param>
    ''' <returns>Itemオブジェクトのリスト</returns>
    Public Function GetEventsForPriceRange(minPrice As Double, maxPrice As Double) As IList(Of Item)
        Return ExecuteInTransaction(Function(session, table)
                                        Dim results = New List(Of Item)()

                                        Api.JetSetCurrentIndex(session, table, "price_index")
                                        Api.MakeKey(session, table, minPrice, MakeKeyGrbit.NewKey)

                                        If Api.TrySeek(session, table, SeekGrbit.SeekGE) Then
                                            Api.MakeKey(session, table, maxPrice, MakeKeyGrbit.NewKey)
                                            Api.JetSetIndexRange(session, table, SetIndexRangeGrbit.RangeUpperLimit Or SetIndexRangeGrbit.RangeInclusive)

                                            Do
                                                results.Add(GetEvent(session, table))
                                            Loop While Api.TryMoveNext(session, table)
                                        End If
                                        Return results

                                    End Function)
    End Function
End Class

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================

Imports System.Diagnostics
Imports System.Linq
Imports Microsoft.Isam.Esent.Collections.Generic

Module Module1

    Sub Main()
        ' データベースがすでに存在している場合には削除する
        If PersistentDictionaryFile.Exists("PlantLogs") Then
            PersistentDictionaryFile.DeleteFiles("PlantLogs")
        End If

        ' ｎ件追加する。
        Dim logs = New PersistentDictionary(Of Integer, PlantLog)("PlantLogs")

        Dim n As Integer = 1000
        Console.WriteLine("{0}件データ追加", n)
        Dim watch = Stopwatch.StartNew()
        Enumerable.Range(1, n).ToList().ForEach(Sub(i)
                                                    Dim nowTime = DateTime.Now
                                                    Dim log As PlantLog
                                                    log.AleartTime = nowTime
                                                    log.AlarmLevel = i
                                                    log.Area = i + 1
                                                    log.Description = "Message " + i.ToString()
                                                    'Addメソッドだと例外が発生するので、以下の構文で。
                                                    logs(i) = log

                                                End Sub)
        Console.WriteLine("データ追加経過 {0}ms", watch.ElapsedMilliseconds)
        watch.[Stop]()
        watch.Reset()

        ' n件取得する。
        Console.WriteLine("{0}件データ取得", n)
        watch.Start()
        Dim plantLogs = logs.[Select](Function(l) l.Value).ToList()
        Console.WriteLine("データ取得経過時間 {0}ms", watch.ElapsedMilliseconds)
        watch.[Stop]()

        ' 件数の確認。
        Dim q = logs.ToArray()
        Console.WriteLine("{0}件", q.Length)

        Console.Write("Enterをおしてください。>")
        Console.ReadLine()
    End Sub

End Module


''' <summary>
''' 構造体はSerializable属性を付ける。
''' 構造体も以下のクラスとInt32やDoubleといったプリミティブな値型のみが許される
''' string, Uri, IPAddress
''' </summary>
<Serializable> _
Public Structure PlantLog
    Public AleartTime As Date
    Public AlarmLevel As Integer
    Public Area As Integer
    Public Description As String
    ' 以下のような配列も使用できない
    ' public byte[] Data;
End Structure

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================
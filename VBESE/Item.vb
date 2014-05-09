''' <summary>
''' テーブル構造のサンプル
''' </summary>
''' <remarks></remarks>
Public Class Item
    Public Property Id As Guid
    Public Property Description As String
    Public Property Price As Double
    Public Property StartTIme As DateTime

    Public Overrides Function ToString() As String
        Return String.Format("{0}, {1}, {2:F2}, {3}", Id, Description, Price, StartTIme)
    End Function
End Class

Public Class Connection
    Private cn As System.Data.SqlClient.SqlConnection

    Public Function Connect()
        Try
            '接続する端末名
            Dim ServerName As String = "192.168.0.173"
            '接続するデータベース名
            Dim DBName As String = "master"
            'ユーザ名
            Dim UserID As String = "sa"
            'パスワード
            Dim Password As String = "SaPassword2017"

            cn = New System.Data.SqlClient.SqlConnection()
            'SQL Server認証を利用して接続
            cn.ConnectionString =
         "Data Source = " & ServerName &
         ";Initial Catalog = " & DBName &
         ";User ID = " & UserID &
         ";Password = " & Password

            cn.Open()

            Return cn

        Catch ex As Exception
            Console.WriteLine("Error! {0}", ex.Message)
            Return Nothing
        End Try
    End Function
End Class

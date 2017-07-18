Public Class DbConnectionTest
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cn As System.Data.SqlClient.SqlConnection

        Try

            Dim serverName As String = "192.168.0.173"
            Dim dbName As String = "master"
            Dim userName As String = "sa"
            Dim password As String = "SaPassword2017"

            cn = New System.Data.SqlClient.SqlConnection()

            cn.ConnectionString =
                "Data Source = " & serverName &
                ";Initial Catalog = " & dbName &
                ";User ID = " & userName &
                ";Password = " & password

            cn.Open()

            Label1.Text = "接続完了"

        Catch ex As Exception

            Label1.Text = ex.Message

        Finally

            If cn.State <> ConnectionState.Closed Then

                cn.Close()

            End If

            cn.Dispose()


        End Try
    End Sub
End Class
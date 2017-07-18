Public Class Login
    Public Shared cn As System.Data.SqlClient.SqlConnection

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.PasswordChar = "*"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim name As String = TextBox1.Text
        Dim pass As String = TextBox2.Text


        Try

            Connect()

            Dim cmd As New SqlClient.SqlCommand

            'コネクションの指定
            cmd.Connection = cn
            'コマンドの種類をテキストにする（省略可）
            cmd.CommandType = CommandType.Text
            '実行するSQLを指定
            cmd.CommandText = "SELECT USER_NAME,PASSWORD FROM M_USER WHERE USER_NAME = @name"
            cmd.Parameters.AddWithValue("@name", name)

            'SQLの結果を取得する
            Dim sr As SqlClient.SqlDataReader
            sr = cmd.ExecuteReader()

            Dim cReader As System.Data.SqlClient.SqlDataReader = sr

            cmd.Dispose()
            Dim sqlname As String = String.Empty
            Dim sqlpass As String = String.Empty
            '入力した氏名が正しいか判断
            If (cReader.Read()) Then
                ' 列名を元に値を取得する
                sqlname = cReader("USER_NAME").ToString()
                sqlpass = cReader("PASSWORD").ToString()
                

                '氏名があっていたらパスワードがあっているか判断
                If pass = sqlpass Then

                    '氏名、パスがあっていたらメニューに遷移
                    MainMenu.Show()
                    My.Application.ApplicationContext.MainForm = MainMenu
                    Me.Close()

                Else
                    MsgBox("パスワードが正しくありません")
                End If

            Else
                MsgBox("入力された氏名が正しくありません")
            End If

        Catch ex As Exception

            MsgBox("エラー")

        Finally
            Disconnect()
        End Try
    End Sub

    Private Sub Connect()

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

    End Sub

    Private Sub Disconnect()
        'コネクションを閉じ、リソースを解放する
        If cn.State <> ConnectionState.Closed Then
            cn.Close()
        End If
        cn.Dispose()
    End Sub

End Class
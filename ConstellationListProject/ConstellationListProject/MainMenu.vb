Imports System.Data.SqlClient

Public Class MainMenu
    Public Shared cn As System.Data.SqlClient.SqlConnection

    Private length As Integer
    Private count As Integer = 0
    Private conlist As New List(Of String)
    Private ab As New List(Of String)
    Private pageCount As Integer = 1
    Private page As Double
    Private forCount As Integer


    Private Sub MainMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For i = 1 To 17
            ab.Add(i)
        Next

        'DB接続
        DbCon()

        Dim cmd As New SqlClient.SqlCommand

        '現在月取得
        Dim dtNow = GetMonth()
        Console.WriteLine(dtNow)

        '季節コード取得
        Dim seasonCode = FindSeasonalCode(dtNow)

        Console.WriteLine(seasonCode)

        'Label1表示処理
        DispLabel1(seasonCode)

        Try
            DbCon()

            'コネクションの指定
            cmd.Connection = cn
            'コマンドの種類をテキストにする（省略可）
            cmd.CommandType = CommandType.Text
            '実行するSQLを指定
            cmd.CommandText = "SELECT CONSTELLATION_NAME FROM M_CONSTELLATION WHERE SEASON_ID = @id"
            cmd.Parameters.AddWithValue("@id", seasonCode)

            'SQLの結果を取得する
            Dim sr As SqlClient.SqlDataReader
            sr = cmd.ExecuteReader()

            Dim cReader As System.Data.SqlClient.SqlDataReader = sr

            cmd.Dispose()

            While (cReader.Read())

                conlist.Add(cReader("CONSTELLATION_NAME"))
                Label2.Text = Label2.Text + cReader("CONSTELLATION_NAME")
            End While
            'ページの枚数を計算
            page = Math.Ceiling(conlist.Count / 5)
            'リストのサイズを取得
            length = conlist.Count
            'ページ数表示
            pageNum()
            'listを表示
            viewlist(conlist)


        Catch ex As Exception

            MsgBox("エラー")

        Finally

            DbDisCon()

        End Try




    End Sub
    Private Sub DbCon()

        '端末
        Dim ServerName As String = "192.168.0.173"
        'DB名
        Dim DBName As String = "master"
        'ユーザー名
        Dim UserId As String = "sa"
        'パスワード
        Dim Password As String = "SaPassword2017"

        cn = New System.Data.SqlClient.SqlConnection

        cn.ConnectionString =
        "Data Source = " & ServerName &
        ";Initial Catalog = " & DBName &
        ";User ID = " & UserId &
        ";Password = " & Password
        cn.Open()

        Console.WriteLine("{0}の{1}に接続しました", ServerName, DBName)

    End Sub
    Private Sub DbDisCon()

        If cn.State <> ConnectionState.Closed Then
            cn.Close()
        End If
        cn.Dispose()

    End Sub
    Private Function GetMonth() As Integer
        Dim dtNow As DateTime = DateTime.Now
        Return dtNow.Month
    End Function
    Private Function FindSeasonalCode(ByVal dtNow As Integer) As String
        Dim seasonCode As String = ""

        Select Case dtNow
            Case 3 To 5
                seasonCode = "0"
            Case 6 To 8
                seasonCode = "1"
            Case 9 To 11
                seasonCode = "2"
            Case 12, 1, 2
                seasonCode = "3"

        End Select

        Return seasonCode

    End Function
    Private Sub DispLabel1(ByRef seasonCode)
        Select Case seasonCode
            Case "0"
                Label1.Text = "春の星座"
            Case "1"
                Label1.Text = "夏の星座"
            Case "2"
                Label1.Text = "秋の星座"
            Case "3"
                Label1.Text = "冬の星座"
        End Select
    End Sub
    'ページ数に応じたリストの内容表示メソッド
    Private Sub viewlist(list As List(Of String))

        If count < length And count >= 0 Then

            For i As Integer = 1 To 5
                TextBox1.Text = TextBox1.Text + list(count) & vbCrLf
                count = count + 1
                forCount = i

                If length <= count Then

                    Exit For

                End If

            Next

        End If

    End Sub
    'ページ数を表示するメソッド
    Private Sub pageNum()

        Label2.Text = pageCount & "/" & page
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '星座検索画面にとぶ
        Dim searchForm As New Search
        searchForm.Show()
        My.Application.ApplicationContext.MainForm = searchForm
        Me.Dispose()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        MaintenanceMenu.Show()
        My.Application.ApplicationContext.MainForm = MaintenanceMenu
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If pageCount < page Then
            TextBox1.Text = ""
            pageCount = pageCount + 1
            forCount = 0
            pageNum()

            viewlist(conlist)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If pageCount > 1 Then
            TextBox1.Text = ""
            If pageCount = page Then
                count = count - forCount - 5
            Else
                count = count - 10
            End If
            pageCount = pageCount - 1
            pageNum()
            viewlist(conlist)
        End If
    End Sub


End Class
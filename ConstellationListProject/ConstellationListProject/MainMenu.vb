Public Class MainMenu
    Private Sub MainMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'DB接続
        DbCon()

        '現在月取得
        'Dim dtNow = GetMonth()
        Dim dtNow = 11
        Console.WriteLine(dtNow)

        '季節コード取得
        Dim seasonCode = FindSeasonalCode(dtNow)
        Console.WriteLine(seasonCode)

        'Label1表示処理
        DispLabel1(seasonCode)



    End Sub
    Private Sub DbCon()
        Dim cn As System.Data.SqlClient.SqlConnection
        Try
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
        Catch ex As Exception
            Console.WriteLine("Error! {0}", ex.Message)
        Finally
            'コネクションを閉じリソースを開放する
            If cn.State <> ConnectionState.Closed Then
                cn.Close()
            End If
            cn.Dispose()
        End Try
    End Sub
    Private Function GetMonth() As Integer
        Dim dtNow As DateTime = DateTime.Now
        Return dtNow.Month
    End Function
    Private Function FindSeasonalCode(ByVal dtNow As Integer) As String
        Dim seasonCode As String = ""

        Select Case dtNow
            Case 3 To 5
                seasonCode = "1"
            Case 6 To 8
                seasonCode = "2"
            Case 9 To 11
                seasonCode = "3"
            Case 12, 1, 2
                seasonCode = "4"

        End Select

        Return seasonCode

    End Function
    Private Sub DispLabel1(ByRef seasonCode)
        Select Case seasonCode
            Case "1"
                Label1.Text = "春の星座"
            Case "2"
                Label1.Text = "夏の星座"
            Case "3"
                Label1.Text = "秋の星座"
            Case "4"
                Label1.Text = "冬の星座"
        End Select
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
End Class
Public Class ConstellationDataAdder

    Dim endType As Integer
    Dim cn As System.Data.SqlClient.SqlConnection
    Private Sub DataAdder_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        endType = 1

        MasterMaintenance.Label2.Text = "星座"

        Me.Close()

    End Sub

    Private Sub DataAdder_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        If endType = 0 Then

        ElseIf endType = 1 Then

            MasterMaintenance.Show()

            My.Application.ApplicationContext.MainForm = MasterMaintenance

        ElseIf endType = 2 Then

            DoneMaintenance.Show()

            My.Application.ApplicationContext.MainForm = DoneMaintenance

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try

            If lengthCheck() And comboCheck() Then

                connectDb()

                insertProcess(getNextConstellationId())

                endType = 2

                Me.Close()

            End If


        Catch ex As Exception

            MsgBox(ex.Message())

        End Try
    End Sub

    Function lengthCheck() As Boolean

        Dim flag As Boolean = True

        If TextBox1.Text.Length > 45 Or TextBox1.Text.Length = 0 Then

            MsgBox("星座名は全角15文字以内で入力してください")
            flag = False

        End If

        Return flag

    End Function

    Function comboCheck() As Boolean

        Dim flag As Boolean = True

        If ComboBox1.Text.Length = 0 Then

            MsgBox("季節を選択してください")
            flag = False

        End If

        Return flag

    End Function

    Sub connectDb()

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

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try

    End Sub

    Function getNextConstellationId() As String

        Try

            Dim cmd As New System.Data.SqlClient.SqlCommand
            Dim sdr As SqlClient.SqlDataReader

            Dim sqlStr As String
            Dim nextId As String

            cmd.Connection = cn
            cmd.CommandType = CommandType.Text

            sqlStr = "SELECT MAX(CONSTELLATION_ID) + 1 FROM M_CONSTELLATION"

            cmd.CommandText = sqlStr

            sdr = cmd.ExecuteReader()

            cmd.Dispose()

            sdr.Read()

            nextId = sdr(0).ToString()

            sdr.Close()

            If nextId.Length = 1 Then

                Return 0 & nextId

            Else

                Return nextId

            End If

        Catch ex As Exception

            MsgBox(ex.Message())
            Throw New Exception

        End Try

    End Function

    Sub insertProcess(ByVal id As String)

        Dim cmd As New System.Data.SqlClient.SqlCommand
        Dim sqlStr As String

        cmd.Connection = cn

        cmd.CommandType = CommandType.Text

        sqlStr = "INSERT INTO M_CONSTELLATION(CONSTELLATION_ID, CONSTELLATION_NAME,
                    SEASON_ID, ZODIAC_FLAG, DELETE_FLAG, CREATE_DATE, CREATE_USER)"
        sqlStr += "VALUES('" & id & "','" & TextBox1.Text & "','" & getSeasonId()
        sqlStr += "', 0, 0, GETDATE(), '00')"

        cmd.CommandText = sqlStr

        cmd.ExecuteNonQuery()

        cmd.Dispose()

    End Sub

    Function getSeasonId() As Integer

        Dim seasonId As Integer

        If ComboBox1.Text = "春" Then

            seasonId = 0

        ElseIf ComboBox1.Text = "夏" Then

            seasonId = 1

        ElseIf ComboBox1.Text = "秋" Then

            seasonId = 2

        ElseIf ComboBox1.Text = "冬" Then

            seasonId = 3

        End If

        Return seasonId

    End Function
End Class
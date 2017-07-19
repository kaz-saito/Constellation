Public Class DataUpdater
    Dim maintenanceType As Integer
    Dim endType As Integer
    Dim cn As System.Data.SqlClient.SqlConnection

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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        endType = 1

        If maintenanceType = 0 Then

            MasterMaintenance.Label2.Text = "ユーザー"

        Else

            MasterMaintenance.Label2.Text = "星座"

        End If

        Me.Close()

    End Sub

    Private Sub DataUpdater_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        connectDb()

        If Label2.Text = "ユーザー" Then

            maintenanceType = 0

        Else

            maintenanceType = 1

        End If

    End Sub

    Private Sub DataUpdater_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        If endType = 0 Then

        ElseIf endType = 1 Then

            MasterMaintenance.Show()

            My.Application.ApplicationContext.MainForm = MasterMaintenance

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        If TextBox1.Text.Length > 2 Then

            MsgBox("IDは2桁で入力してください")

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        searchData()

        createRadioAndTextbox()

    End Sub

    Sub searchData()

        Dim cmd As New System.Data.SqlClient.SqlCommand
        Dim sdr As SqlClient.SqlDataReader
        Dim sqlStr As String

        cmd.Connection = cn
        cmd.CommandType = CommandType.Text

        If TextBox1.Text.Length = 0 Then

            sqlStr = "SELECT USER_ID, USER_NAME FROM M_USER WHERE USER_NAME LIKE'%" & TextBox2.Text & "%'"

        Else

            sqlStr = "SELECT USER_ID, USER_NAME FROM M_USER WHERE USER_ID = '" & TextBox1.Text & "'"

        End If

        cmd.CommandText = sqlStr

        sdr = cmd.ExecuteReader()

        While sdr.Read

            MsgBox(sdr(0) & "AND" & sdr(1))

        End While

        cmd.Dispose()

        sdr.Close()

    End Sub

    Sub createRadioAndTextbox()

        Dim radioVertical As Integer = 49
        Dim textVertical As Integer = 110
        Dim horizontal As Integer = 181

        For i = 0 To 2

            Dim radio As New System.Windows.Forms.RadioButton
            Dim textbox As New System.Windows.Forms.TextBox

            radio.Name = "addRadio" & i
            radio.Size = New Size(14, 13)
            radio.Location = New Point(radioVertical, horizontal)
            radio.Text = "00"

            TextBox.Name = "addText" & i
            TextBox.Location = New Point(textVertical, horizontal)
            TextBox.Text = "ooo"

            Me.Controls.Add(radio)
            Me.Controls.Add(textbox)

            horizontal += 42

        Next

    End Sub
End Class
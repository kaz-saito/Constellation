Public Class DataUpdater
    Dim maintenanceType As Integer
    Dim endType As Integer
    Dim cn As System.Data.SqlClient.SqlConnection

    Dim nowPage As System.Windows.Forms.Label

    Dim idList As New List(Of String)
    Dim nameList As New List(Of String)

    Dim adderRadio(2) As System.Windows.Forms.RadioButton
    Dim adderTextbox(2) As System.Windows.Forms.TextBox

    Dim totalPage As Integer

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

        createFooter()

        setData()

    End Sub

    Sub searchData()

        Dim cmd As New System.Data.SqlClient.SqlCommand
        Dim sdr As SqlClient.SqlDataReader
        Dim sqlStr As String

        Dim pageCount As Integer

        cmd.Connection = cn
        cmd.CommandType = CommandType.Text

        If TextBox1.Text.Length = 0 Then

            If maintenanceType = 0 Then

                sqlStr = "SELECT G_ROW.USER_ID, G_ROW.USER_NAME FROM("
                sqlStr += "SELECT USER_ID, USER_NAME, ROW_NUMBER() OVER(ORDER BY USER_ID ASC) AS ROW FROM M_USER WHERE USER_NAME LIKE'%" & TextBox2.Text & "%'"
                sqlStr += ")G_ROW "
                sqlStr += "WHERE G_ROW.ROW BETWEEN 1 AND 3"

            Else

                sqlStr = "SELECT G_ROW.CONSTELLATION_ID, G_ROW.CONSTELLATION_NAME FROM("
                sqlStr += "SELECT CONSTELLATION_ID, CONSTELLATION_NAME, ROW_NUMBER() OVER(ORDER BY CONSTELLATION_ID ASC) AS ROW FROM M_CONSTELLATION WHERE CONSTELLATION_NAME LIKE'%" & TextBox2.Text & "%'"
                sqlStr += ")G_ROW "
                sqlStr += "WHERE G_ROW.ROW BETWEEN 1 AND 3"

            End If

        Else

            If maintenanceType = 0 Then

                sqlStr = "SELECT G_ROW.USER_ID, G_ROW.USER_NAME FROM("
                sqlStr += "SELECT USER_ID, USER_NAME, ROW_NUMBER() OVER(ORDER BY USER_ID ASC) AS ROW FROM M_USER WHERE USER_ID = '" & TextBox1.Text & "'"
                sqlStr += ")G_ROW "
                sqlStr += "WHERE G_ROW.ROW BETWEEN 1 AND 3"

            Else

                sqlStr = "SELECT G_ROW.CONSTELLATION_ID, G_ROW.CONSTELLATIOIN_NAME FROM("
                sqlStr += "SELECT CONSTELLATION_ID, CONSTELLATION_NAME, ROW_NUMBER() OVER(ORDER BY CONSTELLATION_ID ASC) AS ROW FROM M_CONSTELLATION WHERE CONSTELLATION_ID = '" & TextBox1.Text & "'"
                sqlStr += ")G_ROW "
                sqlStr += "WHERE G_ROW.ROW BETWEEN 1 AND 3"

            End If

        End If

        cmd.CommandText = sqlStr

        sdr = cmd.ExecuteReader()

        While sdr.Read

            idList.Add(sdr(0).ToString())
            nameList.Add(sdr(1).ToString())

            pageCount += 1

        End While

        totalPage = Math.Ceiling(pageCount / 3)

        cmd.Dispose()

        sdr.Close()

    End Sub

    Sub createRadioAndTextbox()

        Dim radioVertical As Integer = 49
        Dim textVertical As Integer = 110
        Dim radioHorizontal As Integer = 165
        Dim textHorizontal As Integer = 175

        For i = 0 To 2

            Dim radio As New System.Windows.Forms.RadioButton
            Dim textbox As New System.Windows.Forms.TextBox

            radio.Name = "addRadio" & i
            radio.Size = New Size(45, 45)
            radio.Location = New Point(radioVertical, radioHorizontal)

            adderRadio(i) = radio

            textbox.Name = "addText" & i
            textbox.Location = New Point(textVertical, textHorizontal)

            adderTextbox(i) = textbox

            Me.Controls.Add(radio)
            Me.Controls.Add(textbox)

            radioHorizontal += 42
            textHorizontal += 42

        Next

    End Sub

    Sub createFooter()

        Dim backButton As New System.Windows.Forms.Button
        Dim nextButton As New System.Windows.Forms.Button
        Dim child As New System.Windows.Forms.Label
        Dim slash As New System.Windows.Forms.Label
        Dim mother As New System.Windows.Forms.Label
        Dim update As New System.Windows.Forms.Button

        Dim buttonSizeV As Integer = 27
        Dim buttonSizeH As Integer = 23

        Dim buttonHorizontal As Integer = 295
        Dim labelHorizontal As Integer = 297

        Dim labelSizeV As Integer = 14
        Dim labelsizeH As Integer = 13


        backButton.Name = "backButton"
        backButton.Text = "＜"
        backButton.Size = New Size(buttonSizeV, buttonSizeH)
        backButton.Location = New Point(104, buttonHorizontal)

        nextButton.Name = "nextButton"
        nextButton.Text = "＞"
        nextButton.Size = New Size(buttonSizeV, buttonSizeH)
        nextButton.Location = New Point(216, buttonHorizontal)

        child.Name = "child"
        child.Text = 1
        child.Size = New Size(labelSizeV, labelsizeH)
        child.Location = New Point(147, labelHorizontal)

        slash.Name = "slash"
        slash.Text = "/"
        slash.Size = New Size(labelSizeV, labelsizeH)
        slash.Location = New Point(168, labelHorizontal)

        mother.Name = "mother"
        mother.Text = totalPage
        mother.Size = New Size(labelSizeV, labelsizeH)
        mother.Location = New Point(189, labelHorizontal)

        update.Name = "updateButton"
        update.Text = "更新"
        update.Size = New Size(56, 35)
        update.Location = New Point(248, 219)

        Me.Controls.Add(nextButton)
        Me.Controls.Add(backButton)
        Me.Controls.Add(child)
        Me.Controls.Add(slash)
        Me.Controls.Add(mother)
        Me.Controls.Add(update)

        nowPage = child

    End Sub

    Sub setData()

        For i = 0 To 2

            adderRadio(i).Text = idList(i)
            adderTextbox(i).Text = nameList(i)

        Next

    End Sub

End Class
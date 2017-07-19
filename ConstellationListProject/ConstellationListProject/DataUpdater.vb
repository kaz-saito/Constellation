Public Class DataUpdater
    Dim maintenanceType As Integer
    Dim endType As Integer
    Dim cn As System.Data.SqlClient.SqlConnection

    Dim nowPageLabel As System.Windows.Forms.Label
    Dim totalPageLabel As System.Windows.Forms.Label

    Dim nowPage As Integer = 1
    Dim totalPage As Integer = 1

    Dim idList As New List(Of String)
    Dim nameList As New List(Of String)

    Dim backButton As System.Windows.Forms.Button
    Dim nextButton As System.Windows.Forms.Button

    Dim adderRadio(2) As System.Windows.Forms.RadioButton
    Dim adderTextbox(2) As System.Windows.Forms.TextBox

    Dim totalData As Integer

    Dim createRadioFlag As Boolean

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

        ElseIf endType = 2 Then

            DoneMaintenance.Show()

            My.Application.ApplicationContext.MainForm = DoneMaintenance

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        If TextBox1.Text.Length > 2 Then

            MsgBox("IDは2桁で入力してください")

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        resetRadioAndTextbox()

        getTotalData()

        searchData()

        createFooter()

        createRadioAndTextbox()

        setData()

        nowPage = 1
        nowPageLabel.Text = nowPage

        totalPageLabel.Text = totalPage

        If totalPage = 1 Then

            nextButton.Enabled = False

        Else

            nextButton.Enabled = True

        End If

        backButton.Enabled = False

    End Sub

    Sub getTotalData()

        Dim cmd As New SqlClient.SqlCommand
        Dim sdr As SqlClient.SqlDataReader
        Dim sqlStr As String
        Dim dataCount As Integer

        cmd.Connection = cn
        cmd.CommandType = CommandType.Text

        If TextBox1.Text.Length = 0 Then

            If maintenanceType = 0 Then

                sqlStr = "SELECT COUNT(G_ROW.USER_ID) FROM(SELECT USER_ID FROM M_USER WHERE USER_NAME LIKE '%" & TextBox2.Text & "%'"
                sqlStr += ")G_ROW"

            Else

                sqlStr = "SELECT COUNT(G_ROW.CONSTELLATION_ID) FROM(SELECT CONSTELLATION_ID FROM M_CONSTELLATION WHERE CONSTELLATION_NAME LIKE '%" & TextBox2.Text & "%')G_ROW"

            End If

        Else

            totalData = 1

        End If

        cmd.CommandText = sqlStr

        sdr = cmd.ExecuteReader()

        sdr.Read()

        dataCount = sdr(0)

        cmd.Dispose()
        sdr.Close()

        totalData = dataCount

    End Sub

    Sub resetRadioAndTextbox()

        If createRadioFlag Then

            idList = New List(Of String)
            nameList = New List(Of String)

            For i = 0 To 2

                adderRadio(i).Text = ""
                adderTextbox(i).Text = ""

            Next

        End If

    End Sub

    Sub searchData()

        Dim cmd As New System.Data.SqlClient.SqlCommand
        Dim sdr As SqlClient.SqlDataReader
        Dim sqlStr As String

        Dim endData As Integer = nowPage * 3
        Dim startData As Integer = endData - 2

        cmd.Connection = cn
        cmd.CommandType = CommandType.Text

        If TextBox1.Text.Length = 0 Then

            If maintenanceType = 0 Then

                sqlStr = "SELECT G_ROW.USER_ID, G_ROW.USER_NAME FROM("
                sqlStr += "SELECT USER_ID, USER_NAME, ROW_NUMBER() OVER(ORDER BY USER_ID ASC) AS ROW FROM M_USER WHERE USER_NAME Like'%" & TextBox2.Text & "%'"
                sqlStr += ")G_ROW "
                sqlStr += "WHERE G_ROW.ROW BETWEEN " & startData & " AND " & endData

            Else

                sqlStr = "SELECT G_ROW.CONSTELLATION_ID, G_ROW.CONSTELLATION_NAME FROM("
                sqlStr += "SELECT CONSTELLATION_ID, CONSTELLATION_NAME, ROW_NUMBER() OVER(ORDER BY CONSTELLATION_ID ASC) AS ROW FROM M_CONSTELLATION WHERE CONSTELLATION_NAME LIKE'%" & TextBox2.Text & "%'"
                sqlStr += ")G_ROW "
                sqlStr += "WHERE G_ROW.ROW BETWEEN " & startData & " AND " & endData

            End If

        Else

            If maintenanceType = 0 Then

                sqlStr = "SELECT G_ROW.USER_ID, G_ROW.USER_NAME FROM("
                sqlStr += "SELECT USER_ID, USER_NAME, ROW_NUMBER() OVER(ORDER BY USER_ID ASC) AS ROW FROM M_USER WHERE USER_ID = '" & TextBox1.Text & "'"
                sqlStr += ")G_ROW "
                sqlStr += "WHERE G_ROW.ROW BETWEEN " & startData & " AND " & endData

            Else

                sqlStr = "SELECT G_ROW.CONSTELLATION_ID, G_ROW.CONSTELLATION_NAME FROM("
                sqlStr += "SELECT CONSTELLATION_ID, CONSTELLATION_NAME, ROW_NUMBER() OVER(ORDER BY CONSTELLATION_ID ASC) AS ROW FROM M_CONSTELLATION WHERE CONSTELLATION_ID = '" & TextBox1.Text & "'"
                sqlStr += ")G_ROW "
                sqlStr += "WHERE G_ROW.ROW BETWEEN " & startData & " AND " & endData

            End If

        End If

        cmd.CommandText = sqlStr

        sdr = cmd.ExecuteReader()

        While sdr.Read

            idList.Add(sdr(0).ToString())
            nameList.Add(sdr(1).ToString())

        End While

        If (Math.Ceiling(totalData / 3)) = 0 Then

            totalPage = 1

        Else

            totalPage = Math.Ceiling(totalData / 3)
        End If

        cmd.Dispose()

        sdr.Close()

    End Sub

    Sub createRadioAndTextbox()

        If Not createRadioFlag Then

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

            createRadioFlag = True

        End If

    End Sub

    Sub createFooter()

        If Not createRadioFlag Then

            Dim backB As New System.Windows.Forms.Button
            Dim nextB As New System.Windows.Forms.Button
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


            backB.Name = "backButton"
            backB.Text = "＜"
            backB.Size = New Size(buttonSizeV, buttonSizeH)
            backB.Location = New Point(104, buttonHorizontal)

            nextB.Name = "nextButton"
            nextB.Text = "＞"
            nextB.Size = New Size(buttonSizeV, buttonSizeH)
            nextB.Location = New Point(216, buttonHorizontal)

            child.Name = "child"
            child.Text = 1
            child.Size = New Size(labelSizeV, labelsizeH)
            child.Location = New Point(147, labelHorizontal)

            slash.Name = "slash"
            slash.Text = "/"
            slash.Size = New Size(labelSizeV, labelsizeH)
            slash.Location = New Point(168, labelHorizontal)

            mother.Name = "mother"
            mother.Size = New Size(labelSizeV, labelsizeH)
            mother.Location = New Point(189, labelHorizontal)

            update.Name = "updateButton"
            update.Text = "更新"
            update.Size = New Size(56, 35)
            update.Location = New Point(248, 219)

            Me.Controls.Add(nextB)
            Me.Controls.Add(backB)
            Me.Controls.Add(child)
            Me.Controls.Add(slash)
            Me.Controls.Add(mother)
            Me.Controls.Add(update)

            nowPageLabel = child
            totalPageLabel = mother

            backButton = backB
            nextButton = nextB

            AddHandler backButton.Click, AddressOf backData
            AddHandler nextButton.Click, AddressOf nextData
            AddHandler update.Click, AddressOf updateData

        End If

    End Sub

    Sub setData()

        For i = 0 To idList.Count - 1

            adderRadio(i).Text = idList(i)
            adderTextbox(i).Text = nameList(i)

        Next

    End Sub

    Private Sub nextData()

        backButton.Enabled = True

        nowPage += 1

        resetRadioAndTextbox()

        setData()

        nowPageLabel.Text = nowPage

        If nowPage = totalPage Then

            nextButton.Enabled = False

        End If

        searchData()

        setData()

    End Sub

    Private Sub backData()

        nextButton.Enabled = True

        nowPage -= 1

        resetRadioAndTextbox()

        setData()

        nowPageLabel.Text = nowPage

        If nowPage = 1 Then

            backButton.Enabled = False

        End If

        searchData()

        setData()

    End Sub

    Sub updateData()

        Try

            Dim checkedId As Integer = searchCheckedRadio()

            Dim cmd As New SqlClient.SqlCommand
            Dim sqlStr As String

            cmd.Connection = cn
            cmd.CommandType = CommandType.Text

            If adderTextbox(checkedId).Text.Length > 48 Or adderTextbox(checkedId).Text.Length = 0 Then

                MsgBox("名前は全角16文字以内で入力してください")

            End If

            If maintenanceType = 0 Then

                sqlStr = "UPDATE M_USER SET USER_NAME = '" & adderTextbox(checkedId).Text
                sqlStr += "' WHERE USER_ID = '" & adderRadio(checkedId).Text & "'"

            Else

                sqlStr = "UPDATE M_CONSTELLATION SET CONSTELLATION_NAME = '" & adderTextbox(checkedId).Text
                sqlStr = "' WHERE CONSTELLATION_ID = '" & adderRadio(checkedId).Text & "'"

            End If

            cmd.CommandText = sqlStr

            If updateCoution() Then

                cmd.ExecuteNonQuery()

                endType = 2

                Me.Close()

            End If

        Catch ex As Exception

            MsgBox("ラジオボタンにチェックを入れてください")

        End Try

    End Sub

    Function updateCoution() As Boolean

        Dim result As DialogResult = MessageBox.Show("更新を実行してもよろしいでしょうか",
                                                           "確認",
                                                         MessageBoxButtons.YesNo,
                                                         MessageBoxIcon.Exclamation,
                                                         MessageBoxDefaultButton.Button2)

        If result = DialogResult.Yes Then

            Return True

        End If

        Return False

    End Function

    Function searchCheckedRadio() As Integer

        Dim checked As Integer = 3

        For i = 0 To 2

            If adderRadio(i).Checked Then

                checked = i

            End If
        Next

        Return checked

    End Function
End Class
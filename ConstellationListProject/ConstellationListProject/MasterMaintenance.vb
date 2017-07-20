Public Class MasterMaintenance

    Dim endType As Integer
    Dim label(4) As System.Windows.Forms.Label
    Dim radioButton(4) As System.Windows.Forms.RadioButton
    Dim nowPage As Integer = 1
    '0がユーザ、1が星座
    Dim maintenanceType As Integer
    Dim adderText As System.Windows.Forms.TextBox
    Dim adderRadio As System.Windows.Forms.RadioButton
    Dim adderPass As System.Windows.Forms.TextBox

    Dim cn As System.Data.SqlClient.SqlConnection

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        endType = 1

        Me.Close()

    End Sub

    Private Sub MasterMaintenance_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        If endType = 0 Then

            '戻るボタン
        ElseIf endType = 1 Then

            MaintenanceMenu.Show()

            My.Application.ApplicationContext.MainForm = MaintenanceMenu

            '削除ボタン
        ElseIf endType = 2 Then

            DoneMaintenance.Show()

            My.Application.ApplicationContext.MainForm = DoneMaintenance

            '追加ボタン(ユーザー)
        ElseIf endType = 3 Then

            UserDataAdder.Show()

            My.Application.ApplicationContext.MainForm = UserDataAdder

            '更新ボタン
        ElseIf endType = 4 Then

            DataUpdater.Show()

            My.Application.ApplicationContext.MainForm = DataUpdater

            '追加ボタン(星座)
        ElseIf endType = 5 Then

            ConstellationDataAdder.Show()

            My.Application.ApplicationContext.MainForm = ConstellationDataAdder

        End If

        closeConnection()

    End Sub

    Private Sub MasterMaintenance_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        createRadioButtonList()
        createLabelList()
        connectDb()

        Button5.Enabled = False

        If Label2.Text = "ユーザー" Then

            maintenanceType = 0

        ElseIf Label2.Text = "星座" Then

            maintenanceType = 1

        End If

        getTotalPage()

        setData()

    End Sub

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

    Sub getTotalPage()

        Dim cmd As New System.Data.SqlClient.SqlCommand
        Dim sdr As SqlClient.SqlDataReader
        Dim total As Integer
        Dim totalPage As Integer
        Dim sqlStr As String

        cmd.Connection = cn
        cmd.CommandType = CommandType.Text

        If maintenanceType = 0 Then

            sqlStr = "SELECT COUNT(USER_ID) FROM M_USER"

        Else

            sqlStr = "SELECT COUNT(CONSTELLATION_ID) FROM M_CONSTELLATION"

        End If

        cmd.CommandText = sqlStr

        sdr = cmd.ExecuteReader()

        sdr.Read()

        total = Int(sdr(0).ToString())

        totalPage = Math.Ceiling(total / 5)

        Label12.Text = totalPage

        If Label12.Text = 1 Then

            Button6.Enabled = False

        End If

        cmd.Dispose()

        sdr.Close()

    End Sub

    Sub setData()

        Dim cmd As New SqlClient.SqlCommand

        Dim fastId As Integer
        Dim lastId As Integer

        lastId = nowPage * 5
        fastId = lastId - 4

        cmd.Connection = cn
        cmd.CommandType = CommandType.Text

        Dim sr As SqlClient.SqlDataReader

        If maintenanceType = 0 Then

            cmd.CommandText = "SELECT G_ROW.USER_NAME, G_ROW.USER_ID FROM(
                                SELECT ROW_NUMBER() OVER(ORDER BY USER_ID ASC) AS ROW,
                                    USER_NAME,USER_ID
                                FROM M_USER)G_ROW
                                WHERE G_ROW.ROW BETWEEN " & fastId & " AND " & lastId

        ElseIf maintenanceType = 1 Then

            cmd.CommandText = "SELECT G_ROW.CONSTELLATION_NAME, G_ROW.CONSTELLATION_ID FROM(
                                SELECT ROW_NUMBER() OVER(ORDER BY CONSTELLATION_ID ASC) AS ROW,
                                    CONSTELLATION_NAME, CONSTELLATION_ID
                                FROM M_CONSTELLATION)G_ROW
                                WHERE G_ROW.ROW BETWEEN " & fastId & " AND " & lastId
        End If

        sr = cmd.ExecuteReader()

        cmd.Dispose()

        Dim i As Integer = 0

        While sr.Read

            label(i).Text = sr(0).ToString()
            radioButton(i).Text = sr(1).ToString()
            i = i + 1

        End While

        sr.Close()

    End Sub

    Sub closeConnection()

        If cn.State <> ConnectionState.Closed Then

            cn.Close()
            cn.Dispose()

        End If
    End Sub

    Sub createLabelList()

        label(0) = Label5
        label(1) = Label6
        label(2) = Label7
        label(3) = Label8
        label(4) = Label9

    End Sub

    Sub createRadioButtonList()

        radioButton(0) = RadioButton1
        radioButton(1) = RadioButton2
        radioButton(2) = RadioButton3
        radioButton(3) = RadioButton4
        radioButton(4) = RadioButton5

    End Sub

    Function getTextBoxCnt(ByVal ctrl As Control) As Integer

        If ctrl.Controls.Count = 0 Then
            If TypeOf ctrl Is TextBox Then

                Return 1

            Else

                Return 0
            End If
        End If
        Dim cnt As Integer

        For Each con As Control In ctrl.Controls

            cnt += getTextBoxCnt(con)

        Next

        Return cnt
    End Function

    Sub checkeWhereRadio()

        For i = 0 To 4

            If (radioButton(i).Checked()) Then

                adderRadio = radioButton(i)

            End If
        Next

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Try

            Dim result As DialogResult = MessageBox.Show("削除を実行してもよろしいでしょうか",
                                                           "確認",
                                                         MessageBoxButtons.YesNo,
                                                         MessageBoxIcon.Exclamation,
                                                         MessageBoxDefaultButton.Button2)

            If result = DialogResult.Yes Then

                checkeWhereRadio()

                deleteProcess()

            End If

        Catch ex As Exception

        End Try

    End Sub

    Sub deleteProcess()

        Dim cmd As New SqlClient.SqlCommand
        Dim sqlStr As String
        Dim deleteId As String

        cmd.Connection = cn

        cmd.CommandType = CommandType.Text

        If adderRadio.Text.Length = 1 Then

            deleteId = "0" + adderRadio.Text

        Else

            deleteId = adderRadio.Text

        End If

        If maintenanceType = 0 Then

            sqlStr = "DELETE M_USER WHERE USER_ID = '" & deleteId & "'"

        Else

            sqlStr = "DELETE M_CONSTELLATION WHERE CONSTELLATION_ID = '" & deleteId & "'"

        End If

        cmd.CommandText = sqlStr

        cmd.ExecuteNonQuery()

        cmd.Dispose()

        endType = 2

        Me.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try

            If maintenanceType = 0 Then

                endType = 3

            Else

                endType = 5

            End If

            Me.Close()

        Catch ex As Exception

            MsgBox(ex.Message())

        End Try
    End Sub

    Sub resetData()

        For i = 0 To 4

            radioButton(i).Text = ""
            label(i).Text = ""

        Next

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Button5.Enabled = True

        nowPage += 1

        resetData()

        setData()

        Label10.Text = nowPage

        If Label10.Text = Label12.Text Then

            Button6.Enabled = False

        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Button6.Enabled = True

        nowPage -= 1

        resetData()

        setData()

        Label10.Text = nowPage

        If Label10.Text = 1 Then

            Button5.Enabled = False

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If maintenanceType = 0 Then

            DataUpdater.Label2.Text = "ユーザー"

        Else

            DataUpdater.Label2.Text = "星座"

        End If

        endType = 4

        Me.Close()

    End Sub
End Class
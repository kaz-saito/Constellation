Public Class MasterMaintenance

    Dim endType As Integer
    Dim label(4) As System.Windows.Forms.Label
    Dim radioButton(4) As System.Windows.Forms.RadioButton
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

            '追加ボタン
        ElseIf endType = 3 Then

            DataAdder.Show()

            My.Application.ApplicationContext.MainForm = DataAdder

        End If

        closeConnection()

    End Sub

    Private Sub MasterMaintenance_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        createRadioButtonList()
        createLabelList()
        connectDb()

        If Label2.Text = "ユーザー" Then

            maintenanceType = 0

        ElseIf Label2.Text = "星座" Then

            maintenanceType = 1

        End If

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

    Sub setData()

        Dim cmd As New SqlClient.SqlCommand

        cmd.Connection = cn
        cmd.CommandType = CommandType.Text

        Dim sr As SqlClient.SqlDataReader

        If maintenanceType = 0 Then

            cmd.CommandText = "SELECT G_ROW.USER_NAME, G_ROW.USER_ID FROM(
                                SELECT ROW_NUMBER() OVER(ORDER BY USER_ID ASC) AS ROW,
                                    USER_NAME,USER_ID
                                FROM M_USER)G_ROW
                                WHERE G_ROW.ROW BETWEEN 1 AND 5"

        ElseIf maintenanceType = 1 Then

            cmd.CommandText = "SELECT G_ROW.CONSTELLATION_NAME, G_ROW.CONSTELLATION_ID FROM(
                                SELECT ROW_NUMBER() OVER(ORDER BY CONSTELLATION_ID ASC) AS ROW,
                                    CONSTELLATION_NAME, CONSTELLATION_ID
                                FROM M_CONSTELLATION)G_ROW
                                WHERE G_ROW.ROW BETWEEN 1 AND 5"
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

        sqlStr = "DELETE M_USER WHERE USER_ID = '" & deleteId & "'"

        cmd.CommandText = sqlStr

        cmd.ExecuteNonQuery()

        cmd.Dispose()

        endType = 2

        Me.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try

            If maintenanceType = 0 Then

                DataAdder.Label2.Text = "ユーザー"



            Else

                DataAdder.Label2.Text = "星座"

            End If

            endType = 3

            Me.Close()

        Catch ex As Exception

            MsgBox(ex.Message())

        End Try
    End Sub
End Class
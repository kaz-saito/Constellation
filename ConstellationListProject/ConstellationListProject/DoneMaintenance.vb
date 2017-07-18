Public Class DoneMaintenance

    Dim endType As Integer

    Private Sub DoneMaintenance_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        endType = 1

        Me.Close()

    End Sub

    Private Sub DoneMaintenance_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        If endType = 0 Then

        ElseIf endType = 1 Then

            MainMenu.Show()

            My.Application.ApplicationContext.MainForm = MainMenu

        ElseIf endType = 2 Then

            MasterMaintenance.Show()

            My.Application.ApplicationContext.MainForm = MasterMaintenance

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        endType = 2

        MasterMaintenance.Label2.Text = "ユーザー"

        Me.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        endType = 2

        MasterMaintenance.Label2.Text = "星座"

        Me.Close()

    End Sub
End Class
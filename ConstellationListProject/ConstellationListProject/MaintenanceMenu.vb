Public Class MaintenanceMenu
    Dim endType As Integer

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        endType = 0

        Me.Close()

    End Sub

    Private Sub MaintenanceMenu_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        If endType = 0 Then

        ElseIf endType = 1 Then

            MasterMaintenance.Label2.Text = "ユーザー"

            MasterMaintenance.Show()

            My.Application.ApplicationContext.MainForm = MasterMaintenance

        ElseIf endType = 2 Then

            MasterMaintenance.Label2.Text = "星座"

            MasterMaintenance.Show()

            My.Application.ApplicationContext.MainForm = MasterMaintenance

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        endType = 1

        Me.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        endType = 2

        Me.Close()

    End Sub

    Private Sub MaintenanceMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
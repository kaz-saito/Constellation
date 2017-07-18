Public Class DataAdder

    Dim maintenanceType As Integer
    Dim endType As Integer
    Private Sub DataAdder_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Label2.Text = "ユーザー" Then

            maintenanceType = 0

        Else

            maintenanceType = 1

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        endType = 1

        If maintenanceType = 0 Then

            MasterMaintenance.Label2.Text = "ユーザ"

        Else

            MasterMaintenance.Label2.Text = "星座"

        End If

        Me.Close()

    End Sub

    Private Sub DataAdder_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        If endType = 0 Then

        ElseIf endType = 1 Then

            MasterMaintenance.Show()

            My.Application.ApplicationContext.MainForm = MasterMaintenance

        End If
    End Sub
End Class
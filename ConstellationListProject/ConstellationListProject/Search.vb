Public Class Search

    Private constellationList As List(Of String)
    Private page As Integer = 0
    Private startPage As Integer = 0
    Private countPage As Integer
    Private maxPage As Integer

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles searchButton.Click

        Dim season As String = condition.SelectedValue
        Dim dao As New SearchDao

        constellationList = dao.searchConstellation(season)

        resultView.Clear()

        resultView.View = View.Details

        resultView.Columns.Add("星座名", 135, HorizontalAlignment.Center)

        For i As Integer = 0 To 4

            If i >= 0 And i < constellationList.Count() Then
                resultView.Items.Add(constellationList(i))

                If i Mod 2 = 0 Then
                    resultView.Items(i).BackColor = Color.Cyan
                End If

            End If

        Next

        page = 0

        maxPage = Math.Ceiling(constellationList.Count() / 5)

        countPage = 1

        Label2.Text = countPage & "/" & maxPage

        resultView.GridLines = True


    End Sub

    Private Sub Search_Load(sender As Object, e As EventArgs) Handles Me.Load

        'DataTableオブジェクトの作成
        Dim seasonTable As New DataTable()

        'DataTableに列追加
        seasonTable.Columns.Add("ID", GetType(String))
        seasonTable.Columns.Add("NAME", GetType(String))

        '配列の用意
        Dim rowDataArray As String(,) = {{"0", "春"},
                                        {"1", "夏"},
                                        {"2", "秋"},
                                        {"3", "冬"}}

        For i As Integer = 0 To rowDataArray.GetLength(0) - 1
            '新しい行を作成
            Dim row As DataRow = seasonTable.NewRow()

            '各列に値をセット
            row("ID") = rowDataArray(i, 0)
            row("NAME") = rowDataArray(i, 1)

            'DataTableに行を追加
            seasonTable.Rows.Add(row)
        Next

        'コンボボックスのDataSourceにDataTableを割り当てる
        condition.DataSource = seasonTable

        '表示される値はDataTableのNAME列
        condition.DisplayMember = "NAME"

        '対応する値はDataTableのID列
        condition.ValueMember = "ID"

    End Sub

    Private Sub resultView_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles resultView.ColumnWidthChanging
        e.NewWidth = Me.resultView.Columns(e.ColumnIndex).Width
        e.Cancel = True
    End Sub

    Private Sub nextButton_Click(sender As Object, e As EventArgs) Handles nextButton.Click


        If page + 5 >= 0 And page + 5 < constellationList.Count() Then

            startPage = page + 5

            resultView.Clear()

            resultView.View = View.Details

            resultView.Columns.Add("星座名", 135, HorizontalAlignment.Center)

            Dim i As Integer = 0

            For page = page + 5 To page + 9

                If constellationList.Count() > page Then
                    resultView.Items.Add(constellationList(page))

                    i = i + 1

                Else

                    Exit For

                End If


            Next

            For j As Integer = 0 To i
                If j Mod 2 = 0 Then
                    resultView.Items(j).BackColor = Color.Cyan
                End If
            Next

            page = startPage

            countPage = countPage + 1

            Label2.Text = countPage & "/" & maxPage

            resultView.GridLines = True

        End If

    End Sub

    Private Sub backButton_Click(sender As Object, e As EventArgs) Handles backButton.Click

        If page - 5 >= 0 And page - 5 < constellationList.Count() Then

            startPage = page - 5

            resultView.Clear()

            resultView.View = View.Details

            resultView.Columns.Add("星座名", 135, HorizontalAlignment.Center)

            Dim i As Integer = 0

            For page = page - 5 To page - 1

                If constellationList.Count() > page Then
                    resultView.Items.Add(constellationList(page))

                    i = i + 1

                Else

                    Exit For

                End If

            Next

            For j As Integer = 0 To i
                If j Mod 2 = 0 Then
                    resultView.Items(j).BackColor = Color.Cyan
                End If
            Next

            page = startPage

            countPage = countPage - 1

            Label2.Text = countPage & "/" & maxPage

            resultView.GridLines = True

        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim f As New MainMenu
        f.Show()
        My.Application.ApplicationContext.MainForm = f
        Me.Close()

    End Sub
End Class
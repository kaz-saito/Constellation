Public Class SearchDao

    Private con As New System.Data.SqlClient.SqlConnection
    Private cn As New Connection()

    Function searchConstellation(season As String) As List(Of String)


        con = cn.Connect()

        Dim sc As New SqlClient.SqlCommand
        Dim sdr As System.Data.SqlClient.SqlDataReader
        Dim list As New List(Of String)

        sc.Connection = con

        sc.CommandType = CommandType.Text

        Try

            sc.CommandText = "SELECT CONSTELLATION_NAME FROM M_CONSTELLATION WHERE SEASON_ID = @Key1"

            sc.Parameters.Clear()

            sc.Parameters.AddWithValue("@Key1", season)

            sdr = sc.ExecuteReader()

            sc.Dispose()

            While sdr.Read()

                list.Add(sdr("CONSTELLATION_NAME").ToString)

            End While

            Return list

        Catch e As Exception
            Console.WriteLine("Error! {0}", e.Message)
            Return list

        Finally
            If sc IsNot Nothing Then
                sc.Dispose()
            End If

        End Try


    End Function

End Class

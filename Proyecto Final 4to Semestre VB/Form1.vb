﻿Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports System.Configuration
Imports System.IO
Imports Newtonsoft.Json
Imports YamlDotNet.Serialization
Imports YamlDotNet.Serialization.NamingConventions

Namespace Proyecto_Final_4to_Semestre
    Partial Public Class Form1
        Inherits Form

        Private connectionString As String = ConfigurationManager.ConnectionStrings("MyConnectionString").ConnectionString
        Private dataTable As DataTable
        Private bindingSource As BindingSource
        Private dataAdapter As SqlDataAdapter

        Public Sub New()
            InitializeComponent()
            InitializeDataGridView()
        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)
            DataGridView.RowsDefaultCellStyle.ForeColor = Color.Black
            DataGridView.DefaultCellStyle.BackColor = Color.FromArgb(220, 220, 220)
            DataGridView.DefaultCellStyle.ForeColor = Color.Black
            DataGridView.DefaultCellStyle.SelectionBackColor = Color.FromArgb(173, 216, 230)
            DataGridView.DefaultCellStyle.SelectionForeColor = Color.Black

            LoadData()
        End Sub

        Private Sub InitializeDataGridView()
            bindingSource = New BindingSource()
            DataGridView.DataSource = bindingSource
        End Sub
    End Class

#Region "Metodo para cargar los datos de la base de datos en el DataGridView"
    Private Sub LoadData()
        Try
            DataTable = New DataTable()
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim dataAdapter As New SqlDataAdapter("SELECT * FROM spotify_songs", connection)
                dataAdapter.Fill(DataTable)
            End Using

            ' Enlazar el DataTable con el DataGridView
            DataGridView.DataSource = DataTable
        Catch ex As Exception
            MessageBox.Show("Error al cargar datos desde la base de datos: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "BOTON PARA GUARDAR EN FORMATO CSV"
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        Dim dialogoGuardar As New SaveFileDialog()
        dialogoGuardar.Filter = "Archivos CSV (*.csv)|*.csv"
        dialogoGuardar.DefaultExt = "csv"

        If dialogoGuardar.ShowDialog() = DialogResult.OK Then
            Try
                GuardarEnCSV(dialogoGuardar.FileName)

                MessageBox.Show("Datos guardados correctamente en el archivo: " & dialogoGuardar.FileName)
            Catch ex As Exception
                MessageBox.Show("Error al guardar el archivo: " & ex.Message)
            End Try
        End If
    End Sub

    'METODO PARA GUARDAR EN CSV
    Private Sub GuardarEnCSV(ByVal rutaArchivo As String)
        Using writer As New StreamWriter(rutaArchivo)
            Dim filaEncabezados As New StringBuilder()
            For Each columna As DataGridViewColumn In DataGridView.Columns
                filaEncabezados.Append(columna.HeaderText & ",")
            Next
            writer.WriteLine(filaEncabezados.ToString())

            For Each fila As DataGridViewRow In DataGridView.Rows
                Dim filaDatos As New StringBuilder()
                For Each celda As DataGridViewCell In fila.Cells
                    If celda.Value IsNot Nothing Then
                        filaDatos.Append(celda.Value.ToString() & ",")
                    Else
                        filaDatos.Append("[CELDA VACÍA],")
                    End If
                Next
                writer.WriteLine(filaDatos.ToString())
            Next
        End Using
    End Sub
#End Region

#Region "BOTON PARA LEER FORMATOS CSV"
    Private Sub Button8_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button8.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Archivos CSV (*.csv)|*.csv|Todos los archivos (*.*)|*.*"
        openFileDialog.Title = "Seleccionar archivo CSV"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = openFileDialog.FileName
            Dim dataTable As DataTable = LeerDesdeCSV(filePath)
            If dataTable.Rows.Count > 0 Then
                DataGridView.DataSource = dataTable
            End If
        End If
    End Sub

    'METODO PARA LEER DESDE CSV
    Private Function LeerDesdeCSV(ByVal filePath As String) As DataTable
        Dim dataTable As New DataTable()
        Try
            Using sr As New StreamReader(filePath)
                Dim headers As String() = sr.ReadLine().Split(","c)
                For Each header As String In headers
                    dataTable.Columns.Add(header)
                Next
                While Not sr.EndOfStream
                    Dim rows As String() = sr.ReadLine().Split(","c)
                    Dim dataRow As DataRow = dataTable.NewRow()
                    For i As Integer = 0 To headers.Length - 1
                        dataRow(i) = rows(i)
                    Next
                    dataTable.Rows.Add(dataRow)
                End While
            End Using
        Catch ex As Exception
            MessageBox.Show("Error al leer archivo CSV: " & ex.Message)
        End Try
        Return dataTable
    End Function
#End Region

#Region "BOTON PARA GUARDAR EN FORMATO JSON"
    Private Sub btnGuardarJSON_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnGuardarJSON.Click
        Dim dataList As New List(Of Object)()

        For Each row As DataGridViewRow In DataGridView.Rows
            If Not row.IsNewRow Then
                Dim rowData As New Dictionary(Of String, Object)()

                ' Iterar a través de las celdas de la fila
                For Each cell As DataGridViewCell In row.Cells
                    ' Obtener el nombre de la columna y el valor de la celda
                    Dim columnName As String = DataGridView.Columns(cell.ColumnIndex).Name
                    Dim cellValue As Object = cell.Value

                    ' Agregar el nombre de la columna y el valor a rowData
                    rowData(columnName) = cellValue
                Next

                ' Agregar rowData a la lista dataList
                dataList.Add(rowData)
            End If
        Next

        ' Serializar la lista de objetos a JSON usando Newtonsoft.Json
        Dim json As String = JsonConvert.SerializeObject(dataList, Formatting.Indented)

        ' Guardar el archivo JSON usando SaveFileDialog
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "Archivos JSON (*.json)|*.json|Todos los archivos (*.*)|*.*"
        saveFileDialog1.Title = "Guardar datos JSON"
        saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        saveFileDialog1.FileName = "spotify_songs_data.json"

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = saveFileDialog1.FileName
            File.WriteAllText(filePath, json)

            MessageBox.Show("Datos guardados en formato JSON correctamente en:" & vbCrLf & filePath)
        End If
    End Sub
#End Region

#Region "BOTON PARA LEER FORMATO JSON"
    Private Sub btnLeerJSON_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnLeerJSON.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Archivos JSON (*.json)|*.json|Todos los archivos (*.*)|*.*"
        openFileDialog.Title = "Seleccionar archivo JSON"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = openFileDialog.FileName
            Dim dataTable As DataTable = LeerDesdeJSON(filePath)
            If dataTable IsNot Nothing Then
                DataGridView.DataSource = dataTable
            End If
        End If
    End Sub

    Private Function LeerDesdeJSON(ByVal filePath As String) As DataTable
        Dim tempDataTable As New DataTable()
        Try
            If File.Exists(filePath) Then
                ' Leer el contenido del archivo JSON
                Dim jsonContent As String = File.ReadAllText(filePath)

                ' Deserializar el JSON a una lista de objetos anónimos
                Dim dataList = JsonConvert.DeserializeObject(Of List(Of Dictionary(Of String, Object)))(jsonContent)

                ' Verificar si dataList tiene elementos
                If dataList IsNot Nothing AndAlso dataList.Count > 0 Then
                    ' Agregar columnas al DataTable basado en las claves del primer diccionario
                    For Each key As String In dataList(0).Keys
                        tempDataTable.Columns.Add(key)
                    Next

                    ' Agregar filas al DataTable
                    For Each Data In dataList
                        Dim row As DataRow = tempDataTable.NewRow()
                        For Each key As String In Data.Keys
                            row(key) = If(Data(key) IsNot Nothing, Data(key).ToString(), Nothing) ' Convertir valores a cadena
                        Next
                        tempDataTable.Rows.Add(row)
                    Next
                Else
                    MessageBox.Show($"El archivo JSON ""{Path.GetFileName(filePath)}"" está vacío o no tiene el formato esperado.")
                End If
            Else
                MessageBox.Show($"El archivo JSON ""{Path.GetFileName(filePath)}"" no existe.")
            End If
        Catch ex As Exception
            MessageBox.Show($"Error al leer archivo JSON ""{Path.GetFileName(filePath)}"": {ex.Message}")
        End Try
        Return tempDataTable
    End Function
#End Region

#Region "BOTON PARA GUARDAR EN FORMATO XML"
    Private Sub btnGuardarXML_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnGuardarXML.Click
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "Archivos XML (*.xml)|*.xml|Todos los archivos (*.*)|*.*"
        saveFileDialog1.Title = "Guardar datos XML"
        saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        saveFileDialog1.FileName = "spotify_songs_data.xml"

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            Try
                ' Obtener el DataTable desde el DataSource del DataGridView
                Dim dataTable As DataTable = DirectCast(DataGridView.DataSource, DataTable)

                If dataTable IsNot Nothing Then
                    ' Asignar nombre de la tabla si está vacío
                    If String.IsNullOrEmpty(dataTable.TableName) Then
                        dataTable.TableName = "spotify_songs"
                    End If

                    ' Guardar el DataTable como XML
                    Dim filePath As String = saveFileDialog1.FileName
                    dataTable.WriteXml(filePath, XmlWriteMode.WriteSchema)

                    MessageBox.Show("Datos guardados en formato XML correctamente en:" & vbCrLf & filePath)
                Else
                    MessageBox.Show("No hay datos en el DataGridView para guardar.")
                End If
            Catch ex As Exception
                MessageBox.Show("Error al guardar el archivo XML: " & ex.Message & vbCrLf & ex.StackTrace)
            End Try
        End If
    End Sub
#End Region

#Region "Boton para leer en XML"
    Private Sub btnLeerXML_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnLeerXML.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Archivos XML (*.xml)|*.xml|Todos los archivos (*.*)|*.*"
        openFileDialog.Title = "Seleccionar archivo XML"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = openFileDialog.FileName
            Dim dataTable As DataTable = LeerDesdeXML(filePath)
            If dataTable.Rows.Count > 0 Then
                DataGridView.DataSource = dataTable
            End If
        End If
    End Sub

    Private Function LeerDesdeXML(ByVal filePath As String) As DataTable
        Dim dataTable As New DataTable()
        Try
            ' Leer el archivo XML en un DataSet
            Dim dataSet As New DataSet()
            dataSet.ReadXml(filePath)

            ' Verificar si el DataSet tiene al menos una tabla y esta tiene al menos una fila
            If dataSet.Tables.Count > 0 AndAlso dataSet.Tables(0).Rows.Count > 0 Then
                ' Asignar la tabla del DataSet al DataTable
                dataTable = dataSet.Tables(0)
            End If
        Catch ex As Exception
            MessageBox.Show("Error al leer archivo XML: " & ex.Message)
        End Try
        Return dataTable
    End Function
#End Region
#Region "BOTON PARA GUARDAR EN FORMATO YAML"
    Private Sub btnGuardarYaml_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnGuardarYaml.Click
        ' Obtener el DataTable desde el DataSource del DataGridView
        Dim dataTable As DataTable = DirectCast(DataGridView.DataSource, DataTable)

        ' Convertir DataTable a una lista de diccionarios para facilitar la serialización
        Dim dataList As New List(Of Dictionary(Of String, Object))()
        For Each row As DataRow In dataTable.Rows
            Dim dict As New Dictionary(Of String, Object)()
            For Each col As DataColumn In row.Table.Columns
                dict(col.ColumnName) = row(col)
            Next
            dataList.Add(dict)
        Next

        ' Serializar a YAML usando YamlDotNet
        Dim serializer = New SerializerBuilder() _
        .WithNamingConvention(CamelCaseNamingConvention.Instance) _
        .Build()

        Dim yaml As String = serializer.Serialize(dataList)

        ' Configurar SaveFileDialog
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "Archivos YAML (*.yaml)|*.yaml|Todos los archivos (*.*)|*.*"
        saveFileDialog1.Title = "Guardar datos YAML"
        saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        saveFileDialog1.FileName = "spotify_songs_data.yaml"

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            Try
                ' Guardar el archivo YAML
                Dim filePath As String = saveFileDialog1.FileName
                File.WriteAllText(filePath, yaml)

                MessageBox.Show("Datos guardados en formato YAML correctamente en:" & vbCrLf & filePath)
            Catch ex As Exception
                MessageBox.Show("Error al guardar el archivo YAML: " & ex.Message)
            End Try
        End If
    End Sub
#End Region

#Region "Boton para Leer en YAML"
    Private Sub btnLeerYAML_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnLeerYAML.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Archivos YAML (*.yaml;*.yml)|*.yaml;*.yml|Todos los archivos (*.*)|*.*"
        openFileDialog.Title = "Seleccionar archivo YAML"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = openFileDialog.FileName
            Dim dataTable As DataTable = LeerDesdeYAML(filePath)
            If dataTable.Rows.Count > 0 Then
                DataGridView.DataSource = dataTable
            End If
        End If
    End Sub

    Private Function LeerDesdeYAML(ByVal filePath As String) As DataTable
        Dim dataTable As New DataTable()
        Try
            ' Crear un deserializador YAML
            Dim deserializer As New Deserializer()

            ' Leer el contenido del archivo YAML
            Dim yamlContent As String = File.ReadAllText(filePath)

            ' Deserializar el contenido YAML a una lista de diccionarios
            Dim dataList As List(Of Dictionary(Of String, Object)) = deserializer.Deserialize(Of List(Of Dictionary(Of String, Object))))(yamlContent)

        ' Si dataList tiene elementos
        If dataList.Count > 0 Then
                ' Agregar columnas al DataTable basado en las claves del primer diccionario
                For Each key As String In dataList(0).Keys
                    dataTable.Columns.Add(key)
                Next

                ' Agregar filas al DataTable
                For Each data As Dictionary(Of String, Object) In dataList
                    Dim row As DataRow = dataTable.NewRow()
                    For Each key As String In data.Keys
                        row(key) = If(data(key) IsNot Nothing, data(key).ToString(), Nothing) ' Convertir valores a cadena
                    Next
                    dataTable.Rows.Add(row)
                Next
            End If
        Catch ex As Exception
            MessageBox.Show("Error al leer archivo YAML: " & ex.Message)
        End Try
        Return dataTable
    End Function
#End Region

#Region "BOTONES PARA MANEJAR EL TAMAÑO DE LA VENTANA"
    Private Sub button1_Click_1(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
        ' Maximizar o normalizar la ventana
        If WindowState = FormWindowState.Normal Then
            WindowState = FormWindowState.Maximized
        Else
            WindowState = FormWindowState.Normal
        End If
    End Sub

    Private Sub button4_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button4.Click
        ' Cerrar la ventana
        Me.Close()
    End Sub

    Private Sub button9_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button9.Click
        ' Minimizar la ventana
        WindowState = FormWindowState.Minimized
    End Sub
#End Region
    Private Sub btnGuardar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnGuardar.Click
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim dataAdapter As New SqlDataAdapter()

                ' Configurar el comando DELETE
                Dim deleteCommand As New SqlCommand(
                "DELETE FROM spotify_songs WHERE track_id = @track_id", connection)
                deleteCommand.Parameters.Add("@track_id", SqlDbType.NVarChar, 50, "track_id")
                dataAdapter.DeleteCommand = deleteCommand

                ' Configurar el comando INSERT
                Dim insertCommand As New SqlCommand(
                "INSERT INTO spotify_songs (track_id, track_name, track_artist, track_popularity, " &
                "track_album_id, track_album_name, track_album_release_date, playlist_name, playlist_id, " &
                "playlist_genre, playlist_subgenre, danceability, energy, [key], loudness, mode, speechiness, " &
                "acousticness, instrumentalness, liveness, valence, tempo, duration_ms) " &
                "VALUES (@track_id, @track_name, @track_artist, @track_popularity, @track_album_id, " &
                "@track_album_name, @track_album_release_date, @playlist_name, @playlist_id, @playlist_genre, " &
                "@playlist_subgenre, @danceability, @energy, @key, @loudness, @mode, @speechiness, " &
                "@acousticness, @instrumentalness, @liveness, @valence, @tempo, @duration_ms); " &
                "SELECT * FROM spotify_songs WHERE track_id = SCOPE_IDENTITY();", connection)
                insertCommand.Parameters.Add("@track_id", SqlDbType.NVarChar, 50, "track_id").Value = DBNull.Value ' Valor generado automáticamente por la base de datos
                insertCommand.Parameters.Add("@track_name", SqlDbType.VarChar, -1, "track_name")
                insertCommand.Parameters.Add("@track_artist", SqlDbType.VarChar, -1, "track_artist")
                insertCommand.Parameters.Add("@track_popularity", SqlDbType.TinyInt, 1, "track_popularity")
                insertCommand.Parameters.Add("@track_album_id", SqlDbType.NVarChar, 50, "track_album_id")
                insertCommand.Parameters.Add("@track_album_name", SqlDbType.VarChar, -1, "track_album_name")
                insertCommand.Parameters.Add("@track_album_release_date", SqlDbType.DateTime2, 7, "track_album_release_date")
                insertCommand.Parameters.Add("@playlist_name", SqlDbType.VarChar, -1, "playlist_name")
                insertCommand.Parameters.Add("@playlist_id", SqlDbType.NVarChar, 50, "playlist_id")
                insertCommand.Parameters.Add("@playlist_genre", SqlDbType.NVarChar, 50, "playlist_genre")
                insertCommand.Parameters.Add("@playlist_subgenre", SqlDbType.NVarChar, 50, "playlist_subgenre")
                insertCommand.Parameters.Add("@danceability", SqlDbType.Float, 8, "danceability")
                insertCommand.Parameters.Add("@energy", SqlDbType.Float, 8, "energy")
                insertCommand.Parameters.Add("@key", SqlDbType.TinyInt, 1, "key")
                insertCommand.Parameters.Add("@loudness", SqlDbType.Float, 8, "loudness")
                insertCommand.Parameters.Add("@mode", SqlDbType.Bit, 1, "mode")
                insertCommand.Parameters.Add("@speechiness", SqlDbType.Float, 8, "speechiness")
                insertCommand.Parameters.Add("@acousticness", SqlDbType.Float, 8, "acousticness")
                insertCommand.Parameters.Add("@instrumentalness", SqlDbType.Float, 8, "instrumentalness")
                insertCommand.Parameters.Add("@liveness", SqlDbType.Float, 8, "liveness")
                insertCommand.Parameters.Add("@valence", SqlDbType.Float, 8, "valence")
                insertCommand.Parameters.Add("@tempo", SqlDbType.Float, 8, "tempo")
                insertCommand.Parameters.Add("@duration_ms", SqlDbType.Int, 4, "duration_ms")
                dataAdapter.InsertCommand = insertCommand

                ' Configurar el comando UPDATE
                Dim updateCommand As New SqlCommand(
                "UPDATE spotify_songs SET " &
                "track_name = @track_name, " &
                "track_artist = @track_artist, " &
                "track_popularity = @track_popularity, " &
                "track_album_id = @track_album_id, " &
                "track_album_name = @track_album_name, " &
                "track_album_release_date = @track_album_release_date, " &
                "playlist_name = @playlist_name, " &
                "playlist_id = @playlist_id, " &
                "playlist_genre = @playlist_genre, " &
                "playlist_subgenre = @playlist_subgenre, " &
                "danceability = @danceability, " &
                "energy = @energy, " &
                "[key] = @key, " &
                "loudness = @loudness, " &
                "mode = @mode, " &
                "speechiness = @speechiness, " &
                "acousticness = @acousticness, " &
                "instrumentalness = @instrumentalness, " &
                "liveness = @liveness, " &
                "valence = @valence, " &
                "tempo = @tempo, " &
                "duration_ms = @duration_ms " &
                "WHERE track_id = @track_id", connection)
                updateCommand.Parameters.Add("@track_id", SqlDbType.NVarChar, 50, "track_id")
                updateCommand.Parameters.Add("@track_name", SqlDbType.VarChar, -1, "track_name")
                updateCommand.Parameters.Add("@track_artist", SqlDbType.VarChar, -1, "track_artist")
                updateCommand.Parameters.Add("@track_popularity", SqlDbType.TinyInt, 1, "track_popularity")
                updateCommand.Parameters.Add("@track_album_id", SqlDbType.NVarChar, 50, "track_album_id")
                updateCommand.Parameters.Add("@track_album_name", SqlDbType.VarChar, -1, "track_album_name")
                updateCommand.Parameters.Add("@track_album_release_date", SqlDbType.DateTime2, 7, "track_album_release_date")
                updateCommand.Parameters.Add("@playlist_name", SqlDbType.VarChar, -1, "playlist_name")
                updateCommand.Parameters.Add("@playlist_id", SqlDbType.NVarChar, 50, "playlist_id")
                updateCommand.Parameters.Add("@playlist_genre", SqlDbType.NVarChar, 50, "playlist_genre")
                updateCommand.Parameters.Add("@playlist_subgenre", SqlDbType.NVarChar, 50, "playlist_subgenre")
                updateCommand.Parameters.Add("@danceability", SqlDbType.Float, 8, "danceability")
                updateCommand.Parameters.Add("@energy", SqlDbType.Float, 8, "energy")
                updateCommand.Parameters.Add("@key", SqlDbType.TinyInt, 1, "key")
                updateCommand.Parameters.Add("@loudness", SqlDbType.Float, 8, "loudness")
                updateCommand.Parameters.Add("@mode", SqlDbType.Bit, 1, "mode")
                updateCommand.Parameters.Add("@speechiness", SqlDbType.Float, 8, "speechiness")
                updateCommand.Parameters.Add("@acousticness", SqlDbType.Float, 8, "acousticness")
                updateCommand.Parameters.Add("@instrumentalness", SqlDbType.Float, 8, "instrumentalness")
                updateCommand.Parameters.Add("@liveness", SqlDbType.Float, 8, "liveness")
                updateCommand.Parameters.Add("@valence", SqlDbType.Float, 8, "valence")
                updateCommand.Parameters.Add("@tempo", SqlDbType.Float, 8, "tempo")
                updateCommand.Parameters.Add("@duration_ms", SqlDbType.Int, 4, "duration_ms")
                dataAdapter.UpdateCommand = updateCommand

                ' Actualizar la base de datos con los cambios realizados en el DataTable
                dataAdapter.Update(DataTable)

                ' Confirmación de cambios guardados
                MessageBox.Show("Cambios guardados correctamente en la base de datos.")
            End Using
        Catch ex As Exception
            MessageBox.Show("Error al guardar cambios en la base de datos: " & ex.Message)
        End Try
    End Sub

    Private Sub DataGridView_UserDeletingRow(ByVal sender As Object, ByVal e As DataGridViewRowCancelEventArgs) Handles DataGridView.UserDeletingRow
        ' Eliminar la fila de la base de datos
        Dim rowView As DataRowView = DirectCast(e.Row.DataBoundItem, DataRowView)
        Dim row As DataRow = rowView.Row
        row.Delete()
    End Sub

    Private Sub DataGridView_RowValidating(ByVal sender As Object, ByVal e As DataGridViewCellCancelEventArgs) Handles DataGridView.RowValidating
        ' Manejar la inserción o actualización de una fila en el DataGridView
        Dim dataGridViewRow As DataGridViewRow = DataGridView.Rows(e.RowIndex)

        ' Verificar si es una nueva fila o una fila existente
        If dataGridViewRow.IsNewRow Then Return

        ' Actualizar la fila en el DataTable
        Dim rowView As DataRowView = DirectCast(dataGridViewRow.DataBoundItem, DataRowView)
        Dim row As DataRow = rowView.Row
        row("track_id") = dataGridViewRow.Cells("track_id").Value
        row("track_name") = dataGridViewRow.Cells("track_name").Value
        row("track_artist") = dataGridViewRow.Cells("track_artist").Value
        row("track_popularity") = dataGridViewRow.Cells("track_popularity").Value
        row("track_album_id") = dataGridViewRow.Cells("track_album_id").Value
        row("track_album_name") = dataGridViewRow.Cells("track_album_name").Value
        row("track_album_release_date") = dataGridViewRow.Cells("track_album_release_date").Value
        row("playlist_name") = dataGridViewRow.Cells("playlist_name").Value
        row("playlist_id") = dataGridViewRow.Cells("playlist_id").Value
        row("playlist_genre") = dataGridViewRow.Cells("playlist_genre").Value
        row("playlist_subgenre") = dataGridViewRow.Cells("playlist_subgenre").Value
        row("danceability") = dataGridViewRow.Cells("danceability").Value
        row("energy") = dataGridViewRow.Cells("energy").Value
        row("key") = dataGridViewRow.Cells("key").Value
        row("loudness") = dataGridViewRow.Cells("loudness").Value
        row("mode") = dataGridViewRow.Cells("mode").Value
        row("speechiness") = dataGridViewRow.Cells("speechiness").Value
        row("acousticness") = dataGridViewRow.Cells("acousticness").Value
        row("instrumentalness") = dataGridViewRow.Cells("instrumentalness").Value
        row("liveness") = dataGridViewRow.Cells("liveness").Value
        row("valence") = dataGridViewRow.Cells("valence").Value
        row("tempo") = dataGridViewRow.Cells("tempo").Value
        row("duration_ms") = dataGridViewRow.Cells("duration_ms").Value
    End Sub


End Namespace

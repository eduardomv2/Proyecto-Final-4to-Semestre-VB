Imports System
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


End Namespace

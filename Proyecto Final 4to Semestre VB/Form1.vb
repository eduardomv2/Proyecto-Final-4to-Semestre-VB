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

End Namespace

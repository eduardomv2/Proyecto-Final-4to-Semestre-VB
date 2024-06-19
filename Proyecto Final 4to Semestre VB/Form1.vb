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
End Namespace

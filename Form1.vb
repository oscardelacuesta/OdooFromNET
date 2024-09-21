Imports System.Data
Imports OfficeOpenXml
Imports System.IO
' Función para exportar el DataGridView a un archivo Excel
Imports System.Diagnostics ' Para abrir el archivo automáticamente
Imports System.Windows.Forms ' Para SaveFileDialog

Public Class Form1
    Private userId As Integer
    Private selectedPartnerId As Integer = -1 ' Para almacenar el ID del socio seleccionado

    ' Al cargar el formulario, conectarse a Odoo
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        userId = OdooConnector.ConnectToOdoo()
        If userId > 0 Then
            MessageBox.Show("Conectado a Odoo")
            ' Cargar la lista de socios
            CargarSocios()
        Else
            MessageBox.Show("Error al conectar con Odoo")
        End If
    End Sub

    ' Cargar datos en el DataGridView
    Private Sub CargarSocios()
        Dim dataTable As DataTable = OdooConnector.GetPartnerList(userId)
        DataGridView1.DataSource = dataTable
        ' Ocultar la columna ID
        DataGridView1.Columns("ID").Visible = False
    End Sub

    ' Evento que se dispara al seleccionar una fila en el DataGridView
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        ' Verificar si hay alguna fila seleccionada
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Validar que el ID del socio no sea nulo
            If Not IsDBNull(row.Cells("ID").Value) Then
                selectedPartnerId = Convert.ToInt32(row.Cells("ID").Value)
            Else
                selectedPartnerId = -1 ' Si es nulo, establecer un valor predeterminado
            End If

            ' Verificar si el nombre y el email no son nulos antes de asignarlos
            If Not IsDBNull(row.Cells("Nombre").Value) Then
                txtNombre.Text = row.Cells("Nombre").Value.ToString()
            Else
                txtNombre.Text = ""
            End If

            If Not IsDBNull(row.Cells("Email").Value) Then
                txtEmail.Text = row.Cells("Email").Value.ToString()
            Else
                txtEmail.Text = ""
            End If
        End If
    End Sub

    ' Crear un nuevo socio
    Private Sub btnCrear_Click(sender As Object, e As EventArgs) Handles btnCrear.Click
        If Not String.IsNullOrEmpty(txtNombre.Text) AndAlso Not String.IsNullOrEmpty(txtEmail.Text) Then
            Try
                ' Llamar a la función de creación en OdooConnector
                OdooConnector.CreatePartner(userId, txtNombre.Text, txtEmail.Text)
                MessageBox.Show("Nuevo socio creado con éxito.")
                ' Recargar la lista de socios
                CargarSocios()
            Catch ex As Exception
                MessageBox.Show("Error al crear el socio: " & ex.Message)
            End Try
        Else
            MessageBox.Show("El nombre y el email no pueden estar vacíos.")
        End If
    End Sub

    ' Actualizar un socio
    Private Sub btnActualizar_Click(sender As Object, e As EventArgs) Handles btnActualizar.Click
        If selectedPartnerId <> -1 AndAlso Not String.IsNullOrEmpty(txtNombre.Text) Then
            Try
                ' Llamar a la función de actualización en OdooConnector
                OdooConnector.UpdatePartner(userId, selectedPartnerId, txtNombre.Text)
                MessageBox.Show("Socio actualizado con éxito.")
                ' Recargar la lista de socios
                CargarSocios()
            Catch ex As Exception
                MessageBox.Show("Error al actualizar el socio: " & ex.Message)
            End Try
        Else
            MessageBox.Show("Seleccione un socio válido y asegúrese de que el nombre no esté vacío.")
        End If
    End Sub

    ' Eliminar un socio
    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        If selectedPartnerId <> -1 Then
            ' Confirmar eliminación
            Dim result = MessageBox.Show("¿Está seguro de que desea eliminar este socio?", "Confirmación", MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then
                Try
                    ' Llamar a la función de eliminación en OdooConnector
                    OdooConnector.DeletePartner(userId, selectedPartnerId)
                    MessageBox.Show("Socio eliminado con éxito.")
                    ' Recargar la lista de socios
                    CargarSocios()
                Catch ex As Exception
                    MessageBox.Show("Error al eliminar el socio: " & ex.Message)
                End Try
            End If
        Else
            MessageBox.Show("Seleccione un socio válido para eliminar.")
        End If
    End Sub

    ' Exportar a Excel usando EPPlus
    Private Sub btnExportar_Click(sender As Object, e As EventArgs) Handles btnExportar.Click
        ExportarDataGridViewAExcel(DataGridView1)
    End Sub


    ' Función para exportar el DataGridView a un archivo Excel con control de errores
    Public Sub ExportarDataGridViewAExcel(dataGridView As DataGridView)
        ' Crear el diálogo para que el usuario seleccione dónde guardar el archivo
        Using saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel Files|*.xlsx"
            saveFileDialog.Title = "Guardar archivo Excel"
            saveFileDialog.FileName = "Informe_Socios.xlsx"

            ' Si el usuario selecciona un lugar y un nombre válido
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                ' Obtener la ruta seleccionada por el usuario
                Dim filePath As String = saveFileDialog.FileName

                ' Verificar si el archivo ya está en uso
                If IsFileInUse(filePath) Then
                    MessageBox.Show("El archivo ya está abierto. Cierre el archivo antes de exportar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Try
                    ' Establecer la licencia como uso no comercial
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial

                    ' Crear el archivo Excel
                    Dim fileInfo As New FileInfo(filePath)
                    Using package As New ExcelPackage(fileInfo)
                        Dim worksheet = package.Workbook.Worksheets.Add("Socios")

                        ' Escribir las cabeceras
                        For i As Integer = 0 To dataGridView.Columns.Count - 1
                            worksheet.Cells(1, i + 1).Value = dataGridView.Columns(i).HeaderText
                        Next

                        ' Escribir los datos del DataGridView
                        For i As Integer = 0 To dataGridView.Rows.Count - 1
                            For j As Integer = 0 To dataGridView.Columns.Count - 1
                                worksheet.Cells(i + 2, j + 1).Value = dataGridView.Rows(i).Cells(j).Value
                            Next
                        Next

                        ' Guardar el archivo Excel
                        package.Save()
                    End Using

                    ' Mostrar un mensaje de éxito
                    MessageBox.Show("Datos exportados a Excel con éxito")

                    ' Abrir el archivo automáticamente
                    Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})

                Catch ex As IOException
                    MessageBox.Show("Error al escribir en el archivo. Asegúrese de que el archivo no esté abierto.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Catch ex As Exception
                    MessageBox.Show("Se produjo un error inesperado: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    ' Función para verificar si un archivo está en uso
    Private Function IsFileInUse(filePath As String) As Boolean
        Try
            ' Intentar abrir el archivo con acceso exclusivo
            Using fileStream As FileStream = File.Open(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None)
            End Using
            Return False ' Si no hay excepción, el archivo no está en uso
        Catch ex As IOException
            Return True ' El archivo está en uso si se lanza una excepción
        End Try
    End Function

End Class

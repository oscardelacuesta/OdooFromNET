Imports CookComputing.XmlRpc

' Define la interfaz para la conexión XML-RPC
Public Interface IOdooRpcCommon
    <XmlRpcMethod("login")>
    Function Login(db As String, user As String, pass As String) As Integer
End Interface

Public Interface IOdooRpcObject
    <XmlRpcMethod("execute")>
    Function Execute(db As String, uid As Integer, pwd As String, model As String, method As String, ParamArray args As Object()) As Object
End Interface

Public Class OdooConnector
    ' Ajustar la URL, base de datos, usuario y contraseña
    Private Const OdooUrl As String = "http://pruebas.publicvm.com/xmlrpc/2/"
    Private Const OdooDb As String = "odoo"
    Private Const OdooUser As String = "demo@demoesdemo.com"
    Private Const OdooPassword As String = "demostracion"

    ' Conexión y autenticación con Odoo
    Public Shared Function ConnectToOdoo() As Integer
        Try
            ' Crear el proxy para el servicio común (common)
            Dim proxyCommon As IOdooRpcCommon = CType(XmlRpcProxyGen.Create(GetType(IOdooRpcCommon)), IOdooRpcCommon)
            CType(proxyCommon, IXmlRpcProxy).Url = OdooUrl & "common"

            ' Realizar la autenticación
            Dim userId As Integer = proxyCommon.Login(OdooDb, OdooUser, OdooPassword)

            If userId > 0 Then
                Console.WriteLine("Conexión exitosa. ID de usuario: " & userId)
                Return userId
            Else
                Console.WriteLine("Error de autenticación.")
                Return -1
            End If
        Catch ex As Exception
            MessageBox.Show("Error al conectar con Odoo: " & ex.Message)
            Return -1
        End Try
    End Function

    ' Crear nuevo socio
    Public Shared Sub CreatePartner(userId As Integer, name As String, email As String)
        Try
            ' Crear el proxy para el servicio de objetos (object)
            Dim proxyObject As IOdooRpcObject = CType(XmlRpcProxyGen.Create(GetType(IOdooRpcObject)), IOdooRpcObject)
            CType(proxyObject, IXmlRpcProxy).Url = OdooUrl & "object"

            ' Crear un XmlRpcStruct para los datos del nuevo socio
            Dim newPartner As New XmlRpcStruct From {
                {"name", name},
                {"email", email}
            }

            ' Llamar al método 'create' de Odoo para insertar el nuevo socio
            Dim partnerId = proxyObject.Execute(OdooDb, userId, OdooPassword, "res.partner", "create", New Object() {newPartner})

            If partnerId > 0 Then
                MessageBox.Show("Nuevo socio creado con ID: " & partnerId)
            Else
                MessageBox.Show("Error al crear el socio.")
            End If
        Catch ex As Exception
            MessageBox.Show("Error al conectar con Odoo: " & ex.Message)
        End Try
    End Sub

    ' Obtener lista de socios
    Public Shared Function GetPartnerList(userId As Integer) As DataTable
        Try
            ' Crear el proxy para el servicio de objetos (object)
            Dim proxyObject As IOdooRpcObject = CType(XmlRpcProxyGen.Create(GetType(IOdooRpcObject)), IOdooRpcObject)
            CType(proxyObject, IXmlRpcProxy).Url = OdooUrl & "object"

            ' Obtener la lista de socios
            Dim result = proxyObject.Execute(OdooDb, userId, OdooPassword, "res.partner", "search_read", New Object() {}, New Object() {"id", "name", "email"})

            ' Crear DataTable
            Dim dataTable As New DataTable()
            dataTable.Columns.Add("ID", GetType(Integer)) ' Columna oculta para el ID del socio
            dataTable.Columns.Add("Nombre", GetType(String))
            dataTable.Columns.Add("Email", GetType(String))

            ' Rellenar la tabla con los datos obtenidos
            For Each partner As XmlRpcStruct In result
                Dim row As DataRow = dataTable.NewRow()
                row("ID") = partner("id") ' Asegúrate de obtener el ID del socio
                row("Nombre") = partner("name")
                row("Email") = partner("email")
                dataTable.Rows.Add(row)
            Next

            Return dataTable
        Catch ex As Exception
            MessageBox.Show("Error al obtener la lista de socios: " & ex.Message)
            Return Nothing
        End Try
    End Function

    ' Actualizar un socio
    Public Shared Sub UpdatePartner(userId As Integer, partnerId As Integer, newName As String)
        Try
            ' Crear el proxy para el servicio de objetos (object)
            Dim proxyObject As IOdooRpcObject = CType(XmlRpcProxyGen.Create(GetType(IOdooRpcObject)), IOdooRpcObject)
            CType(proxyObject, IXmlRpcProxy).Url = OdooUrl & "object"

            ' Crear un diccionario para pasar los datos de actualización
            Dim updateValues As New XmlRpcStruct From {
                {"name", newName}
            }

            ' Llamar al método 'write' de Odoo para actualizar el socio
            Dim result = proxyObject.Execute(OdooDb, userId, OdooPassword, "res.partner", "write", New Object() {New Integer() {partnerId}, updateValues})

            If result Then
                MessageBox.Show("Socio actualizado correctamente.")
            Else
                MessageBox.Show("Error al actualizar el socio.")
            End If
        Catch ex As Exception
            MessageBox.Show("Error al conectar con Odoo: " & ex.Message)
        End Try
    End Sub

    ' Eliminar un socio
    Public Shared Sub DeletePartner(userId As Integer, partnerId As Integer)
        Try
            ' Crear el proxy para el servicio de objetos (object)
            Dim proxyObject As IOdooRpcObject = CType(XmlRpcProxyGen.Create(GetType(IOdooRpcObject)), IOdooRpcObject)
            CType(proxyObject, IXmlRpcProxy).Url = OdooUrl & "object"

            Dim result = proxyObject.Execute(OdooDb, userId, OdooPassword, "res.partner", "unlink", New Object() {partnerId})

            If result Then
                MessageBox.Show("Socio eliminado correctamente.")
            Else
                MessageBox.Show("Error al eliminar el socio.")
            End If
        Catch ex As Exception
            MessageBox.Show("Error al conectar con Odoo: " & ex.Message)
        End Try
    End Sub
End Class

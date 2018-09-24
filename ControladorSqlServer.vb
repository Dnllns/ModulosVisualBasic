Imports System.Data.SqlClient

Module ControladorSqlServer

    ''' <summary>
    ''' Clase que facilita el acceso a una base de datos de SQLServer
    ''' </summary>
    Public Class ControladorBD

        'PROPIEDADES
        Property ConnectionString As String
        Property Connection As SqlConnection

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="connectionString">La cadena de conexion de acceso a la BD</param>
        Public Sub New(ByVal connectionString As String)
            Me.ConnectionString = connectionString
            Me.Connection = New SqlConnection(connectionString)
        End Sub

        ''' <summary>
        ''' Abre la conexion
        ''' </summary>
        Public Sub AbrirConexion()
            Try
                Connection.Open()
            Catch ex As Exception
                System.Console.WriteLine("No se ha podido abrir la conexion con la base de datos: " & ex.Message)
            End Try

        End Sub

        ''' <summary>
        ''' Cierra la conexion
        ''' </summary>
        Public Sub CerrarConexion()
            Try
                Connection.Close()
            Catch ex As Exception
                System.Console.WriteLine("No se ha podido cerrar la conexion con la base de datos: " & ex.Message)
            End Try
        End Sub

        '-----------------------------------------------------
        '------------------METODOS DE SELECT------------------

        ''' <summary>
        ''' Facilita el control de una determinada tabla ya que permite obtener el objeto dataSet
        ''' </summary>
        ''' <param name="query">La consulta que generara el dataset</param>
        ''' <returns>Devuelve el objeto Dataset equivalente</returns>
        Public Function GetDataset(query As String) As DataSet
            Dim dataAdapter = New SqlDataAdapter(query, Connection)
            Dim dataSet As New DataSet
            dataAdapter.Fill(dataSet)           'Cargar registros en el dataSet
            Return dataSet
        End Function


        ''' <summary>
        ''' Permite realizar una consulta de tipo Select
        ''' </summary>
        ''' <param name="query">Es la instrucion SQL</param>
        ''' <returns>
        ''' Devuelve un List de DataRow que contiene los registros resultado de la consulta
        ''' </returns>
        Public Function ExecuteSelect(query As String) As List(Of DataRow)
            Dim registros As New List(Of DataRow)
            Dim ds As DataSet = GetDataset(query)
            Dim dt As DataTable = ds.Tables.Item(0)     'Seleccionamos la tabla que nos interesa del dataset
            For Each row As DataRow In dt.Rows          'Almacenar cada registro en el List
                registros.Add(row)
            Next
            ds.Dispose()                                'Cerrar el dataSet una vez usado
            Return registros
        End Function


        ''' <summary>
        ''' Ejecuta una consulta a la Bd de un solo dato
        ''' </summary>
        ''' <param name="query"> La consuta Sql </param>
        ''' <returns>devuelve un string que contiene el valor del escalar buscado</returns>
        Public Function ExecuteEscalar(query As String) As String
            Dim comando As New SqlCommand(query, Connection)
            Connection.Open()
            Dim busqueda As String = comando.ExecuteScalar().ToString
            Connection.Close()
            Return busqueda
        End Function



        '------------------------------------------------------
        '------------------METODOS DE CONTROL------------------
        ''' <summary>
        ''' Ejecuta una sentencia sql que no devuelve ningun dato
        ''' </summary>
        ''' <param name="sql">La instruccion sql</param>
        Public Sub ExecuteNonQuery(sql As String)
            Dim comando As New SqlCommand(sql, Connection)
            Connection.Open()
            Try
                comando.ExecuteNonQuery()
            Catch ex As Exception
                System.Console.WriteLine("Error ejecutando la instruccion sql : " & ex.Message)
            End Try
            Connection.Close()
        End Sub



        'EJEMPLO DE LECTURA
        'Public Sub ExecuteRead(sql As String)
        '    Dim comando As New SqlCommand(sql, Connection)
        '    Connection.Open()
        '    Dim reader As SqlDataReader = comando.ExecuteReader
        '    If reader.HasRows Then
        '        While reader.Read
        '            'realizar lo querido
        '        End While
        '    End If
        '    Connection.Close()
        'End Sub


    End Class
End Module


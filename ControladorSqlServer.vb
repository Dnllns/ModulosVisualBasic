Imports System.Data.SqlClient

Module ControladorSqlServer

    ''' <summary>
    ''' Clase que facilita el acceso a una base de datos de SQLServer
    ''' </summary>
    Public Class ControladorBD

        'PROPIEDADES
        Property ConnectionString As String
        Property Connection As SqlConnection
        Property DataAdapter As SqlDataAdapter

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
            Connection.Open()
        End Sub

        ''' <summary>
        ''' Cierra la conexion
        ''' </summary>
        Public Sub CerrarConexion()
            Connection.Close()
        End Sub


        ''' <summary>
        ''' Facilita el control de una determinada tabla ya que permite obtener el objeto dataSet
        ''' </summary>
        ''' <param name="query">La consulta que generara el dataset</param>
        ''' <returns>Devuelve el objeto Dataset equivalente</returns>
        Public Function GetDataset(query As String) As DataSet
            Dim dataSet As New DataSet
            DataAdapter = New SqlDataAdapter(query, Connection)
            DataAdapter.Fill(dataSet)           'Cargar registros en el dataSet
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


    End Class
End Module



Imports System.Data.Common
Imports System.Runtime.Remoting.Lifetime

Public Class clsDBCommand
    Inherits MarshalByRefObject
    Implements System.Data.IDbCommand

    Private prvLisStrAlias As List(Of String)
    Private prvObjDBCommand As DbCommand
    Private prvIntParamsRetorno As Integer
    Private prvIntTiempoVida As Integer
    Private prvLease As ILease

    ' ''Public Sub ParametrosRemover(ByVal pvparametro As DbParameter)
    ' ''    prvLisStrAlias.RemoveAt(Me.Parameters.IndexOf(pvparametro.ParameterName))
    ' ''    Me.Parameters.Remove(pvparametro)

    ' ''End Sub

    Public Sub ParametrosRemover(ByVal pvparametro As DbParameter)
        prvLisStrAlias.RemoveAt(Me.Parameters.IndexOf(pvparametro.ParameterName.ToUpper))
        Me.Parameters.Remove(pvparametro)

    End Sub

    Public Sub New(ByVal pvCommand As DbCommand)
        prvObjDBCommand = pvCommand
        prvLisStrAlias = New List(Of String)
        prvIntParamsRetorno = 0
        prvIntTiempoVida = 0
    End Sub
    Public ReadOnly Property pubobjCommand() As DbCommand
        Get
            Return prvObjDBCommand
        End Get
    End Property

    Public Sub addParamRetorno()
        prvIntParamsRetorno = prvIntParamsRetorno + 1
    End Sub

    Public Sub Cancel() Implements System.Data.IDbCommand.Cancel

    End Sub

    Public Property CommandText() As String Implements System.Data.IDbCommand.CommandText
        Get
            Return prvObjDBCommand.CommandText

        End Get
        Set(ByVal value As String)
            prvObjDBCommand.CommandText = value
        End Set
    End Property

    Public Property CommandTimeout() As Integer Implements System.Data.IDbCommand.CommandTimeout
        Get
            Return prvObjDBCommand.CommandTimeout
        End Get
        Set(ByVal value As Integer)
            prvObjDBCommand.CommandTimeout = value
        End Set
    End Property

    Public Property CommandType() As System.Data.CommandType Implements System.Data.IDbCommand.CommandType
        Get
            Return prvObjDBCommand.CommandType
        End Get
        Set(ByVal value As System.Data.CommandType)
            prvObjDBCommand.CommandType = value
        End Set
    End Property

    Public Property Connection() As System.Data.IDbConnection Implements System.Data.IDbCommand.Connection
        Get
            Return prvObjDBCommand.Connection
        End Get
        Set(ByVal value As System.Data.IDbConnection)
            prvObjDBCommand.Connection = value
        End Set
    End Property

    Public Function CreateParameter() As System.Data.IDbDataParameter Implements System.Data.IDbCommand.CreateParameter
        Return prvObjDBCommand.CreateParameter()
    End Function

    Public Function ExecuteNonQuery() As Integer Implements System.Data.IDbCommand.ExecuteNonQuery
        Return prvObjDBCommand.ExecuteNonQuery
    End Function

    Public Function ExecuteReader() As System.Data.IDataReader Implements System.Data.IDbCommand.ExecuteReader
        Return prvObjDBCommand.ExecuteReader
    End Function

    Public Function ExecuteReader(ByVal behavior As System.Data.CommandBehavior) As System.Data.IDataReader Implements System.Data.IDbCommand.ExecuteReader
        Return prvObjDBCommand.ExecuteReader(behavior)
    End Function

    Public Function ExecuteScalar() As Object Implements System.Data.IDbCommand.ExecuteScalar
        Return prvObjDBCommand.ExecuteScalar
    End Function

    Public ReadOnly Property Parameters() As System.Data.IDataParameterCollection Implements System.Data.IDbCommand.Parameters
        Get
            Return prvObjDBCommand.Parameters
        End Get
    End Property

    Public Property Parameters(ByVal vStrNombre As String) As DbParameter
        Get
            'Tenemos el nombre real
            Dim pvIntPosParam = GetParamPos(vStrNombre)
            Return prvObjDBCommand.Parameters(pvIntPosParam)
        End Get
        Set(ByVal value As DbParameter)

        End Set
    End Property
    Public Property Parameters(ByVal vIntIndex As Integer) As DbParameter
        Get
            'Tenemos el nombre real
            Return prvObjDBCommand.Parameters(vIntIndex)
        End Get
        Set(ByVal value As DbParameter)

        End Set
    End Property

    Public ReadOnly Property CuentaParams() As Integer
        Get
            CuentaParams = Parameters.Count - prvIntParamsRetorno
        End Get
    End Property


    Public Sub Prepare() Implements System.Data.IDbCommand.Prepare
        prvObjDBCommand.Prepare()
    End Sub

    Public Property Transaction() As System.Data.IDbTransaction Implements System.Data.IDbCommand.Transaction
        Get
            Return prvObjDBCommand.Transaction
        End Get
        Set(ByVal value As System.Data.IDbTransaction)
            prvObjDBCommand.Transaction = value
        End Set
    End Property

    Public Property UpdatedRowSource() As System.Data.UpdateRowSource Implements System.Data.IDbCommand.UpdatedRowSource
        Get
            Return prvObjDBCommand.UpdatedRowSource
        End Get
        Set(ByVal value As System.Data.UpdateRowSource)
            prvObjDBCommand.UpdatedRowSource = value
        End Set
    End Property

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free other state (managed objects).
                prvObjDBCommand = Nothing
            End If

            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    Public Sub Dispose() Implements System.IDisposable.Dispose

        If Not prvObjDBCommand Is Nothing Then
            prvObjDBCommand.Connection = Nothing
            prvObjDBCommand.Transaction = Nothing

            prvObjDBCommand.Dispose()
            prvObjDBCommand = Nothing

        End If
    End Sub

    ''#Region "Metodo nuevo"
    Public Sub AddAlias(ByVal pvStParamAlias As String)
        prvLisStrAlias.Add(pvStParamAlias.ToUpper)
    End Sub

    Public Function GetParamPos(ByVal pvStrAliasName As String) As Integer
        Dim pvIntPosRes = prvLisStrAlias.IndexOf(pvStrAliasName.ToUpper)
        If pvIntPosRes >= 0 Then
            GetParamPos = pvIntPosRes
        Else
            Err.Raise(1, "UnoeeDatos.clsDBCommand", "Parametro no encontrado en el command PARAMETRO: " & pvStrAliasName)
        End If
    End Function
    '#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class





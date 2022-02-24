'Clase especial para cada ENT, Para usos generales, como el leer versión de la DLL
Public Class Application

    ' **************************************************************
    ' Nombre de función: ObtenerVersion
    ' Proposito: Obtener la versión
    ' Retorna: La versión como un string
    ' Creado por:     dfo
    ' Modificado por:  dfo
    ' **************************************************************
    Public Shared Function ObtenerVersion() As String
        Dim version As System.Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
        ObtenerVersion = version.Major & "." & version.Minor & "." & version.Build
    End Function

    '********************************************************
    '******************** VB6Migracion **********************
    '********************************************************
    '<Serializable()>
    'Public Class clsError

    '    Public ReadOnly Property NativeError() As Long

    'End Class

    '<Serializable()>
    'Public Class clsErrors

    '    Public Sub Clear()
    '    End Sub

    '    Public ReadOnly Property Count() As Long

    '    Public Property Item(Index As Int16) As clsError
    '        Get
    '        End Get
    '        Set(value As clsError)
    '        End Set
    '    End Property

    '    Public Sub Refresh()
    '    End Sub

    'End Class

    '********************************************************
    '********************************************************
End Class

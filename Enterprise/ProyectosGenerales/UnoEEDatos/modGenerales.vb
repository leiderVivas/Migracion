Imports System.IO
Imports System.Reflection

Module modGenerales
    Public Function LeerNombreAppLlamado() As String

        'Return ""
        'Exit Function

        Dim vstrCallStack As String()
        ''vstrCallStack = Environment.StackTrace.ToString().Split(vbCrLf, options:=StringSplitOptions.RemoveEmptyEntries)
        Dim vstrCallStack3 As New System.Diagnostics.StackTrace
        vstrCallStack = vstrCallStack3.ToString().Split(vbCrLf, options:=StringSplitOptions.RemoveEmptyEntries)


        'MsgBox(Environment.StackTrace.ToString())

        Dim vstrCallStack2 = (From vstrMetodo In vstrCallStack
                              Where (Trim(vstrMetodo).StartsWith("at UnoEE") Or Trim(vstrMetodo).StartsWith("at Srv")) And Not Trim(vstrMetodo).StartsWith("at UnoEEDatos") And Not Trim(vstrMetodo).StartsWith("at UnoEEGeneral") And Not Trim(vstrMetodo).StartsWith("at UnoEEDatos")
                              Select Trim(vstrMetodo)).ToList.FirstOrDefault


        ''' Grabs the entry assembly.
        'Dim vobjentryAssembly As Assembly = Assembly.GetCallingAssembly()
        '' Grabs its name
        'Dim vbojentryAssemblyName As AssemblyName = vobjentryAssembly.GetName()

        'LeerNombreAppLlamado = vbojentryAssemblyName.Name


        If Not vstrCallStack2 Is Nothing Then
            'vstrCallStack2 = vstrCallStack2.Substring(3, Len(vstrCallStack2) - 4)
            ' EscribirArchivo("UnoeeDatos.clsrecordset", vbojentryAssemblyName.Name)
            'EscribirArchivo("UnoeeDatos.clsrecordset", vstrCallStack3.ToString())

            'EscribirArchivo("UnoeeDatos.clsrecordset", Now.ToString)
            Return vstrCallStack2.Substring(3, vstrCallStack2.IndexOf(".") - 3)
        Else
            Return ""
        End If


        'vobjentryAssembly = Nothing
        'vbojentryAssemblyName = Nothing

    End Function

    Public Function LlamadoDesdeServidor() As Boolean
        'LlamadoDesdeServidor = False
        Return LeerNombreAppLlamado.StartsWith("Srv")
    End Function


    ''' <summary>
    ''' Escribe un archivo en el directorio especificado y nombre de archivo 
    ''' </summary>
    ''' <param name="pvstrNombreSistema">Nombre del sistema</param>
    ''' <param name="pvstrLinea">Mensaje a escribir</param>
    ''' <param name="pvstrNombreArchivo">Nombre del archivo, si no se envia o escribe en el archivo, SiesaEELog.log</param>
    ''' <param name="pvstrDirectorio"></param>
    Friend Sub EscribirArchivo(pvstrNombreSistema As String, pvstrLinea As String, Optional pvstrNombreArchivo As String = "SiesaEELog.log", Optional pvstrDirectorio As String = "")

        Dim vstrDirectorio As String = pvstrDirectorio
        Try
            ' Si no viene directorio asume el directorio actual
            If vstrDirectorio = "" Then
                vstrDirectorio = System.AppDomain.CurrentDomain.BaseDirectory
            End If
            'Dim sw As StreamWriter = New StreamWriter(System.AppDomain.CurrentDomain.BaseDirectory & "\" & pvstrNombreArchivo, True)
            Dim sw As StreamWriter = New StreamWriter(vstrDirectorio & "\" & pvstrNombreArchivo, True)
            sw.WriteLine("Sistema: " & pvstrNombreSistema & "; " & pvstrLinea)
            sw.Flush()
            sw.Close()
        Catch ex As Exception

        End Try

    End Sub

    Friend Function LeerDatoADO(ByVal VRec As UnoEEDatos.clsRecordset, ByVal sVName As String, Optional ByVal vntVDefault As Object = Nothing,
                                         Optional ByVal vntVVariableAsignar As Object = Nothing) As Object
        On Error GoTo errLeerDatoADO

        Dim vintType As DbType



        vintType = VRec.Fields(sVName).Type

        If IsDBNull(VRec.Fields(sVName).Value) Then

            ' Por jairc 20210112
            ' Se comenta porque la logica en vb6 no asigna nunca el valor defecto enviado

            'If vntVDefault IsNot Nothing Then
            '    Select Case VRec.Fields(sVName).Type

            '        Case DbType.Int16, DbType.Int32, DbType.Int64, DbType.Double, DbType.Decimal
            '            ' If Not IsNumeric(vntVDefault) Then
            '            Return 0
            '            'Else
            '    'Return vntVDefault
            ''End If

            '        Case DbType.Date, DbType.DateTime
            '            If Not IsDate(vntVDefault) Then
            '                Return Nothing
            '            Else
            '                Return vntVDefault
            '            End If

            '        Case Else
            '            Return vntVDefault
            '    End Select

            'Else

            ' Por jairc 2020126
            ' Si la variable a asignar el tipo de dato es fecha, garantiza que se devuelva una fecha
            ' para que se asigne a la variable correcta
            Select Case True
                Case TypeOf vntVVariableAsignar Is Date
                    ' Tipo Fecha
                    vintType = DbType.Date
                Case TypeOf vntVVariableAsignar Is String
                    ' Tipo String
                    vintType = DbType.String
            End Select

            Select Case vintType
                Case DbType.String, DbType.StringFixedLength
                    Return String.Empty
                Case DbType.Int16, DbType.Int32, DbType.Int64
                    Return 0
                Case DbType.Double, DbType.Decimal
                    Return 0.0#
                Case DbType.Date, DbType.DateTime
                    Return Nothing
                Case DbType.Guid
                    Return "" ' Guid.Empty
                Case Else
                    Return Nothing
            End Select

            'Select Case rec.Fields(strName).Type
            '    Case adChar '1   Fixed-length character string.  Length set by Size property.
            '        GblLeerDatoADOExt = Empty
            '    Case adNumeric, adInteger, adSmallInt, adBigInt, adTinyInt
            '        GblLeerDatoADOExt = 0
            '    Case adDecimal, adDouble
            '        GblLeerDatoADOExt = 0#
            '    Case adDate, adVarChar 'rdTypeLONGVARCHAR, rdTypeBINARY,rdTypeVARBINARY, rdTypeLONGVARBINARY
            '        GblLeerDatoADOExt = Empty
            '    Case Else
            '        GblLeerDatoADOExt = Empty
            'End Select
            'End If
        Else
            Select Case vintType
                Case DbType.Guid
                    ' por jairc 20201204
                    ' Se tipifica el tipo de dato GUID porque la logica en vb6 retorna un string
                    Return VRec.Fields(sVName).Value.ToString
                Case Else
                    Return VRec.Fields(sVName).Value
            End Select

        End If

        'Select Case VarType(VRec.Fields(sVName).Value)
        '    Case vbEmpty  '0   Empty (uninitialized).
        '        GblLeerDatoADO = Nothing
        '    Case vbNull   '1   Null (no valid VRec .rdoColumns(sVname)a).
        '        Select Case VRec.Fields(sVName).Type

        '            Case DbType.StringFixedLength ' adChar   ' Fixed-length character string.  Length set by Size property.
        '                GblLeerDatoADO = String.Empty
        '            Case DbType.Int16, DbType.Int32, DbType.Int64 '   adNumeric, adInteger, adSmallInt, adBigInt, adTinyInt
        '                GblLeerDatoADO = 0
        '            Case DbType.Decimal, DbType.Double ' adDecimal, adDouble
        '                GblLeerDatoADO = 0.0#
        '            Case DbType.DateTime, DbType.String, DbType.Binary   ' adDate, adVarChar 'rdTypeLONGVARCHAR, rdTypeBINARY,rdTypeVARBINARY, rdTypeLONGVARBINARY
        '                GblLeerDatoADO = String.Empty
        '            Case Else
        '                GblLeerDatoADO = Nothing ' Null
        '        End Select
        '    Case Else
        '        GblLeerDatoADO = VRec.Fields(sVName).Value
        'End Select
        Exit Function

errLeerDatoADO:
        LeerDatoADO = vntVDefault
    End Function

    Friend Function LeerDatoADO(ByVal VRec As UnoEEDatos.clsRecordset, ByVal pvIntColumnIndex As Integer, Optional ByVal vntVDefault As Object = Nothing,
                                          Optional ByVal vntVVariableAsignar As Object = Nothing) As Object
        On Error GoTo errLeerDatoADO
        Dim vintType As DbType

        vintType = VRec.Fields(pvIntColumnIndex).Type

        If IsDBNull(VRec.Fields(pvIntColumnIndex).Value) Then
            ' Por jairc 20210112
            ' Se comenta porque la logica en vb6 no asigna nunca el valor defecto enviado
            'If vntVDefault IsNot Nothing Then
            '    Return vntVDefault
            'Else
            ' Por jairc 2020126
            ' Si la variable a asignar el tipo de dato es fecha, garantiza que se devuelva una fecha
            ' para que se asigne a la variable correcta
            Select Case True
                Case TypeOf vntVVariableAsignar Is Date
                    ' Tipo Fecha
                    vintType = DbType.Date
                Case TypeOf vntVVariableAsignar Is String
                    ' Tipo String
                    vintType = DbType.String
            End Select

            Select Case vintType
                Case DbType.String, DbType.StringFixedLength
                    Return String.Empty
                Case DbType.Int16, DbType.Int32, DbType.Int64
                    Return 0
                Case DbType.Double, DbType.Decimal
                    Return 0.0#
                Case DbType.Date, DbType.DateTime
                    Return Nothing
                Case DbType.Guid
                    Return "" ' Guid.Empty
                Case Else
                    Return Nothing
            End Select
            'End If
        Else
            Select Case vintType
                Case DbType.Guid
                    ' por jairc 20201204
                    ' Se tipifica el tipo de dato GUID porque la logica en vb6 retorna un string
                    Return VRec.Fields(pvIntColumnIndex).Value.ToString
                Case Else
                    Return VRec.Fields(pvIntColumnIndex).Value
            End Select

        End If

        Exit Function


errLeerDatoADO:
        LeerDatoADO = vntVDefault
    End Function

    'LRV Req. 191472
    'funcion que determina el tipo de dato de la columna en un ViewVista
    'cuando es Nothing debe devolver el valor 0 para simular el comportamiento en vb6
    'solo se valida cuando la columna es de tipo integer
    Friend Function LeerDatoDataView(ByVal pvObjRecordSet As UnoEEDatos.clsRecordset, ByVal pvRow As Long, ByVal pvCol As Object, prvblnServer As Boolean, Optional pvblnReturnNullSinCondicion As Boolean = False) As Object
        Try
            Dim vintType As DbType
            vintType = pvObjRecordSet.Fields(pvCol, pvblnServer:=prvblnServer).Type

            If pvObjRecordSet.prvlngPosicion = UnoEEDatos.clsRecordset.pubPositionEnum.adPosUnknown And pvObjRecordSet.PrvBlnTablaFiltrada Then

                If IsDBNull(pvObjRecordSet.pubObjDtTabla(pvObjRecordSet.IndiceRowFilter).Item(pvCol)) Then
                    ' por jairc 20210528
                    ' Cuadno se evalua un numero debe retornar nulo en algunos casos para ser utilizado con IsDBNull
                    If pvblnReturnNullSinCondicion Then
                        Return pvObjRecordSet.pubObjDtTabla(pvObjRecordSet.IndiceRowFilter).Item(pvCol)
                    Else
                        Select Case vintType
                            Case DbType.Int16, DbType.Int32, DbType.Int64, DbType.Decimal
                                Return 0
                            Case Else
                                Return pvObjRecordSet.pubObjDtTabla(pvObjRecordSet.IndiceRowFilter).Item(pvCol)
                        End Select
                    End If

                Else
                    ' por jaic 20121210. OP 210409
                    ' Cuando es GUId debe retornar como string, porque el manejo de tipo de dato GUI puro no se maneja en vb6
                    Select Case vintType
                        Case DbType.Guid
                            Return pvObjRecordSet.pubObjDtTabla(pvObjRecordSet.IndiceRowFilter).Item(pvCol).ToString
                        Case Else
                            Return pvObjRecordSet.pubObjDtTabla(pvObjRecordSet.IndiceRowFilter).Item(pvCol)
                    End Select

                End If
            Else

                If IsDBNull(pvObjRecordSet.PrvDtViewVista(pvRow).Item(pvCol)) Then
                    ' por jairc 20210528
                    ' Cuadno se evalua un numero debe retornar nulo en algunos casos para ser utilizado con IsDBNull
                    If pvblnReturnNullSinCondicion Then
                        Return pvObjRecordSet.PrvDtViewVista(pvRow).Item(pvCol)
                    Else
                        Select Case vintType
                            Case DbType.Int16, DbType.Int32, DbType.Int64, DbType.Decimal
                                Return 0
                            Case Else
                                Return pvObjRecordSet.PrvDtViewVista(pvRow).Item(pvCol)
                        End Select
                    End If
                Else
                    ' por jaic 20121210. OP 210409
                    ' Cuando es GUId debe retornar como string, porque el manejo de tipo de dato GUI puro no se maneja en vb6
                    Select Case vintType
                        Case DbType.Guid
                            Return pvObjRecordSet.PrvDtViewVista(pvRow).Item(pvCol).ToString
                        Case Else
                            Return pvObjRecordSet.PrvDtViewVista(pvRow).Item(pvCol)
                    End Select

                End If
            End If

        Catch ex As Exception
            Try ' Puede sacar error si el campo no existe
                LeerDatoDataView = pvObjRecordSet.PrvDtViewVista(pvRow).Item(pvCol)
            Catch ex1 As Exception
                LeerDatoDataView = Nothing
            End Try

        End Try


    End Function
End Module

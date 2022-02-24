Option Compare Text
Imports System.Data
Imports System.Reflection
Imports System.Runtime.Serialization
Imports UnoEEDatos
Imports System.Text.RegularExpressions

<Serializable()>
Public Class clsRecordset
    Implements ISerializable
    'Inherits MarshalByRefObject

    'Private prvobValor As UnoEEDatos.clsValor
    Private prvstrPath As String = ""
    Private _prvStrSortClient As String = ""
    Private _prvStrSortServer As String = ""

    ' Por jairc, 164499
    ' Cuando llega el recordset desde el clienteb filtrado, en vb6 al pasarlo al servidor se elimina el filter
    ' en .net no lo elimina, se va a validar si el filtro fue realizado en el cliente o el servidor
    Private _prvStrFilterClient As String = ""
    Private _prvStrFilterServer As String = ""

    Private prvintCursorLocation As clsRecordset.pubEnumCursorLocationEnum
    Private prvintCursorType As clsRecordset.pubEnumCursorTypeEnum
    Private prvintLockType As clsRecordset.pubEnumLockTypeEnum

    '#Region "Local Variables"

    Private prvDtableTabla As DataTable
    Private prvstrNombreDTable As String = ""
    Private prvintEstado As pubEnumAction
    Private prvblnAccesoPorNombre As Boolean
    Private _prvlngPosicionCliente As Long = -1
    Private _prvlngPosicionServidor As Long = -1

    Private _prvDtViewVistaClient As DataView = Nothing
    Private _prvDtViewVistaServer As DataView = Nothing
    Private prvintAction As pubEnumRstProp
    Private _prvBlnTablaFiltradaCliente As Boolean = False
    Private _prvBlnTablaFiltradaServidor As Boolean = False


    Private _prvBlnTablaOrdenCliente As Boolean = False
    Private _prvBlnTablaOrdenServidor As Boolean = False
    ' Por jairc 20210407
    ' Para manejo de Posiciones cuando se afecta una columna que afecta el filter
    Private _prvintIndexRowFilterCliente As Int32 = -1
    Private _prvintIndexRowFilterServer As Int32 = -1


    Private prvblnIsClone As Boolean
    Private prvobDNewRow As DataRow
    'Private prvobDTable As DataTable
    Private prvBlnServer As Boolean = False

    Friend prvObjRecOrigenClone As UnoEEDatos.clsRecordset
    ' MDI (Manuel Diaz) OP 203166: Booleana para el manejo de los eventos afterupdate y afterinsert de la grilla en un recordset clonado
    Public pubBlnAcceptChangesDesdeValue As Boolean = False

    Public pubBlnCreacionEnServidor As Boolean = False
    'jpa req.199641 20210622
    Public pubBlnDesdectlAdodc As Boolean = False

    'lrv req.208379
    Public pubBlnDesdectlAdodcPorNavegacion As Boolean = False

    Public Enum pubEnumRstProp
        pubEnumRstPropDescartar
        pubEnumRstPropUpdate
    End Enum

    Public Enum pubEnumAction
        pubEnumActionStandBy
        pubEnumActionNuevoRegistro
    End Enum

    Public Enum pubEnumCursorTypeEnum
        adOpenUnspecified = -1
        adOpenForwardOnly = 0
        adOpenKeyset
        adOpenDynamic
        adOpenStatic
    End Enum

    'se hace por efectos de migración...
    Public Enum pubEnumLockTypeEnum
        adLockUnspecified = -1
        adLockReadOnly = 1
        adLockPessimistic
        adLockOptimistic
        adLockBatchOptimistic
    End Enum

    Public Enum pubEnumCursorLocationEnum
        adUseServer = 0
        adUseClient
    End Enum

    Public Enum pubPositionEnum
        adPosEOF = -3
        adPosBOF
        adPosUnknown
    End Enum

    Public Enum SearchDirectionEnum
        adSearchBackward = -1
        adSearchForward = 1
    End Enum
    Public Enum pubEnumStart
        adBookmarkCurrent = 0
        adBookmarkFirst = 1
        adBookmarkLast = 2
    End Enum

    Public Enum pubEnumFieldAttributeEnum
        adFldUnspecified = -1
        adFldUpdatable = 4
        adFldFixed = 16
        adFldIsNullable = 32
        adFldMayBeNull = 64
        adFldKeyColumn = 32768

    End Enum

    Public Enum pubComparacionOp
        pubEnumNotOP = -1
        pubEnumDif = 0
        pubEnumMayI = 1
        pubEnumMenI = 2
        pubEnumMay = 3
        pubEnumMen = 4
        pubEnumIgual = 5
        pubEnumLike = 6
    End Enum
    'Enum usado para saber q tipo de like se hace
    'LikeInic = trozo* donde la frase original es: trozolargo
    'likeFin = *trozo donde la frase original es: completotrozo
    'likeMid = *trozo* donde la frase original sería esuntrozodefrase
    Public Enum pubTipoLike
        pubLikeInic = 0 '*Like
        pubLikeFin = 1
        pubLikeMid = 2
    End Enum

    '#End Region
    '***********************************Migracion****************************************
    Public Enum AffectEnum
        adAffectCurrent = 1
        adAffectGroup
        adAffectAll
        adAffectAllChapters
    End Enum
    ' Por jairc 20200506
    ' Para darle soporte al codigo en migracion
    Public Enum pubEnumObjectStateEnum
        adStateClosed = 0
        adStateOpen = 1
        adStateConnecting = 2
        adStateExecuting = 4
        adStateFetching = 8
    End Enum

    Public Enum EventReasonEnumRecordset
        adRsnMoveFirst = 12
        adRsnMoveNext = 13
        adRsnMovePrevious = 14
        adRsnMoveLast = 15
    End Enum
    Public Enum EventStatusEnumRecordset
        adStatusOK = 1
        adStatusErrorsOccurred = 2
        adStatusCantDeny = 3
        adStatusCancel = 4
        adStatusUnwantedEvent = 5
    End Enum

    Public Enum pubEnumEnMetodo
        pubEnumEnMetodoNinguno = 0
        pubEnumEnMetodoAddNew
        pubEnumEnMetodoFilter
    End Enum
    '************************************************************************************
    Public pubIntEnMetodo As pubEnumEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoNinguno
    Public Event RecordAdded()
    Public Event RecordChangeComplete(ByVal adReason As clsRecordset.EventReasonEnumRecordset, ByVal cRecords As Integer, ByVal pError As Data.Common.DbException, adStatus As clsRecordset.EventStatusEnumRecordset, ByVal pRecordset As UnoEEDatos.clsRecordset)
    Public Event MoverPosicion()
    Public Event FieldChangeComplete(ByVal cFields As Long, ByRef Fields As Object, ByVal pError As Data.Common.DbException, ByRef adStatus As clsRecordset.EventStatusEnumRecordset, ByVal pRecordset As clsRecordset)

    '#Region "Propiedades"

    Public Function GetPath() As String
        'Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Return System.Reflection.Assembly.GetExecutingAssembly().FullName
    End Function
    ' Por jairc 20210423
    ' Puede ser BindingSource o CurrencyManager
    Public pubobjBinding As Object
    Public pubintEventoActivoGrilla As Integer
    ' por jairc 20210601
    ' para tener acceso a la grilla desde el recordset
    Public pubobjGrilla As Object
    'MDI (Manuel Diaz) OP 208382: Para tener acceso a la flex desde el recordset, se utiliza para mover la posicion de la flex cuando
    'se mueva en el recordset
    Public pubobjFlexGrid As Object
    '***********************Migracion*****************************
    ''' <summary>
    ''' Mamr:Req 144828,144829,144827
    ''' Asigna y Retorna un Conjunto de elementos que pasan a usar como un RecortSet
    '''  </summary>
    ''' <returns></returns>
    Public Overloads Property DataSource() As DataTable
        Get
            DataSource = prvDtableTabla
        End Get
        Set(value As DataTable)
            prvDtableTabla = value
            prvstrNombreDTable = prvDtableTabla.TableName
            PrvDtViewVista = New DataView(prvDtableTabla)
            'prvDtViewVista.Table = prvDtableTabla
            MoveFirst()
        End Set
    End Property
    ''' <summary>
    '''maneja el indice cuando se afecta una columna que interviene en el filtro
    '''  </summary>
    ''' <returns>Posicion</returns>
    Public Property IndiceRowFilter(Optional pvblnServer As Boolean = False) As Int32
        Get
            'prvBlnServer = pvblnServer
            'If prvlngPosicion = pubPositionEnum.adPosUnknown Then
            If pvblnServer Then
                Return _prvintIndexRowFilterCliente
            Else
                Return _prvintIndexRowFilterServer
            End If
            'Else
            '    Return -1
            'End If
        End Get
        Set(value As Int32)
            prvBlnServer = pvblnServer
            If pvblnServer Then
                _prvintIndexRowFilterCliente = value
            Else
                _prvintIndexRowFilterServer = value
            End If
        End Set
    End Property
    '*************************************************************

    Public ReadOnly Property GetDataSet() As DataSet
        Get
            If PrvBlnTablaFiltrada Or PrvBlnTablaOrden Then
                Dim vdtsDataset As New DataSet
                vdtsDataset.Tables.Add(PrvDtViewVista.ToTable())
                GetDataSet = vdtsDataset
                vdtsDataset = Nothing
            Else
                GetDataSet = prvDtableTabla.DataSet
            End If
        End Get
    End Property

    Public Property pubObjDtTabla() As DataTable
        Get
            pubObjDtTabla = prvDtableTabla
        End Get
        Set(ByVal value As DataTable)


            'prvobValor = New UnoEEDatos.clsValor(value)
            prvDtableTabla = value
            prvstrNombreDTable = prvDtableTabla.TableName
            PrvDtViewVista = New DataView(prvDtableTabla)
            'prvDtViewVista.Table = prvDtableTabla
            MoveFirst()
        End Set
    End Property

    Public Property pubOjbDtView() As DataView
        Get
            Return PrvDtViewVista
        End Get
        Set(ByVal value As DataView)
            PrvDtViewVista = value
        End Set
    End Property
    'Public Property AbsolutePosition2() As Long
    '    Get
    '        If (prvlngPosicion >= RecordCount()) Then
    '            AbsolutePosition2 = pubPositionEnum.adPosEOF
    '        ElseIf prvlngPosicion <= -1 Then
    '            AbsolutePosition2 = pubPositionEnum.adPosBOF
    '        ElseIf RecordCount() <= 0 Then
    '            AbsolutePosition2 = pubPositionEnum.adPosUnknown
    '        Else
    '            AbsolutePosition2 = prvlngPosicion
    '        End If
    '    End Get
    '    Set(ByVal value As Long)
    '        prvlngPosicion = value
    '    End Set
    'End Property


    Public Property AbsolutePosition() As Long
        Get
            If prvlngPosicion >= RecordCount() Then
                Return pubPositionEnum.adPosEOF
            ElseIf prvlngPosicion < 0 Then
                Return pubPositionEnum.adPosBOF
            ElseIf RecordCount() <= 0 Then
                Return pubPositionEnum.adPosUnknown
            Else
                'AbsolutePosition2 = prvlngPosicion
                Return prvlngPosicion + 1
            End If
        End Get
        Set(ByVal value As Long)
            If value >= 0 Then
                prvlngPosicion = value - 1
            Else
                prvlngPosicion = -1
            End If
        End Set
    End Property
    Public Property AbsolutePositionServer() As Long
        Get
            prvBlnServer = True
            If prvlngPosicion >= RecordCount() Then
                AbsolutePositionServer = pubPositionEnum.adPosEOF
            ElseIf prvlngPosicion < 0 Then
                AbsolutePositionServer = pubPositionEnum.adPosBOF
            ElseIf RecordCount() <= 0 Then
                AbsolutePositionServer = pubPositionEnum.adPosUnknown
            Else
                'AbsolutePosition2 = prvlngPosicion
                AbsolutePositionServer = prvlngPosicion + 1
            End If
            prvBlnServer = False
        End Get
        Set(ByVal value As Long)
            prvBlnServer = True
            If value >= 0 Then
                prvlngPosicion = value - 1
            Else
                prvlngPosicion = -1
            End If
            prvBlnServer = False
        End Set
    End Property

    Public ReadOnly Property BOF(Optional pvblnServer As Boolean = False) As Boolean
        Get
            prvBlnServer = pvblnServer
            'lsdt: 21/05/2020 req: 162873
            'el bof debe de dar true si el recordset no tiene datos.
            ' por jairc 20210616
            ' Se cambia condicion de la posicion para que evalue si la posicion es -1 o -2, si es -3 ya esta ubicado al final
            ' OP 197356
            BOF = (prvlngPosicion = -1 Or prvlngPosicion = pubPositionEnum.adPosBOF) Or (RecordCount(pvblnServer:=pvblnServer) = 0)
        End Get
    End Property
    Public ReadOnly Property BOFServer() As Boolean
        Get

            BOFServer = BOF(pvblnServer:=True)
            prvBlnServer = False
        End Get
    End Property
    'Funcion bookmar por confirmar:

    Public Property Bookmark(Optional pvblnServer As Boolean = False) As Object
        Get
            Dim vObjBookmark As Object
            'retorna la posicion pero en la tabla original
            prvBlnServer = pvblnServer
            If prvDtableTabla.Rows.Count <> 0 Then
                'lsdt: 05/05/2021 req: 196029 para manejar el bookmark de acuerdo a vb6.
                'jpa,lsdt 07/05/2021 req. 196342 se agrega validacion pendiente
                vObjBookmark = prvDtableTabla.Rows.IndexOf(PrvDtViewVista(prvlngPosicion).Row)
                If IsNumeric(vObjBookmark) Then vObjBookmark = vObjBookmark + 1
                Return vObjBookmark
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Object)
            prvBlnServer = pvblnServer
            'lsdt: 05/05/2021 req: 196029 
            'jpa,lsdt 07/05/2021 req. 196342 se agrega validacion pendiente
            If IsNumeric(value) Then value = value - 1
            prvlngPosicion = buscarRowenView(value)
        End Set
    End Property
    Public Property BookmarkServer() As Object
        Get

            BookmarkServer = Bookmark(pvblnServer:=True)
            prvBlnServer = False
        End Get
        Set(ByVal value As Object)
            Bookmark(pvblnServer:=True) = value
            prvBlnServer = False
        End Set
    End Property

    Public ReadOnly Property EOF(Optional pvblnServer As Boolean = False) As Boolean
        Get
            prvBlnServer = pvblnServer
            EOF = (prvlngPosicion >= RecordCount(pvblnServer:=pvblnServer)) Or (RecordCount(pvblnServer:=pvblnServer) = 0 And prvlngPosicion <= -1) Or prvlngPosicion = pubPositionEnum.adPosEOF
        End Get
    End Property
    Public ReadOnly Property EOFServer() As Boolean
        Get
            EOFServer = EOF(pvblnServer:=True)
            prvBlnServer = False
        End Get
    End Property
    'El filtro teoricamente funciona con un variant no se si dejarlo como lo hizo jairc
    'Ademas si es string debe ser un string de procedimiento como: id<=i3023
    'Property Filter AS object
    '   Get
    '   Set
    'Public Property Filter() As String
    '    Get
    '        Return pvstrFilter
    '    End Get
    '    Set(ByVal value As String)
    '        pvstrFilter = value
    '        pubObjDtTabla.Select(value)
    '    End Set
    'End Property
    'Nueva version del Filtro:
    Public Property Filter(Optional pvblnServer As Boolean = False) As String
        Get
            prvBlnServer = pvblnServer
            Return PrvStrFilter
        End Get
        Set(ByVal value As String)
            Try
                pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoFilter
                prvBlnServer = pvblnServer
                Update()
                PrvStrFilter = value
                If PrvStrFilter = "" AndAlso PrvBlnTablaFiltrada Then
                    PrvBlnTablaFiltrada = False
                    PrvDtViewVista.RowFilter = ""
                    'MDI (Manuel Diaz) OP 208382: Para filtrar el recordset cuando este en una flexgrid
                    If Not IsNothing(pubobjFlexGrid) Then
                        prvDtableTabla.DefaultView.RowFilter = ""
                    End If 
                    MoveFirst()
                ElseIf PrvStrFilter <> "" Then
                    'lsdt: 03/05/2020 req: 168362
                    'Se agrega por que a grilla no estaba tomando los cambios cuando se guarda
                    'Cuando se realizaba el filtro no traia lo ultimo
                    'prvDtViewVista.Table.Has()
                    'prvDtViewVista.Table.GetChanges(DataRowState.)
                    'lsdt: 15/06/2021 req: 198840
                    PrvDtViewVista.Table.AcceptChanges()
                    PrvDtViewVista.RowFilter = PrvStrFilter
                    'MDI (Manuel Diaz) OP 208382: Para filtrar el recordset cuando este en una flexgrid
                    If Not IsNothing(pubobjFlexGrid) Then
                        prvDtableTabla.DefaultView.RowFilter = PrvStrFilter
                    End If
                    prvlngPosicion = 0
                    PrvBlnTablaFiltrada = True
                End If
                pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoNinguno
            Catch ex As Exception
                pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoNinguno
                Throw ex
            End Try
        End Set
    End Property
    Public Property FilterServer() As String
        Get
            FilterServer = Filter(pvblnServer:=True)
            prvBlnServer = False
        End Get
        Set(ByVal value As String)
            Filter(pvblnServer:=True) = value

            prvBlnServer = False

        End Set
    End Property

    Public Property Sort(Optional pvblnServer As Boolean = False) As String
        Get
            prvBlnServer = pvblnServer
            Return PrvStrSort
        End Get
        Set(ByVal value As String)

            Dim vblnRefresh As Boolean

            vblnRefresh = Not GblCadenasIguales(PrvStrSort, value)

            'lsdt: 17/03/2021 req: 191467  al realizar el ordenar es necesario que esten actualizados los cambion.
            prvBlnServer = pvblnServer
            If Not IsNothing(prvDtableTabla.GetChanges()) Then prvDtableTabla.AcceptChanges()
            PrvStrSort = value

            If PrvStrSort = "" AndAlso PrvBlnTablaOrden Then
                PrvBlnTablaOrden = False
                PrvDtViewVista.Sort = ""
            ElseIf PrvStrSort <> "" Then
                PrvBlnTablaOrden = True
                PrvDtViewVista.Sort = PrvStrSort
            End If

            MoveFirst()

            If Not prvBlnServer And vblnRefresh Then
                If Not pubobjGrilla Is Nothing Then
                    pubobjGrilla.rebind()
                End If
            End If
        End Set
    End Property
    Public Property SortServer() As String
        Get
            SortServer = Sort(pvblnServer:=True)
            prvBlnServer = False
        End Get
        Set(ByVal value As String)
            Sort(pvblnServer:=True) = value
            prvBlnServer = False
        End Set
    End Property

    Public ReadOnly Property GetField(ByVal pvIntRecordIndex As Integer, ByVal pvStrColName As String) As Object
        Get
            Return PrvDtViewVista(pvIntRecordIndex).Item(pvStrColName)
        End Get
    End Property

    Public ReadOnly Property getFields() As Object()
        Get
            Return PrvDtViewVista.Item(prvlngPosicion).Row.ItemArray
        End Get
    End Property

    Public WriteOnly Property SetField(ByVal pvStrNombreCol As String) As Object
        Set(ByVal value As Object)
            PrvDtViewVista(prvlngPosicion)(pvStrNombreCol) = value
        End Set
    End Property

    Public WriteOnly Property SetField(ByVal pvIntRecordIndex As Integer, ByVal pvStrNombreCol As String) As Object
        Set(ByVal value As Object)
            PrvDtViewVista(pvIntRecordIndex)(pvStrNombreCol) = value
        End Set
    End Property

    '*********************************VB6Migracion****************************************

    ''' <summary>
    ''' lsdt req: 141166 se crea propiedad por que no existe como propiedad de la clase recordset en .net y esta es utilizada en vb6
    ''' </summary>
    ''' <returns></returns>
    Public Property LockType() As pubEnumLockTypeEnum

    ''' <summary>
    ''' ### Importante ###
    ''' mamr: req 144829,144828,144827
    ''' En VB6 el UpdateBatch realiza una actualizacion al recorset y a la BD
    ''' Pero no obedece a la arquitectura de la empresa porque es usado en el cliente
    ''' En vb.Net al parecer al realizar cualquier cambio al recorset no hay necesidad de realizar un Update
    ''' Se mantiene en observacion el Update y UpdateBatch
    ''' </summary>
    ''' <param name="AffectRecords"></param>
    Public Sub UpdateBatch(ByRef Optional AffectRecords As AffectEnum = AffectEnum.adAffectAll)
        Update()
    End Sub
    '***************************************************************************************

    Public Sub Deserialize()
        PrvDtViewVista.Table = prvDtableTabla
        PrvDtViewVista.RowFilter = PrvStrFilter
        PrvDtViewVista.Sort = PrvStrSort
    End Sub
    Private Function buscarRowenView(ByVal pvObjBookmark As Object) As Long
        Dim vLongPos As Long = -1
        'lsdt: 21/05/2020 req:162873 no estab comparando correctamente las rows a pesar de que eran iguales.
        Dim comparer As IEqualityComparer(Of DataRow) = DataRowComparer.Default
        Dim vDtRowtmp As DataRow
        'lsdt: 21/04/2021 req: 194355 se mejora rendimiento de consulta del index
        If Not PrvStrFilter.Equals("") Or Not PrvStrSort.Equals("") Then
            vDtRowtmp = prvDtableTabla(pvObjBookmark)

            vLongPos = Enumerable.Range(0, PrvDtViewVista.Count).Where(
                       Function(i) comparer.Equals(PrvDtViewVista(i).Row, vDtRowtmp)).First
        Else
            vLongPos = pvObjBookmark
        End If
        Return vLongPos
    End Function

    Private Function buscarRowenView(ByRef pvObjDtRow As DataRow) As Long
        Dim vLongPos As Long = -1
        Dim vObjRow As DataRow = pvObjDtRow
        'lsdt: 21/04/2021 req: 194355 se mejora rendimiento de consulta del index
        'cdrb: 08/07/2021 req: 200979 DataRowComparer.Equals por DataRowComparer.Default.Equals para instancias
        vLongPos = Enumerable.Range(0, PrvDtViewVista.Count).Where(
                   Function(i) DataRowComparer.Default.Equals(PrvDtViewVista(i).Row, vObjRow)).First

        Return vLongPos
    End Function

    Public Function getDataFields() As DataRow
        If prvlngPosicion >= 0 Then
            Return PrvDtViewVista.Item(prvlngPosicion).Row
        Else
            Return Nothing
        End If
    End Function
    Public Function Find(ByVal pvStrCriterio As String, Optional ByVal SkipRecords As Integer = 0,
                         Optional ByVal SearchDirection As SearchDirectionEnum = SearchDirectionEnum.adSearchForward,
                         Optional ByVal Start As pubEnumStart = pubEnumStart.adBookmarkCurrent) As Integer
        'Se busca el timpo de comparación a realizar, buscando el operador, y las posiciones para extraer la columna y el operando de comparación
        Dim vIntPosresult As Integer = -1
        Dim tipo As pubComparacionOp
        Dim arrOperando(6) As String
        Dim intIndice As Integer
        Dim strOperando As String
        Dim blnContinuar As Boolean
        Dim arrResult() As String = Nothing
        Dim varrResultRegx() As String = Nothing
        Dim tipoLike As pubTipoLike



        arrOperando(0) = "<>"
        arrOperando(1) = ">="
        arrOperando(2) = "<="
        arrOperando(3) = ">"
        arrOperando(4) = "<"
        arrOperando(5) = "="
        arrOperando(6) = "like"




        strOperando = arrOperando(0)
        intIndice = 0
        blnContinuar = True
        While intIndice < arrOperando.Count And blnContinuar
            ' Por jairc, parte la cadena separando por ''

            varrResultRegx = Regex.Split(pvStrCriterio, "(\'.*\')")
            ' Si tiene caracteres con comilla sencilla, analiza de manera diferente
            If varrResultRegx.Count > 2 Then
                arrResult = Split(varrResultRegx(0), strOperando)
            Else
                arrResult = Split(pvStrCriterio, strOperando)
            End If

            intIndice = intIndice + 1
            If arrResult.Count > 1 Then
                If varrResultRegx.Count > 2 Then
                    ' Si tiene caracteres con comilla sencilla, analiza de manera diferente
                    arrResult(1) = varrResultRegx(1)
                End If
                tipo = intIndice - 1
                blnContinuar = False
            Else
                strOperando = arrOperando(intIndice)
            End If
        End While
        If (arrResult.Count <= 1) Then
            Err.Raise(1, "Unoeedatos.clsRecordSet", "Cadena de Find, mal realizada")
        End If

        'Se obtiene la columna y el objetivo de comparación
        Dim vStrColumna As String
        Dim vStrTarget As String
        vStrColumna = Trim(arrResult(0))
        vStrTarget = Trim(arrResult(1))

        If (vStrTarget.IndexOf("'") = 0 Or vStrTarget.IndexOf("#") = 0) Then
            vStrTarget = vStrTarget.Substring(1, vStrTarget.Length - 2)
        End If
        'Console.WriteLine("palabras " & vStrColumna & " " & vStrTarget)
        pubOjbDtView.Table.Select()
        'Zona especial si es busqueda like, saber queu tipo de busqueda segun la pos del ' * '
        If tipo = pubComparacionOp.pubEnumLike Then
            Dim vIntPosA As Integer = vStrTarget.IndexOf("*")
            Dim vIntPosF As Integer = vStrTarget.LastIndexOf("*")
            If Not (vIntPosA = 0) And vIntPosA = vIntPosF Then
                tipoLike = pubTipoLike.pubLikeInic
                vStrTarget = vStrTarget.Substring(0, (vStrTarget.Length - 1))
            ElseIf vIntPosA = 0 And vIntPosA = vIntPosF Then
                tipoLike = pubTipoLike.pubLikeFin
                vStrTarget = vStrTarget.Substring(1, (vStrTarget.Length - 1))
            Else
                tipoLike = pubTipoLike.pubLikeMid
                vStrTarget = vStrTarget.Substring(1, (vStrTarget.Length - 2))
            End If

        End If

        'Definimos de donde a donde debe ir el ciclo de búsqueda:
        Dim vIntInicio As Integer
        Dim vIntFinal As Integer
        Dim vIntIncremento As Integer
        Select Case Start
            Case pubEnumStart.adBookmarkCurrent
                vIntInicio = prvlngPosicion
            Case pubEnumStart.adBookmarkFirst
                vIntInicio = 0
            Case pubEnumStart.adBookmarkLast
                vIntInicio = PrvDtViewVista.Count - 1
        End Select
        vIntInicio = vIntInicio + SkipRecords

        If (SearchDirection = SearchDirectionEnum.adSearchForward) Then
            vIntFinal = PrvDtViewVista.Count - 1
            vIntIncremento = 1
        Else
            vIntFinal = 0
            vIntIncremento = -1
        End If

        'Seleccionamos el tipo de la columna para saber el tipo de comparacion (numerica, string o fecha)
        Dim vTypeDato As System.Type
        vTypeDato = prvDtableTabla.Columns(vStrColumna).DataType
        Dim vObjOperando As Object
        vObjOperando = vStrTarget
        Select Case vTypeDato.Name
            Case "Double"
                vObjOperando = CType(vObjOperando, Double)
            Case "Int32", "Int16", "Int64"
                vObjOperando = CType(vObjOperando, Integer)
            Case "String", "Char"
                vObjOperando = CType(vObjOperando, String)
            Case "Decimal"
                vObjOperando = CType(vObjOperando, Decimal)
            Case "DateTime"
                vObjOperando = CType(vObjOperando, DateTime)
            Case Else
                Err.Raise(1, "unoEEDatos.ClsRecordSet", "Tipo de dato no soportado")
        End Select

        'Se realiza el for general, dado que sin importar el tipo de dato se puede comparar igual:
        'lsdt: 18/11/2021 req: 209330 cuando no hay registros en la tabla no se debe de consultar ningun registro ademas en vb6 no generar error el find
        ' al realizar el mismo proceso sin registros.
        If PrvDtViewVista.Count > 0 Then
            Select Case vTypeDato.Name
                Case "DateTime"
                    For i = vIntInicio To vIntFinal Step vIntIncremento
                        'Dim vDtimeTmp As DateTimes
                        'vDtimeTmp = prvDtViewVista(i).Item(vStrColumna)
                        'vObjOperando.AddHours(vDtimeTmp.Hour
                        'Dim vTmpSComparacion As TimeSpan = vObjOperando.Subtract(prvDtViewVista(i).Item(vStrColumna))
                        Dim vTmpSComparacion As TimeSpan = PrvDtViewVista(i).Item(vStrColumna).Subtract(vObjOperando)
                        Select Case tipo
                            Case pubComparacionOp.pubEnumDif
                                If (vTmpSComparacion.Days <> 0 And vObjOperando.DayOfWeek <> PrvDtViewVista(i).Item(vStrColumna).DayOfWeek) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                End If
                            Case pubComparacionOp.pubEnumIgual
                                If (vTmpSComparacion.Days = 0 And vObjOperando.DayOfWeek = PrvDtViewVista(i).Item(vStrColumna).DayOfWeek) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                End If
                            Case pubComparacionOp.pubEnumMay
                                If ((vTmpSComparacion.Days > 0 Or vTmpSComparacion.TotalDays > 0) And vObjOperando.DayOfWeek <> PrvDtViewVista(i).Item(vStrColumna).DayOfWeek) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For

                                End If
                            Case pubComparacionOp.pubEnumMayI
                                If (vTmpSComparacion.Days > 0 And vTmpSComparacion.TotalDays >= 0) Or (vTmpSComparacion.Days = 0 And vObjOperando.DayOfWeek = PrvDtViewVista(i).Item(vStrColumna).DayOfWeek) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For

                                End If
                            Case pubComparacionOp.pubEnumMen
                                If ((vTmpSComparacion.Days < 0 Or vTmpSComparacion.TotalDays < 0) And vObjOperando.DayOfWeek <> PrvDtViewVista(i).Item(vStrColumna).DayOfWeek) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For

                                End If
                            Case pubComparacionOp.pubEnumMenI
                                If (vTmpSComparacion.Days < 0 And vTmpSComparacion.TotalDays <= 0) Or (vTmpSComparacion.Days = 0 And vObjOperando.DayOfWeek = PrvDtViewVista(i).Item(vStrColumna).DayOfWeek) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For

                                End If
                        End Select
                    Next
                Case "String", "Char"
                    For i = vIntInicio To vIntFinal Step vIntIncremento

                        Select Case tipo
                            Case pubComparacionOp.pubEnumDif
                                If Not GblCadenasIguales(PrvDtViewVista(i).Item(vStrColumna), vObjOperando) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                End If
                            Case pubComparacionOp.pubEnumIgual
                                If GblCadenasIguales(PrvDtViewVista(i).Item(vStrColumna), vObjOperando) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                End If

                            Case pubComparacionOp.pubEnumMay
                                If (Trim$(PrvDtViewVista(i).Item(vStrColumna)) > vObjOperando) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For

                                End If
                            Case pubComparacionOp.pubEnumMayI
                                If (Trim$(PrvDtViewVista(i).Item(vStrColumna)) >= vObjOperando) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For

                                End If
                            Case pubComparacionOp.pubEnumMen
                                If (Trim$(PrvDtViewVista(i).Item(vStrColumna)) < vObjOperando) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For

                                End If
                            Case pubComparacionOp.pubEnumMenI
                                If (Trim$(PrvDtViewVista(i).Item(vStrColumna)) <= vObjOperando) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                End If
                            Case pubComparacionOp.pubEnumLike
                                Dim posSubStr As Integer = PrvDtViewVista(i).Item(vStrColumna).IndexOf(vObjOperando, StringComparison.OrdinalIgnoreCase) 'Req.192924-amg: se agrega el stringComparison porque en la búsqueda solo estaba tomando en cuenta las mayúsculas.
                                If posSubStr = 0 And tipoLike = pubTipoLike.pubLikeInic Then
                                    'ok
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                ElseIf posSubStr > 0 And vObjOperando.Length = PrvDtViewVista(i).Item(vStrColumna).Length - posSubStr And tipoLike = pubTipoLike.pubLikeFin Then
                                    'ok
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                ElseIf tipoLike = pubTipoLike.pubLikeMid And posSubStr > 0 And vObjOperando.Length <> PrvDtViewVista(i).Item(vStrColumna).Length - posSubStr Then
                                    'ok
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                End If

                        End Select
                    Next
                Case Else
                    For i = vIntInicio To vIntFinal Step vIntIncremento

                        Select Case tipo
                            Case pubComparacionOp.pubEnumDif
                                If (PrvDtViewVista(i).Item(vStrColumna) <> vObjOperando) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                End If
                            Case pubComparacionOp.pubEnumIgual
                                If (PrvDtViewVista(i).Item(vStrColumna) = vObjOperando) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                End If
                            Case pubComparacionOp.pubEnumMay
                                If (PrvDtViewVista(i).Item(vStrColumna) > vObjOperando) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For

                                End If
                            Case pubComparacionOp.pubEnumMayI
                                If (PrvDtViewVista(i).Item(vStrColumna) >= vObjOperando) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For

                                End If
                            Case pubComparacionOp.pubEnumMen
                                If (PrvDtViewVista(i).Item(vStrColumna) < vObjOperando) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For

                                End If
                            Case pubComparacionOp.pubEnumMenI
                                If (PrvDtViewVista(i).Item(vStrColumna) <= vObjOperando) Then
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                End If
                            Case pubComparacionOp.pubEnumLike
                                Dim posSubStr As Integer = PrvDtViewVista(i).Item(vStrColumna).IndexOf(vObjOperando)
                                If posSubStr = 0 And tipoLike = pubTipoLike.pubLikeInic Then
                                    'ok
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                ElseIf posSubStr > 0 And vObjOperando.Length = PrvDtViewVista(i).Item(vStrColumna).Length - posSubStr And tipoLike = pubTipoLike.pubLikeFin Then
                                    'ok
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                ElseIf tipoLike = pubTipoLike.pubLikeMid And posSubStr > 0 And vObjOperando.Length <> PrvDtViewVista(i).Item(vStrColumna).Length - posSubStr Then
                                    'ok
                                    vIntPosresult = i
                                    prvlngPosicion = i
                                    Exit For
                                End If

                        End Select
                    Next
            End Select
        End If

        If (SearchDirection = SearchDirectionEnum.adSearchBackward And vIntPosresult = -1) Then
            vIntPosresult = pubPositionEnum.adPosBOF
        ElseIf (SearchDirection = SearchDirectionEnum.adSearchForward And vIntPosresult = -1) Then
            vIntPosresult = pubPositionEnum.adPosEOF
        Else
            vIntPosresult = vIntPosresult
        End If

        prvlngPosicion = vIntPosresult
        'MDI (Manuel Diaz) OP 208382: Para mover la posicion de la Row de la flexgrid cuando se haga un find en un recordset que sea datasource de la flex
        If Not IsNothing(pubobjFlexGrid) Then
            pubobjFlexGrid.row = prvlngPosicion + 1
        End If
        Return vIntPosresult
        'NO SOPORTADOS EN FIN:
        'LeerDataTypeCampo = System.Type.GetType("System.Byte")
        'LeerDataTypeCampo = System.Type.GetType("System.Boolean")


    End Function
    Public Function FindServer(ByVal pvStrCriterio As String, Optional ByVal SkipRecords As Integer = 0,
                         Optional ByVal SearchDirection As SearchDirectionEnum = SearchDirectionEnum.adSearchForward,
                         Optional ByVal Start As pubEnumStart = pubEnumStart.adBookmarkCurrent) As Integer
        prvBlnServer = True

        FindServer = Find(pvStrCriterio, SkipRecords, SearchDirection, Start)
        prvBlnServer = False

    End Function

    Public Function RecordCount(Optional pvblnServer As Boolean = False) As Long
        Try
            'prvBlnServer = pvblnServer
            If PrvDtViewVista.Count = 0 And Not PrvBlnTablaFiltrada Then
                Return prvDtableTabla.Rows.Count
            Else
                Return PrvDtViewVista.Count
            End If

        Catch ex As Exception
            RecordCount = 0
        End Try
    End Function
    Public Function RecordCountServer() As Long
        Try
            prvBlnServer = True
            RecordCountServer = RecordCount(pvblnServer:=True)
            prvBlnServer = False

        Catch ex As Exception
            RecordCountServer = 0
            prvBlnServer = False
            'Throw ex
        End Try
    End Function

    '#End Region


    '#Region "Funciones tipicas RecordSet"
    'Funcion por revisión, se debe revisar si el arreglo de valores siempre se pasa en orden
    'Si se peude pasar en desorden, entonces se debe implementar un ciclo para insertar en cada columna con nombre
    'pvObjDatos(i) el valor pvObjValores(i)
    Public Function AddNew(ByVal pvObjDatos() As Object, ByVal pvObjValores() As Object) As Boolean
        Try
            pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoAddNew
            If pvObjDatos.Count = pvObjValores.Count Then
                Dim vobDRow As DataRow
                vobDRow = prvDtableTabla.NewRow()
                For vshtIndiceDatos As Short = 0 To pvObjDatos.Count - 1
                    ' Por jairc 20210621
                    ' Se hace esta validacion porque cuando solo viene nothing y es fecha se cae la asignacion
                    ' Es necesario convertir a CDate(Nothing)
                    If vobDRow.Table.Columns(vshtIndiceDatos).DataType.Name = "DateTime" Then
                        If pvObjValores(vshtIndiceDatos) Is Nothing Then
                            pvObjValores(vshtIndiceDatos) = CDate(Nothing)
                        End If
                    End If
                    vobDRow(pvObjDatos(vshtIndiceDatos)) = pvObjValores(vshtIndiceDatos)
                Next
                prvDtableTabla.Rows.Add(vobDRow)
                prvDtableTabla.AcceptChanges()
                ' Por jairc 20210316
                ' Al clone hay que adicionar una nueva fila, o sino sacar error que la fila ya existe en otro datatable
                If prvblnIsClone Then
                    vobDRow = prvObjRecOrigenClone.prvDtableTabla.NewRow()
                    For vshtIndiceDatos As Short = 0 To pvObjDatos.Count - 1
                        vobDRow(pvObjDatos(vshtIndiceDatos)) = pvObjValores(vshtIndiceDatos)
                    Next
                    vobDRow.ItemArray = pvObjValores
                    prvObjRecOrigenClone.prvDtableTabla.Rows.Add(vobDRow)
                    prvObjRecOrigenClone.prvDtableTabla.AcceptChanges()
                End If
                ' Por jairc, 20210303
                ' No es necesario este estado ya que por la implementacion lo hace efectivo inmediatamente 
                ' en el datatable, y se vuelve una modificacion al registro, ya que queda ubicado en el ultimo
                'prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro
                If PrvBlnTablaFiltrada Then
                    prvlngPosicion = prvDtableTabla.Rows.Count
                Else
                    MoveLast()
                    MoveLastServer()
                End If
                RaiseEvent RecordAdded()
                'Dim vlngPosTempo = prvlngPosicion
                'prvlngPosicion = buscarRowenView(prvobDRow)
                'If prvlngPosicion = -1 Then prvlngPosicion = vlngPosTempo
                'prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro
                'prvobValor.pubDRDataRow = prvobDRow
                pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoNinguno
                Return True
            Else
                Err.Raise(0, , "Numero de valores y columnas no coinciden", , )
            End If
        Catch ex As Exception
            pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoNinguno

            Throw ex
        End Try
    End Function
    Public Function AddNew(ByVal pvObjValores() As Object) As Boolean
        Try
            Dim vobDRow As DataRow
            pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoAddNew
            vobDRow = prvDtableTabla.NewRow()
            vobDRow.ItemArray = pvObjValores
            prvDtableTabla.Rows.Add(vobDRow)
            prvDtableTabla.AcceptChanges()
            ' Por jairc 20210316
            ' Al clone hay que adicionar una nueva fila, o sino sacar error que la fila ya existe en otro datatable
            If prvblnIsClone Then
                vobDRow = prvObjRecOrigenClone.prvDtableTabla.NewRow()
                vobDRow.ItemArray = pvObjValores
                prvObjRecOrigenClone.prvDtableTabla.Rows.Add(vobDRow)
                prvObjRecOrigenClone.prvDtableTabla.AcceptChanges()
            End If
            ' Por jairc, 20210303
            ' No es necesario este estado ya que por la implementacion lo hace efectivo inmediatamente 
            ' en el datatable, y se vuelve una modificacion al registro, ya que queda ubicado en el ultimo
            'prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro
            If PrvBlnTablaFiltrada Then
                prvlngPosicion = prvDtableTabla.Rows.Count
            Else
                MoveLast()
                MoveLastServer()
            End If
            RaiseEvent RecordAdded()
            pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoNinguno

            Return True
        Catch ex As Exception
            pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoNinguno
            Throw ex
        End Try
    End Function

    Public Function InsertFirst(ByVal pvObjValores() As Object) As Boolean
        Try
            Dim prvobDRow As DataRow
            prvobDRow = prvDtableTabla.NewRow()
            prvobDRow.ItemArray = pvObjValores
            prvDtableTabla.Rows.InsertAt(prvobDRow, 0)
            prvDtableTabla.AcceptChanges()
            MoveLast()
            MoveLastServer()
            RaiseEvent RecordAdded()
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function AddNew(ByVal pvObjDato As String, ByVal pvObjValor As Object) As Boolean
        Try
            Dim vobDRow As DataRow
            pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoAddNew
            vobDRow = prvDtableTabla.NewRow()
            vobDRow(pvObjDato) = pvObjValor
            prvDtableTabla.Rows.Add(vobDRow)
            prvDtableTabla.AcceptChanges()
            ' Por jairc 20210316
            ' Al clone hay que adicionar una nueva fila, o sino sacar error que la fila ya existe en otro datatable
            If prvblnIsClone Then
                vobDRow = prvObjRecOrigenClone.prvDtableTabla.NewRow()
                vobDRow(pvObjDato) = pvObjValor
                prvObjRecOrigenClone.prvDtableTabla.Rows.Add(vobDRow)
                prvObjRecOrigenClone.prvDtableTabla.AcceptChanges()
            End If
            ' Por jairc, 20210303
            ' No es necesario este estado ya que por la implementacion lo hace efectivo inmediatamente 
            ' en el datatable, y se vuelve una modificacion al registro, ya que queda ubicado en el ultimo
            'prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro
            If PrvBlnTablaFiltrada Then
                prvlngPosicion = prvDtableTabla.Rows.Count
            Else
                MoveLast()
                MoveLastServer()
            End If
            RaiseEvent RecordAdded()
            'Dim vlngPosTempo = prvlngPosicion
            'prvlngPosicion = buscarRowenView(prvobDRow)
            'If prvlngPosicion = -1 Then prvlngPosicion = vlngPosTempo
            'prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro
            'prvobValor.pubDRDataRow = prvobDRow
            pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoNinguno
            Return True
        Catch ex As Exception
            pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoNinguno
            Throw ex
        End Try
    End Function
    Public Function AddNew() As Boolean
        Try
            'Dim prvobDRow As DataRow
            pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoAddNew

            prvobDNewRow = prvDtableTabla.NewRow()
            prvDtableTabla.Rows.Add(prvobDNewRow)
            prvDtableTabla.AcceptChanges()
            ' Por jairc 20210316
            ' Al clone hay que adicionar una nueva fila, o sino sacar error que la fila ya existe en otro datatable
            If prvblnIsClone Then
                prvobDNewRow = prvObjRecOrigenClone.prvDtableTabla.NewRow()
                prvObjRecOrigenClone.prvDtableTabla.Rows.Add(prvobDNewRow)
                prvObjRecOrigenClone.prvDtableTabla.AcceptChanges()
            End If
            ' Por jairc, 20210303
            ' No es necesario este estado ya que por la implementacion lo hace efectivo inmediatamente 
            ' en el datatable, y se vuelve una modificacion al registro, ya que queda ubicado en el ultimo
            ' por jairc 20210628
            ' Se elimina el comentario porque en el caso en que este filtrado y le dicen addnew el .fields.value
            ' va a afectar la fila nueva del objeto prvobDNewRow
            prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro
            If PrvBlnTablaFiltrada Then
                prvlngPosicion = prvDtableTabla.Rows.Count
            Else
                AbsolutePosition = buscarRowenView(prvobDNewRow) + 1
                AbsolutePositionServer = AbsolutePosition
                'MoveLast()
                'MoveLastServer()
            End If

            RaiseEvent RecordAdded()
            'Dim vlngPosTempo = prvlngPosicion
            'prvlngPosicion = buscarRowenView(prvobDRow)
            'If prvlngPosicion = -1 Then prvlngPosicion = vlngPosTempo
            'prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro
            'prvobValor.pubDRDataRow = prvobDRow
            pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoNinguno
            Return True
        Catch ex As Exception
            pubIntEnMetodo = pubEnumEnMetodo.pubEnumEnMetodoNinguno
            Throw ex
        End Try
    End Function
    'Metodo bajo prueba para una funcion especial del tiquete
    Public Sub Append(ByVal pvRstComplemento As UnoEEDatos.clsRecordset)
        pubObjDtTabla.Merge(pvRstComplemento.pubObjDtTabla)
    End Sub
    Public Function Clone(Optional pvintTypeLock As clsRecordset.pubEnumLockTypeEnum = clsRecordset.pubEnumLockTypeEnum.adLockOptimistic) As UnoEEDatos.clsRecordset
        Dim vRcSetCloned As UnoEEDatos.clsRecordset
        vRcSetCloned = New UnoEEDatos.clsRecordset
        If Not Me.pubObjDtTabla Is Nothing Then

            'lsdt: 03/03/2021 req: 192193
            'El defaultView trae los cambio realizado en la tabla y no requiere hacer AcceptChanges() para reflejarlos en el copy.
            vRcSetCloned.pubObjDtTabla = Me.pubObjDtTabla.DefaultView.ToTable().Copy
            ' Por jairc 20210609
            ' Se asigna la grilla para evitar que si se modifica un valor en el recordet, dispare el evento RowColChange_vb6
            ' En teoria no es necesario estos eventos cuando se trata de sincronizacion
            vRcSetCloned.pubobjGrilla = Me.pubobjGrilla
            'lsdt: 27/05/2020 req: 164088
            'Se realiza cambio ya que al hacer comparaciones de row no estaba realiando bein la comparacion 
            'ya que apesar de que eran las mismas rows en realidad no pertenecian a la misma tabla.
            'Este caso sucedio en el Bookmark al buscar la row. en el get
            vRcSetCloned.pubOjbDtView = New DataView(vRcSetCloned.pubObjDtTabla)
            vRcSetCloned.MoveFirst(pvblnSincronizarGrilla:=False)
            vRcSetCloned.MoveFirstServer()
        End If
        vRcSetCloned.prvblnIsClone = True
        vRcSetCloned.prvObjRecOrigenClone = Me
        Return vRcSetCloned
    End Function
    Public Function CloneServer(Optional pvintTypeLock As clsRecordset.pubEnumLockTypeEnum = clsRecordset.pubEnumLockTypeEnum.adLockOptimistic) As UnoEEDatos.clsRecordset
        Dim vrecclone As UnoEEDatos.clsRecordset = Nothing
        Try
            prvBlnServer = True
            vrecclone = Clone(pvintTypeLock)
            vrecclone.MoveFirstServer()
            CloneServer = vrecclone
        Catch ex As Exception
        Finally
            CloneServer = vrecclone
            vrecclone = Nothing
            prvBlnServer = False
        End Try


    End Function
    ''' <summary>
    ''' Esta funcion crea una copia del recordset, es parecida al clone pero el clone es la simulacion de vb6
    ''' </summary>
    ''' <returns></returns>
    Public Function Copy() As UnoEEDatos.clsRecordset
        Dim vRcSetCloned As UnoEEDatos.clsRecordset
        vRcSetCloned = New UnoEEDatos.clsRecordset
        vRcSetCloned.pubObjDtTabla = Me.pubObjDtTabla.Copy
        'lsdt: 27/05/2020 req: 164088
        'Se realiza cambio ya que al hacer comparaciones de row no estaba realiando bein la comparacion 
        'ya que apesar de que eran las mismas rows en realidad no pertenecian a la misma tabla.
        'Este caso sucedio en el Bookmark al buscar la row. en el get
        vRcSetCloned.pubOjbDtView = New DataView(vRcSetCloned.pubObjDtTabla)
        vRcSetCloned.MoveFirst()
        vRcSetCloned.MoveFirstServer()
        Return vRcSetCloned
    End Function
    'este aun  no

    Public Sub Close()

        Dim vLngError As Long
        Dim vStrError As String
        Dim vStrSource As String
        'jac: Se guarda el error en variabes, porque el On Error Resume Next limpia Objeto err, entonces si trae un error se pierde.
        'El caso se daba cuando generaba un error y antes de llamar el GenerarErrorParaCliente, se hacia un .close de un recordset.
        vLngError = Err.Number
        vStrError = Err.Description
        vStrSource = Err.Source

        On Error Resume Next

        PrvDtViewVista.Dispose()
        PrvDtViewVista = Nothing

        'Dim vlstViews As Object
        'vlstViews = prvDtableTabla.GetType().GetProperty("_dataViewListeners", BindingFlags.Instance Or BindingFlags.NonPublic)
        ' Por jairc 20210713
        ' Si las vistas cliente y servidor quedan nothig, se cierra el datatable
        If _prvDtViewVistaClient Is Nothing And _prvDtViewVistaServer Is Nothing Then
            prvDtableTabla.Clear()
            prvDtableTabla.Dispose()
            prvDtableTabla = Nothing
        End If
        prvstrPath = ""
        PrvStrFilter = ""
        PrvStrSort = ""
        prvintCursorLocation = Nothing
        prvintCursorType = Nothing
        prvintLockType = Nothing

        '#Region "Local Variables"
        prvstrNombreDTable = ""
        prvintEstado = pubEnumAction.pubEnumActionStandBy
        'prvblnAccesoPorNombre = Nothing
        prvlngPosicion = -1
        prvintAction = Nothing
        'prvobDRow = Nothing

        If vLngError <> 0 Then
            Err.Number = vLngError
            Err.Description = vStrError
            Err.Source = vStrSource
        End If
        ' Por jairc, 20210218
        ' Si es clone, cierra el recordset
        If prvblnIsClone Then
            Me.prvObjRecOrigenClone = Nothing

        End If
    End Sub

    Public Function Collect(ByVal pvstrColumnName As String, Optional pvBlnServer As Boolean = False) As Object
        Try
            prvBlnServer = pvBlnServer
            'If prvBlnTablaFiltrada Or prvBlnTablaOrden Then
            Return PrvDtViewVista(prvlngPosicion).Item(pvstrColumnName)
            'Else
            '    Return prvDtableTabla.Rows(prvlngPosicion).Item(pvstrColumnName)
            'End If

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function Collect(ByVal pvintColumnIndex As Integer, Optional pvBlnServer As Boolean = False) As Object
        Try
            prvBlnServer = pvBlnServer
            ' If prvBlnTablaFiltrada Or prvBlnTablaOrden Then
            Return PrvDtViewVista(prvlngPosicion).Item(pvintColumnIndex)
            'Else
            '    Return prvDtableTabla.Rows(prvlngPosicion).Item(pvintColumnIndex)
            'End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function CollectServer(ByVal pvstrColumnName As String) As Object
        Try
            ' If prvBlnTablaFiltrada Or prvBlnTablaOrden Then
            CollectServer = Collect(pvstrColumnName, pvBlnServer:=True)
            prvBlnServer = False
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function CollectServer(ByVal pvintColumnIndex As Integer) As Object
        Try

            ' If prvBlnTablaFiltrada Or prvBlnTablaOrden Then
            CollectServer = Collect(pvintColumnIndex, pvBlnServer:=True)
            prvBlnServer = False
            'Else
            '    Return prvDtableTabla.Rows(prvlngPosicion).Item(pvintColumnIndex)
            'End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    'Existe una variación de este método, q se basa en borrar los records q cumplen con el filtro
    'Revisar si se necesita esta implementacion al parecer luego de un delete no se necesita un update, por eso se usarpa removeAt.
    'Lo dijo OME, BAnano y Chichilin
    Public Function Delete(Optional pvblnServer As Boolean = False) As Boolean

        Dim vblnEjecutarDelete As Boolean = True

        Try
            prvBlnServer = pvblnServer
            'vlngPosicion = prvlngPosicion
            'if prvBlnTablaFiltrada Or prvBlnTablaOrden Then
            If Not Me.pubobjGrilla Is Nothing Then
                ' Si es llamado desde BeforeDelete no ejecuta el remove de la fila porque lo ejecuta la grilla
                ' No se puede colocar el enum porqur no esta referenciado unoeegeneral
                vblnEjecutarDelete = Not (Me.pubobjGrilla.pubIntEventoActivo = 4)
            End If
            'MDI(Manuel Diaz) OP 206445, 206448: En caso que el recorset sea de solo lectura no haga el delete
            If Me.LockType = pubEnumLockTypeEnum.adLockReadOnly
                vblnEjecutarDelete = False
            End If

            If vblnEjecutarDelete Then

                prvDtableTabla.Rows.Remove(PrvDtViewVista(prvlngPosicion).Row)

                prvlngPosicion -= 1
                'If prvlngPosicion < 0 Then
                '    prvlngPosicion = 
                'End If
                prvDtableTabla.AcceptChanges()
            End If
            'Else
            '    prvDtableTabla.Rows.RemoveAt(prvlngPosicion)
            'End If
            'NO SE DEBE PONER EN STAND BY dado que no se hace UPDATE
            'prvintEstado = pubEnumAction.pubEnumActionStandBy
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function DeleteServer() As Boolean

        Try
            Delete(pvblnServer:=True)
            prvBlnServer = False
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Sub Clear()
        PrvStrFilter = String.Empty
        PrvStrSort = String.Empty
        prvDtableTabla.Clear()
        prvlngPosicion = 0
        PrvBlnTablaFiltrada = False
        PrvBlnTablaOrden = False
    End Sub

    '################# FUNCIONES DE DESPLAZAMIENTO ##########################
    Public Function MoveNext(Optional pvblnServer As Boolean = False) As Boolean
        Try
            Dim vintPosicionActual As Integer
            Dim vblnDesdeRecordsetAnterior As Boolean

            prvBlnServer = pvblnServer
            Update()

            'lsdt: 06/08/2021 req: 202507  actualiza el estado del EOF de la grilla si no esta clonado el recordset
            If SincronizarGrilla() And Not prvblnIsClone And Not pvblnServer Then
                vintPosicionActual = prvlngPosicion
                ' Por jairc 20210908, para evitar que el movimiento genere eventos
                vblnDesdeRecordsetAnterior = pubobjGrilla.pubblnDesdeRecordset
                pubobjGrilla.pubblnDesdeRecordset = True
                pubobjGrilla.MoveNext()
                pubobjGrilla.pubblnDesdeRecordset = vblnDesdeRecordsetAnterior
                If vintPosicionActual = prvlngPosicion Then
                    prvlngPosicion += 1
                End If

            Else
                prvlngPosicion += 1
            End If
            If prvlngPosicion < RecordCount(pvblnServer:=pvblnServer) Then
                Return True
            Else
                prvlngPosicion = RecordCount(pvblnServer:=pvblnServer)
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function MoveNextServer() As Boolean
        Try
            MoveNextServer = MoveNext(pvblnServer:=True)
            prvBlnServer = False
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function MovePrevious(Optional pvblnServer As Boolean = False) As Boolean
        Try
            Dim vintPosicionActual As Integer
            Dim vblnDesdeRecordsetAnterior As Boolean
            prvBlnServer = pvblnServer
            Update()

            'MDI (Manuel Diaz): OP 201928 - Cuando se mueve el bookmark del recordset debe mover el de la grilla
            If SincronizarGrilla() And Not prvblnIsClone And Not pvblnServer Then

                vintPosicionActual = prvlngPosicion
                ' Por jairc 20210908, para evitar que el movimiento genere eventos
                vblnDesdeRecordsetAnterior = pubobjGrilla.pubblnDesdeRecordset
                pubobjGrilla.pubblnDesdeRecordset = True
                pubobjGrilla.MovePrevious()
                pubobjGrilla.pubblnDesdeRecordset = vblnDesdeRecordsetAnterior
                If vintPosicionActual = prvlngPosicion Then
                    prvlngPosicion -= 1
                End If
            Else
                prvlngPosicion -= 1
            End If
            If prvlngPosicion >= 0 Then
                Return True

            Else
                prvlngPosicion = -1
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function MovePreviousServer() As Boolean
        Try

            MovePreviousServer = MovePrevious(pvblnServer:=True)
            prvBlnServer = False
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function MoveLast(Optional pvblnServer As Boolean = False) As Boolean
        Try
            Dim vintPosicionActual As Integer
            Dim vblnDesdeRecordsetAnterior As Boolean

            prvBlnServer = pvblnServer

            'MDI (Manuel Diaz): OP 201928 - Cuando se mueve el bookmark del recordset debe mover el de la grilla
            If SincronizarGrilla() And Not prvblnIsClone And Not pvblnServer Then
                vintPosicionActual = prvlngPosicion
                ' Por jairc 20210908, para evitar que el movimiento genere eventos
                vblnDesdeRecordsetAnterior = pubobjGrilla.pubblnDesdeRecordset
                pubobjGrilla.pubblnDesdeRecordset = True
                pubobjGrilla.MoveLast()
                pubobjGrilla.pubblnDesdeRecordset = vblnDesdeRecordsetAnterior
                If vintPosicionActual = prvlngPosicion Then
                    prvlngPosicion = RecordCount(pvblnServer:=pvblnServer) - 1
                End If
            Else
                prvlngPosicion = RecordCount(pvblnServer:=pvblnServer) - 1
            End If
            Update()
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function MoveLastServer() As Boolean
        Try

            MoveLastServer = MoveLast(pvblnServer:=True)
            prvBlnServer = False
        Catch ex As Exception
            prvBlnServer = False
            Throw ex
        End Try
    End Function
    Public Function MoveFirst(Optional pvblnServer As Boolean = False, Optional pvblnSincronizarGrilla As Boolean = True) As Boolean
        Try
            Dim vintPosicionActual As Integer
            Dim vblnDesdeRecordsetAnterior As Boolean

            prvBlnServer = pvblnServer

            'MDI (Manuel Diaz): OP 201928 - Cuando se mueve el bookmark del recordset debe mover el de la grilla
            If SincronizarGrilla() And Not prvblnIsClone And Not pvblnServer And pvblnSincronizarGrilla Then

                vintPosicionActual = prvlngPosicion
                ' Por jairc 20210908, para evitar que el movimiento genere eventos
                vblnDesdeRecordsetAnterior = pubobjGrilla.pubblnDesdeRecordset
                pubobjGrilla.pubblnDesdeRecordset = True
                pubobjGrilla.MoveFirst()
                pubobjGrilla.pubblnDesdeRecordset = vblnDesdeRecordsetAnterior
                If vintPosicionActual = prvlngPosicion Then
                    prvlngPosicion = 0
                End If
            Else
                prvlngPosicion = 0
            End If
            'MDI (Manuel Diaz) OP 208382: Para mover la posicion del Row de una flexgrid cuando se hace un movefirst en un recordset que sea datasource 
            'de una flexgrid
            If Not IsNothing(pubobjFlexGrid) Then
                pubobjFlexGrid.row = prvlngPosicion + 1
            End If
            Update()
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function MoveFirstServer() As Boolean
        Try

            MoveFirstServer = MoveFirst(pvblnServer:=True)
            prvBlnServer = False
        Catch ex As Exception
            prvBlnServer = False
            Throw ex
        End Try
    End Function
    'VB6Migracion
    ''' <summary>
    ''' cpp req: 120234 se crea propiedad por que no existe como propiedad de la clase recordset en .net y esta es utilizada en vb6. Validar!
    ''' </summary>
    ''' <returns></returns>

    Public Function Move(NumRecords As Long, Optional pvblnServer As Boolean = False) As Boolean
        Try
            prvBlnServer = pvblnServer
            'prvlngPosicion = RecordCount() + NumRecords
            prvlngPosicion = prvlngPosicion + NumRecords
            Update()
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function MoveServer(NumRecords As Long) As Boolean
        Try

            MoveServer = Move(NumRecords, pvblnServer:=True)
            prvBlnServer = False
        Catch ex As Exception
            prvBlnServer = False
            Throw ex
        End Try
    End Function
    '################### FUNCIONES DE APERTURA #################
    Public Function Open(ByVal pvstrNombreDataTable As String) As Boolean
        Try
            prvstrNombreDTable = pvstrNombreDataTable
            ' LLama funcion de OPen sin nombre
            ' Se deja asi por contabilidad
            Open = Open()
            Return True
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function Open() As Boolean

        If prvstrNombreDTable.Trim.Length <> 0 Then
            Try
                LimpiarPosicion()
                Return True
            Catch ex As Exception
                Throw ex
            End Try
        Else
            Throw New Exception("No table name given")
        End If
    End Function

    ''' <summary>
    ''' Este Open es para efectos de migración, los parametros de pvSource y pvActiveConnection, solo se toman pero no hace nada con esto y 
    ''' los parametros de cursortype y locktype se asignan a las respectivas propiedades.
    ''' cpp;16/06/2019
    ''' </summary>
    ''' <param name="pvSource"></param>
    ''' <param name="pvActiveConnection"></param>
    ''' <param name="pvCursorType"></param>
    ''' <param name="pvLockType"></param>
    ''' <returns></returns>
    Public Function Open(Optional ByVal pvSource As String = "", Optional ByVal pvActiveConnection As String = "", Optional ByVal pvCursorType As pubEnumCursorTypeEnum = pubEnumCursorTypeEnum.adOpenUnspecified, Optional ByVal pvLockType As pubEnumLockTypeEnum = pubEnumLockTypeEnum.adLockUnspecified) As Boolean
        prvintCursorLocation = pvCursorType
        prvintLockType = pvLockType

        If prvstrNombreDTable.Trim.Length <> 0 Then
            Try
                LimpiarPosicion()
                Return True
            Catch ex As Exception
                Throw ex
            End Try
        Else
            Throw New Exception("No table name given")
        End If
    End Function


    Public Function Update() As Boolean
        Try
            Select Case prvintEstado
                Case pubEnumAction.pubEnumActionNuevoRegistro
                    'prvDtableTabla.Rows.Add(prvobDRow)
                    prvintEstado = pubEnumAction.pubEnumActionStandBy
                    Return True
                Case Else
                    'prvobDTable.Rows(prvlngPosicion).Item(pvintColumnIndex) = Value
                    prvintEstado = pubEnumAction.pubEnumActionStandBy
                    Return True

            End Select
        Catch ex As Exception
            Throw ex
        End Try

    End Function
    'jac 31/12/2020
    ' NOTA: Se modifica el Parametro para que reciba un object y no un string en arreglo, para los movimientos contables genera error de conversión de tipo de dato
    ' Arreglo de object a arreglo de string.
    ' Se quita el movelast() ya que al modificar un registro diferente al ultimo, al momento de hacer un proceso posterior al update, lo hara con el ultimo
    ' registro del recordset.
    Public Function Update(ByVal pvObjDatos() As Object, ByVal pvObjValores() As Object) As Boolean
        Try
            If pvObjDatos.Count = pvObjValores.Count Then
                Dim prvobDRow As DataRow
                prvobDRow = PrvDtViewVista(prvlngPosicion).Row

                For vshtIndiceDatos As Short = 0 To pvObjDatos.Count - 1
                    prvobDRow(pvObjDatos(vshtIndiceDatos)) = pvObjValores(vshtIndiceDatos)
                Next
                prvDtableTabla.AcceptChanges()

                Return True
            Else
                Err.Raise(0, , "Numero de valores y columnas no coinciden", , )
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '#End Region

    '#Region "Funciones de limpieza"

    Private Sub LimpiarPosicion(Optional pvblnServer As Boolean = False)
        prvBlnServer = pvblnServer
        prvlngPosicion = -1
        If Not Me Is Nothing Then
            If RecordCount(pvblnServer:=prvBlnServer) > 0 Then
                MoveFirst(pvblnServer:=prvBlnServer)
            End If
        End If
        prvintEstado = pubEnumAction.pubEnumActionStandBy
    End Sub
    Private Sub LimpiarPosicionServer()
        prvBlnServer = True
        LimpiarPosicion(pvblnServer:=True)
        prvBlnServer = False
    End Sub
    Public Property CursorLocation() As clsRecordset.pubEnumCursorLocationEnum
        Get
            Return prvintCursorLocation
        End Get
        Set(ByVal value As clsRecordset.pubEnumCursorLocationEnum)
            prvintCursorLocation = value
        End Set
    End Property
    Public Property Cursortype() As UnoEEDatos.clsRecordset.pubEnumCursorTypeEnum
        Get
            Return prvintCursorType
        End Get
        Set(ByVal value As UnoEEDatos.clsRecordset.pubEnumCursorTypeEnum)
            prvintCursorType = value
        End Set
    End Property

    'Public Property Fields(ByVal pvstrColumnName As String) As clsValor
    '    Get
    '        prvobValor.pubblnAccesoPorNombre = True
    '        prvobValor.pubstrColumnName = pvstrColumnName
    '        prvobValor.publngPosicion = prvlngPosicion
    '        prvobValor.pubintEstado = prvintEstado
    '        Fields = prvobValor
    '    End Get

    '    Set(ByVal value As clsValor)
    '        If prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro Then
    '            prvobDRow.Item(pvstrColumnName) = value
    '            prvobValor.Value = value
    '        Else
    '            prvDtableTabla.Rows(prvlngPosicion).Item(pvstrColumnName) = value
    '            prvobValor.Value = value
    '        End If

    '    End Set

    'End Property
    'Public Property Fields(ByVal pvintColumnIndex As Integer) As clsValor
    '    Get
    '        prvobValor.pubintColumnIndex = pvintColumnIndex
    '        prvobValor.pubblnAccesoPorNombre = False
    '        prvobValor.publngPosicion = prvlngPosicion
    '        prvobValor.pubintEstado = prvintEstado
    '        Fields = prvobValor
    '    End Get

    '    Set(ByVal value As clsValor)
    '        If prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro Then
    '            prvobDRow.Item(pvintColumnIndex) = value
    '            prvobValor.Value = value
    '        Else
    '            prvDtableTabla.Rows(prvlngPosicion).Item(pvintColumnIndex) = value
    '            prvobValor.Value = value
    '        End If

    '    End Set
    'End Property


    Public Property Fields(ByVal pvstrColumnName As String, Optional pvblnServer As Boolean = False) As clsField
        Get
            prvBlnServer = pvblnServer
            prvblnAccesoPorNombre = True
            Dim prvObValor As New clsField
            prvObValor.pubstrColumnName = pvstrColumnName
            prvObValor.ObjRecordSetOri = Me
            Fields = prvObValor
            prvObValor = Nothing
        End Get
        Set(ByVal value As clsField)
            'If prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro Then
            '    prvobDRow.Item(pvstrColumnName) = value
            '    'prvobValor.Value = value
            'Else
            prvBlnServer = pvblnServer
            prvDtableTabla.Rows(prvlngPosicion).Item(pvstrColumnName) = value
            'prvobValor.Value = value
            'End If
        End Set
    End Property
    Public Property FieldsServer(ByVal pvstrColumnName As String) As clsField
        Get

            FieldsServer = Fields(pvstrColumnName, pvblnServer:=True)
            ' prvBlnServer = False
        End Get
        Set(ByVal value As clsField)
            'If prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro Then
            '    prvobDRow.Item(pvstrColumnName) = value
            '    'prvobValor.Value = value
            'Else
            Fields(pvstrColumnName, pvblnServer:=True) = value

            'prvBlnServer = False
            'prvobValor.Value = value
            'End If
        End Set
    End Property

    Public Property Fields(ByVal pvintColumnIndex As Integer, Optional pvblnServer As Boolean = False) As clsField
        Get
            prvBlnServer = pvblnServer
            prvblnAccesoPorNombre = False
            Dim prvObValor As New clsField
            prvObValor.pubintColumnIndex = pvintColumnIndex
            prvObValor.ObjRecordSetOri = Me
            Fields = prvObValor
            prvObValor = Nothing
        End Get
        Set(ByVal value As clsField)
            prvBlnServer = pvblnServer
            'If prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro Then
            '    prvobDRow.Item(pvintColumnIndex) = value
            '    'prvobValor.Value = value
            'Else
            'prvDtViewVista.Rows(prvlngPosicion).Item(pvintColumnIndex) = value
            prvDtableTabla.Rows(prvlngPosicion).Item(pvintColumnIndex) = value
            'prvobValor.Value = value
            'End If
        End Set
    End Property
    Public Property FieldsServer(ByVal pvintColumnIndex As Integer) As clsField
        Get

            FieldsServer = Fields(pvintColumnIndex, pvblnServer:=True)
            'prvBlnServer = False
        End Get
        Set(ByVal value As clsField)
            'If prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro Then
            '    prvobDRow.Item(pvintColumnIndex) = value
            '    'prvobValor.Value = value
            'Else
            'prvDtViewVista.Rows(prvlngPosicion).Item(pvintColumnIndex) = value
            Fields(pvintColumnIndex, pvblnServer:=True) = value


            'prvobValor.Value = value
            'End If
        End Set
    End Property
    Public ReadOnly Property Fields(Optional pvblnServer As Boolean = False) As clsFields
        Get
            prvBlnServer = pvblnServer
            Dim prvObValor As New clsFields
            prvObValor.objRecordSetOri = Me
            prvObValor.Inicializar()
            Fields = prvObValor
            prvObValor = Nothing
        End Get
        'Set(ByVal value As clsField)
        '    If prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro Then
        '        prvobDRow.Item(pvstrColumnName) = value
        '        'prvobValor.Value = value
        '    Else
        '        prvDtableTabla.Rows(prvlngPosicion).Item(pvstrColumnName) = value
        '        'prvobValor.Value = value
        '    End If
        'End Set
    End Property
    Public ReadOnly Property FieldsServer() As clsFields
        Get

            FieldsServer = Fields(pvblnServer:=True)
            'prvBlnServer = False

        End Get
        'Set(ByVal value As clsField)
        '    If prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro Then
        '        prvobDRow.Item(pvstrColumnName) = value
        '        'prvobValor.Value = value
        '    Else
        '        prvDtableTabla.Rows(prvlngPosicion).Item(pvstrColumnName) = value
        '        'prvobValor.Value = value
        '    End If
        'End Set
    End Property


    Public Function FieldsCount(Optional pvblnServer As Boolean = False) As Integer

        FieldsCount = 0
        Try
            prvBlnServer = pvblnServer
            FieldsCount = prvDtableTabla.Columns.Count
        Catch ex As Exception

        End Try
    End Function
    Public Function FieldsCountServer() As Integer

        Try
            FieldsCountServer = FieldsCount(pvblnServer:=True)
            prvBlnServer = False
        Catch ex As Exception

        End Try
    End Function

    '#End Region

    '#Region "Creacion campos"
    'Migracion: req.137891; cpp ; se adiciona el parametro opcional pvAttrib para el manejo de ciertos atributos de la columna. Se adiconan en el momento los que se han encontrado 
    'en vb6. 
    Public Sub Append(ByVal pvstrNombreCampo As String, ByVal pvintDataType As DbType,
                                Optional pvintSize As Integer = 0, Optional pvAttrib As pubEnumFieldAttributeEnum = pubEnumFieldAttributeEnum.adFldUnspecified)
        Dim obcolumn As DataColumn
        obcolumn = New DataColumn()
        obcolumn.ColumnName = pvstrNombreCampo
        obcolumn.DataType = LeerDataTypeCampo(pvintDataType)

        'Manejo de los parametros opcionales...
        If (pvintSize > 0) Then
            Select Case pvintDataType
                Case DbType.Binary, DbType.Byte, DbType.String, DbType.StringFixedLength
                    obcolumn.MaxLength = pvintSize
            End Select
        End If
        If (pvAttrib > pubEnumFieldAttributeEnum.adFldUnspecified) Then
            Select Case pvAttrib
                Case pubEnumFieldAttributeEnum.adFldIsNullable, pubEnumFieldAttributeEnum.adFldMayBeNull
                    obcolumn.AllowDBNull = True
                Case pubEnumFieldAttributeEnum.adFldKeyColumn
                    prvDtableTabla.PrimaryKey = New DataColumn() {obcolumn}
                Case pubEnumFieldAttributeEnum.adFldUpdatable
                    obcolumn.ReadOnly = False
                Case (pubEnumFieldAttributeEnum.adFldKeyColumn + pubEnumFieldAttributeEnum.adFldUpdatable)
                    prvDtableTabla.PrimaryKey = New DataColumn() {obcolumn}
                    obcolumn.ReadOnly = False
            End Select
        End If

        ' Add the Column to the DataColumnCollection.
        prvDtableTabla.Columns.Add(obcolumn)
    End Sub

    Public Sub Append(ByVal pvstrNombreCampo As String, ByVal pvintDataType As DbType)
        Dim obcolumn As DataColumn
        obcolumn = New DataColumn()
        obcolumn.ColumnName = pvstrNombreCampo
        obcolumn.DataType = LeerDataTypeCampo(pvintDataType)
        ' Add the Column to the DataColumnCollection.
        prvDtableTabla.Columns.Add(obcolumn)

    End Sub
    Private Function LeerDataTypeCampo(ByVal pvintDataType As DbType) As System.Type
        Select Case pvintDataType
            Case DbType.Binary
                LeerDataTypeCampo = System.Type.GetType("System.Byte[]")
            Case DbType.Byte
                LeerDataTypeCampo = System.Type.GetType("System.Byte")
            Case DbType.Boolean
                LeerDataTypeCampo = System.Type.GetType("System.Boolean")
            Case DbType.Currency, DbType.Decimal
                LeerDataTypeCampo = System.Type.GetType("System.Decimal")
            Case DbType.Date, DbType.DateTime, DbType.DateTime2
                LeerDataTypeCampo = System.Type.GetType("System.DateTime")
            Case DbType.Double
                LeerDataTypeCampo = System.Type.GetType("System.Double")
            Case DbType.Int16
                LeerDataTypeCampo = System.Type.GetType("System.Int16")
            Case DbType.Int32
                LeerDataTypeCampo = System.Type.GetType("System.Int32")
            Case DbType.Int64
                LeerDataTypeCampo = System.Type.GetType("System.Int64")
            Case DbType.String
                LeerDataTypeCampo = System.Type.GetType("System.String")
            Case DbType.StringFixedLength
                LeerDataTypeCampo = System.Type.GetType("System.Char")
            Case DbType.Guid
                LeerDataTypeCampo = System.Type.GetType("System.Guid")
            Case Else
                LeerDataTypeCampo = System.Type.GetType("System.String")
        End Select
    End Function

    '#End Region
    Public Sub New(ByVal SI As Runtime.Serialization.SerializationInfo, ByVal SC As Runtime.Serialization.StreamingContext)
        prvstrPath = SI.GetString("dat1")
        PrvStrSort = SI.GetString("dat2")
        PrvStrFilter = SI.GetString("dat3")
        prvintCursorLocation = SI.GetInt32("dat4")
        prvintCursorType = SI.GetInt32("dat5")

        '#Region "Local Variables"

        prvDtableTabla = SI.GetValue("dat6", GetType(DataTable))
        prvstrNombreDTable = SI.GetString("dat7")
        prvintEstado = SI.GetInt32("dat8")
        prvblnAccesoPorNombre = SI.GetBoolean("dat9")
        prvlngPosicion = SI.GetInt64("dat10")
        prvintAction = SI.GetInt32("dat11")
        PrvBlnTablaFiltrada = SI.GetBoolean("dat12")
        PrvBlnTablaOrden = SI.GetBoolean("dat13")

        'prvDtViewVista = New DataView(prvDtableTabla)
        ' prvDtViewVista.Table = prvDtableTabla
        PrvDtViewVista.RowFilter = PrvStrFilter
        PrvDtViewVista.Sort = PrvStrSort


    End Sub
    Public Sub GetObjectData(ByVal info As SerializationInfo, ByVal context As StreamingContext) Implements System.Runtime.Serialization.ISerializable.GetObjectData

        info.AddValue("dat1", Me.prvstrPath)
        info.AddValue("dat2", Me.PrvStrSort)
        info.AddValue("dat3", Me.PrvStrFilter)
        info.AddValue("dat4", Me.prvintCursorLocation)
        info.AddValue("dat5", Me.prvintCursorType)
        info.AddValue("dat6", Me.prvDtableTabla)
        info.AddValue("dat7", Me.prvstrNombreDTable)
        info.AddValue("dat8", Me.prvintEstado)
        info.AddValue("dat9", Me.prvblnAccesoPorNombre)
        info.AddValue("dat10", Me.prvlngPosicion)
        info.AddValue("dat11", Me.prvintAction)
        info.AddValue("dat12", Me.PrvBlnTablaFiltrada)
        info.AddValue("dat13", Me.PrvBlnTablaOrden)

    End Sub

    'Metodo para darle las capacidades de convertir a XML con los campos que tenga:
    Public Function toXML(ByVal pvStrNombreTabla As String, Optional ByVal pvStrNombreRaiz As String = "root") As String
        Dim vXmlDocumento As Xml.XmlDocument
        Dim vXmlInfoTabla As Xml.XmlElement
        Dim vXmlRaiz As Xml.XmlElement

        If PrvDtViewVista.Count <= 0 Then
            toXML = String.Empty
            Exit Function
        End If

        Try
            vXmlDocumento = New Xml.XmlDocument
            vXmlRaiz = vXmlDocumento.CreateElement(pvStrNombreRaiz)

            For vIntIndice As Integer = 0 To PrvDtViewVista.Count - 1
                vXmlInfoTabla = vXmlDocumento.CreateElement(pvStrNombreTabla)
                For vIntIndiceColumnas As Integer = 0 To prvDtableTabla.Columns.Count - 1
                    If Not IsDBNull(PrvDtViewVista(vIntIndice).Item(vIntIndiceColumnas)) Then
                        vXmlInfoTabla.SetAttribute(prvDtableTabla.Columns(vIntIndiceColumnas).ColumnName.ToLower, FormatearCampoXML(PrvDtViewVista(vIntIndice).Item(vIntIndiceColumnas)))
                    End If
                Next
                vXmlRaiz.AppendChild(vXmlInfoTabla)
            Next
            vXmlDocumento.AppendChild(vXmlRaiz)
            toXML = vXmlDocumento.InnerXml
        Catch ex As Exception
            toXML = String.Empty
        End Try
    End Function
#Region "clsField"
    '###############################################
    '###########CLASE ESPECIAL PARA LOS .values
    '###############################################
    'contiene la referencia al recordset original de donde salen los datos, antes se habia utilizado variables shared, pero todas las instancias
    'recordset entonces se modificabana un mismo dato ya q al ser shared todas tienen el mismo dato en cada instante
    <Serializable()>
    Public Class clsField
        ' Inherits MarshalByRefObject
        Public pubintColumnIndex As Integer
        Public pubstrColumnName As String
        Private _objRecordSetOri As UnoEEDatos.clsRecordset
        ''Public objDatos As DataRow
        Private prvshtNumericScale As Short
        Private prvshtPrecision As Short
        Protected Friend pubobjParentFields As clsFields

        ''' <summary>
        ''' Lee el valor del campo
        ''' </summary>
        ''' <returns></returns>
        Public Function GetValue() As Object
            Return Me.Value
        End Function
        Public ReadOnly Property ValueNull() As Object

            Get


                If ObjRecordSetOri.prvblnAccesoPorNombre Then
                    ' Accesa por nombre de columna
                    'Value = prvDtableTabla.Rows(prvlngPosicion).Item(pubstrColumnName)
                    'Value = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubstrColumnName)
                    Return LeerDatoDataView(ObjRecordSetOri, ObjRecordSetOri.prvlngPosicion, pubstrColumnName, ObjRecordSetOri.prvBlnServer, pvblnReturnNullSinCondicion:=True)
                    '' ''Value = objDatos(pubstrColumnName)
                Else
                    ' Accesa por indice de columna
                    'Value = prvDtableTabla.Rows(prvlngPosicion).Item(pubintColumnIndex)
                    'Value = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubintColumnIndex)
                    Return LeerDatoDataView(ObjRecordSetOri, ObjRecordSetOri.prvlngPosicion, pubintColumnIndex, ObjRecordSetOri.prvBlnServer, pvblnReturnNullSinCondicion:=True)
                    ''Value = objDatos(pubintColumnIndex)
                End If

            End Get
        End Property
        Public Property Value() As Object

            Get
                ' Por jairc 20210924
                ' Si esta filtrada y es nuevo accede al dato por objeto prvnewrow
                If ObjRecordSetOri.prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro And Not ObjRecordSetOri.prvobDNewRow Is Nothing And ObjRecordSetOri.PrvBlnTablaFiltrada Then

                    If ObjRecordSetOri.prvblnAccesoPorNombre Then
                        Return ObjRecordSetOri.prvobDNewRow.Item(pubstrColumnName)
                    Else
                        ' Accesa por indice de columna
                        Return ObjRecordSetOri.prvobDNewRow.Item(pubintColumnIndex)
                    End If
                Else

                    If ObjRecordSetOri.prvblnAccesoPorNombre Then
                        ' Accesa por nombre de columna
                        'Value = prvDtableTabla.Rows(prvlngPosicion).Item(pubstrColumnName)
                        'Value = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubstrColumnName)
                        Return LeerDatoDataView(ObjRecordSetOri, ObjRecordSetOri.prvlngPosicion, pubstrColumnName, ObjRecordSetOri.prvBlnServer)
                        '' ''Value = objDatos(pubstrColumnName)
                    Else
                        ' Accesa por indice de columna
                        'Value = prvDtableTabla.Rows(prvlngPosicion).Item(pubintColumnIndex)
                        'Value = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubintColumnIndex)
                        Return LeerDatoDataView(ObjRecordSetOri, ObjRecordSetOri.prvlngPosicion, pubintColumnIndex, ObjRecordSetOri.prvBlnServer)
                        ''Value = objDatos(pubintColumnIndex)
                    End If
                End If

            End Get
            Set(ByVal value As Object)
                Dim vobjDRow As DataRow = Nothing
                Dim vobjDRowClone As DataRow = Nothing
                Dim vlngNumRegistros As Long
                Dim vEnumStatus As clsRecordset.EventStatusEnumRecordset
                Dim vblnAcceptChanges As Boolean = True

                ' Por jairc 20210303
                ''' Reasigna el valor del maxlenght de la columna de acuerdo al dato enviado, si el dato enviado es mayor reasigna el maxlenght
                ''' Solo lo hace cuando el tipo de dato de la columna es string
                ReasignarMaxLength(value)

                ' Por jairc 20210601
                ' Para evitar que genere el codigo del evento rowcolchangevb6 en la grilla
                If Not ObjRecordSetOri.pubobjGrilla Is Nothing Then
                    ObjRecordSetOri.pubobjGrilla.pubblnDesdeRecordset = True
                    ' Por gvm, rrf, jairc, 20210708
                    ' Se evalua si el cambio viene desde el evento BeforeColUpdate, si es asi el acceptchanges no se debe ejecutar
                    ' porque no ha terminado la sincronizacion, este acceptchanges se hace al terminar el evento BeforColUpdate Nativo
                    vblnAcceptChanges = Not (ObjRecordSetOri.pubintEventoActivoGrilla = 1) ' No se puede quemar el enum porque no tiene referencia a unoeegeneral
                End If

                If ObjRecordSetOri.PrvBlnTablaFiltrada Then
                    ' Lee el numero de registros para saber si con el valor que se cambia afecta el filtro
                    vlngNumRegistros = ObjRecordSetOri.RecordCount
                    If ObjRecordSetOri.prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro And Not ObjRecordSetOri.prvobDNewRow Is Nothing Then
                        If ObjRecordSetOri.prvblnAccesoPorNombre Then
                            ObjRecordSetOri.prvobDNewRow.Item(pubstrColumnName) = value
                            vEnumStatus = EventStatusEnumRecordset.adStatusOK
                            ' Si es clone afecta el dato de la tabla origina;
                            If ObjRecordSetOri.prvblnIsClone Then
                                ObjRecordSetOri.prvObjRecOrigenClone.prvobDNewRow.Item(pubstrColumnName) = value
                            End If
                        Else
                            ' Accesa por indice de columna
                            ObjRecordSetOri.prvobDNewRow.Item(pubintColumnIndex) = value
                            vEnumStatus = EventStatusEnumRecordset.adStatusOK
                            ' Si es clone afecta el dato de la tabla origina;
                            If ObjRecordSetOri.prvblnIsClone Then
                                ObjRecordSetOri.prvObjRecOrigenClone.prvobDNewRow.Item(pubintColumnIndex) = value
                            End If
                        End If
                        ' por jairc 20210628
                        ' Se habilita la condicion de validacion de estado de nuevo registro para que afecte el objeto
                        ' creado de la fila
                        'ObjRecordSetOri.prvlngPosicion = ObjRecordSetOri.PrvDtViewVista.Count - 1


                    Else

                        ' Por jairc 20210407
                        ' Deduce la fila correcta a afectar 
                        If ObjRecordSetOri.prvlngPosicion = clsRecordset.pubPositionEnum.adPosUnknown Then
                            vobjDRow = ObjRecordSetOri.pubObjDtTabla(ObjRecordSetOri.IndiceRowFilter)
                        Else
                            vobjDRow = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Row
                        End If
                        ObjRecordSetOri.IndiceRowFilter = ObjRecordSetOri.prvDtableTabla.Rows.IndexOf(vobjDRow)
                        ' Por jairc 20210407
                        ' Deduce la fila correcta a afectar cuando del datatable del clone
                        ' Esta el caso especificamente cuando es filter porque existen casos en los cuales
                        ' se afecta una columna que esta en el fitro, y la posicion manejada en prvlngPosicion no aplica
                        ' en este caso toca acceder por el indice del datatable no de la vista
                        If ObjRecordSetOri.prvblnIsClone Then
                            If ObjRecordSetOri.prvlngPosicion = clsRecordset.pubPositionEnum.adPosUnknown Then
                                vobjDRowClone = ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla(ObjRecordSetOri.IndiceRowFilter)
                                ObjRecordSetOri.IndiceRowFilter = ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.Rows.IndexOf(vobjDRowClone)
                            Else
                                'vobjDRowClone = ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla(ObjRecordSetOri.prvlngPosicion)
                                vobjDRowClone = ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla(ObjRecordSetOri.prvDtableTabla.Rows.IndexOf(vobjDRow))
                                ObjRecordSetOri.IndiceRowFilter = ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.Rows.IndexOf(vobjDRowClone)
                            End If

                        End If


                        If ObjRecordSetOri.prvblnAccesoPorNombre Then
                            ' Accesa por nombre de columna
                            'prvDtableTabla.Rows(prvlngPosicion).Item(pubstrColumnName) = value

                            If ObjRecordSetOri.prvblnIsClone Then
                                'vobjDRow = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Row
                                'ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla(ObjRecordSetOri.prvDtableTabla.Rows.IndexOf(vobjDRow)).Item(pubstrColumnName) = value
                                vobjDRowClone.Item(pubstrColumnName) = value
                                ' Por jairc 20210902 Op 201571
                                ' Afecta en cascada las filas que fueron creadas como clone
                                UpdateCloneFromDataRowInCascade(ObjRecordSetOri.prvObjRecOrigenClone, vobjDRowClone, value)
                            End If
                            vobjDRow.Item(pubstrColumnName) = value
                            vEnumStatus = EventStatusEnumRecordset.adStatusOK

                        Else
                            ' Accesa por indice de columna
                            'prvDtableTabla.Rows(prvlngPosicion).Item(pubintColumnIndex) = value
                            'lsdt: 25/11/2020 req: 185526  Se quita la propiedad .Table por que trai los datos de la tabla sin el filtro

                            If ObjRecordSetOri.prvblnIsClone Then
                                'vobjDRow = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Row
                                'ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla(ObjRecordSetOri.prvDtableTabla.Rows.IndexOf(vobjDRow)).Item(pubstrColumnName) = value
                                vobjDRowClone.Item(pubintColumnIndex) = value
                                ' Por jairc 20210902, Op 201571
                                ' Afecta en cascada las filas que fueron creadas como clone
                                UpdateCloneFromDataRowInCascade(ObjRecordSetOri.prvObjRecOrigenClone, vobjDRowClone, value)
                            End If
                            vobjDRow.Item(pubintColumnIndex) = value
                            vEnumStatus = EventStatusEnumRecordset.adStatusOK
                        End If
                    End If

                    If vblnAcceptChanges Then
                        If Not ObjRecordSetOri.prvDtableTabla.GetChanges() Is Nothing Then
                            'vobjDRow.AcceptChanges()
                            ObjRecordSetOri.prvDtableTabla.AcceptChanges()
                        End If
                    End If

                    If ObjRecordSetOri.prvblnIsClone Then
                        'If Not ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.GetChanges() Is Nothing Then
                        'vobjDRowClone.AcceptChanges()
                        ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.AcceptChanges()
                        'vobjDRowClone.AcceptChanges()
                        'End If
                        'End If
                    End If
                    ' Si despues de fitrado el numero de registros disminuye, signifca que afecto un campo que esta en el filtro
                    ' entonces se deje ubicado en la posicion -1 que es el inicio
                    ' Si se hace movenext se colocaria en la primera posicion
                    If ObjRecordSetOri.RecordCount < vlngNumRegistros Then
                        ObjRecordSetOri.prvlngPosicion = clsRecordset.pubPositionEnum.adPosUnknown
                    End If
                Else ' Cuando la tabla no esta filtrada

                    ' Deduce la fila a afectar
                    vobjDRow = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Row
                    If ObjRecordSetOri.prvblnIsClone Then
                        'vobjDRowClone = ObjRecordSetOri.prvObjRecOrigenClone.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Row
                        vobjDRowClone = ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla(ObjRecordSetOri.prvDtableTabla.Rows.IndexOf(vobjDRow))
                    End If

                    If ObjRecordSetOri.prvblnAccesoPorNombre Then
                        ' Accesa por nombre de columna
                        'prvDtableTabla.Rows(prvlngPosicion).Item(pubstrColumnName) = value
                        'vobjDRow = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Row
                        'ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubstrColumnName) = value
                        'rrf, jairc req:191025 20210216 Se cambia porque al dar clic 3 veces en el headclic de grupos impositivos fallaba
                        'ObjRecordSetOri.prvDtableTabla.Rows(ObjRecordSetOri.prvlngPosicion).AcceptChanges()
                        'ObjRecordSetOri.prvDtableTabla.AcceptChanges()
                        vobjDRow.Item(pubstrColumnName) = value
                        vEnumStatus = EventStatusEnumRecordset.adStatusOK
                        If ObjRecordSetOri.prvblnIsClone Then
                            ' MDI (Manuel Diaz) OP 203166: Se coloca para que cuando se este agregando una fila nueva y no se hayan actualizado los datos de un clone tome
                            ' los datos de la vista
                            If Not ObjRecordSetOri.pubobjGrilla Is Nothing Then
                                If ObjRecordSetOri.pubobjGrilla.AddNewMode = 2 Then
                                    ObjRecordSetOri.prvObjRecOrigenClone.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubstrColumnName) = value
                                Else
                                    vobjDRowClone.Item(pubstrColumnName) = value
                                End If
                            Else
                                vobjDRowClone.Item(pubstrColumnName) = value
                            End If
                            'ObjRecordSetOri.prvObjRecOrigenClone.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubintColumnIndex) = value
                            'ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.AcceptChanges()
                        End If
                    Else
                        ' Accesa por indice de columna
                        'prvDtableTabla.Rows(prvlngPosicion).Item(pubintColumnIndex) = value
                        'ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubintColumnIndex) = value
                        'rrf, jairc req:191025 20210216 Se cambia porque al dar clic 3 veces en el headclic de grupos impositivos fallaba
                        'ObjRecordSetOri.prvDtableTabla.Rows(ObjRecordSetOri.prvlngPosicion).AcceptChanges()
                        vobjDRow.Item(pubintColumnIndex) = value
                        vEnumStatus = EventStatusEnumRecordset.adStatusOK
                        'ObjRecordSetOri.prvDtableTabla.AcceptChanges()
                        If ObjRecordSetOri.prvblnIsClone Then
                            ' MDI (Manuel Diaz) OP 203166: Se coloca para que cuando se este agregando una fila nueva y no se hayan actualizado los datos de un clone tome
                            ' los datos de la vista
                            If Not ObjRecordSetOri.pubobjGrilla Is Nothing Then
                                If ObjRecordSetOri.pubobjGrilla.AddNewMode = 2 Then
                                    ObjRecordSetOri.prvObjRecOrigenClone.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubintColumnIndex) = value
                                Else
                                    vobjDRowClone.Item(pubintColumnIndex) = value
                                End If
                            Else
                                vobjDRowClone.Item(pubintColumnIndex) = value
                            End If
                            'ObjRecordSetOri.prvObjRecOrigenClone.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubintColumnIndex) = value
                            'ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.AcceptChanges()
                        End If
                    End If

                    If vblnAcceptChanges Then
                        If Not ObjRecordSetOri.prvDtableTabla.GetChanges() Is Nothing Then
                            'vobjDRow.AcceptChanges()
                            ObjRecordSetOri.prvDtableTabla.AcceptChanges()
                        End If
                    End If
                    If ObjRecordSetOri.prvblnIsClone Then
                        If Not ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.GetChanges() Is Nothing Then
                            'vobjDRowClone.AcceptChanges()
                            ' MDI (Manuel Diaz) OP 203166: Se hace para que no entre a los eventos afterupdate y afterinsert cuando este haciendo este acceptchanges
                            If Not ObjRecordSetOri.pubobjGrilla Is Nothing Then
                                ObjRecordSetOri.prvObjRecOrigenClone.pubBlnAcceptChangesDesdeValue = (ObjRecordSetOri.pubobjGrilla.pubIntEventoActivo = 3 Or ObjRecordSetOri.pubobjGrilla.pubIntEventoActivo = 8)
                            End If
                            'ObjRecordSetOri.prvObjRecOrigenClone.pubBlnAcceptChangesDesdeValue = True
                            ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.AcceptChanges()
                            ObjRecordSetOri.prvObjRecOrigenClone.pubBlnAcceptChangesDesdeValue = False
                        End If
                    End If


                End If
                If Not ObjRecordSetOri.pubobjGrilla Is Nothing Then
                    ObjRecordSetOri.pubobjGrilla.pubblnDesdeRecordset = False
                End If
                ObjRecordSetOri.llamarEventoFieldChangeComplete(1, Me, Nothing, vEnumStatus, ObjRecordSetOri)
            End Set
        End Property


        Public Property ValueWithoutAcceptChanges() As Object
            Get
                ' Por jairc 20210924
                ' Si esta filtrada y es nuevo accede al dato por objeto prvnewrow
                If ObjRecordSetOri.prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro And Not ObjRecordSetOri.prvobDNewRow Is Nothing And ObjRecordSetOri.PrvBlnTablaFiltrada Then

                    If ObjRecordSetOri.prvblnAccesoPorNombre Then
                        Return ObjRecordSetOri.prvobDNewRow.Item(pubstrColumnName)
                    Else
                        ' Accesa por indice de columna
                        Return ObjRecordSetOri.prvobDNewRow.Item(pubintColumnIndex)
                    End If
                Else

                    If ObjRecordSetOri.prvblnAccesoPorNombre Then
                        ' Accesa por nombre de columna
                        'Value = prvDtableTabla.Rows(prvlngPosicion).Item(pubstrColumnName)
                        'Value = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubstrColumnName)
                        Return LeerDatoDataView(ObjRecordSetOri, ObjRecordSetOri.prvlngPosicion, pubstrColumnName, ObjRecordSetOri.prvBlnServer)
                        '' ''Value = objDatos(pubstrColumnName)
                    Else
                        ' Accesa por indice de columna
                        'Value = prvDtableTabla.Rows(prvlngPosicion).Item(pubintColumnIndex)
                        'Value = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubintColumnIndex)
                        Return LeerDatoDataView(ObjRecordSetOri, ObjRecordSetOri.prvlngPosicion, pubintColumnIndex, ObjRecordSetOri.prvBlnServer)
                        ''Value = objDatos(pubintColumnIndex)
                    End If
                End If

            End Get
            Set(ByVal value As Object)
                Dim vobjDRow As DataRow = Nothing
                Dim vobjDRowClone As DataRow = Nothing
                Dim vlngNumRegistros As Long
                Dim vEnumStatus As clsRecordset.EventStatusEnumRecordset
                Dim vblnAcceptChanges As Boolean = True

                ' Por jairc 20210303
                ''' Reasigna el valor del maxlenght de la columna de acuerdo al dato enviado, si el dato enviado es mayor reasigna el maxlenght
                ''' Solo lo hace cuando el tipo de dato de la columna es string
                ReasignarMaxLength(value)

                ' Por jairc 20210601
                ' Para evitar que genere el codigo del evento rowcolchangevb6 en la grilla
                If Not ObjRecordSetOri.pubobjGrilla Is Nothing Then
                    ObjRecordSetOri.pubobjGrilla.pubblnDesdeRecordset = True
                    ' Por gvm, rrf, jairc, 20210708
                    ' Se evalua si el cambio viene desde el evento BeforeColUpdate, si es asi el acceptchanges no se debe ejecutar
                    ' porque no ha terminado la sincronizacion, este acceptchanges se hace al terminar el evento BeforColUpdate Nativo
                    vblnAcceptChanges = Not (ObjRecordSetOri.pubintEventoActivoGrilla = 1) ' No se puede quemar el enum porque no tiene referencia a unoeegeneral
                End If

                If ObjRecordSetOri.PrvBlnTablaFiltrada Then
                    ' Lee el numero de registros para saber si con el valor que se cambia afecta el filtro
                    vlngNumRegistros = ObjRecordSetOri.RecordCount
                    If ObjRecordSetOri.prvintEstado = pubEnumAction.pubEnumActionNuevoRegistro And Not ObjRecordSetOri.prvobDNewRow Is Nothing Then
                        If ObjRecordSetOri.prvblnAccesoPorNombre Then
                            ObjRecordSetOri.prvobDNewRow.Item(pubstrColumnName) = value
                            vEnumStatus = EventStatusEnumRecordset.adStatusOK
                            ' Si es clone afecta el dato de la tabla origina;
                            If ObjRecordSetOri.prvblnIsClone Then
                                ObjRecordSetOri.prvObjRecOrigenClone.prvobDNewRow.Item(pubstrColumnName) = value
                            End If
                        Else
                            ' Accesa por indice de columna
                            ObjRecordSetOri.prvobDNewRow.Item(pubintColumnIndex) = value
                            vEnumStatus = EventStatusEnumRecordset.adStatusOK
                            ' Si es clone afecta el dato de la tabla origina;
                            If ObjRecordSetOri.prvblnIsClone Then
                                ObjRecordSetOri.prvObjRecOrigenClone.prvobDNewRow.Item(pubintColumnIndex) = value
                            End If
                        End If
                        ' por jairc 20210628
                        ' Se habilita la condicion de validacion de estado de nuevo registro para que afecte el objeto
                        ' creado de la fila
                        'ObjRecordSetOri.prvlngPosicion = ObjRecordSetOri.PrvDtViewVista.Count - 1
                    Else
                        ' Por jairc 20210407
                        ' Deduce la fila correcta a afectar 
                        If ObjRecordSetOri.prvlngPosicion = clsRecordset.pubPositionEnum.adPosUnknown Then
                            vobjDRow = ObjRecordSetOri.pubObjDtTabla(ObjRecordSetOri.IndiceRowFilter)
                        Else
                            vobjDRow = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Row
                        End If
                        ObjRecordSetOri.IndiceRowFilter = ObjRecordSetOri.prvDtableTabla.Rows.IndexOf(vobjDRow)
                        ' Por jairc 20210407
                        ' Deduce la fila correcta a afectar cuando del datatable del clone
                        ' Esta el caso especificamente cuando es filter porque existen casos en los cuales
                        ' se afecta una columna que esta en el fitro, y la posicion manejada en prvlngPosicion no aplica
                        ' en este caso toca acceder por el indice del datatable no de la vista
                        If ObjRecordSetOri.prvblnIsClone Then
                            If ObjRecordSetOri.prvlngPosicion = clsRecordset.pubPositionEnum.adPosUnknown Then
                                vobjDRowClone = ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla(ObjRecordSetOri.IndiceRowFilter)
                                ObjRecordSetOri.IndiceRowFilter = ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.Rows.IndexOf(vobjDRowClone)
                            Else
                                'vobjDRowClone = ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla(ObjRecordSetOri.prvlngPosicion)
                                vobjDRowClone = ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla(ObjRecordSetOri.prvDtableTabla.Rows.IndexOf(vobjDRow))
                                ObjRecordSetOri.IndiceRowFilter = ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.Rows.IndexOf(vobjDRowClone)
                            End If

                        End If


                        If ObjRecordSetOri.prvblnAccesoPorNombre Then
                            ' Accesa por nombre de columna
                            'prvDtableTabla.Rows(prvlngPosicion).Item(pubstrColumnName) = value

                            If ObjRecordSetOri.prvblnIsClone Then
                                'vobjDRow = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Row
                                'ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla(ObjRecordSetOri.prvDtableTabla.Rows.IndexOf(vobjDRow)).Item(pubstrColumnName) = value
                                vobjDRowClone.Item(pubstrColumnName) = value
                                ' Por jairc 20210902 Op 201571
                                ' Afecta en cascada las filas que fueron creadas como clone
                                UpdateCloneFromDataRowInCascade(ObjRecordSetOri.prvObjRecOrigenClone, vobjDRowClone, value)
                            End If
                            vobjDRow.Item(pubstrColumnName) = value
                            vEnumStatus = EventStatusEnumRecordset.adStatusOK

                        Else
                            ' Accesa por indice de columna
                            'prvDtableTabla.Rows(prvlngPosicion).Item(pubintColumnIndex) = value
                            'lsdt: 25/11/2020 req: 185526  Se quita la propiedad .Table por que trai los datos de la tabla sin el filtro

                            If ObjRecordSetOri.prvblnIsClone Then
                                'vobjDRow = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Row
                                'ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla(ObjRecordSetOri.prvDtableTabla.Rows.IndexOf(vobjDRow)).Item(pubstrColumnName) = value
                                vobjDRowClone.Item(pubintColumnIndex) = value
                                ' Por jairc 20210902, Op 201571
                                ' Afecta en cascada las filas que fueron creadas como clone
                                UpdateCloneFromDataRowInCascade(ObjRecordSetOri.prvObjRecOrigenClone, vobjDRowClone, value)
                            End If
                            vobjDRow.Item(pubintColumnIndex) = value
                            vEnumStatus = EventStatusEnumRecordset.adStatusOK
                        End If
                    End If

                    'If vblnAcceptChanges Then
                    '    If Not ObjRecordSetOri.prvDtableTabla.GetChanges() Is Nothing Then
                    '        'vobjDRow.AcceptChanges()
                    '        'ObjRecordSetOri.prvDtableTabla.AcceptChanges()
                    '    End If
                    'End If

                    'If ObjRecordSetOri.prvblnIsClone Then
                    '    'If Not ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.GetChanges() Is Nothing Then
                    '    'vobjDRowClone.AcceptChanges()
                    '    'ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.AcceptChanges()
                    '    'vobjDRowClone.AcceptChanges()
                    '    'End If
                    '    'End If
                    'End If
                    ' Si despues de fitrado el numero de registros disminuye, signifca que afecto un campo que esta en el filtro
                    ' entonces se deje ubicado en la posicion -1 que es el inicio
                    ' Si se hace movenext se colocaria en la primera posicion
                    If ObjRecordSetOri.RecordCount < vlngNumRegistros Then
                        ObjRecordSetOri.prvlngPosicion = clsRecordset.pubPositionEnum.adPosUnknown
                    End If
                Else ' Cuando la tabla no esta filtrada

                    ' Deduce la fila a afectar
                    vobjDRow = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Row
                    If ObjRecordSetOri.prvblnIsClone Then
                        'vobjDRowClone = ObjRecordSetOri.prvObjRecOrigenClone.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Row
                        vobjDRowClone = ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla(ObjRecordSetOri.prvDtableTabla.Rows.IndexOf(vobjDRow))
                    End If

                    If ObjRecordSetOri.prvblnAccesoPorNombre Then
                        ' Accesa por nombre de columna
                        'prvDtableTabla.Rows(prvlngPosicion).Item(pubstrColumnName) = value
                        'vobjDRow = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Row
                        'ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubstrColumnName) = value
                        'rrf, jairc req:191025 20210216 Se cambia porque al dar clic 3 veces en el headclic de grupos impositivos fallaba
                        'ObjRecordSetOri.prvDtableTabla.Rows(ObjRecordSetOri.prvlngPosicion).AcceptChanges()
                        'ObjRecordSetOri.prvDtableTabla.AcceptChanges()
                        vobjDRow.Item(pubstrColumnName) = value
                        vEnumStatus = EventStatusEnumRecordset.adStatusOK
                        If ObjRecordSetOri.prvblnIsClone Then
                            ' MDI (Manuel Diaz) OP 203166: Se coloca para que cuando se este agregando una fila nueva y no se hayan actualizado los datos de un clone tome
                            ' los datos de la vista
                            If Not ObjRecordSetOri.pubobjGrilla Is Nothing Then
                                If ObjRecordSetOri.pubobjGrilla.AddNewMode = 2 Then
                                    ObjRecordSetOri.prvObjRecOrigenClone.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubstrColumnName) = value
                                Else
                                    vobjDRowClone.Item(pubstrColumnName) = value
                                End If
                            Else
                                vobjDRowClone.Item(pubstrColumnName) = value
                            End If
                            'ObjRecordSetOri.prvObjRecOrigenClone.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubintColumnIndex) = value
                            'ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.AcceptChanges()
                        End If
                    Else
                        ' Accesa por indice de columna
                        'prvDtableTabla.Rows(prvlngPosicion).Item(pubintColumnIndex) = value
                        'ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubintColumnIndex) = value
                        'rrf, jairc req:191025 20210216 Se cambia porque al dar clic 3 veces en el headclic de grupos impositivos fallaba
                        'ObjRecordSetOri.prvDtableTabla.Rows(ObjRecordSetOri.prvlngPosicion).AcceptChanges()
                        vobjDRow.Item(pubintColumnIndex) = value
                        vEnumStatus = EventStatusEnumRecordset.adStatusOK
                        'ObjRecordSetOri.prvDtableTabla.AcceptChanges()
                        If ObjRecordSetOri.prvblnIsClone Then
                            ' MDI (Manuel Diaz) OP 203166: Se coloca para que cuando se este agregando una fila nueva y no se hayan actualizado los datos de un clone tome
                            ' los datos de la vista
                            If Not ObjRecordSetOri.pubobjGrilla Is Nothing Then
                                If ObjRecordSetOri.pubobjGrilla.AddNewMode = 2 Then
                                    ObjRecordSetOri.prvObjRecOrigenClone.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubintColumnIndex) = value
                                Else
                                    vobjDRowClone.Item(pubintColumnIndex) = value
                                End If
                            Else
                                vobjDRowClone.Item(pubintColumnIndex) = value
                            End If
                            'ObjRecordSetOri.prvObjRecOrigenClone.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Item(pubintColumnIndex) = value
                            'ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.AcceptChanges()
                        End If
                    End If

                    'If vblnAcceptChanges Then
                    '    If Not ObjRecordSetOri.prvDtableTabla.GetChanges() Is Nothing Then
                    '        'vobjDRow.AcceptChanges()
                    '        'ObjRecordSetOri.prvDtableTabla.AcceptChanges()
                    '    End If
                    'End If
                    'If ObjRecordSetOri.prvblnIsClone Then
                    '    If Not ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.GetChanges() Is Nothing Then
                    '        'vobjDRowClone.AcceptChanges()
                    '        ' MDI (Manuel Diaz) OP 203166: Se hace para que no entre a los eventos afterupdate y afterinsert cuando este haciendo este acceptchanges
                    '        If Not ObjRecordSetOri.pubobjGrilla Is Nothing Then
                    '            ObjRecordSetOri.prvObjRecOrigenClone.pubBlnAcceptChangesDesdeValue = (ObjRecordSetOri.pubobjGrilla.pubIntEventoActivo = 3 Or ObjRecordSetOri.pubobjGrilla.pubIntEventoActivo = 8)
                    '        End If
                    '        'ObjRecordSetOri.prvObjRecOrigenClone.pubBlnAcceptChangesDesdeValue = True
                    '        'ObjRecordSetOri.prvObjRecOrigenClone.prvDtableTabla.AcceptChanges()
                    '        ObjRecordSetOri.prvObjRecOrigenClone.pubBlnAcceptChangesDesdeValue = False
                    '    End If
                    'End If

                End If
                If Not ObjRecordSetOri.pubobjGrilla Is Nothing Then
                    ObjRecordSetOri.pubobjGrilla.pubblnDesdeRecordset = False
                End If
                ObjRecordSetOri.llamarEventoFieldChangeComplete(1, Me, Nothing, vEnumStatus, ObjRecordSetOri)
            End Set
        End Property


        ''' <summary>
        ''' Devuelve el tipo de dato del campo
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Type() As DbType
            Get
                Try
                    If ObjRecordSetOri.prvblnAccesoPorNombre Then
                        ' Accesa por nombre de columna
                        Type = DevolverDataTypeCampo(ObjRecordSetOri.prvDtableTabla.Columns(pubstrColumnName).DataType)
                    Else
                        ' Accesa por indice de columna
                        Type = DevolverDataTypeCampo(ObjRecordSetOri.prvDtableTabla.Columns(pubintColumnIndex).DataType)
                    End If
                Catch ex As Exception
                    ' Se puede caer cuando el campo no exists
                    Type = DbType.String
                End Try

            End Get
        End Property
        ''' <summary>
        ''' Devuelve el tipo de dato del campo, retorna nothing si existe algun error
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property TypeReturnNothingIfError As Int16
            Get
                Try
                    If ObjRecordSetOri.prvblnAccesoPorNombre Then
                        ' Accesa por nombre de columna
                        Return DevolverDataTypeCampo(ObjRecordSetOri.prvDtableTabla.Columns(pubstrColumnName).DataType)
                    Else
                        ' Accesa por indice de columna
                        Return DevolverDataTypeCampo(ObjRecordSetOri.prvDtableTabla.Columns(pubintColumnIndex).DataType)
                    End If
                Catch ex As Exception
                    ' Se puede caer cuando el campo no exists
                    Return -1
                End Try

            End Get
        End Property
        Public ReadOnly Property Name() As String
            Get
                If ObjRecordSetOri.prvblnAccesoPorNombre Then
                    ' Accesa por nombre de columna
                    Dim pvDtColDato As DataColumn = ObjRecordSetOri.prvDtableTabla.Columns(pubstrColumnName)
                    Name = pvDtColDato.ColumnName

                Else
                    ' Accesa por indice de columna
                    Dim pvDtColDato As DataColumn = ObjRecordSetOri.prvDtableTabla.Columns(pubintColumnIndex)
                    Name = pvDtColDato.ColumnName
                End If
            End Get
        End Property

        ''' <summary>
        ''' Devuelve el tamano definido de la tabla, en vb6 lo trae en la estructura del recordset
        ''' pero en .Net toca deducirlo de la data por eso se saca un Max de los datos
        ''' Es readonly
        ''' </summary>
        ''' <returns>Tamano del campo</returns>
        Public ReadOnly Property DefinedSize() As Integer
            Get
                Dim vlngSize As Long

                If ObjRecordSetOri.RecordCount > 0 Then

                    If ObjRecordSetOri.prvblnAccesoPorNombre Then
                        ' Accesa por nombre de columna
                        ' DefinedSize = Len(ObjRecordSetOri.Fields(pubstrColumnName).Value)
                        'DefinedSize = Len(LeerDatoADO(ObjRecordSetOri, pubstrColumnName))
                        'DefinedSize = ObjRecordSetOri.prvDtableTabla.AsEnumerable().Max(Function(row) row.Field(Of String)(pubstrColumnName).Length)
                        vlngSize = (From row In ObjRecordSetOri.prvDtableTabla
                                    Select (row(pubstrColumnName).ToString.Length)).Max()
                    Else
                        ' Accesa por indice de columna
                        'DefinedSize = Len(ObjRecordSetOri.Fields(pubintColumnIndex).Value)

                        vlngSize = (From row In ObjRecordSetOri.prvDtableTabla
                                    Select (row(pubintColumnIndex).ToString.Length)).Max()

                        'DefinedSize = Len(LeerDatoADO(ObjRecordSetOri, pubintColumnIndex))
                        'Dim vstrColumnName As String
                        'vstrColumnName = ObjRecordSetOri.Fields(pubintColumnIndex).Name

                        'DefinedSize = ObjRecordSetOri.prvDtableTabla.AsEnumerable().Max(Function(row) row(vstrColumnName).Length)
                    End If


                Else
                    ' se hace el proceso anterior si no tiene datos, si aun asi se cae, devuelve -1
                    Try

                        If ObjRecordSetOri.prvblnAccesoPorNombre Then
                            ' Accesa por nombre de columna
                            vlngSize = ObjRecordSetOri.prvDtableTabla.Columns(pubstrColumnName).MaxLength
                        Else
                            ' Accesa por indice de columna
                            vlngSize = ObjRecordSetOri.prvDtableTabla.Columns(pubintColumnIndex).MaxLength
                        End If
                    Catch ex As Exception
                        vlngSize = -1
                    End Try

                End If

                Return vlngSize
            End Get
        End Property
        Private Function DevolverDataTypeCampo(ByVal pvSTypeTipo As System.Type) As DbType
            Select Case pvSTypeTipo.Name
                Case "Byte"
                    DevolverDataTypeCampo = DbType.Byte
                Case "Decimal"
                    DevolverDataTypeCampo = DbType.Decimal
                Case "Boleean"
                    DevolverDataTypeCampo = DbType.Boolean
                Case "DateTime"
                    DevolverDataTypeCampo = DbType.DateTime
                Case "Double"
                    DevolverDataTypeCampo = DbType.Double
                Case "Int16"
                    DevolverDataTypeCampo = DbType.Int16
                Case "Int32"
                    DevolverDataTypeCampo = DbType.Int32
                Case "Int64"
                    DevolverDataTypeCampo = DbType.Int64
                Case "String"
                    DevolverDataTypeCampo = DbType.String
                Case "Char"
                    DevolverDataTypeCampo = DbType.StringFixedLength
                Case "Guid"
                    DevolverDataTypeCampo = DbType.Guid
                Case "Byte[]"
                    DevolverDataTypeCampo = DbType.Binary
            End Select
        End Function
        ' Por jairc , se coloca como soporte para el campo clsfield
        ' pero no sirve porque en .Net no sabe de precisiones
        Public Property Precision() As Short

            Get
                Return prvshtPrecision

            End Get
            Set(ByVal value As Short)
                prvshtPrecision = value
            End Set
        End Property

        Public Property NumericScale() As Short

            Get
                Return prvshtNumericScale

            End Get
            Set(ByVal value As Short)
                prvshtNumericScale = value
            End Set
        End Property

        Public ReadOnly Property OriginalValue() As Object

            Get
                Dim vobjDRowVersion As DataRow
                Dim vobjValue As Object

                vobjDRowVersion = ObjRecordSetOri.PrvDtViewVista(ObjRecordSetOri.prvlngPosicion).Row
                If ObjRecordSetOri.prvblnAccesoPorNombre Then
                    ' Accesa por nombre de columna
                    vobjValue = vobjDRowVersion(pubstrColumnName, DataRowVersion.Original)
                Else
                    ' Accesa por indice de columna
                    vobjValue = vobjDRowVersion(pubintColumnIndex, DataRowVersion.Original)
                End If
                Return vobjValue
            End Get


        End Property

        Public Property ObjRecordSetOri As clsRecordset
            Get
                If pubobjParentFields Is Nothing Then
                    Return _objRecordSetOri
                Else
                    Return pubobjParentFields.objRecordSetOri
                End If
            End Get
            Set(value As clsRecordset)
                If pubobjParentFields Is Nothing Then
                    _objRecordSetOri = value
                Else
                    ' No hay necesidad de asignarlo porqueya lo tiene el papa
                    'pubobjParentFields.objRecordSetOri = value
                End If

            End Set
        End Property

        ''' <summary>
        ''' Reasigna el valor del maxlenght de la columna de acuerdo al dato enviado, si el dato enviado es mayor reasigna el maxlenght
        ''' Solo lo hace cuando el tipo de dato de la columna es string
        ''' </summary>
        ''' <param name="pvobjValor">Valor a asignar.</param>
        Public Sub ReasignarMaxLength(ByVal pvobjValor As Object)
            ' por jairc 20210621
            ' Se vuelve publico para poder ser llamado desde la truedbgrid para que recalcule la longitud
            On Error Resume Next
            ' Por jairc 20211126
            ' Si el valor es nulo no tiene que hacer ningun recalculo
            If IsDBNull(pvobjValor) Then Exit Sub

            If ObjRecordSetOri.prvblnAccesoPorNombre Then
                ' Accesa por nombre de columna
                Select Case DevolverDataTypeCampo(ObjRecordSetOri.prvDtableTabla.Columns(pubstrColumnName).DataType)
                    Case DbType.String, DbType.StringFixedLength
                        If Len(pvobjValor) > ObjRecordSetOri.pubObjDtTabla.Columns(pubstrColumnName).MaxLength Then
                            ObjRecordSetOri.pubObjDtTabla.Columns(pubstrColumnName).MaxLength = Len(pvobjValor)

                        End If
                        If ObjRecordSetOri.prvblnIsClone Then
                            If Len(pvobjValor) > ObjRecordSetOri.prvObjRecOrigenClone.pubObjDtTabla.Columns(pubstrColumnName).MaxLength Then
                                ObjRecordSetOri.prvObjRecOrigenClone.pubObjDtTabla.Columns(pubstrColumnName).MaxLength = Len(pvobjValor)

                            End If
                        End If
                End Select
            Else
                ' Accesa por indice de columna
                Select Case DevolverDataTypeCampo(ObjRecordSetOri.prvDtableTabla.Columns(pubintColumnIndex).DataType)
                    Case DbType.String, DbType.StringFixedLength
                        If Len(pvobjValor) > ObjRecordSetOri.pubObjDtTabla.Columns(pubintColumnIndex).MaxLength Then
                            ObjRecordSetOri.pubObjDtTabla.Columns(pubintColumnIndex).MaxLength = Len(pvobjValor)
                        End If
                        If ObjRecordSetOri.prvblnIsClone Then
                            If Len(pvobjValor) > ObjRecordSetOri.prvObjRecOrigenClone.pubObjDtTabla.Columns(pubintColumnIndex).MaxLength Then
                                ObjRecordSetOri.prvObjRecOrigenClone.pubObjDtTabla.Columns(pubintColumnIndex).MaxLength = Len(pvobjValor)

                            End If
                        End If
                End Select

            End If

        End Sub
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
            If Not pubobjParentFields Is Nothing Then
                ObjRecordSetOri = Nothing
            End If
        End Sub

        ''' <summary>
        ''' Afecta el dato de la fila en cascada cuando ha sido en creado con Clone
        ''' </summary>
        ''' <param name="probjRecordSetClone">Recordset base</param>
        ''' <param name="pvobjDataRow">DataRow base</param>
        ''' <param name="pvobjvalue">Valor a actualizar</param>
        Private Sub UpdateCloneFromDataRowInCascade(ByRef probjRecordSetClone As UnoEEDatos.clsRecordset, pvobjDataRow As DataRow, pvobjvalue As Object)
            Dim vobjDRowClone As DataRow
            ' Si el recordset que envian es un clone, quiere decir que debe afectar la fila del datatable en que se baso

            If probjRecordSetClone.prvblnIsClone Then
                ' Deduce la fila del recordset en cual se baso par el clone
                vobjDRowClone = probjRecordSetClone.prvObjRecOrigenClone.prvDtableTabla(probjRecordSetClone.prvDtableTabla.Rows.IndexOf(pvobjDataRow))
                If ObjRecordSetOri.prvblnAccesoPorNombre Then
                    vobjDRowClone.Item(pubstrColumnName) = pvobjvalue
                Else
                    vobjDRowClone.Item(pubintColumnIndex) = pvobjvalue
                End If

                If Not probjRecordSetClone.prvObjRecOrigenClone.prvDtableTabla.GetChanges Is Nothing Then
                    probjRecordSetClone.prvObjRecOrigenClone.prvDtableTabla.AcceptChanges()
                End If

                UpdateCloneFromDataRowInCascade(probjRecordSetClone.prvObjRecOrigenClone, vobjDRowClone, pvobjvalue)
            Else

            End If

        End Sub

    End Class

#End Region
#Region "clsFields"
    '########################################################
    '######### Clase especial para usar Fileds como grupo
    '#########################################################
    <Serializable()>
    Public Class clsFields
        Implements IEnumerable
        'Inherits MarshalByRefObject
        Public objRecordSetOri As UnoEEDatos.clsRecordset
        Private _fields As ArrayList
        Private _Item As clsField

        Public Sub New()
            'Se tiene una lista de clsField y luego lo unico q se tiene q hacer es con la implementacion del IEnumerable devolverla
            _fields = New ArrayList()

        End Sub

        Public Sub Inicializar()

            For vintColumna As Integer = 0 To objRecordSetOri.prvDtableTabla.Columns.Count - 1
                Dim objField As clsField
                objField = New clsField
                objField.pubintColumnIndex = vintColumna
                objField.pubstrColumnName = objRecordSetOri.prvDtableTabla.Columns(vintColumna).ColumnName
                objField.pubobjParentFields = Me
                ' objField.objRecordSetOri = objRecordSetOri
                objRecordSetOri.prvblnAccesoPorNombre = False
                _fields.Add(objField)
                objField = Nothing
            Next
        End Sub

        'Default Public ReadOnly Property field(ByVal pvStrColum As String) As clsField
        '    Get

        '    End Get
        'End Property
        Public ReadOnly Property Count() As Integer
            Get
                Count = 0
                Try
                    Count = objRecordSetOri.prvDtableTabla.Columns.Count
                Catch ex As Exception
                End Try
            End Get
        End Property

        'Migracion: req.137891; cpp ; se adiciona el parametro opcional pvAttrib para el manejo de ciertos atributos de la columna. Se adiconan en el momento los que se han encontrado 
        'en vb6. 
        Public Sub Append(ByVal pvstrNombreCampo As String, ByVal pvintDataType As DbType,
                                Optional pvintSize As Integer = 0, Optional pvAttrib As pubEnumFieldAttributeEnum = pubEnumFieldAttributeEnum.adFldUnspecified)
            Dim obcolumn As DataColumn

            obcolumn = New DataColumn()
            obcolumn.ColumnName = pvstrNombreCampo
            obcolumn.DataType = LeerDataTypeCampo(pvintDataType)
            If (pvintSize > 0) Then
                Select Case pvintDataType
                    'cdrb op 194827, propiedad MaxLength DataColumn se omite para las columnas que no son de texto, se borran del case Byte y Binary
                    Case DbType.String, DbType.StringFixedLength
                        obcolumn.MaxLength = pvintSize
                End Select
            End If
            If (pvAttrib > pubEnumFieldAttributeEnum.adFldUnspecified) Then
                Select Case pvAttrib
                    Case pubEnumFieldAttributeEnum.adFldIsNullable, pubEnumFieldAttributeEnum.adFldMayBeNull
                        obcolumn.AllowDBNull = True
                    Case pubEnumFieldAttributeEnum.adFldKeyColumn
                        On Error Resume Next
                        ' Queda pendinte manejo cuando el datatable esta vacio
                        objRecordSetOri.prvDtableTabla.PrimaryKey = New DataColumn() {obcolumn}
                    Case pubEnumFieldAttributeEnum.adFldUpdatable
                        obcolumn.ReadOnly = False
                    Case (pubEnumFieldAttributeEnum.adFldKeyColumn + pubEnumFieldAttributeEnum.adFldUpdatable)
                        On Error Resume Next
                        ' Queda pendinte manejo cuando el datatable esta vacio
                        objRecordSetOri.prvDtableTabla.PrimaryKey = New DataColumn() {obcolumn}
                        obcolumn.ReadOnly = False
                End Select
            End If

            ' Add the Column to the DataColumnCollection.
            objRecordSetOri.prvDtableTabla.Columns.Add(obcolumn)


        End Sub

        Public Sub Append(ByVal pvstrNombreCampo As String, ByVal pvintDataType As DbType)
            Dim obcolumn As DataColumn

            obcolumn = New DataColumn(pvstrNombreCampo, LeerDataTypeCampo(pvintDataType))
            'obcolumn.ColumnName = pvstrNombreCampo
            'obcolumn.DataType = LeerDataTypeCampo(pvintDataType)
            ' Add the Column to the DataColumnCollection.
            objRecordSetOri.prvDtableTabla.Columns.Add(obcolumn)

        End Sub

        Private Function LeerDataTypeCampo(ByVal pvintDataType As DbType) As System.Type
            Select Case pvintDataType
                Case DbType.Binary
                    LeerDataTypeCampo = System.Type.GetType("System.Byte[]")
                Case DbType.Byte
                    LeerDataTypeCampo = System.Type.GetType("System.Byte")
                Case DbType.Boolean
                    LeerDataTypeCampo = System.Type.GetType("System.Boolean")
                Case DbType.Currency, DbType.Decimal
                    LeerDataTypeCampo = System.Type.GetType("System.Decimal")
                Case DbType.Date, DbType.DateTime, DbType.DateTime2
                    LeerDataTypeCampo = System.Type.GetType("System.DateTime")
                Case DbType.Double
                    LeerDataTypeCampo = System.Type.GetType("System.Double")
                Case DbType.Int16
                    LeerDataTypeCampo = System.Type.GetType("System.Int16")
                Case DbType.Int32
                    LeerDataTypeCampo = System.Type.GetType("System.Int32")
                Case DbType.Int64
                    LeerDataTypeCampo = System.Type.GetType("System.Int64")
                Case DbType.String
                    LeerDataTypeCampo = System.Type.GetType("System.String")
                Case DbType.StringFixedLength
                    LeerDataTypeCampo = System.Type.GetType("System.String")
                Case DbType.Guid
                    LeerDataTypeCampo = System.Type.GetType("System.Guid")
                Case DbType.Object
                    LeerDataTypeCampo = System.Type.GetType("System.Object")
                Case Else
                    LeerDataTypeCampo = System.Type.GetType("System.String")
            End Select
        End Function

        Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return (TryCast(_fields, IEnumerable)).GetEnumerator()
        End Function
        'jairc: 17/06/2020 req: 170012
        ''' <summary>
        ''' Trae el campo clsField de la coleccion de fields por nombre
        ''' </summary>
        ''' <param name="pvstrNombreCampo">Nombre del campo</param>
        Public Overloads ReadOnly Property Item(pvstrNombreCampo As String) As clsField
            Get
                Dim vobjFieldReturn As clsField = Nothing
                Dim vobjArrayFields = (From vobjField As clsField In _fields.ToArray
                                       Where vobjField.pubstrColumnName.Equals(pvstrNombreCampo, StringComparison.InvariantCultureIgnoreCase)
                                       Select vobjField).ToArray()

                If vobjArrayFields.Count > 0 Then
                    vobjFieldReturn = vobjArrayFields(0)

                End If
                Return vobjFieldReturn
            End Get

        End Property
        'jairc: 17/06/2020 req: 170012
        ''' <summary>
        ''' Trae el campo clsField de la coleccion de fields por indice
        ''' </summary>
        ''' <param name="pvshIndiceCampo">Indice del campo</param>
        Public Overloads ReadOnly Property Item(pvintIndiceCampo As Integer) As clsField
            Get
                Return _fields(pvintIndiceCampo)
            End Get

        End Property

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
            objRecordSetOri = Nothing
            If _fields IsNot Nothing Then
                _fields.Clear()
            End If
            _fields = Nothing
        End Sub
    End Class
#End Region

    '<Serializable()> _
    'Public Class clsFields
    '    'Inherits MarshalByRefObject
    '    Public objRecordSetOri As UnoEEDatos.clsRecordset

    '    Public ReadOnly Property Count() As Integer
    '        Get
    '            Count = 0
    '            Try
    '                Count = objRecordSetOri.prvDtableTabla.Columns.Count
    '            Catch ex As Exception
    '            End Try
    '        End Get
    '    End Property

    '    Public Sub Append(ByVal pvstrNombreCampo As String, ByVal pvintDataType As DbType, _
    '                            ByVal pvintSize As Integer)
    '        Dim obcolumn As DataColumn

    '        obcolumn = New DataColumn()
    '        obcolumn.ColumnName = pvstrNombreCampo
    '        obcolumn.DataType = LeerDataTypeCampo(pvintDataType)
    '        Select Case pvintDataType
    '            Case DbType.Binary, DbType.Byte, DbType.String, DbType.StringFixedLength
    '                obcolumn.MaxLength = pvintSize
    '        End Select
    '        ' Add the Column to the DataColumnCollection.
    '        objRecordSetOri.prvDtableTabla.Columns.Add(obcolumn)


    '    End Sub
    '    Public Sub Append(ByVal pvstrNombreCampo As String, ByVal pvintDataType As DbType)
    '        Dim obcolumn As DataColumn

    '        obcolumn = New DataColumn()
    '        obcolumn.ColumnName = pvstrNombreCampo
    '        obcolumn.DataType = LeerDataTypeCampo(pvintDataType)
    '        ' Add the Column to the DataColumnCollection.
    '        objRecordSetOri.prvDtableTabla.Columns.Add(obcolumn)

    '    End Sub
    '    Private Function LeerDataTypeCampo(ByVal pvintDataType As DbType) As System.Type
    '        Select Case pvintDataType
    '            Case DbType.Binary, DbType.Byte
    '                LeerDataTypeCampo = System.Type.GetType("System.Byte")
    '            Case DbType.Boolean
    '                LeerDataTypeCampo = System.Type.GetType("System.Boolean")
    '            Case DbType.Currency, DbType.Decimal
    '                LeerDataTypeCampo = System.Type.GetType("System.Decimal")
    '            Case DbType.Date, DbType.DateTime, DbType.DateTime2
    '                LeerDataTypeCampo = System.Type.GetType("System.DateTime")
    '            Case DbType.Double
    '                LeerDataTypeCampo = System.Type.GetType("System.Double")
    '            Case DbType.Int16
    '                LeerDataTypeCampo = System.Type.GetType("System.Int16")
    '            Case DbType.Int32
    '                LeerDataTypeCampo = System.Type.GetType("System.Int32")
    '            Case DbType.Int64
    '                LeerDataTypeCampo = System.Type.GetType("System.Int64")
    '            Case DbType.String
    '                LeerDataTypeCampo = System.Type.GetType("System.String")
    '            Case DbType.StringFixedLength
    '                LeerDataTypeCampo = System.Type.GetType("System.Char")
    '            Case Else
    '                LeerDataTypeCampo = System.Type.GetType("System.String")
    '        End Select
    '    End Function
    'End Class

    Public Sub New()
        Dim vDtsDataset As New DataSet
        prvDtableTabla = New DataTable("table")
        vDtsDataset.Tables.Add(prvDtableTabla)
        prvstrNombreDTable = "table"
        'prvDtViewVista = New DataView(prvDtableTabla)
        'prvDtViewVista.Table = prvDtableTabla       
    End Sub

    '/***************************************************
    'Req. 193492, 194327
    'Se adiciona estos 2 constructores para identificar cuando 
    'un recordset es creado desde el servidor
    'parametro pvBlnDesdeServidor
    '/***************************************************
    Public Sub New(pvBlnDesdeServidor As Boolean)
        Dim vDtsDataset As New DataSet
        prvDtableTabla = New DataTable("table")
        vDtsDataset.Tables.Add(prvDtableTabla)
        prvstrNombreDTable = "table"
        pubBlnCreacionEnServidor = pvBlnDesdeServidor
    End Sub

    Public Sub New(ByVal pvDtsDatos As DataSet, Optional pvBlnDesdeServidor As Boolean = False)
        prvDtableTabla = pvDtsDatos.Tables(0)
        prvstrNombreDTable = prvDtableTabla.TableName
        pubBlnCreacionEnServidor = pvBlnDesdeServidor
        MoveFirst()
    End Sub
    '/***************************************************
    '/***************************************************

    Public Sub New(ByVal pvDttDatos As DataTable)
        prvDtableTabla = pvDttDatos
        If prvDtableTabla.DataSet Is Nothing Then
            Dim vDtsDataset As New DataSet
            vDtsDataset.Tables.Add(prvDtableTabla)
        End If
        prvstrNombreDTable = pvDttDatos.TableName
        'prvDtViewVista = New DataView(prvDtableTabla)
        'prvDtViewVista.Table = prvDtableTabla
        MoveFirst()

    End Sub

    Public Sub New(ByVal pvDtsDatos As DataSet, ByVal pvStrNombreTabla As String)
        prvDtableTabla = pvDtsDatos.Tables(pvStrNombreTabla)
        prvstrNombreDTable = prvDtableTabla.TableName
        'prvDtViewVista = New DataView(prvDtableTabla)
        'prvDtViewVista.Table = prvDtableTabla
        MoveFirst()
    End Sub

    Public Sub New(ByVal pvDtsDatos As DataSet)
        prvDtableTabla = pvDtsDatos.Tables(0)
        prvstrNombreDTable = prvDtableTabla.TableName
        ' prvDtViewVista = New DataView(prvDtableTabla)
        'prvDtViewVista.Table = prvDtableTabla
        MoveFirst()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        If prvDtableTabla IsNot Nothing Then
            prvDtableTabla.Dispose()
        End If
        If _prvDtViewVistaClient IsNot Nothing Then
            _prvDtViewVistaClient.Dispose()
        End If
        If _prvDtViewVistaServer IsNot Nothing Then
            _prvDtViewVistaServer.Dispose()
        End If
        prvDtableTabla = Nothing
        PrvDtViewVista = Nothing

    End Sub


    Private Function FormatearCampoXML(ByVal PvObjDato As Object) As String

        If TypeOf PvObjDato Is String Then
            Return CType(PvObjDato, String)
        ElseIf TypeOf PvObjDato Is Date Then
            Return Format$(CType(PvObjDato, Date), "yyyy-MM-dd HH:mm:ss")

        ElseIf TypeOf PvObjDato Is Decimal Then
            Dim vStrValorTexto As String
            vStrValorTexto = CType(PvObjDato, Decimal).ToString()
            If vStrValorTexto.Contains(",") Then
                vStrValorTexto = vStrValorTexto.Replace(",", ".")
            End If
            FormatearCampoXML = vStrValorTexto
            vStrValorTexto = Nothing
        ElseIf TypeOf PvObjDato Is Double Then
            Dim vStrValorTexto As String
            vStrValorTexto = CType(PvObjDato, Double).ToString()
            If vStrValorTexto.Contains(",") Then
                vStrValorTexto = vStrValorTexto.Replace(",", ".")
            End If
            FormatearCampoXML = vStrValorTexto
            vStrValorTexto = Nothing
        ElseIf TypeOf PvObjDato Is Integer Then
            Return CType(PvObjDato, Integer).ToString
        ElseIf TypeOf PvObjDato Is Short Then
            Return CType(PvObjDato, Short).ToString
        ElseIf TypeOf PvObjDato Is Guid Then
            Return CType(PvObjDato, Guid).ToString
        ElseIf TypeOf PvObjDato Is Long Then
            Return CType(PvObjDato, Long).ToString
        ElseIf TypeOf PvObjDato Is Single Then
            Return CType(PvObjDato, Single).ToString
        ElseIf TypeOf PvObjDato Is Byte() Then
            Return ByteArrayToString(PvObjDato)
        Else
            Err.Raise(9999, " ", "Tipo de dato no implementado para xml")
            FormatearCampoXML = Nothing
        End If
    End Function


    Private Function ByteArrayToString(ByVal vBtyDato() As Byte) As String
        Dim vStrHexResult As String
        'Dim hex As New StringBuilder(ba.Length * 2)
        ''For Each b As Byte In ba
        ''    hex.AppendFormat("{0:x2}", b)
        ''Next b
        ''Return hex.ToString()

        vStrHexResult = BitConverter.ToString(vBtyDato)
        Return vStrHexResult.Replace("-", "")
    End Function
    ''' <summary>
    ''' Se copia de unoeegeneral para conservar la funcionalidad del cadenas iguales
    ''' </summary>
    ''' <param name="strVCadena1"></param>
    ''' <param name="strVCadena2"></param>
    ''' <returns></returns>
    Private Function GblCadenasIguales(ByVal strVCadena1 As String, ByVal strVCadena2 As String) As Boolean
        GblCadenasIguales = (StrComp(UCase$(Trim$(strVCadena1)), UCase$(Trim$(strVCadena2))) = 0)
    End Function
    ''' <summary>
    ''' jairc:Req 165295
    ''' propiedad par darle soporte a la migracion
    '''  </summary>
    ''' <returns></returns>
    Public ReadOnly Property State() As pubEnumObjectStateEnum
        Get
            Try
                If prvstrNombreDTable.Trim.Length <> 0 Then
                    Return pubEnumObjectStateEnum.adStateOpen
                Else
                    Return pubEnumObjectStateEnum.adStateClosed
                End If
            Catch ex As Exception
                Return pubEnumObjectStateEnum.adStateClosed
            End Try

        End Get

    End Property
    ''' <summary>
    ''' Soporte de metodo CancelBatch del recordset
    ''' </summary>
    ''' <returns></returns>
    Public Sub CancelBatch()
        prvDtableTabla.RejectChanges()
        prvintEstado = pubEnumAction.pubEnumActionStandBy
    End Sub

    ''' <summary>
    ''' Esta propiedad es utizada para acceder al la vista solo cuando se llama desde servidor
    ''' </summary>
    ''' <returns></returns>
    Friend Property PrvDtViewVistaServer As DataView
        Get
            If _prvDtViewVistaServer Is Nothing Then
                _prvDtViewVistaServer = New DataView(prvDtableTabla)
            End If
            Return _prvDtViewVistaServer
        End Get
        Set(value As DataView)
            _prvDtViewVistaServer = value
        End Set
    End Property
    '' CODIGO PARA DEVOLVER CAMBIOS DE FILTERSERVER
    'Friend Property PrvDtViewVistaServer As DataView
    '    Get
    '        If _prvDtViewVistaClient Is Nothing Then
    '            _prvDtViewVistaClient = New DataView(prvDtableTabla)
    '        End If
    '        Return _prvDtViewVistaClient
    '    End Get
    '    Set(value As DataView)
    '        _prvDtViewVistaClient = value
    '    End Set
    'End Property
    'Private Property prvlngPosicion As Long
    '    Get
    '        If prvBlnServer Then

    '            Return _prvlngPosicionCliente
    '        Else
    '            Return _prvlngPosicionCliente
    '        End If

    '    End Get
    '    Set(value As Long)
    '        If prvBlnServer Then
    '            _prvlngPosicionCliente = value
    '        Else
    '            _prvlngPosicionCliente = value
    '        End If

    '    End Set
    'End Property

    '' FIN CODIGO PARA DEVOLVER CAMBIOS DE FILTERSERVER

    Friend Property PrvDtViewVista As DataView
        Get
            If prvBlnServer Then
                Return PrvDtViewVistaServer
            Else
                If _prvDtViewVistaClient Is Nothing And Not prvDtableTabla Is Nothing Then
                    _prvDtViewVistaClient = New DataView(prvDtableTabla) '  prvDtableTabla.DefaultView
                End If
                Return _prvDtViewVistaClient
            End If


        End Get
        Set(value As DataView)
            If prvBlnServer Then
                PrvDtViewVistaServer = value
            Else
                _prvDtViewVistaClient = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' Encapsula el funcionamiento de la posicion, esto para el manejo del recordset
    ''' </summary>
    ''' <returns></returns>
    Friend Property prvlngPosicion As Long
        Get
            If prvBlnServer Then

                Return _prvlngPosicionServidor
            Else
                Return _prvlngPosicionCliente
            End If

        End Get
        Set(value As Long)
            If prvBlnServer Then
                _prvlngPosicionServidor = value
            Else
                'lsdt: 17/06/2021 req: 199155 no se debe asignar el valor de nuevo si es igual - esto ya no es necesario.
                'If _prvlngPosicionCliente = value Then Exit Property  lsdt: 22/06/2021 - se comenta linea por el caso de la Op johan.199241 y se puede ejecutar dos veces ya no daña ningun dato. req 199155
                _prvlngPosicionCliente = value

                ' Evento que se llama cada que la posicion en el recordset cambia
                'If _prvlngPosicionCliente <> value Then
                RaiseEvent MoverPosicion()
                'End If
                'If Not pubobjBinding Is Nothing Then
                '    pubobjBinding.Position = value
                'End If

            End If

        End Set
    End Property


    Friend Property PrvBlnTablaFiltrada As Boolean
        Get
            If prvBlnServer Then
                Return _prvBlnTablaFiltradaServidor
            Else
                Return _prvBlnTablaFiltradaCliente
            End If

        End Get
        Set(value As Boolean)
            If prvBlnServer Then
                _prvBlnTablaFiltradaServidor = value
            Else
                _prvBlnTablaFiltradaCliente = value
            End If
        End Set
    End Property

    Friend Property PrvStrFilter As String
        Get
            If prvBlnServer Then
                Return _prvStrFilterServer
            Else
                Return _prvStrFilterClient
            End If

        End Get
        Set(value As String)
            If prvBlnServer Then
                _prvStrFilterServer = value
            Else
                _prvStrFilterClient = value
            End If

        End Set
    End Property

    Friend Property PrvStrSort As String
        Get
            If prvBlnServer Then
                Return _prvStrSortServer
            Else
                Return _prvStrSortClient
            End If

        End Get
        Set(value As String)
            If prvBlnServer Then
                _prvStrSortServer = value
            Else
                _prvStrSortClient = value
            End If

        End Set
    End Property

    Friend Property PrvBlnTablaOrden As Boolean
        Get
            If prvBlnServer Then
                Return _prvBlnTablaOrdenServidor
            Else
                Return _prvBlnTablaOrdenCliente
            End If

        End Get
        Set(value As Boolean)

            If prvBlnServer Then
                _prvBlnTablaOrdenServidor = value
            Else
                _prvBlnTablaOrdenCliente = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' Dispara el evento RecordChangeComplete del recordset
    ''' </summary>
    ''' <param name="adReason"></param>
    ''' <param name="cRecords"></param>
    ''' <param name="pError"></param>
    ''' <param name="adStatus"></param>
    ''' <param name="pRecordset"></param>
    Public Sub DispararEvento_RecordChangeComplete(ByVal adReason As clsRecordset.EventReasonEnumRecordset, ByVal cRecords As Integer, ByVal pError As Data.Common.DbException, adStatus As clsRecordset.EventStatusEnumRecordset, ByVal pRecordset As UnoEEDatos.clsRecordset)
        RaiseEvent RecordChangeComplete(adReason, cRecords, pError, adStatus, pRecordset)
    End Sub

    ''' <summary>
    ''' MDI (Manuel Diaz) req: 199783 Se utiliza para llamar el evento FildChangeComplete
    ''' </summary>
    ''' <returns></returns>
    Public Sub llamarEventoFieldChangeComplete(ByVal cFields As Long, ByRef Fields As Object, ByVal pError As Data.Common.DbException, ByRef adStatus As clsRecordset.EventStatusEnumRecordset, ByVal pRecordset As clsRecordset)
        RaiseEvent FieldChangeComplete(cFields, Fields, pError, adStatus, pRecordset)
    End Sub

    ''' <summary>
    ''' Lee la vista cliente 
    ''' </summary>
    ''' <returns></returns>
    Public Function LeerVistaCliente() As DataView

        If _prvDtViewVistaClient Is Nothing And Not prvDtableTabla Is Nothing Then
            _prvDtViewVistaClient = New DataView(prvDtableTabla) '  prvDtableTabla.DefaultView
        End If
        Return _prvDtViewVistaClient

    End Function

    ''' <summary>
    ''' Lee la vista servidor 
    ''' </summary>
    ''' <returns></returns>
    Public Function LeerVistaServidor() As DataView

        Return _prvDtViewVistaServer


    End Function

    ''' <summary>
    ''' Indica si sincroniza la grilla o no cuando se mueve el recordset
    ''' </summary>
    ''' <returns></returns>
    Private Function SincronizarGrilla() As Boolean
        Dim blnSincronizar As Boolean = False
        blnSincronizar = Not Me.pubobjGrilla Is Nothing
        If blnSincronizar Then
            blnSincronizar = Not (Me.pubobjGrilla.pubIntEventoActivo = 7) ' Evento AfterDelete
        End If
        Return blnSincronizar
    End Function
End Class


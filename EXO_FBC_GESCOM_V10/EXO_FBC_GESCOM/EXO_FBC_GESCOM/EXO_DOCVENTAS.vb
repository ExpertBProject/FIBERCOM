Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_DOCVENTAS
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
        If actualizar Then
            cargaCampos()
        End If
    End Sub

    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            'Pantalla Clientes - Campos
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_QUT1.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDF QUT1.xml", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        '  Case "139", "133", "179", "149", "140", "234234234567", "180", "65303"
                        Case "139", "149", "140", "179", "180", "133"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_Form_VALIDATE_AFTER(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "139", "149", "140", "179", "180", "133"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "139", "149", "140", "179", "180", "133"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "139", "149", "140", "179", "180", "133"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_Form_Load(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        'Dim EXO_Functions As New EXO_BasicDLL.EXO_Generic_Forms_Functions(Me.objGlobal.conexionSAP)
        Dim Path As String = ""
        Dim XmlDoc As New System.Xml.XmlDocument

        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oForm.Freeze(True)

            'Buscar XML de update
            objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'Ponemos a visible el dto.            
            CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Visible = True
            CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Width = 80
            'Ponemos a visible el % comisión          
            CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("28").Visible = True
            CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("28").Width = 80

            'Ponemos a visible el % comisión          
            CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_DTOCAL").Visible = True
            CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_DTOCAL").Width = 80

            EventHandler_Form_Load = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_Form_VALIDATE_AFTER(ByRef pVal As ItemEvent) As Boolean
        EventHandler_Form_VALIDATE_AFTER = False
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ColUID
                Case "1" 'articulo
                    'buscar si tiene limiente de descuento, y marcar en el desplegable si o no
                    sSQL = "SELECT COALESCE(""U_EXO_LIMDTO"",'N') ""U_EXO_LIMDTO""  FROM ""OCRD"" WHERE ""CardCode""='" & CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value.ToString & "' "
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount > 0 Then

                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_APLILIMDTO").Cells.Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Select(oRs.Fields.Item("U_EXO_LIMDTO").Value.ToString)
                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("11").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Active = True
                    End If

                    'Case "15" '%Dto. Controlamos si puede introducir Dto.
                    '    sSQL = "SELECT ""U_EXO_LIMDTO"" FROM ""OCRD"" WHERE ""CardCode""='" & CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value.ToString & "' "
                    '    oRs.DoQuery(sSQL)
                    '    If oRs.RecordCount > 0 Then
                    '        If oRs.Fields.Item("U_EXO_LIMDTO").Value.ToString <> "Y" And CDbl(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value) > 0 Then
                    '            'SboApp.StatusBar.SetText("(EXO) - No es posible otorgar descuentos a este cliente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            'SboApp.MessageBox("No es posible otorgar descuentos a este cliente.")
                    '            'CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = CType(0, String)
                    '        Else
                    '            ValidarDescuentos(oForm, pVal)
                    '            'Valida el grupo de artículo al cual pertenece el artículo seleccionado
                    '            'Comparará el porcentaje de descuento indicado en el documento contra la información registrada en el objeto “Descuentos y comisiones”.
                    '            'Basándose en el descuento otorgado al artículo, el sistema identificará el porcentaje de comisión correspondiente.
                    '            'SAP escribirá el porcentaje de comisión en el campo % de comisión (Commission).
                    '        End If
                    '    Else
                    '        SboApp.StatusBar.SetText("(EXO) - Error inesperado. No se encuentra el interlocutor.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        SboApp.MessageBox("Error inesperado. No se encuentra el interlocutor.")
                    '        CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Active = True
                    '        Exit Function
                    '    End If

            End Select


            CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Visible = True
            CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Width = 80

            EventHandler_Form_VALIDATE_AFTER = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_FormDataEvent(infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "139", "149", "140"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                ''Antes de actualizar comprobamos  los datos.
                                'If ComprobarDescuentos(oForm) = False Then
                                '    Return False
                                'End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                ''Antes de añadir comprobamos  los datos.
                                'If ComprobarDescuentos(oForm) = False Then
                                '    Return False
                                'End If
                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "139", "149", "140"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select
                End Select
            End If

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sCancel As String = ""
        Dim Status As String = ""
        Dim sTabla As String = ""
        EventHandler_ItemPressed_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            'TODO Quitar - Descomentar para volver atrás
            If pVal.ItemUID = "1" Then

                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    'TODO Quitar - Descomentar para volver atrás

                    objGlobal.SBOApp.SetStatusBarMessage("..Comprobando Descuentos y Comsiones..", BoMessageTime.bmt_Short, False)
                    If oForm.TypeEx = "140" OrElse oForm.TypeEx = "133" OrElse oForm.TypeEx = "179" OrElse oForm.TypeEx = "180" OrElse oForm.TypeEx = "149" OrElse oForm.TypeEx = "139" Then
                        Select Case oForm.TypeEx
                            Case "140"
                                sTabla = "ODLN"
                            Case "133"
                                sTabla = "OINV"
                            Case "179"
                                sTabla = "ORIN"
                            Case "180"
                                sTabla = "ORDN"
                            Case "149"
                                sTabla = "OQUT"
                            Case "139"
                                sTabla = "ORDR"
                        End Select
                        sCancel = oForm.DataSources.DBDataSources.Item("" & sTabla & "").GetValue("CANCELED", 0).Trim
                        Status = oForm.DataSources.DBDataSources.Item("" & sTabla & "").GetValue("DocStatus", 0).Trim
                    Else
                        sCancel = "C"
                        Status = "C"
                    End If
                    If sCancel <> "C" Then
                        If Status <> "C" Then
                            'si el documento se está actualizando, sólo entrar si es oferta y pedido.
                            If ComprobarDescuentos(oForm) = False Then
                                Exit Function
                            End If
                        End If

                    End If

                End If

                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objGlobal.SBOApp.SetStatusBarMessage("..Comprobando Descuentos y Comsiones..", BoMessageTime.bmt_Short, False)
                    If oForm.TypeEx = "149" OrElse oForm.TypeEx = "139" Then
                        Select Case oForm.TypeEx
                            Case "139"
                                sTabla = "ORDR"
                            Case "149"
                                sTabla = "OQUT"

                        End Select
                        sCancel = oForm.DataSources.DBDataSources.Item("" & sTabla & "").GetValue("CANCELED", 0).Trim
                        Status = oForm.DataSources.DBDataSources.Item("" & sTabla & "").GetValue("DocStatus", 0).Trim
                    Else
                        sCancel = "C"
                        Status = "C"
                    End If
                    If sCancel <> "C" Then
                        If Status <> "C" Then
                            'si el documento se está actualizando, sólo entrar si es oferta y pedido.
                            If ComprobarDescuentos(oForm) = False Then
                                Exit Function
                            End If
                        End If

                    End If

                End If


            End If

            EventHandler_ItemPressed_Before = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function ComprobarDescuentos(ByRef oForm As SAPbouiCOM.Form) As Boolean

        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing

        Dim sGrupo As String = ""
        Dim dDto As Double = 0
        Dim dDtoGlobal As Double = 0
        Dim dDtoLin As Double = 0

        'variable para guardar el precio antes de descuento global y despues de desc. linea
        Dim dblPrecioDtoLinea As Double = 0

        'variable para guardar el precio con descuento global
        Dim dblPrecioDtoLineaTotal As Double = 0

        'variable para guardar por linea precio*cantidad
        Dim dblImpLinSinDto As Double = 0

        'variable para calcular el importe de descuento global añadido por línea
        Dim dblImpComisionLinea As Double = 0

        'Precio despues del descuento line + global.
        Dim dblPrecioDespDescGlobal As Double = 0

        'cantidad
        Dim dblCan As Double = 0

        'precio
        Dim dblPrecio As Double = 0

        ComprobarDescuentos = False
        Try
            oForm.Freeze(True)
            sSQL = "SELECT ""U_EXO_LIMDTO"" FROM ""OCRD"" WHERE ""CardCode""='" & CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value.ToString & "' "
            oRs.DoQuery(sSQL)

            Dim statusDoc As String = CType(oForm.Items.Item("81").Specific, SAPbouiCOM.ComboBox).Value.ToString
            ' statusDoc = 1 --> Abiertos
            ' statusDoc = 2 --> Abrir - Impreso
            ' statusDoc = 3 --> Cerrado
            'Dim sCancel As String = ""
            'sCancel = oForm.DataSources.DBDataSources.Item("ODLN").GetValue("CANCELED", 0).Trim
            If oRs.RecordCount > 0 Then
                '((oRs.Fields.Item("U_EXO_LIMDTO").Value.ToString = "Y")  esto lo quito, y ahora se calcula en funcion de las lineas
                If (statusDoc = "1" Or statusDoc = "2") Then

                    'dblTotSinDtoGlobal = CDbl(CType(oForm.Items.Item("22").Specific, SAPbouiCOM.EditText).String.Substring(0, CInt(CType(oForm.Items.Item("22").Specific, SAPbouiCOM.EditText).String.Length - 4)).Replace(".", ""))
                    'If CType(oForm.Items.Item("42").Specific, SAPbouiCOM.EditText).String <> "" Then
                    ' dblImpDtoTotalGlobal = CDbl(CType(oForm.Items.Item("42").Specific, SAPbouiCOM.EditText).String.Substring(0, CInt(CType(oForm.Items.Item("42").Specific, SAPbouiCOM.EditText).String.Length - 4)).Replace(".", ""))
                    'Else
                    '   dblImpDtoTotalGlobal = 0
                    'End If

                    ' Recorrer lineas

                    oForm.PaneLevel = 1

                    For i As Integer = 1 To CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).RowCount - 1
                        'si la cantidad del documento es difernte de la open qty, no entro por aqui
                        'column 11 cantidad
                        'column 32 cantidad pendiente
                        'column 257 tipo linea
                        If CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("11").Cells.Item(i).Specific, SAPbouiCOM.EditText).String = CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("32").Cells.Item(i).Specific, SAPbouiCOM.EditText).String And CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("257").Cells.Item(i).Specific, SAPbouiCOM.ComboBox).Value = "" Then

                            sSQL = "SELECT ""ItmsGrpCod"" FROM ""OITM"" WHERE ""ItemCode""='" & CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("1").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value & "' "
                            oRs.DoQuery(sSQL)
                            If oRs.RecordCount > 0 Then
                                sGrupo = oRs.Fields.Item("ItmsGrpCod").Value.ToString
                            End If

                            sSQL = "SELECT T0.""Code"", T0.""U_EXO_MIN"", T0.""U_EXO_MAX"", T1.""U_EXO_DTOD"", T1.""U_EXO_DTOH"", T1.""U_EXO_COM"" " _
                            & " FROM ""@EXO_DTCOMC""  T0 " _
                            & " INNER JOIN ""@EXO_DTOCOML""  T1 ON T1.""Code"" = T0.""Code"" " _
                            & " where  T0.""Code""='" & sGrupo & "'"
                            oRs.DoQuery(sSQL)
                            'si hay valor, y la linea está marcada como limitacion de descuentos, entro al proceso y sino todo a 0
                            If oRs.RecordCount > 0 And CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_APLILIMDTO").Cells.Item(i).Specific, SAPbouiCOM.ComboBox).Value = "Y" Then
                                If CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Cells.Item(i).Specific, SAPbouiCOM.EditText).String <> "" Then
                                    dDtoLin = CDbl(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Cells.Item(i).Specific, SAPbouiCOM.EditText).String.Replace(".", ""))
                                Else
                                    dDtoLin = 0
                                End If
                                If CType(oForm.Items.Item("24").Specific, SAPbouiCOM.EditText).String <> "" Then
                                    dDtoGlobal = CDbl(CType(oForm.Items.Item("24").Specific, SAPbouiCOM.EditText).String.Replace(".", ""))
                                End If

                                dblCan = CDbl(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("11").Cells.Item(i).Specific, SAPbouiCOM.EditText).String.Replace(".", ""))

                                'columna 14 precio
                                If CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("14").Cells.Item(i).Specific, SAPbouiCOM.EditText).String <> "" Then

                                    dblPrecio = CDbl(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("14").Cells.Item(i).Specific, SAPbouiCOM.EditText).String.Substring(0, CInt(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("14").Cells.Item(i).Specific, SAPbouiCOM.EditText).String.Length - 4)).Replace(".", ""))


                                    dblPrecioDtoLinea = dblPrecio - ((dblPrecio * dDtoLin) / 100)
                                    dblPrecioDtoLineaTotal = dblPrecioDtoLinea - ((dblPrecioDtoLinea * dDtoGlobal) / 100)

                                    dDto = 100 - ((dblPrecioDtoLineaTotal * 100) / dblPrecio)
                                    dDto = Math.Round(dDto, 2)
                                    'redondear el descuento

                                    'Me.SboApp.MessageBox("dDto: " & dDto)

                                    sSQL = "SELECT T0.""Code"", T0.""U_EXO_MIN"", T0.""U_EXO_MAX"", T1.""U_EXO_DTOD"", T1.""U_EXO_DTOH"", T1.""U_EXO_COM"" " _
                                    & " FROM ""@EXO_DTCOMC""  T0 " _
                                    & " INNER JOIN ""@EXO_DTOCOML""  T1 ON T1.""Code"" = T0.""Code"" " _
                                    & " where  T0.""Code""='" & sGrupo & "' and " & dDto.ToString.Replace(",", ".") & " BETWEEN  ""U_EXO_MIN"" AND T0.""U_EXO_MAX"" "
                                    oRs.DoQuery(sSQL)
                                    If oRs.RecordCount > 0 Then
                                        CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).Active = True
                                        CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).Value = CType(0, String)
                                        'si no hay descuento globlal, pasamos a obtener los descuento por linea
                                        oXml.LoadXml(oRs.GetAsXML())
                                        oNodes = oXml.SelectNodes("//row")

                                        For j As Integer = 0 To oNodes.Count - 1
                                            oNode = oNodes.Item(j)
                                            If dDto >= CDbl(oNode.SelectSingleNode("U_EXO_DTOD").InnerText.Replace(".", ",")) And dDto <= CDbl(oNode.SelectSingleNode("U_EXO_DTOH").InnerText.Replace(".", ",")) Then
                                                Try
                                                    CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).Active = True
                                                    CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).String = CType(CDbl(oNode.SelectSingleNode("U_EXO_COM").InnerText.Replace(".", ",")), String)
                                                    CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).Active = True
                                                    CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).String = CType(CDbl(dDto.ToString.Replace(".", ",")), String)
                                                    CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).Active = False
                                                    CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).Active = False

                                                    dblPrecioDespDescGlobal = dblPrecio - (dblPrecio * dDto / 100)

                                                    dblImpLinSinDto = dblCan * dblPrecioDespDescGlobal

                                                    'Me.SboApp.MessageBox("Lineas: " & i & " dblCan: " & dblCan & " - dblPrecio: " & dblPrecio & " - dblImpLinSinDto: " & dblImpLinSinDto)
                                                    'dblImpComisionLinea = dblImpLinSinDto * CDbl(dDto.ToString.Replace(".", ",")) / 100
                                                    dblImpComisionLinea = dblImpLinSinDto * CDbl(oNode.SelectSingleNode("U_EXO_COM").InnerText.Replace(".", ",")) / 100
                                                    'si el documento es una devolución o un abono, multiplicar por -1
                                                    If oForm.TypeEx = "180" OrElse oForm.TypeEx = "179" Then
                                                        dblImpComisionLinea = dblImpComisionLinea * -1
                                                    End If

                                                    CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_IMPCOM").Cells.Item(i).Specific, EditText).String = CType(CDbl(dblImpComisionLinea.ToString.Replace(".", ",")), String)

                                                    Exit For
                                                Catch ex As Exception

                                                End Try
                                            Else
                                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).Active = True
                                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).String = CType(0, String)
                                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).Active = True
                                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_IMPCOM").Cells.Item(i).Specific, EditText).Active = True
                                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).String = CType(0, String)
                                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_IMPCOM").Cells.Item(i).Specific, EditText).String = CType(0, String)
                                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).Active = False
                                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).Active = False
                                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_IMPCOM").Cells.Item(i).Specific, EditText).Active = False
                                            End If
                                        Next
                                    Else
                                        objGlobal.SBOApp.MessageBox("Lineas:      " & i & " - El % Dto no cumple o no está dado de alta en las Comisiones-Descuentos por Familia" & " Desc. Linea: " & dDtoLin & " Desc. global: " & dDtoGlobal & " Decs. Total: " & dDto)
                                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value = CType(0, String)
                                        CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).Active = True
                                        CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).String = CType(0, String)
                                        CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).Active = True
                                        CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_IMPCOM").Cells.Item(i).Specific, EditText).Active = True
                                        CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).String = CType(0, String)
                                        CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_IMPCOM").Cells.Item(i).Specific, EditText).String = CType(0, String)
                                        CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).Active = False
                                        CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).Active = False
                                        CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_IMPCOM").Cells.Item(i).Specific, EditText).Active = False
                                        CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("15").Cells.Item(i).Specific, EditText).Active = True
                                        Exit Function
                                    End If

                                    'If dDtoGlobal <> 0 Then 

                                    '    dblCan = CDbl(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("11").Cells.Item(i).Specific, SAPbouiCOM.EditText).String.Replace(".", ""))
                                    '    dblPrecio = CDbl(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("14").Cells.Item(i).Specific, SAPbouiCOM.EditText).String.Substring(0, CInt(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("14").Cells.Item(i).Specific, SAPbouiCOM.EditText).String.Length - 4)).Replace(".", ""))
                                    '    dblImpLinSinDto = dblCan * dblPrecio
                                    '    dblImpLinDto = (dblImpLinSinDto * CDbl(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Cells.Item(i).Specific, SAPbouiCOM.EditText).String.Replace(".", ""))) / 100
                                    '    dblImpLinConDto = CDbl(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("21").Cells.Item(i).Specific, SAPbouiCOM.EditText).String.Substring(0, CInt(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("21").Cells.Item(i).Specific, SAPbouiCOM.EditText).String.Length - 4)).Replace(".", ""))
                                    '    dblImpDtoReparto = dblImpLinConDto * dblImpDtoTotalGlobal / dblTotSinDtoGlobal
                                    '    dblPorDtoCalculado = ((dblImpLinDto + dblImpDtoReparto) * 100) / dblImpLinSinDto

                                    '    'este descuento calculado será el que tenemos que controlar que esté dentro del minimo y máximo, y obtendremos el % de comision
                                    '    sSQL = "SELECT T0.""Code"", T0.""U_EXO_MIN"", T0.""U_EXO_MAX"", T1.""U_EXO_DTOD"", T1.""U_EXO_DTOH"", T1.""U_EXO_COM"" " _
                                    '    & " FROM ""@EXO_DTCOMC""  T0 " _
                                    '    & " INNER JOIN ""@EXO_DTOCOML""  T1 ON T1.""Code"" = T0.""Code"" " _
                                    '    & " where  T0.""Code""='" & sGrupo & "' and " & dDto & " BETWEEN  ""U_EXO_MIN"" AND T0.""U_EXO_MAX"" "
                                    '    oRs.DoQuery(sSQL)
                                    '    If oRs.RecordCount > 0 Then
                                    '        'si no hay descuento globlal, pasamos a obtener los descuento por linea
                                    '        oXml.LoadXml(oRs.GetAsXML())
                                    '        oNodes = oXml.SelectNodes("//row")
                                    '        CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).Value = CType(0, String)
                                    '        For j As Integer = 0 To oNodes.Count - 1
                                    '            oNode = oNodes.Item(j)
                                    '            If CDbl(dblPorDtoCalculado.ToString.Replace(".", ",")) >= CDbl(oNode.SelectSingleNode("U_EXO_DTOD").InnerText.Replace(".", ",")) And CDbl(dblPorDtoCalculado.ToString.Replace(".", ",")) <= CDbl(oNode.SelectSingleNode("U_EXO_DTOH").InnerText.Replace(".", ",")) Then
                                    '                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).String = CType(CDbl(oNode.SelectSingleNode("U_EXO_COM").InnerText.Replace(".", ",")), String)
                                    '                dblPorDtoCalculado = Math.Round(dblPorDtoCalculado, 2)

                                    '                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).String = CType(CDbl(dblPorDtoCalculado.ToString.Replace(".", ",")), String)
                                    '            End If
                                    '        Next

                                    '    Else
                                    '        Me.SboApp.MessageBox("El % Dto no cumple 44 o no está dado de alta en las Comisiones-Descuentos por Familia")
                                    '        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Cells.Item(i).Specific, SAPbouiCOM.EditText).String = CType(0, String)
                                    '        Exit Function
                                    '    End If
                                    'End If
                                End If
                            Else
                                'todo a 0 comsion de sap, y campos propios

                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).Active = True
                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).String = CType(0, String)
                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).Active = True
                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_IMPCOM").Cells.Item(i).Specific, EditText).Active = True
                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).String = CType(0, String)
                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_IMPCOM").Cells.Item(i).Specific, EditText).String = CType(0, String)
                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).Active = False
                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).Active = False
                                CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_IMPCOM").Cells.Item(i).Specific, EditText).Active = False

                            End If
                        End If
                    Next
                Else
                    'recorrer lineas y poner la comisión a 0

                    For i As Integer = 1 To CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).RowCount - 1
                        If CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("257").Cells.Item(i).Specific, SAPbouiCOM.ComboBox).Value = "" Then

                            CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("15").Cells.Item(i).Specific, EditText).Active = True
                            CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value = CType(0, String)
                            CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).Active = True
                            CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).String = CType(0, String)
                            CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).Active = True
                            CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_IMPCOM").Cells.Item(i).Specific, EditText).Active = True
                            CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).String = CType(0, String)
                            CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_IMPCOM").Cells.Item(i).Specific, EditText).String = CType(0, String)
                            CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(i).Specific, EditText).Active = False
                            CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(i).Specific, EditText).Active = False
                            CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_IMPCOM").Cells.Item(i).Specific, EditText).Active = False
                            CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("15").Cells.Item(i).Specific, EditText).Active = True
                        End If
                    Next

                    'silvia comento esto, si está en N, les dejamos hacer lo que quieran
                    'If CType(oForm.Items.Item("24").Specific, SAPbouiCOM.EditText).String <> "" Then
                    '    If (CDbl(CType(oForm.Items.Item("24").Specific, SAPbouiCOM.EditText).String.Replace(".", "")) > 0) Then
                    '        SboApp.StatusBar.SetText("(EXO) - No es posible otorgar descuentos a este cliente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        CType(oForm.Items.Item("24").Specific, SAPbouiCOM.EditText).Value = CType(0, String)
                    '        SboApp.MessageBox("No es posible otorgar descuentos a este cliente.")
                    '        Exit Function
                    '    End If
                    'End If
                End If
            End If
            ComprobarDescuentos = True
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try


    End Function
    Private Sub ValidarDescuentos(ByRef oForm As SAPbouiCOM.Form, ByRef pVal As ItemEvent)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sGrupo As String = ""
        Dim dDto As Double = 0
        Dim dDtoGlobal As Double = 0
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing

        Try
            sSQL = "SELECT ""U_EXO_LIMDTO"" FROM ""OCRD"" WHERE ""CardCode""='" & CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value.ToString & "' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                If oRs.Fields.Item("U_EXO_LIMDTO").Value.ToString = "Y" Then
                    'no descuentos
                Else
                    sSQL = "SELECT ""ItmsGrpCod"" FROM ""OITM"" WHERE ""ItemCode""='" & CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value & "' "
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount > 0 Then
                        sGrupo = oRs.Fields.Item("ItmsGrpCod").Value.ToString
                    End If

                    If CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).String <> "" Then
                        dDto = CDbl(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).String.Replace(".", ""))
                    Else
                        dDto = 0
                    End If
                    If CType(oForm.Items.Item("24").Specific, SAPbouiCOM.EditText).String <> "" Then
                        dDtoGlobal = CDbl(CType(oForm.Items.Item("24").Specific, SAPbouiCOM.EditText).String.Replace(".", ""))
                    End If

                    sSQL = "SELECT T0.""Code"", T0.""U_EXO_MIN"", T0.""U_EXO_MAX"", T1.""U_EXO_DTOD"", T1.""U_EXO_DTOH"", T1.""U_EXO_COM"" " _
                    & " FROM ""@EXO_DTCOMC""  T0 " _
                    & " INNER JOIN ""@EXO_DTOCOML""  T1 ON T1.""Code"" = T0.""Code"" " _
                    & " where  T0.""Code""='" & sGrupo & "' and " & dDto & " BETWEEN  ""U_EXO_MIN"" AND T0.""U_EXO_MAX"" "
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount > 0 Then
                        'si no hay descuento globlal, pasamos a obtener los descuento por linea
                        oXml.LoadXml(oRs.GetAsXML())
                        oNodes = oXml.SelectNodes("//row")
                        If dDtoGlobal = 0 Then
                            For i As Integer = 0 To oNodes.Count - 1
                                oNode = oNodes.Item(i)
                                If dDto >= CDbl(oNode.SelectSingleNode("U_EXO_DTOD").InnerText.Replace(".", ",")) And dDto <= CDbl(oNode.SelectSingleNode("U_EXO_DTOH").InnerText.Replace(".", ",")) Then
                                    CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("28").Cells.Item(pVal.Row).Specific, EditText).String = CType(CDbl(oNode.SelectSingleNode("U_EXO_COM").InnerText.Replace(".", ",")), String)
                                    CType(CType(oForm.Items.Item("38").Specific, Matrix).Columns.Item("U_EXO_DTOCAL").Cells.Item(pVal.Row).Specific, EditText).String = CType(CDbl(dDto.ToString.Replace(".", ",")), String)
                                End If
                            Next
                        Else


                        End If
                    Else
                        objGlobal.SBOApp.MessageBox("El % Dto no cumple 55o no está dado de alta en las Comisiones-Descuentos por Familia")
                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("15").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).String = CType(0, String)
                        Exit Sub
                    End If
                End If



            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
End Class

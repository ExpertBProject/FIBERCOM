Imports System.Xml
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports SAPbouiCOM
Imports EXO_FBC_GESCOM.Extensions

Public Class EXO_SELCOM
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaCampos()
        End If
        cargamenu()

    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults

        If objGlobal.SBOApp.Menus.Exists("EXO-MnGCOM") = True Then
            Path = objGlobal.refDi.OGEN.pathGeneral & "\02.Menus"
            If Path <> "" Then
                If IO.File.Exists(Path & "\MnGCOM.png") = True Then
                    objGlobal.SBOApp.Menus.Item("EXO-MnGCOM").Image = Path & "\MnGCOM.png"
                End If
            End If
        End If
    End Sub
    Private Sub cargaCampos()
        'TEST MANU  
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            'Pantalla Clientes - Campos
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_DTCOM.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDO_EXO_DTCOM", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
#Region "Eventos"


    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "EXO-SELCOM"
                        'Cargamos UDO
                        OpenForm(objGlobal)
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

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
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_SELCOM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    'If EventHandler_VALIDATE_After(infoEvento) = False Then
                                    '    GC.Collect()
                                    '    Return False
                                    'End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_SELCOM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                    If EventHandler_Matrix_Link_Press_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_SELCOM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_Form_Visible(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_SELCOM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

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



    Public Overrides Function SBOApp_FormDataEvent(infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "EXO_SELCOM"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                'Antes de actualizar comprobamos  los datos.
                                'If ComprobarDatos(oForm) = False Then
                                '    Return False
                                'End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                'Antes de añadir comprobamos  los datos.
                                'If ComprobarDatos(oForm) = False Then
                                '    Return False
                                'End If
                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "EXO_SELCOM"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select
                End Select
            End If

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
    Private Function EventHandler_Form_Visible(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing

        EventHandler_Form_Visible = False

        Try

            If pVal.ActionSuccess = True Then
                'Recuperar el formulario
                oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
                If oForm.Visible = True Then
                    If CargarComboAgente(oForm) = False Then
                        Exit Function
                    End If

                End If
            End If


            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "btVer" Then
                If pVal.ActionSuccess = True Then
                    'cargarGrid
                    'comprobar si ha metido fechas

                    If CType(oForm.Items.Item("2_U_E").Specific, SAPbouiCOM.EditText).Value = "" Then
                        objGlobal.SBOApp.MessageBox("Antes de consultar los documentos a enviar debe seleccionar una fecha desde.")
                        Exit Function
                    End If

                    If CType(oForm.Items.Item("3_U_E").Specific, SAPbouiCOM.EditText).Value = "" Then
                        objGlobal.SBOApp.MessageBox("Antes de consultar los documentos a enviar debe seleccionar una fecha hasta.")
                        Exit Function
                    End If

                    CargarGrid(objGlobal, oForm)
                End If
                'al acabar el proceso atualizo los datos 


            End If
            If pVal.ItemUID = "btCom" Then
                If pVal.ActionSuccess = True Then
                    If CType(oForm.Items.Item("5_U_E").Specific, SAPbouiCOM.EditText).Value = "" Then
                        objGlobal.SBOApp.MessageBox("La fecha de cobro de comisión no puede estar vacía")
                        CType(oForm.Items.Item("5_U_E").Specific, SAPbouiCOM.EditText).Active = True
                        Exit Function
                    End If
                    If CType(oForm.Items.Item("Check_0").Specific, SAPbouiCOM.CheckBox).Checked = True Then
                        objGlobal.SBOApp.MessageBox("Para asignar los cobros de comisión debe desmarcar ""Ver Comisiones aplicadas"" ")
                        Exit Function
                    End If
                    For i As Integer = 0 To oForm.DataSources.DataTables.Item("DT_GR").Rows.Count - 1
                        If oForm.DataSources.DataTables.Item("DT_GR").GetValue("Sel", i).ToString = "Y" Then
                            'actualizar la fecha cobro, la aplicación de comisión y el importe de la comisión
                            TratarFactura(objGlobal, oForm.DataSources.DataTables.Item("DT_GR"), i, oForm)

                        End If
                    Next
                    objGlobal.SBOApp.MessageBox("Comisiones aplicadas correctamente ")
                    CargarGrid(objGlobal, oForm)
                End If
            End If

            If pVal.ItemUID = "btImp" Then
                GenerarPDF(oForm, "Comisiones.rpt")
            End If

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function EventHandler_Matrix_Link_Press_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_Matrix_Link_Press_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "EXO_GR" Then
                If pVal.ColUID = "DocEntry" Then
                    CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item("DocEntry"), SAPbouiCOM.EditTextColumn).LinkedObjectType = CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).DataTable.GetValue("ObjType", pVal.Row).ToString
                End If
            End If

            EventHandler_Matrix_Link_Press_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
#End Region
#Region "Auxiliares"


    Public Function OpenForm(ByRef OGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        Dim res As Boolean = True

        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim ficheroPantalla As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)


        'abrir formulario
        oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
        oFP.XmlData = EXO_Xml.LoadFormXml(objGlobal.leerEmbebido(GetType(EXO_SELCOM), "EXO_SELCOM.srf"), True).ToString

        Try
            oForm = OGlobal.SBOApp.Forms.AddEx(oFP)

        Catch ex As Exception
            If ex.Message.StartsWith("Form - already exists") = True Then
                OGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
            End If
        End Try

        oForm.Visible = True
        CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.ComboBox).Active = True
        Return res

    End Function

    Private Function CargarComboAgente(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim sSQL As String = ""
        CargarComboAgente = False

        Try
            sSQL = "SELECT ""SlpCode"", ""SlpName""  FROM ""OSLP"" Order by ""SlpCode"" "

            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("1_U_E").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

            CargarComboAgente = True

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Sub CargarGrid(ByRef OGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form)

        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        'Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim strFechaD As String = ""
        Dim strFechaH As String = ""
        Dim strMarcar As String = ""
        Dim strAgenteD As String = ""
        Dim strAgenteH As String = ""
        Dim strCobradas As String = ""

        Try


            'cargar consulta datos formulario edi
            oRs = CType(OGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            strFechaD = CType(oForm.Items.Item("2_U_E").Specific, SAPbouiCOM.EditText).Value
            strFechaH = CType(oForm.Items.Item("3_U_E").Specific, SAPbouiCOM.EditText).Value
            If CType(oForm.Items.Item("Check_0").Specific, SAPbouiCOM.CheckBox).Checked = True Then
                strCobradas = "Y"
            Else
                strCobradas = "N"
            End If

            strAgenteD = CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.ComboBox).Value
            strAgenteH = CType(oForm.Items.Item("1_U_E").Specific, SAPbouiCOM.ComboBox).Value
            ' (t1.""LineTotal"" * t1.""Commission"")/100
            'en esta select relacionar con destiono 14 y restar
            'voy a añadir esto INNER  JOIN "OBOE" T3 ON T3."BoeNum" = t1."BOENumber" and t3."BoeStatus"='P' 
            'hacer un max de esas dos tablas, order de bot1
            sSQL = " Select * FROM " _
            & " (SELECT 'N'  ""Sel"", t0.""DocEntry"", t0.""DocNum"", t0.""TaxDate"",COALESCE(T30.""TaxDate"", MAX(T55.""ReconDate"")) ""FechaUltCobro"", t2.""ItmsGrpCod"", t0.""CardCode"" ,t0.""CardName"" ,t1.""LineNum"", t1.""ItemCode"", t1.""Dscription"", t1.""DiscPrcnt"" ," _
            & " t1.""LineTotal"",t1.""Commission"",t1.""U_EXO_DTOCAL"", T1.""U_EXO_IMPCOM"" + COALESCE(T19.""U_EXO_IMPCOM"",0) ""Importe"", t1.""U_EXO_FECCOM"",t0.""ObjType"", T0.""PaidToDate"" ,T0.""DocTotal"", t0.""SlpCode"",T1.""U_EXO_IMPCOM"",t3.""SlpName""," _
            & " COALESCE(t5.""DocNum"",'0') ""NumOfertaDire"", " _
            & " COALESCE(t11.""DocNum"",'0') ""NumOfertaPed"", COALESCE(t13.""DocNum"",'0') ""NumOfertaAlb"",COALESCE(t9.""DocNum"",'0') ""NumAlb"",COALESCE(t18.""DocNum"",'0') ""NumOfertaTodo"" " _
            & " FROM ""OINV"" t0  " _
            & " INNER Join ""INV1"" t1 ON t0.""DocEntry"" = t1.""DocEntry""  " _
            & " INNER Join ""OITM"" t2 on t1.""ItemCode"" = t2.""ItemCode""  " _
            & " INNER Join ""OSLP"" t3 on t0.""SlpCode"" = t3.""SlpCode"" " _
            & " Left Join ""INV6"" T33 On T0.""DocEntry"" = T33.""DocEntry""    " _
            & " Left Join (Select MAX(""ReconNum"")As ""ReconNum"", ""SrcObjTyp"" , ""SrcObjAbs""   from ""ITR1"" group by ""SrcObjTyp"" , ""SrcObjAbs"" ) T44  " _
            & " On  T44.""SrcObjTyp"" = 13 And T44.""SrcObjAbs"" = T0.""DocEntry"" " _
            & " Left Join ""OITR"" T55 On t44.""ReconNum"" = T55.""ReconNum"" and T55.""Canceled"" = 'N'  " _
            & " Left OUTER JOIN  ""QUT1"" t4 On t1.""BaseEntry"" = t4.""DocEntry"" And T1.""BaseType"" = t4.""ObjType"" And  T1.""BaseLine"" = t4.""LineNum"" " _
            & " Left OUTER JOIN ""OQUT"" t5 On t4.""DocEntry"" = t5.""DocEntry"" " _
            & " Left OUTER JOIN ""RDR1"" t6 On t1.""BaseEntry"" = t6.""DocEntry"" And T1.""BaseType"" = t6.""ObjType"" And  T1.""BaseLine"" = t6.""LineNum"" " _
            & " Left OUTER JOIN ""ORDR"" t7 On t6.""DocEntry"" = t7.""DocEntry""  " _
            & " Left OUTER JOIN ""QUT1"" t12 On t6.""BaseEntry"" = t12.""DocEntry"" And T6.""BaseType"" = t12.""ObjType"" And T6.""BaseLine"" = t12.""LineNum"" " _
            & " Left OUTER JOIN ""OQUT"" t13 On t13.""DocEntry"" = t12.""DocEntry""  " _
            & " Left OUTER JOIN ""DLN1"" t8 On t1.""BaseEntry"" = t8.""DocEntry"" And T1.""BaseType"" = t8.""ObjType"" And  T1.""BaseLine"" = t8.""LineNum"" " _
            & " Left OUTER JOIN ""ODLN"" t9 On t8.""DocEntry"" = t9.""DocEntry""  " _
            & " Left OUTER JOIN ""QUT1"" t10 On t8.""BaseEntry"" = t10.""DocEntry"" And T8.""BaseType"" = t10.""ObjType"" And  T8.""BaseLine"" = t10.""LineNum"" " _
            & " Left OUTER JOIN ""OQUT"" t11 On t10.""DocEntry"" = t11.""DocEntry""  " _
            & " Left OUTER JOIN ""DLN1"" t14 On t1.""BaseEntry"" = t14.""DocEntry"" And T1.""BaseType"" = t14.""ObjType"" And  T1.""BaseLine"" = t14.""LineNum"" " _
            & " Left OUTER JOIN ""RDR1"" t15 On t14.""BaseEntry"" = t15.""DocEntry"" And T14.""BaseType"" = t15.""ObjType"" And  T14.""BaseLine"" = t15.""LineNum"" " _
            & " Left OUTER JOIN ""ORDR"" t16 On t15.""DocEntry"" = t16.""DocEntry""  " _
            & " Left OUTER JOIN ""QUT1"" t17 On t15.""BaseEntry"" = t17.""DocEntry"" And T15.""BaseType"" = t17.""ObjType"" And T15.""BaseLine"" = t17.""LineNum"" " _
            & " Left OUTER JOIN ""OQUT"" t18 On t17.""DocEntry"" = t18.""DocEntry""  " _
            & " Left Join ""RIN1"" T19 On T1.""DocEntry"" = t19.""BaseEntry"" And T1.""ObjType"" = t19.""BaseType"" And T1.""LineNum"" = t19.""BaseLine"" "
            ' --ojo cambio de carlos zea
            'Left Join "ITR1" T23 On  T23."SrcObjTyp" = 24 And T23."ReconNum" = T55."ReconNum" 
            sSQL = sSQL & " Left Join(Select MAX(""ReconNum"")As ""ReconNum"", ""SrcObjTyp"", ""SrcObjAbs""   from ""ITR1"" group by ""SrcObjTyp"" , ""SrcObjAbs"" ) T23 " _
            & " On  T23.""SrcObjTyp"" = 24 And T23.""ReconNum"" = T55.""ReconNum"" "
            ' -- fin camios de carlos zea

            sSQL = sSQL & " Left Join ""ORCT"" T24 On  T23.""SrcObjTyp"" = 24 And T24.""DocEntry"" = T23.""SrcObjAbs"" And T24.""BoeSum"" > 0 " _
            & " Left Join ""OBOE"" T25 On  T25.""BoeNum"" = T24.""BoeNum"" And T25.""BoeStatus"" = 'P'  AND COALESCE(t25.""BoeNum"",0) <>0 " _
            & " LEFT JOIN (select X.""TaxDate"" , X.""BOENumber"" , 'Efectos' as ""ESEFECTO"" from (  " _
            & " Select T2.""TaxDate"", Coalesce (cast( T1.""BOENumber"" As Nvarchar(50)),'SIN EFECTO') as ""BOENumber"" from OBOT T2   " _
            & " Left Join BOT1 T1 ON T1.""AbsEntry"" = T2.""AbsEntry"" and T2.""StatusTo"" = 'P' and T2.""Reconciled"" = 'N'   " _
            & " ) as X  where X.""BOENumber"" <> 'SIN EFECTO' ) t30 on t30.""BOENumber""   = t25.""BoeNum"""
            '& "    Left Join(select X.""TaxDate"" , X.""BOENumber"" from ( " _
            '& " Select  T2.""TaxDate"", T1.""BOENumber"" from ""OBOT"" T2  " _
            '& " Left Join ""BOT1"" T1 ON T1.""AbsEntry"" = T2.""AbsEntry"" And T2.""StatusTo"" = 'P' and T2.""Reconciled"" = 'N') as X ) t30 on t30.""BOENumber""   = t25.""BoeNum"" " _
            '& " WHERE  T33.""PaidToDate"" = T0.""DocTotal"" "
            sSQL = sSQL & "  WHERE  T0.""DocTotal"">0 And (t0.""CANCELED"") = 'N' and t1.""LineTotal"">0  AND t1.""Commission""> 0 " _
            & " And t0.""SlpCode"" >= '" & strAgenteD & "' And t0.""SlpCode"" <='" & strAgenteH & "' " _
            & " And  COALESCE(t1.""U_EXO_COMAPL"",'N') = '" & strCobradas & "' and  T55.""Canceled"" = 'N' " _
            & " Group BY  " _
            & " T24.""DocNum"", " _
            & " t24.""BoeNum"",t30.""BOENumber"", " _
            & " t0.""DocEntry"", t0.""DocNum"", t0.""TaxDate"", t2.""ItmsGrpCod"", t0.""CardCode"" ,t0.""CardName"" ,t1.""LineNum"", t1.""ItemCode"", t1.""Dscription"", t1.""DiscPrcnt"" , " _
            & " t1.""LineTotal"",t1.""Commission"",t1.""U_EXO_DTOCAL"",  t1.""U_EXO_FECCOM"",t0.""ObjType"",T0.""PaidToDate"" ,T0.""DocTotal"", t0.""SlpCode"",T1.""U_EXO_IMPCOM"",t3.""SlpName"", " _
            & " t5.""DocNum"", " _
            & " t11.""DocNum"", t13.""DocNum"", t9.""DocNum"", t18.""DocNum"",T19.""U_EXO_IMPCOM"",T30.""TaxDate"" " _
            & " HAVING T1.""U_EXO_IMPCOM"" + COALESCE(T19.""U_EXO_IMPCOM"", 0) > 0 And COALESCE(T30.""TaxDate"", MAX(T55.""ReconDate"")) >= '" & strFechaD & "' " _
            & " And COALESCE(T30.""TaxDate"", MAX(T55.""ReconDate"")) <= '" & strFechaH & "'  and sum( T33.""PaidToDate"") =  T0.""DocTotal"" " _
            & " and Case when T24.""BoeNum"" is not null and T30.""BOENumber"" is null then 'No MOSTRAR' else 'MOSTRAR' end = 'MOSTRAR' " _
            & " UNION ALL " _
            & " Select 'N'  ""Sel"", t0.""DocEntry"", t0.""DocNum"", t0.""TaxDate"",t0.""TaxDate"",  t2.""ItmsGrpCod"", t0.""CardCode"" ,t0.""CardName"" ,t1.""LineNum"",   t1.""ItemCode"", t1.""Dscription"", t1.""DiscPrcnt"" ,   " _
            & " t1.""LineTotal"",t1.""Commission"",t1.""U_EXO_DTOCAL"", T1.""U_EXO_IMPCOM"" ""Importe"", t1.""U_EXO_FECCOM"", t0.""ObjType"",T0.""PaidToDate"" ,T0.""DocTotal"", t0.""SlpCode"",T1.""U_EXO_IMPCOM"",t3.""SlpName"", " _
            & " '0' ""NumOfertaDire"", " _
            & " '0' ""NumOfertaPed"", '0' ""NumOfertaAlb"",'0' ""NumAlb"",'0' ""NumOfertaTodo"" " _
            & " FROM ""ORIN"" t0   " _
            & " INNER Join ""RIN1"" t1 ON t0.""DocEntry"" = t1.""DocEntry""  And T1.""BaseType"" <> 13     " _
            & " INNER Join ""OITM"" t2 On t1.""ItemCode"" = t2.""ItemCode""   " _
            & " INNER Join ""OSLP"" t3 on t0.""SlpCode"" = t3.""SlpCode"" " _
            & " Left Join ""RIN6"" T4 On T0.""DocEntry"" = T4.""DocEntry""  " _
            & "  WHERE  T0.""DocTotal"" > 0 And t1.""LineTotal"" > 0   And T0.""TaxDate"" between '" & strFechaD & "' and '" & strFechaH & "' And t0.""SlpCode"" >= '" & strAgenteD & "'  and t0.""SlpCode"" <= '" & strAgenteH & "' " _
            & " And  COALESCE(t1.""U_EXO_COMAPL"",'N')='" & strCobradas & "' And (t0.""CANCELED"") = 'N'  AND t1.""Commission""> 0    " _
            & " Group by t0.""ObjType"", t0.""DocEntry"", t0.""DocNum"", t0.""TaxDate"", t2.""ItmsGrpCod"", t0.""CardCode"" , t0.""CardName"" , t1.""LineNum"", t1.""ItemCode"", t1.""Dscription"", t1.""DiscPrcnt"" , " _
            & " t1.""LineTotal"", t1.""Commission"", t1.""U_EXO_DTOCAL"", T1.""U_EXO_IMPCOM"", t1.""U_EXO_FECCOM"",t0.""ObjType"", T0.""PaidToDate"", T0.""DocTotal"", t0.""SlpCode"", T1.""U_EXO_IMPCOM"", t3.""SlpName"" " _
            & " )  " _
            & " ORDER BY ""FechaUltCobro"", ""TaxDate"", ""DocNum"", ""LineNum"" "


            'sSQL = "Select * FROM ("
            'sSQL = sSQL & "Select 'N'  ""Sel"", t0.""DocEntry"", t0.""DocNum"", t0.""TaxDate"", COALESCE(T12.""TaxDate"", MAX(T5.""ReconDate"")) ""FechaUltCobro"", t2.""ItmsGrpCod"", t0.""CardCode"" ,t0.""CardName"" ,t1.""LineNum"", t1.""ItemCode"", t1.""Dscription"", t1.""DiscPrcnt"" , " _
            '& " t1.""LineTotal"",t1.""Commission"",t1.""U_EXO_DTOCAL"",  T1.""U_EXO_IMPCOM"" + COALESCE(T6.""U_EXO_IMPCOM"",0) ""Importe"", t1.""U_EXO_FECCOM"",t0.""ObjType""  " _
            '& " FROM ""OINV"" t0  INNER JOIN ""INV1"" t1 ON t0.""DocEntry"" = t1.""DocEntry""  " _
            '& " INNER JOIN ""OITM"" t2 On t1.""ItemCode"" = t2.""ItemCode""  " _
            '& " LEFT JOIN ""INV6"" T3 On T0.""DocEntry"" = T3.""DocEntry""  " _
            '& " Left Join ""ITR1"" T4 On  T4.""SrcObjTyp"" = 13 And T4.""SrcObjAbs"" = T0.""DocEntry"" And T3.""InstlmntID"" = T4.""InstID"" " _
            '& " Left Join ""OITR"" T5 ON t4.""ReconNum"" = T5.""ReconNum"" AND T5.""Canceled"" ='N'  " _
            '& " LEFT JOIN ""ITR1"" T13 On  T13.""SrcObjTyp"" = 24 And T13.""ReconNum"" = T5.""ReconNum"" " _
            '& " LEFT JOIN ""ORCT"" T10 On  T13.""SrcObjTyp"" = 24 And T10.""DocEntry"" = T13.""SrcObjAbs"" And T10.""BoeSum"" > 0 " _
            '& " LEFT JOIN ""OBOE"" T11 on  T11.""BoeNum"" = T10.""BoeNum"" And T11.""BoeStatus"" = 'P' " _
            '& " LEFT JOIN (select X.""TaxDate"" , X.""BOENumber"" from ( " _
            '& " select T2.""TaxDate"", T1.""BOENumber"" from ""OBOT"" T2 " _
            '& " Left Join ""BOT1"" T1 ON T1.""AbsEntry"" = T2.""AbsEntry"" And T2.""StatusTo"" = 'P' and T2.""Reconciled"" = 'N') as X ) t12 on t12.""BOENumber"" = t11.""BoeNum"" " _
            '& " LEFT JOIN ""RIN1"" T6 On T1.""DocEntry"" = t6.""BaseEntry"" And T1.""ObjType"" = t6.""BaseType"" And T1.""LineNum"" = t6.""BaseLine"" " _
            '& " WHERE  T3.""PaidToDate"" >= T0.""DocTotal"" And T0.""DocTotal"" > 0 And t1.""LineTotal"" > 0"
            ''& " And T5.""ReconDate"" between '" & strFechaD & "' and '" & strFechaH & "' " 'antes esta fecha, lo cambio por la de cobro
            'If strAgenteD <> "" Then
            '    sSQL = sSQL & " And t0.""SlpCode"" >= '" & strAgenteD & "'"
            'End If
            'If strAgenteH <> "" Then
            '    sSQL = sSQL & " and t0.""SlpCode"" <= '" & strAgenteH & "'  "
            'End If
            'sSQL = sSQL & " AND  COALESCE(t1.""U_EXO_COMAPL"",'N')='" & strCobradas & "' " _
            '& " And (t0.""CANCELED"") = 'N' " _
            '& " AND t1.""Commission""> 0" _
            '& " group by t0.""DocEntry"", t0.""DocNum"", t0.""TaxDate"",  t2.""ItmsGrpCod"", t0.""CardCode"" ,t0.""CardName"" ,t1.""LineNum"", t1.""ItemCode"", t1.""Dscription"", t1.""DiscPrcnt"" , " _
            '& " t1.""LineTotal"",t1.""Commission"",t1.""U_EXO_DTOCAL"", T1.""U_EXO_IMPCOM"", t1.""U_EXO_FECCOM"",t0.""ObjType"", T6.""U_EXO_IMPCOM"", T12.""TaxDate"" " _
            '& " HAVING T1.""U_EXO_IMPCOM"" + COALESCE(T6.""U_EXO_IMPCOM"",0) > 0" _
            '& " AND COALESCE(T12.""TaxDate"", MAX(T5.""ReconDate"")) between '" & strFechaD & "' and '" & strFechaH & "' "
            ''& " ORDER BY MAX(T5.""ReconDate""),t0.""TaxDate"", t0.""DocNum"" "

            ''abonos sin documento origen
            'sSQL = sSQL & " UNION ALL " _
            '& "  Select 'N'  ""Sel"", t0.""DocEntry"", t0.""DocNum"", t0.""TaxDate"",t0.""TaxDate"", " _
            '& " t2.""ItmsGrpCod"", t0.""CardCode"" ,t0.""CardName"" ,t1.""LineNum"",  " _
            '& " t1.""ItemCode"", t1.""Dscription"", t1.""DiscPrcnt"" ,   " _
            '& " t1.""LineTotal"",t1.""Commission"",t1.""U_EXO_DTOCAL"", T1.""U_EXO_IMPCOM"" ""Importe"", t1.""U_EXO_FECCOM"",t0.""ObjType""   " _
            '& " FROM ""ORIN"" t0   " _
            '& " INNER JOIN ""RIN1"" t1 ON t0.""DocEntry"" = t1.""DocEntry"" and T1.""BaseType"" <>13    " _
            '& " INNER JOIN ""OITM"" t2 On t1.""ItemCode"" = t2.""ItemCode""  " _
            '& " LEFT JOIN ""RIN6"" T3 On T0.""DocEntry"" = T3.""DocEntry"" " _
            '& " WHERE  T0.""DocTotal"" > 0 And t1.""LineTotal"" > 0  " _
            '& " And T0.""TaxDate"" between '" & strFechaD & "' and '" & strFechaH & "' " _
            '& " And t0.""SlpCode"" >= '" & strAgenteD & "' and t0.""SlpCode"" <= '" & strAgenteH & "' " _
            '& " And  COALESCE(t1.""U_EXO_COMAPL"",'N')='" & strCobradas & "'  And (t0.""CANCELED"") = 'N'  AND t1.""Commission""> 0  " _
            '& " group by t0.""ObjType"",t0.""DocEntry"", t0.""DocNum"", t0.""TaxDate"",  t2.""ItmsGrpCod"", t0.""CardCode"" ,t0.""CardName"" ,t1.""LineNum"", t1.""ItemCode"", t1.""Dscription"", t1.""DiscPrcnt"" ,   " _
            '& " t1.""LineTotal"",t1.""Commission"",t1.""U_EXO_DTOCAL"", T1.""U_EXO_IMPCOM"", t1.""U_EXO_FECCOM"" ) " _
            '& " ORDER BY ""FechaUltCobro"",""TaxDate"",""DocNum"",""LineNum"""

            oRs.DoQuery(sSQL)
            oForm.Freeze(True)
            'columnas
            'Permitimos ordenación por columnas
            'oForm = OGlobal.conexionSAP.SBOApp.Forms.AddEx(oFP)

            oForm.DataSources.DataTables.Item("DT_GR").ExecuteQuery(sSQL)

            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(1).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(2).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(3).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(4).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(5).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(6).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(7).TitleObject.Sortable = True


            'formato columnas
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(0).AffectsFormMode = False
            oColumnChk = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(1), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.LinkedObjectType = "13"
            oColumnTxt.TitleObject.Caption = "Num. Interno" ' "DocEntry"


            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(2), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Factura" ' "DocNum"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(3), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Fecha Doc." ' "TaxDate"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(4), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Fecha Cobro" ' "TaxDate"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(5), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Familia" '"ItmsGrpCod"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(6), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.LinkedObjectType = "2"
            oColumnTxt.TitleObject.Caption = "Cod. Cliente" '"CardCode"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(7), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Nombre" '"CardName"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(8), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "LineNum"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(9), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.LinkedObjectType = "4"
            oColumnTxt.TitleObject.Caption = "Artículo " '"ItemCode"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(10), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Descripción" ' "Dscription"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(11), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "% Descuento" ' DiscPrcnt"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(12), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Importe Doc." '"LineTotal"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(13), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "% Comisión" '"Commission"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(14), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Visible = False
            oColumnTxt.TitleObject.Caption = "U_EXO_DTOCAL"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(15), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Importe Comisión"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(16), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Fecha pago comisión" '"U_EXO_FECCOM"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(17), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "ObjType"
            oColumnTxt.Visible = False


            'no visible
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(18).Visible = False
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(19).Visible = False
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(20).Visible = False
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(21).Visible = False
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(22).Visible = False
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(23).Visible = False
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(24).Visible = False
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(25).Visible = False
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(26).Visible = False
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(27).Visible = False
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(28).Visible = False


            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
        End Try

    End Sub

    Shared Sub TratarFactura(ByRef OExoGenerales As EXO_UIAPI.EXO_UIAPI, ByRef dt As SAPbouiCOM.DataTable, ByVal I As Integer, ByVal oFormFac As SAPbouiCOM.Form)
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim strSql As String = ""
        Dim sTabla As String = ""
        Try
            OExoGenerales.SBOApp.StatusBar.SetText("Aplicando comisión a factura", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If dt.GetValue("ObjType", I).ToString() = "13" Then
                sTabla = """INV1"""
            Else
                sTabla = """RIN1"""
            End If
            oRs = CType(OExoGenerales.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            strSql = "UPDATE " & sTabla & " SET ""U_EXO_COMAPL""='Y' ,   ""U_EXO_FECCOM""='" & CType(oFormFac.Items.Item("5_U_E").Specific, SAPbouiCOM.EditText).Value & "',  ""U_EXO_IMPCOM""=" & dt.GetValue("Importe", I).ToString.Replace(",", ".") & "  WHERE ""DocEntry""=" & dt.GetValue("DocEntry", I).ToString() & " and ""LineNum"" = " & dt.GetValue("LineNum", I).ToString() & ";"
            oRs.DoQuery(strSql)

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            OExoGenerales.SBOApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Throw ex
        Finally

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

        End Try
    End Sub

    Private Sub GenerarPDF(ByRef oForm As SAPbouiCOM.Form, ByVal sRptFileName As String)
        Dim oCRReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        Dim oFileDestino As CrystalDecisions.Shared.DiskFileDestinationOptions = Nothing
        Dim sServer As String = ""
        Dim sDriver As String = ""
        Dim sBBDD As String = ""
        Dim sUser As String = ""
        Dim sPwd As String = ""
        Dim sFileName As String = ""
        Dim sTipoDoc As String = ""
        Dim sDesdeFecha As String = ""
        Dim sHastaFecha As String = ""
        Dim strAgenteD As String = ""
        Dim strAgenteH As String = ""
        Dim strCobradas As String = ""
        Try





            'Establecemos las conexiones a la BBDD
            sServer = objGlobal.compañia.Server
            sBBDD = objGlobal.compañia.CompanyDB
            sUser = objGlobal.refDi.SQL.usuarioSQL
            sPwd = objGlobal.refDi.SQL.claveSQL

            sServer = objGlobal.compañia.Server.Replace("NDB@", "").Replace("HDB@", "").Replace("30013", "30015")
            oCRReport.Load(objGlobal.refDi.OGEN.pathGeneral & "\05.Rpt\" & sRptFileName, OpenReportMethod.OpenReportByDefault)

            If objGlobal.compañia.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                If Right(objGlobal.refDi.OGEN.pathDLL, 6).ToUpper = "DLL_64" Then
                    sDriver = "{B1CRHPROXY}"
                Else
                    sDriver = "{B1CRHPROXY32}"
                End If
                oCRReport.ApplyNewServer(sDriver, sServer, sUser, sPwd, sBBDD)
            Else

                For Idx = 0 To oCRReport.DataSourceConnections.Count - 1
                    oCRReport.DataSourceConnections(Idx).SetConnection(sServer, sBBDD, False)
                    oCRReport.DataSourceConnections(Idx).SetLogon(sUser, sPwd)
                Next

                For Idx = 0 To oCRReport.Subreports.Count - 1
                    For Idx2 = 0 To oCRReport.Subreports.Item(Idx).DataSourceConnections.Count - 1
                        oCRReport.Subreports(Idx).DataSourceConnections(Idx2).SetConnection(sServer, sBBDD, False)
                        oCRReport.Subreports(Idx).DataSourceConnections(Idx2).SetLogon(sUser, sPwd)
                    Next
                Next
            End If


            ''Establecemos los parámetros para el report.           
            sDesdeFecha = Right(CType(oForm.Items.Item("2_U_E").Specific, SAPbouiCOM.EditText).Value.ToString, 2) & "/"
            sDesdeFecha += Mid(CType(oForm.Items.Item("2_U_E").Specific, SAPbouiCOM.EditText).Value.ToString, 5, 2) & "/"
            sDesdeFecha += Left(CType(oForm.Items.Item("2_U_E").Specific, SAPbouiCOM.EditText).Value.ToString, 4)

            sHastaFecha = Right(CType(oForm.Items.Item("3_U_E").Specific, SAPbouiCOM.EditText).Value.ToString, 2) & "/"
            sHastaFecha += Mid(CType(oForm.Items.Item("3_U_E").Specific, SAPbouiCOM.EditText).Value.ToString, 5, 2) & "/"
            sHastaFecha += Left(CType(oForm.Items.Item("3_U_E").Specific, SAPbouiCOM.EditText).Value.ToString, 4)

            If CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.ComboBox).Value <> "" Then
                strAgenteD = (CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.ComboBox).Value)
            End If

            If CType(oForm.Items.Item("1_U_E").Specific, SAPbouiCOM.ComboBox).Value <> "" Then
                strAgenteH = (CType(oForm.Items.Item("1_U_E").Specific, SAPbouiCOM.ComboBox).Value)
            End If


            If CType(oForm.Items.Item("Check_0").Specific, SAPbouiCOM.CheckBox).Checked = True Then
                strCobradas = "Y"
            Else
                strCobradas = "N"
            End If

            'Preparamos para la exportación
            oCRReport.SetParameterValue("Schema@", sBBDD)
            oCRReport.SetParameterValue("fecha_desde@", sDesdeFecha)
            oCRReport.SetParameterValue("fecha_hasta@", sHastaFecha)
            If strAgenteD <> "" Then
                oCRReport.SetParameterValue("Agente_desde@", strAgenteD)
            End If
            If strAgenteH <> "" Then
                oCRReport.SetParameterValue("Agente_hasta@", strAgenteH)
            End If


            oCRReport.SetParameterValue("Cobradas@", strCobradas)


            sFileName = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\COMISION"
            sFileName = sFileName & ".pdf"

            'Compruebo si existe y lo borro
            If IO.File.Exists(sFileName) Then
                IO.File.Delete(sFileName)
            End If
            objGlobal.SBOApp.StatusBar.SetText("Generando pdf para envio impresión...Espere por favor", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
            oCRReport.ExportToDisk(ExportFormatType.PortableDocFormat, sFileName)

            'Cerramos
            oCRReport.Close()
            oCRReport.Dispose()

            'Si es Web
            If objGlobal.SBOApp.ClientType = BoClientType.ct_Browser Then
                objGlobal.SBOApp.SendFileToBrowser(sFileName)
                Exit Sub
            Else
                Process.Start(sFileName)
                Exit Sub
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oCRReport = Nothing
            oFileDestino = Nothing
        End Try
    End Sub
End Class
#End Region
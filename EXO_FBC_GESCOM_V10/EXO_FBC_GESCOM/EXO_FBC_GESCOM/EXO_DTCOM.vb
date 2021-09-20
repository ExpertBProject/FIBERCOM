Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_DTCOM
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
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "EXO-DTCOM"
                        'Cargamos UDO
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_DTCOM")
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
                        Case "UDO_FT_EXO_DTCOM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_VALIDATE_After(infoEvento) = False Then
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
                        Case "UDO_FT_EXO_DTCOM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_DTCOM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_FORM_VISIBLE(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_DTCOM"
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
    Private Function EventHandler_FORM_VISIBLE(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Dim Path As String = ""
        Dim XmlDoc As New System.Xml.XmlDocument
        'Dim EXO_Functions As New EXO_BasicDLL.EXO_Generic_Forms_Functions(Me.objGlobal.conexionSAP)

        EventHandler_FORM_VISIBLE = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                If CargarComboGrupoArticulos(oForm) = False Then
                    Exit Function
                End If
            End If

            EventHandler_FORM_VISIBLE = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function CargarComboGrupoArticulos(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim sSQL As String = ""
        CargarComboGrupoArticulos = False

        Try
            sSQL = "SELECT ""ItmsGrpCod"",""ItmsGrpNam"" FROM ""OITB"" Order by ""ItmsGrpNam"" "

            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

            CargarComboGrupoArticulos = True

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function EventHandler_VALIDATE_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        EventHandler_VALIDATE_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            Valida_Campos(oForm)

            EventHandler_VALIDATE_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function Valida_Campos(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Valida_Campos = False

        Dim dblMin As Double = 0
        Dim dblMax As Double = 0
        Try
            If oForm.Visible = True Then
                dblMin = CDbl(CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).String.Replace(objGlobal.refDi.OADM.separadorMillarB1, ""))
                dblMax = CDbl(CType(oForm.Items.Item("15_U_E").Specific, SAPbouiCOM.EditText).String.Replace(objGlobal.refDi.OADM.separadorMillarB1, ""))
                If dblMax < dblMin And dblMin <> dblMax Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - El Campo ""MAX"" no puede ser inferior al campo ""MIN"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    objGlobal.SBOApp.MessageBox("El Campo ""MAX"" no puede ser inferior al campo ""MIN"".")
                    Exit Function
                End If
            End If

            Valida_Campos = True
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function
    Private Function Valida_Campos_Lineas(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Valida_Campos_Lineas = False

        Dim dblMin As Double = 0
        Dim dblMinLin As Double = 0
        Dim dblMaxLin As Double = 0
        Dim dblMaxLinAnt As Double = 0
        Dim dblMax As Double = 0
        dblMin = CDbl(CType(oForm.Items.Item("13_U_E").Specific, SAPbouiCOM.EditText).String.Replace(objGlobal.refDi.OADM.separadorMillarB1, ""))
        dblMax = CDbl(CType(oForm.Items.Item("15_U_E").Specific, SAPbouiCOM.EditText).String.Replace(objGlobal.refDi.OADM.separadorMillarB1, ""))
        Try
            If oForm.Visible = True Then
                For i = 1 To CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).RowCount
                    If i > 1 Then
                        dblMaxLinAnt = CDbl(CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(i - 1).Specific, SAPbouiCOM.EditText).String)
                    End If
                    dblMaxLin = CDbl(CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(i).Specific, SAPbouiCOM.EditText).String)
                    dblMinLin = CDbl(CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(i).Specific, SAPbouiCOM.EditText).String)
                    If Not (dblMinLin >= dblMin And dblMinLin <= dblMax) Then
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Revise el Campo ""Desde"" de la línea " & i.ToString & " debe estar comprendido entre ""MIN"" y ""MAX"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        objGlobal.SBOApp.MessageBox("Revise el Campo ""Desde"" de la línea " & i.ToString & " debe estar comprendido entre ""MIN"" y ""MAX"".")
                        Exit Function
                    End If
                    If Not (dblMaxLin >= dblMin And dblMaxLin <= dblMax) Then
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Revise el Campo ""Hasta"" de la línea " & i.ToString & " debe estar comprendido entre ""MIN"" y ""MAX"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        objGlobal.SBOApp.MessageBox("Revise el Campo ""Hasta"" de la línea " & i.ToString & " debe estar comprendido entre ""MIN"" y ""MAX"".")
                        Exit Function
                    End If
                    If (dblMinLin <= dblMaxLinAnt) And dblMaxLinAnt <> 0 Then
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Revise el Campo ""Desde"" de la línea " & i.ToString & " debe ser superior a la línea anterior.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        objGlobal.SBOApp.MessageBox("Revise el Campo ""Desde"" de la línea " & i.ToString & " debe ser superior a la línea anterior.")
                        Exit Function
                    End If
                Next
            End If

            Valida_Campos_Lineas = True
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function

    Public Overrides Function SBOApp_FormDataEvent(infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_DTCOM"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                'Antes de actualizar comprobamos  los datos.
                                If ComprobarDatos(oForm) = False Then
                                    Return False
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                'Antes de añadir comprobamos  los datos.
                                If ComprobarDatos(oForm) = False Then
                                    Return False
                                End If
                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_DTCOM"
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
    Private Function ComprobarDatos(ByRef oForm As SAPbouiCOM.Form) As Boolean
        ComprobarDatos = False
        Try
            Dim dblMax As Double = 0
            dblMax = CDbl(CType(oForm.Items.Item("15_U_E").Specific, SAPbouiCOM.EditText).String.Replace(objGlobal.refDi.OADM.separadorMillarB1, ""))
            If dblMax <> 0 Then
                Valida_Campos(oForm)
                If CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).RowCount > 0 Then
                    If Valida_Campos_Lineas(oForm) = False Then
                        Exit Function
                    End If
                End If
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Antes de grabar rellene los datos de la cabecera.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objGlobal.SBOApp.MessageBox("Antes de grabar rellene los datos de la cabecera.")
                Exit Function
            End If

            ComprobarDatos = True

        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function
End Class

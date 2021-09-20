Imports System.Xml
Imports SAPbouiCOM

Public Class SAP_OINV
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)
        If actualizar Then
            cargaCampos()
        End If
    End Sub

#Region "Inicialización"

    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            'Campos de usuario en pedidos
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OINV.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs OINV", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
        End If
    End Sub
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function

#End Region
#Region "Eventos"
    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            'Apaño por un error que da EXO_Basic.dll al consultar infoEvento.FormTypeEx
            Try
                If infoEvento.FormTypeEx <> "" Then

                End If
            Catch ex As Exception
                Return False
            End Try

            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "133"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED


                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE


                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                            End Select

                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "133"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE


                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            End Select

                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "133"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "133"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select

                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function


    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sDocEntry As String = ""
        Dim sCodCli As String = ""

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "133"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD


                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)


                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                'Antes de añadir comprobamos que están rellenos los datos


                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select

                End Select

            Else

                Select Case infoEvento.FormTypeEx
                    Case "133"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                                If infoEvento.ActionSuccess = True Then
                                    oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

                                End If



                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                                If infoEvento.ActionSuccess = True Then
                                    oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                    sCodCli = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value
                                    oXml.LoadXml(infoEvento.ObjectKey)
                                    sDocEntry = oXml.SelectSingleNode("DocumentParams/DocEntry").InnerText
                                    EnviarAlerta(objGlobal, sDocEntry, "", sCodCli, "", "", "")
                                End If


                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If infoEvento.ActionSuccess = True Then
                                    oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

                                End If

                        End Select

                End Select

            End If

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function EventHandler_Form_Load(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Dim Path As String = ""
        Dim XmlDoc As New System.Xml.XmlDocument
        Dim oItem As SAPbouiCOM.Item
        Dim Tabla As String = "OINV"
        Dim sUser As String = ""
        Dim sTipo As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim Valor As String = ""

        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario

            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oForm.Freeze(True)
            If pVal.ActionSuccess = False Then


                oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
                oForm.Freeze(True)


               objGlobal. SboApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)


                'Condiciones especiales

                oItem = oForm.Items.Add("chkCondi", BoFormItemTypes.it_CHECK_BOX)

                oItem.Top = oForm.Items.Item("70").Top + 20
                oItem.Left = oForm.Items.Item("70").Left
                oItem.Height = oForm.Items.Item("151").Height
                oItem.Width = oForm.Items.Item("151").Width
                oItem.LinkTo = "70"
                oItem.FromPane = 0
                oItem.ToPane = 0
                oItem.Enabled = True

                CType(oItem.Specific, SAPbouiCOM.CheckBox).Caption = "Condiciones especiales aplicadas"
                CType(oItem.Specific, SAPbouiCOM.CheckBox).DataBind.SetBound(True, Tabla, "U_EXO_CONDIC")


            End If


            EventHandler_Form_Load = True

        Catch ex As Exception

            oForm.Freeze(False)
            oForm.Visible = True
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

        End Try
    End Function
#End Region
#Region "Auxiliares"
    Public Shared Sub EnviarAlerta(ByRef OExoGenerales As EXO_UIAPI.EXO_UIAPI, ByVal sEntry As String, ByVal sNumFac As String, ByVal sCliente As String, ByVal sUser As String, ByVal sSubject As String, ByVal sText As String)
        Dim oCmpSrv As SAPbobsCOM.CompanyService = Nothing
        Dim oMessageService As SAPbobsCOM.MessagesService = Nothing
        Dim oMessage As SAPbobsCOM.Message = Nothing
        Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns = Nothing
        Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn = Nothing
        Dim oLines As SAPbobsCOM.MessageDataLines = Nothing
        Dim oLine As SAPbobsCOM.MessageDataLine = Nothing
        Dim oRecipientCollection As SAPbobsCOM.RecipientCollection = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim SCli As String = ""
        Dim sEspeci As String = ""
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing

        Try

            oRs = CType(OExoGenerales.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            sSQL = "Select  ""t0"".""DocEntry"", ""t0"".""DocNum"",""t0"".""CardCode"",""t1"".""U_EXO_CONFAC"", ""t1"".""U_EXO_ESPECI"", ""t0"". ""U_EXO_CONDIC""  " _
            & " from ""OINV"" ""t0"" " _
            & "  INNER Join ""OCRD"" ""t1"" On ""t0"".""CardCode"" = ""t1"".""CardCode"" " _
            & " WHERE ""t0"".""DocEntry"" = '" & sEntry & "'"
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                If oRs.Fields.Item("U_EXO_CONFAC").Value.ToString = "Y" And oRs.Fields.Item("U_EXO_CONDIC").Value.ToString <> "Y" Then

                    SCli = oRs.Fields.Item("CardCode").Value.ToString
                    sEspeci = oRs.Fields.Item("U_EXO_ESPECI").Value.ToString
                    sNumFac = oRs.Fields.Item("DocNum").Value.ToString

                    sSQL = "Select ""t1"".""USER_CODE"" " &
                       "FROM ""OUSR"" ""t1"" " &
                       "WHERE ""t1"".""U_EXO_ALERTAFAC""='Y'"
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount > 0 Then

                        oCmpSrv = OExoGenerales.compañia.GetCompanyService()

                        oMessageService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService), SAPbobsCOM.MessagesService)
                        oMessage = CType(oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage), SAPbobsCOM.Message)
                        sSubject = "Este cliente cuenta con condiciones de facturacón: " & sEspeci
                        oMessage.Subject = sSubject
                        oMessage.Text = sSubject
                        oRecipientCollection = oMessage.RecipientCollection

                        sSQL = "Select ""t1"".""USER_CODE"" " &
                       "FROM ""OUSR"" ""t1"" " &
                       "WHERE ""t1"".""U_EXO_ALERTAFAC""='Y'"
                        oRs.DoQuery(sSQL)
                        If oRs.RecordCount > 0 Then
                            oXml.LoadXml(oRs.GetAsXML())
                            oNodes = oXml.SelectNodes("//row")
                            For i As Integer = 0 To oNodes.Count - 1
                                oNode = oNodes.Item(i)

                                oRecipientCollection.Add()
                                oRecipientCollection.Item(i).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                                oRecipientCollection.Item(i).UserCode = oNode.SelectSingleNode("USER_CODE").InnerText
                            Next
                        End If

                        pMessageDataColumns = oMessage.MessageDataColumns

                        If sEntry <> "" Then
                            pMessageDataColumn = pMessageDataColumns.Add()
                            pMessageDataColumn.ColumnName = "Número interno"
                            pMessageDataColumn.Link = SAPbobsCOM.BoYesNoEnum.tYES
                            oLines = pMessageDataColumn.MessageDataLines
                            oLine = oLines.Add()
                            oLine.Value = sEntry
                            oLine.Object = "13" 'llamada
                            oLine.ObjectKey = sEntry
                        End If

                        If sNumFac <> "" Then
                            pMessageDataColumn = pMessageDataColumns.Add()
                            pMessageDataColumn.ColumnName = "Número Factura"
                            oLines = pMessageDataColumn.MessageDataLines
                            oLine = oLines.Add()
                            oLine.Value = sNumFac
                        End If

                        If SCli <> "" Then
                            pMessageDataColumn = pMessageDataColumns.Add()
                            pMessageDataColumn.ColumnName = "Código Cliente"
                            pMessageDataColumn.Link = SAPbobsCOM.BoYesNoEnum.tYES
                            oLines = pMessageDataColumn.MessageDataLines
                            oLine = oLines.Add()
                            oLine.Value = SCli
                            oLine.Object = "2" 'cliente
                            oLine.ObjectKey = SCli
                        End If

                        oMessageService.SendMessage(oMessage)
                        'update
                        sSQL = "UPDATE ""OINV"" Set ""U_EXO_ALERTA"" = 'Y' where ""DocEntry"" = '" & sEntry & "';"
                        oRs.DoQuery(sSQL)

                    End If
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            'If oCmpSrv IsNot Nothing Then
            '    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrv)
            '    oCmpSrv = Nothing
            'End If
            'If pMessageDataColumns IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pMessageDataColumns)
            'If pMessageDataColumn IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pMessageDataColumn)
            'If oLines IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oLines)
            'If oLine IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oLine)
            'If oRecipientCollection IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRecipientCollection)
            ''If oCmpSrv IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrv)
            'If oMessageService IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oMessageService)
            'If oMessage IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oMessage)


            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCmpSrv, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(pMessageDataColumns, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(pMessageDataColumn, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oLines, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oLine, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRecipientCollection, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oMessageService, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oMessage, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
#End Region

End Class

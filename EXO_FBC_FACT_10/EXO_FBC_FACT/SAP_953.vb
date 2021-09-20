Imports SAPbouiCOM

Public Class SAP_953
    Inherits EXO_UIAPI.EXO_DLLBase

#Region "Variables"

    Private Shared _sTipoDocumento As String = ""
    Private Shared _bEsBorrador As Boolean = False
    Private Shared _bEjecutar As Boolean = False

#End Region

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

    End Sub

#Region "Inicialización"

    'Si no hay filtros que añadir sobreescribimos el método así
    'Public Overrides Function filtros() As SAPbouiCOM.EventFilters
    '    Dim filtro As SAPbouiCOM.EventFilters = Nothing
    '    Return filtro
    'End Function
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
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
                        Case "953"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select

                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "953"

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
                        Case "953"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "953"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before_Inner(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

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

    Private Function EventHandler_ItemPressed_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "4" Then
                If oForm.PaneLevel = 3 Then
                    _sTipoDocumento = CType(oForm.Items.Item("14").Specific, SAPbouiCOM.ComboBox).Value.ToString.Trim

                    If CType(oForm.Items.Item("156").Specific, SAPbouiCOM.CheckBox).Checked = True Then
                        _bEsBorrador = True
                    Else
                        _bEsBorrador = False
                    End If
                ElseIf oForm.PaneLevel = 9 Then
                    If CType(oForm.Items.Item("180").Specific, SAPbouiCOM.OptionBtn).Selected = True OrElse
                       CType(oForm.Items.Item("181").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                        _bEjecutar = True
                    Else
                        _bEjecutar = False
                    End If
                End If
            End If

            EventHandler_ItemPressed_Before = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function EventHandler_ItemPressed_Before_Inner(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_Before_Inner = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "4" Then
                If oForm.PaneLevel = 3 Then
                    _sTipoDocumento = CType(oForm.Items.Item("14").Specific, SAPbouiCOM.ComboBox).Value.ToString.Trim

                    If CType(oForm.Items.Item("156").Specific, SAPbouiCOM.CheckBox).Checked = True Then
                        _bEsBorrador = True
                    Else
                        _bEsBorrador = False
                    End If
                ElseIf oForm.PaneLevel = 9 Then
                    If CType(oForm.Items.Item("180").Specific, SAPbouiCOM.OptionBtn).Selected = True OrElse
                       CType(oForm.Items.Item("181").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                        _bEjecutar = True
                    Else
                        _bEjecutar = False
                    End If
                End If
            End If

            EventHandler_ItemPressed_Before_Inner = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux As SAPbobsCOM.Recordset = Nothing
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        Dim sSql As String = ""


        'TODO Comentar para volver atrás
        Dim cRate As Double = 0
        Dim cCantidad As Double = 0
        Dim cPrecio As Double = 0
        Dim sTarifa As String = ""
        Dim sTipoTTE As String = ""
        Dim sTipoPorte As String = ""
        Dim cCosteTTE As Double = 0
        Dim cTTEI01 As Double = 0
        Dim sCYAT01 As String = ""
        Dim cCYAV01 As Double = 0
        Dim cCYAI01 As Double = 0
        Dim sCYAT02 As String = ""
        Dim cCYAV02 As Double = 0
        Dim cCYAI02 As Double = 0
        Dim sCYAT03 As String = ""
        Dim cCYAV03 As Double = 0
        Dim cCYAI03 As Double = 0
        Dim sCYAC01 As String = ""
        Dim sCYAC02 As String = ""
        Dim sCYAC03 As String = ""
        Dim sCYAC04 As String = ""
        Dim sCYAC05 As String = ""
        Dim sCYAC06 As String = ""
        Dim iU_EXO_CYAIDCONT As Integer = 0
        Dim oOINVUpdate As SAPbobsCOM.Documents = Nothing
        'Fin TODO
        Dim bEjecutado As Boolean = False

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "4" Then
                If oForm.PaneLevel = 10 Then
                    If _sTipoDocumento = "13" Then
                        oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                        sSql = "SELECT  ""t0"".""DocEntry"", ""t0"".""DocNum"",""t0"".""CardCode"",""t1"".""U_EXO_CONFAC"", ""t1"".""U_EXO_ESPECI"",""t0"".""DataSource"", ""t0"".""U_EXO_ALERTA"" " _
                        & " from ""OINV"" ""t0"" " _
                        & " INNER JOIN ""OCRD"" ""t1"" On ""t0"".""CardCode"" = ""t1"".""CardCode"" " _
                        & " WHERE ""t0"".""DataSource""='A' and ""t0"".""U_EXO_ALERTA""<>'Y' "
                        oRs.DoQuery(sSql)

                        oXml.LoadXml(oRs.GetAsXML())
                        oNodes = oXml.SelectNodes("//row")

                        'Asiganando ATCs a facturas
                        If _bEsBorrador = False AndAlso _bEjecutar = True Then
                            If oRs.RecordCount > 0 Then
                                For i As Integer = 0 To oNodes.Count - 1
                                    oNode = oNodes.Item(i)

                                    objGlobal.SBOApp.StatusBar.SetText("Generando alerta condiciones especiales Factura - Número " & oNode.SelectSingleNode("DocNum").InnerText & " -> " & CInt(i + 1).ToString & " de " & oRs.RecordCount.ToString & " ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    SAP_OINV.EnviarAlerta(objGlobal, oNode.SelectSingleNode("DocEntry").InnerText, oNode.SelectSingleNode("DocNum").InnerText, oNode.SelectSingleNode("CardCode").InnerText, "", "", "")

                                    If bEjecutado = False Then bEjecutado = True
                                Next
                            End If
                        End If


                        'Fin TODO

                        If bEjecutado = True Then
                            objGlobal.SBOApp.StatusBar.SetText("Operación finalizada con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    End If
                End If
            End If

            EventHandler_ItemPressed_After = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            If pVal.ItemUID = "4" AndAlso oForm.PaneLevel = 10 Then
                _sTipoDocumento = ""
                _bEsBorrador = False
                _bEjecutar = False
            End If
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOINVUpdate, Object))
        End Try
    End Function



#End Region

End Class


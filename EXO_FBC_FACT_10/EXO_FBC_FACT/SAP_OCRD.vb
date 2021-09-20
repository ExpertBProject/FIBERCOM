Imports System.Xml
Imports SAPbouiCOM

Public Class SAP_OCRD
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
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OCRD.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs OCRD", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
                        Case "134"

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
                        Case "134"

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
                        Case "134"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "134"

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
    Private Function EventHandler_Form_Load(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Dim Path As String = ""
        Dim XmlDoc As New System.Xml.XmlDocument
        Dim oItem As SAPbouiCOM.Item
        Dim Tabla As String = "OCRD"
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


                objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)


                'Condiciones especiales

                oItem = oForm.Items.Add("chkCondFac", BoFormItemTypes.it_CHECK_BOX)

                oItem.Top = oForm.Items.Item("434").Top
                oItem.Left = oForm.Items.Item("234000597").Left
                oItem.Height = oForm.Items.Item("234000597").Height
                oItem.Width = oForm.Items.Item("234000597").Width
                oItem.FromPane = 6
                oItem.ToPane = 6
                oItem.LinkTo = "234000597"
                oItem.Enabled = True

                CType(oItem.Specific, SAPbouiCOM.CheckBox).Caption = "Condiciones de Facturación"
                CType(oItem.Specific, SAPbouiCOM.CheckBox).DataBind.SetBound(True, Tabla, "U_EXO_CONFAC")

                oItem = oForm.Items.Add("txtCondi", BoFormItemTypes.it_EXTEDIT)
                oItem.Top = oForm.Items.Item("89").Top
                oItem.Left = oForm.Items.Item("283").Left
                oItem.Height = oForm.Items.Item("283").Height * 8
                oItem.Width = oForm.Items.Item("283").Width
                oItem.LinkTo = "234000597"
                oItem.FromPane = 6
                oItem.ToPane = 6
                oItem.Enabled = True
                CType(oItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, Tabla, "U_EXO_ESPECI")

                oItem = oForm.Items.Add("lblCondi", BoFormItemTypes.it_STATIC)
                oItem = oForm.Items.Item("lblCondi")
                oItem.Top = oForm.Items.Item("89").Top
                oItem.Left = oForm.Items.Item("282").Left
                oItem.Height = oForm.Items.Item("282").Height
                oItem.Width = oForm.Items.Item("282").Width
                oItem.LinkTo = "234000597"
                oItem.FromPane = 6
                oItem.ToPane = 6
                CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Especificación condiciones"


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
End Class

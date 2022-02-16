Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OCRD
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

    End Sub
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
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ITEM_PRESSED_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

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
                                    If EventHandler_VALIDATE_Before(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

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
    Private Function EventHandler_ITEM_PRESSED_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sCIF As String = ""
        Try

            If pVal.ItemUID = "1" Then
                Comprobar_CIF(oForm)
            End If

            EventHandler_ITEM_PRESSED_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_VALIDATE_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sCIF As String = ""
        Try

            If pVal.ItemUID = "41" Then
                Comprobar_CIF(oForm)
            End If
            If oForm.Visible = True Then
                ' Comprobamos el comisionista 2 AgentCode
                Dim sComisionista As String = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("AgentCode", 0).ToUpper
                If sComisionista = "" Then
                    CType(oForm.Items.Item("228").Specific, SAPbouiCOM.ComboBox).Select("NINGUNO", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
            End If

            EventHandler_VALIDATE_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function

    Private Function Comprobar_CIF(ByRef oform As SAPbouiCOM.Form) As Boolean
        Dim sCIF As String = "" : Dim sCardTypeAct As String = "" : Dim sCardCodeAct As String = "" : Dim sCardCode As String = ""
        Dim sMensaje As String = ""
        Dim sSQL As String = ""

        Comprobar_CIF = False
        Try
            If oform.Mode = BoFormMode.fm_ADD_MODE Or oform.Mode = BoFormMode.fm_UPDATE_MODE Then
                sCIF = CType(oform.Items.Item("41").Specific, SAPbouiCOM.EditText).Value.ToString
                If sCIF.Trim = "" Then
                    Comprobar_CIF = True
                    Exit Function
                End If
                sCardCodeAct = CType(oform.Items.Item("5").Specific, SAPbouiCOM.EditText).Value.ToString
                If CType(oform.Items.Item("40").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                    sCardTypeAct = CType(oform.Items.Item("40").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                Else
                    sCardTypeAct = ""
                End If
                Dim sWhere As String = ""
                Select Case sCardTypeAct
                    Case "C", "L" : sWhere = " and ""CardType"" in ('C','L') "
                    Case "S" : sWhere = " and ""CardType""='S' "
                End Select
                sSQL = "SELECT TOP 1 ""CardCode"" FROM ""OCRD"" WHERE ""LicTradNum""='" & sCIF & "' and ""CardCode""<>'" & sCardCodeAct & "' " & sWhere
                sCardCode = objGlobal.refDi.SQL.sqlStringB1(sSQL)


                If sCardCode.Trim <> "" Then
                    sSQL = "SELECT ""CardCode"" ""Código"",""CardName"" ""Nombre"",""LicTradNum"" ""Nº identificación fiscal"", "
                    sSQL &= " CASE WHEN ""CardType""='C' THEN 'Cliente' WHEN ""CardType""='L' THEN 'Leads'  WHEN ""CardType""='S' THEN 'Proveedor' END ""Tipo"" "
                    sSQL &= " FROM ""OCRD"" WHERE ""LicTradNum""='" & sCIF & "' and ""CardCode""<>'" & sCardCodeAct & "' " & sWhere
                    Select Case sCardTypeAct
                        Case "C", "L" : sMensaje = "Existen Clientes o Leads con el mismo Nº identificación fiscal. Revise los datos."
                        Case "S" : sMensaje = "Existen Proveedores con el mismo Nº identificación fiscal. Revise los datos."
                    End Select

                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    If CargarFormICDUPLI(sSQL) = False Then
                        Exit Function
                    End If
                    Exit Function
                End If
            End If
            Comprobar_CIF = True
        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Public Function CargarFormICDUPLI(ByVal sSQL As String) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim Path As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)
        CargarFormICDUPLI = False

        Try

            'abrir formulario
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = EXO_Xml.LoadFormXml(objGlobal.leerEmbebido(GetType(EXO_OCRD), "EXO_ICDUPLI.srf"), True).ToString

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)

            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                End If
            End Try

            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            For i = 0 To 3
                oColumnTxt = CType(CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                If i = 0 Then
                    oColumnTxt.LinkedObjectType = "2"
                End If
            Next


            CargarFormICDUPLI = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
End Class

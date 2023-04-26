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

    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oChk As CheckBox
        Try

            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "1282"
                        'marcar propiedad 1
                        oForm = objGlobal.SBOApp.Forms.ActiveForm
                        If oForm.TypeEx = "134" Then
                            oForm.Freeze(True)
                            CType(oForm.Items.Item("10").Specific, SAPbouiCOM.Folder).Select()
                            Try
                                For i As Integer = 1 To 1
                                    If CType(CType(oForm.Items.Item("136").Specific, Matrix).Columns.Item("1").Cells.Item(i).Specific, EditText).Value <> "" Then
                                        oChk = CType(CType(oForm.Items.Item("136").Specific, SAPbouiCOM.Matrix).Columns.Item("2").Cells.Item(i).Specific, CheckBox)
                                        oChk.Checked = True
                                    Else
                                        Exit For
                                    End If
                                Next
                                CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Folder).Select()
                                CType(oForm.Items.Item("7").Specific, SAPbouiCOM.EditText).Active = True
                            Catch ex As Exception
                                oForm.Freeze(False)
                            Finally
                                oForm.Freeze(False)
                            End Try
                        End If


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
                                    If EventHandler_ITEM_PRESSED_Before(infoEvento) = False Then
                                        Return False
                                    End If
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

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ITEM_PRESSED_Before(infoEvento) = False Then
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
    Private Function EventHandler_ITEM_PRESSED_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sCIF As String = ""
        Try

            If pVal.ItemUID = "1" And (oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_UPDATE_MODE) Then
                'solo cuando sea clientes
                If CType(oForm.Items.Item("40").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString = "C" Then
                    If Comprobar_Datos(oForm) = False Then
                        Return False
                    End If
                End If

                'If CType(oForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value.ToString.StartsWith("CC") Then 'cardcode contado no
                '    'no compruebo
                'Else
                '    If CType(oForm.Items.Item("40").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString = "C" Then
                '        If Comprobar_Datos(oForm) = False Then
                '            Return False
                '        End If
                '    End If


                'End If


            End If

            EventHandler_ITEM_PRESSED_Before = True

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

            If pVal.ItemUID = "41" And (oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_EDIT_MODE) Then
                Comprobar_CIF(oForm)
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
    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument
        Dim sItemCode As String = ""

        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "134"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                'If Comprobar_Datos(oForm) = False Then
                                '    Return False
                                'End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                'If Comprobar_Datos(oForm) = False Then
                                '    Return False
                                'End If
                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "134"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If infoEvento.ActionSuccess Then

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess Then


                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If infoEvento.ActionSuccess Then

                                End If
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
    Private Function Comprobar_Datos(ByRef oform As SAPbouiCOM.Form) As Boolean

        Dim sMensaje As String = ""
        Dim sSQL As String = ""
        Dim oChk As CheckBox
        Dim intCuenta As Integer = 0
        Dim sViaPago As String = ""
        Dim intNumProp As Integer = 0
        Dim sDato As String = ""
        Dim sDire As String = ""
        Dim intFolder As Integer = 0
        Dim strCIF As String = ""
        Dim sLetras As String = ""

        Comprobar_Datos = False
        Try
            If CType(oform.Items.Item("5").Specific, SAPbouiCOM.EditText).Value.ToString.StartsWith("CC") Then 'cardcode contado no
                If oform.Mode = BoFormMode.fm_ADD_MODE Or oform.Mode = BoFormMode.fm_UPDATE_MODE Then
                    'CIF
                    If CType(oform.Items.Item("41").Specific, SAPbouiCOM.EditText).Value.ToString = "" Then
                        sMensaje = "El CIF no puede estar vacío "
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    strCIF = CType(oform.Items.Item("41").Specific, SAPbouiCOM.EditText).Value.ToString
                    sLetras = Left(sLetras, 1)
                    If IsNumeric(sLetras) Then
                        objGlobal.SBOApp.StatusBar.SetText("CIF/NIF debe empezar por los 2 primeros digitos del pais", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    sLetras = Mid(strCIF, 2, 1)
                    If IsNumeric(sLetras) Then
                        objGlobal.SBOApp.StatusBar.SetText("CIF/NIF debe empezar por los 2 primeros digitos del pais", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    'si no empieza con 2 digtiso letras no dejo continuar

                    If Left(strCIF.Trim.ToString.ToUpper, 2) = "ES" And CType(oform.Items.Item("39").Specific, SAPbouiCOM.EditText).Value.ToString.Trim = "1" Then 'CType(oform.Items.Item("530001024").Specific, SAPbouiCOM.ComboBox).Selected.Value = "1" Then
                        'comprobar que 
                        If strCIF.Length < 11 Then
                            objGlobal.SBOApp.StatusBar.SetText("El número de dígitos en lo CIF/NIF debe ser 11 obligatoriamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Function
                        End If
                        If Comprobar_CIF_NIF(objGlobal, strCIF) = False Then
                            Exit Function
                        End If

                    End If



                End If
                Comprobar_Datos = True
            Else

                If oform.Mode = BoFormMode.fm_ADD_MODE Or oform.Mode = BoFormMode.fm_UPDATE_MODE Then
                    objGlobal.SBOApp.StatusBar.SetText("...Comprobando datos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oform.Freeze(True)
                    'If CType(oform.Items.Item("5").Specific, SAPbouiCOM.EditText).Value.ToString.StartsWith("CC") Then 'cardname
                    'card name obligatorio
                    If CType(oform.Items.Item("7").Specific, SAPbouiCOM.EditText).Value.ToString = "" Then
                        sMensaje = "El nombre de cliente no puede estar vacío "
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    If CType(oform.Items.Item("16").Specific, SAPbouiCOM.ComboBox).Selected.Description.ToString = "" Or CType(oform.Items.Item("16").Specific, SAPbouiCOM.ComboBox).Selected.Description.ToString = "Clientes" Then     'grupo
                        sMensaje = "El grupo de cliente no puede estar vacío y no puede asignarse como grupo Clientes "
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    'CIF
                    If CType(oform.Items.Item("41").Specific, SAPbouiCOM.EditText).Value.ToString = "" Then
                        sMensaje = "El CIF no puede estar vacío "
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    strCIF = CType(oform.Items.Item("41").Specific, SAPbouiCOM.EditText).Value.ToString
                    'si no empieza con 2 digtiso letras no dejo continuar

                    sLetras = Left(sLetras, 1)
                    If IsNumeric(sLetras) Then
                        objGlobal.SBOApp.StatusBar.SetText("CIF/NIF debe empezar por los 2 primeros digitos del pais", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    sLetras = Mid(strCIF, 2, 1)
                    If IsNumeric(sLetras) Then
                        objGlobal.SBOApp.StatusBar.SetText("CIF/NIF debe empezar por los 2 primeros digitos del pais", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    If Left(strCIF.Trim.ToString.ToUpper, 2) = "ES" And CType(oform.Items.Item("530001024").Specific, SAPbouiCOM.ComboBox).Selected.Value = "1" Then
                        'comprobar que 
                        If strCIF.Length < 11 Then
                            objGlobal.SBOApp.StatusBar.SetText("El número de dígitos en lo CIF/NIF debe ser 11 obligatoriamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Function
                        End If
                        If Comprobar_CIF_NIF(objGlobal, strCIF) = False Then
                            Exit Function
                        End If

                    End If



                    'telefonos
                    If CType(oform.Items.Item("43").Specific, SAPbouiCOM.EditText).Value.ToString = "" And CType(oform.Items.Item("45").Specific, SAPbouiCOM.EditText).Value.ToString = "" And CType(oform.Items.Item("51").Specific, SAPbouiCOM.EditText).Value.ToString = "" Then
                        sMensaje = "Debe introducir al menos un teléfono"
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    'email
                    If CType(oform.Items.Item("60").Specific, SAPbouiCOM.EditText).Value.ToString = "" And CType(oform.Items.Item("113").Specific, SAPbouiCOM.EditText).Value.ToString = "" Then
                        sMensaje = "Debe introducir al menos un email "
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    'comisionista
                    If CType(oform.Items.Item("52").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString = "-1" Then
                        sMensaje = "Debe introducir el dato de Comisionista 1 "
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    If CType(oform.Items.Item("228").Specific, SAPbouiCOM.ComboBox).Value.ToString = "-1" Or CType(oform.Items.Item("228").Specific, SAPbouiCOM.ComboBox).Value.ToString = "" Then
                        sMensaje = "Debe introducir el dato de Comisionista 2 "
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    'recorrer direciones... como minimo una direccion por defecto de cada tipo (envio y factura)
                    'ShipToDef
                    If oform.DataSources.DBDataSources.Item("OCRD").GetValue("ShipToDef", 0).Trim() = "" Then
                        sMensaje = "Debe fijar como estándar como mínimo una dirección de envío "
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    'BillToDef
                    If oform.DataSources.DBDataSources.Item("OCRD").GetValue("BillToDef", 0).Trim() = "" Then
                        sMensaje = "Debe fijar como estándar como mínimo una dirección de facturación "
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function

                    End If

                    Try
                        For i As Integer = 1 To CType(oform.Items.Item("69").Specific, SAPbouiCOM.Matrix).RowCount
                            sDato = CType(CType(oform.Items.Item("69").Specific, Matrix).Columns.Item("20").Cells.Item(i).Specific, EditText).Value
                            If sDato = "Destinatario de la factura" Or sDato = "Definir nuevo" Or sDato = "Enviar a" Then

                            Else
                                CType(oform.Items.Item("69").Specific, Matrix).Columns.Item("20").Cells.Item(i).Click(BoCellClickType.ct_Regular)

                                'en el item 178 tengo las direcciones
                                For j As Integer = 1 To CType(oform.Items.Item("178").Specific, SAPbouiCOM.Matrix).RowCount
                                    sDire = CType(CType(oform.Items.Item("178").Specific, Matrix).Columns.Item("1").Cells.Item(j).Specific, EditText).Value
                                    If sDire = "" Then
                                        sMensaje = "Debe introducir una ID de dirección "
                                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Function
                                    End If

                                    If CType(CType(oform.Items.Item("178").Specific, Matrix).Columns.Item("2").Cells.Item(j).Specific, EditText).Value = "" Then
                                        sMensaje = "Debe introducir una dirección "
                                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Function
                                    End If

                                    If CType(CType(oform.Items.Item("178").Specific, Matrix).Columns.Item("5").Cells.Item(j).Specific, EditText).Value = "" Then
                                        sMensaje = "Debe introducir un código postal "
                                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Function
                                    End If

                                    If CType(CType(oform.Items.Item("178").Specific, Matrix).Columns.Item("6").Cells.Item(j).Specific, EditText).Value = "" Then
                                        sMensaje = "Debe introducir una provincia "
                                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Function
                                    End If

                                    If CType(CType(oform.Items.Item("178").Specific, Matrix).Columns.Item("8").Cells.Item(j).Specific, ComboBox).Selected.Value.ToString = "ES" Then
                                        If CType(CType(oform.Items.Item("178").Specific, Matrix).Columns.Item("7").Cells.Item(j).Specific, ComboBox).Selected.Value.ToString = "" Then
                                            sMensaje = "Debe introducir el dato Estado cuando el país es España "
                                            objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Exit Function
                                        End If

                                    End If

                                    'If oform.DataSources.DBDataSources.Item("CRD1").GetValue("Address", intCont).Trim() = "" Then
                                    '    sMensaje = "Debe introducir una ID de dirección de " & sTipo
                                    '    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    '    Exit Function
                                    'End If

                                    'If oform.DataSources.DBDataSources.Item("CRD1").GetValue("Street", intCont).Trim() = "" Then
                                    '    sMensaje = "Debe introducir una dirección de " & sTipo
                                    '    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    '    Exit Function
                                    'End If

                                    'If oform.DataSources.DBDataSources.Item("CRD1").GetValue("ZipCode", intCont).Trim() = "" Then
                                    '    sMensaje = "Debe introducir un código postal de " & sTipo
                                    '    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    '    Exit Function
                                    'End If

                                    'If oform.DataSources.DBDataSources.Item("CRD1").GetValue("County", intCont).Trim() = "" Then
                                    '    sMensaje = "Debe introducir una población de " & sTipo
                                    '    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    '    Exit Function
                                    'End If

                                    'If oform.DataSources.DBDataSources.Item("CRD1").GetValue("Country", intCont).Trim() = "ES" Then
                                    '    If oform.DataSources.DBDataSources.Item("CRD1").GetValue("State", intCont).Trim() = "" Then
                                    '        sMensaje = "Debe introducir el dato Estado cuando el país es España en " & sTipo
                                    '        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    '        Exit Function
                                    '    End If
                                    'End If
                                Next

                            End If

                        Next
                    Catch ex As Exception

                    End Try

                    'condiciones de pago

                    If CType(oform.Items.Item("75").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString = "-1" Then
                        sMensaje = "Debe introducir una condición de pago "
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If


                    'intCuenta = CType(oform.Items.Item("217").Specific, SAPbouiCOM.Matrix).RowCount
                    'recorrer todas las vias de pago y marcarlas
                    intFolder = oform.PaneLevel

                    CType(oform.Items.Item("214").Specific, SAPbouiCOM.Folder).Select()
                    oform.Freeze(True)
                    Try
                        For i As Integer = 1 To CType(oform.Items.Item("217").Specific, SAPbouiCOM.Matrix).RowCount - 1
                            If CType(CType(oform.Items.Item("217").Specific, Matrix).Columns.Item("2").Cells.Item(i).Specific, EditText).Value <> "" Then
                                oChk = CType(CType(oform.Items.Item("217").Specific, SAPbouiCOM.Matrix).Columns.Item("5").Cells.Item(i).Specific, CheckBox)
                                oChk.Checked = True
                            Else
                                Exit For
                            End If
                        Next

                    Catch ex As Exception
                        oform.Freeze(False)
                    Finally
                        'oform.Freeze(False)
                    End Try


                    oform.Freeze(False)
                    oform.PaneLevel = intFolder

                    'via pago por defecto
                    sViaPago = oform.DataSources.DBDataSources.Item("OCRD").GetValue("PymCode", 0).Trim()
                    If sViaPago <> "" Then
                        'si es giro o recibo, obligar a rellenar la cuenta bancaria
                        sViaPago = objGlobal.refDi.SQL.sqlStringB1("SELECT T0.""Descript"" FROM OPYM T0 WHERE T0.""PayMethCod""='" & sViaPago & "'  ")
                        If sViaPago = "Giro" Or sViaPago = "Recibo" Then
                            If oform.DataSources.DBDataSources.Item("OCRD").GetValue("BankCode", 0).Trim() = "" Or oform.DataSources.DBDataSources.Item("OCRD").GetValue("BankCode", 0).Trim() = "-1" Then
                                sMensaje = "Cuando la forma de pago es Giro o Recibo es obligatorio introducir una cuenta bancaria desde la pestaña de condiciones de pago"
                                objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Function
                            End If
                        End If

                        'si es transferencia o transferencia por adelantado obligar a indicar un banco propio
                        If sViaPago = "Transferencia" Or sViaPago = "Transferencia por adelantado" Then
                            'HouseBank
                            If oform.DataSources.DBDataSources.Item("OCRD").GetValue("HouseBank", 0).Trim() = "" Or oform.DataSources.DBDataSources.Item("OCRD").GetValue("HouseBank", 0).Trim() = "-1" Then
                                sMensaje = "Cuando la forma de pago es Transferencia o Transferencia por adelantado es obligatorio introducir un banco propio"
                                objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Function
                            End If
                        End If
                    Else
                        sMensaje = "Debe fijar una via de pago por defecto"
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If


                    'al crear un IC marcar la facturacion inmediata, 

                    'y verificar que siempre una de las 3 está marcada y solo una
                    'T0."QryGroup1", T0."QryGroup2", T0."QryGroup3"
                    intNumProp = 0
                    'recorrer el matrix
                    For i As Integer = 1 To 3
                        If CType(CType(oform.Items.Item("136").Specific, Matrix).Columns.Item("2").Cells.Item(i).Specific, CheckBox).Checked = True Then
                            intNumProp = intNumProp + 1

                        End If
                    Next


                    If intNumProp > 1 Then
                        sMensaje = "Sólo debe marcar una de las tres Facturaciones: Facturación INMEDIATA, Facturación QUINCENAL o Facturación MENSUAL"
                        objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    Else
                        If intNumProp = 0 Then
                            sMensaje = "Debe marcar una de las tres Facturaciones: Facturación INMEDIATA, Facturación QUINCENAL o Facturación MENSUAL"
                            objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Function
                        End If
                    End If
                    oform.Freeze(False)
                    Comprobar_Datos = True
                End If

                'oform.Freeze(False)
                'Comprobar_Datos = True

            End If


        Catch ex As Exception
            oform.Freeze(False)
            Throw ex

        Finally
            oform.Freeze(False)
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

    Public Shared Function Comprobar_CIF_NIF(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sValor As String) As Boolean
        Comprobar_CIF_NIF = False
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Try
            oRs = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            'Validamos el CIF o NIF
            If oObjGlobal.compañia.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sSQL = "SELECT ""EXO_VALIDAR_NIF_CIF""(RTRIM(LTRIM('" & sValor & "'))) ""Es_CIFNIF_OK"" FROM DUMMY;"
            Else
                sSQL = "SELECT [dbo].[EXO_VALIDAR_NIF_CIF](RTRIM(LTRIM('" & sValor & "'))) ""Es_CIFNIF_OK"" "
            End If
            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                If CInt(oRs.Fields.Item("Es_CIFNIF_OK").Value.ToString) = 0 Then
                    oObjGlobal.SBOApp.StatusBar.SetText("(EXO) - El CIF/NIF " & sValor & " no es válido.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oObjGlobal.SBOApp.MessageBox("El CIF/NIF " & sValor & " no es válido.")
                    Exit Function
                End If
            Else
                Throw New Exception("No se ha encontrado función EXO_VALIDAR_NIF_CIF")
                Exit Function
            End If
            Comprobar_CIF_NIF = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function

End Class

Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class EXOSUBFAM
    Inherits EXO_UIAPI.EXO_DLLBase

    Private oCompanyService As CompanyService

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaDatos()
            cargaAutorizaciones()
        End If
    End Sub

#Region "Inicialización"

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Dim menuXML As Xml.XmlDocument = New Xml.XmlDocument
        Dim sPath As String = ""

        sPath = objGlobal.refDi.OGEN.pathGeneral & "\02.Menus" & "\XML_MENU.xml"
        menuXML.Load(sPath)
        Return menuXML

    End Function

    Private Sub cargaDatos()
        Dim sXML As String = ""
        Dim res As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then
            'UDO Tarificador 401
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_FAMSUBFAM.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO UDO_EXO_FAMSUBFAM", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            res = objGlobal.SBOApp.GetLastBatchResults
        End If
    End Sub

    Private Sub cargaAutorizaciones()
        Dim autorizacionXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_AUFAMSUBFAM.xml")
        objGlobal.refDi.comunes.LoadBDFromXML(autorizacionXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
    End Sub

#End Region
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "EXO-MnSubfam"
                        If CargarForm() = False Then
                            Return False
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
                        Case "EXOSUBFAM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_ComboSelect_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_Validate_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                Case SAPbouiCOM.BoEventTypes.et_CLICK
                                    If infoEvento.ItemUID = "grdSubfam" Then
                                        'cargar datos en la caja de abajo
                                        If EventHandler_Click_After(infoEvento) = False Then
                                            GC.Collect()
                                            Return False
                                        End If

                                    End If

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXOSUBFAM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT


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
                        Case "EXOSUBFAM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXOSUBFAM"
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
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btCon"
                    'cargar datos
                    CargarGrid(oForm)
                    oForm.Items.Item("btAdd").Enabled = True

                Case "btAdd"
                    AddDatos(oForm)
                    LimpiarDatos(oForm)
                    CargarGrid(oForm)
                Case "btDel"
                    DelDatos(oForm)
                    LimpiarDatos(oForm)
                Case "btSalir"
                    oForm.Close
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ComboSelect_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ComboSelect_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "cmbFam"
                    If pVal.ItemChanged = True Then
                        'cargargrid
                        CargarGrid(oForm)
                    End If

            End Select

            EventHandler_ComboSelect_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function

    Private Function EventHandler_Validate_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_Validate_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "txtCodSF" 'codigo subfamilia
                    If pVal.ItemChanged = True Then
                        'concateno en fabricante la descripion de la familia con el codigo de subfamilia
                        CType(oForm.Items.Item("txtFab").Specific, SAPbouiCOM.EditText).Value = CType(oForm.Items.Item("cmbFam").Specific, SAPbouiCOM.ComboBox).Selected.Description.ToString & "-" & CType(oForm.Items.Item("txtCodSF").Specific, SAPbouiCOM.EditText).Value
                    End If

            End Select

            EventHandler_Validate_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function

    Private Function EventHandler_Click_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim iRow As Integer = -1
        EventHandler_Click_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "grdSubfam"
                    'cargar datos
                    If CType(oForm.Items.Item("grdSubfam").Specific, SAPbouiCOM.Grid).Rows.SelectedRows.Count = 0 Then
                        objGlobal.SBOApp.MessageBox("Debe seleccionar un registro.")
                    Else
                        iRow = CType(oForm.Items.Item("grdSubfam").Specific, SAPbouiCOM.Grid).Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)
                        'cargar valores en las cajas.
                        CType(oForm.Items.Item("txtCodSF").Specific, SAPbouiCOM.EditText).Value = oForm.DataSources.DataTables.Item("dtDatos").GetValue("Subfamilia", iRow).ToString()
                        CType(oForm.Items.Item("txtDesSF").Specific, SAPbouiCOM.EditText).Value = oForm.DataSources.DataTables.Item("dtDatos").GetValue("Descripcion", iRow).ToString()
                        CType(oForm.Items.Item("txtFab").Specific, SAPbouiCOM.EditText).Value = oForm.DataSources.DataTables.Item("dtDatos").GetValue("Fabricante", iRow).ToString()

                    End If

            End Select

            EventHandler_Click_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Public Function CargarForm() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)
        Dim bt As SAPbouiCOM.Item = Nothing
        CargarForm = False

        Try
            '       UPDATE t1 SET
            '"U_EXO_DESLARGA" = t2."U_stec_famdesc"
            ' from "OITB" t1 INNER JOIN "@STEC_FAMSUBFAM" t2  ON t1."ItmsGrpCod" = cast(t2."U_stec_fam" as int)
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXOSUBFAM.srf")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try
            CargaComboFamilia(oForm)

            CargarForm = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function CargaComboFamilia(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaComboFamilia = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)
            sSQL = "SELECT T0.""ItmsGrpCod"", T0.""ItmsGrpNam"" FROM ""OITB"" T0 ORDER BY T0.""ItmsGrpCod"""
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cmbFam").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                'CType(oForm.Items.Item("cmbFam").Specific, SAPbouiCOM.ComboBox).Select("ItmsGrpNam", BoSearchKey.psk_ByValue)
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
            CargaComboFamilia = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Sub CargarGrid(ByRef oForm As SAPbouiCOM.Form)
        Dim sSql As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try

            oForm.Freeze(True)

            'cargo en la caja de texto la descripcion larga de la familia.
            sSql = "SELECT  T0.""U_EXO_DESLARGA"" FROM  ""OITB"" T0 where T0.""ItmsGrpCod"" = '" & CType(oForm.Items.Item("cmbFam").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
            oRs.DoQuery(sSql)
            If oRs.RecordCount > 0 Then
                CType(oForm.Items.Item("txtDesL").Specific, SAPbouiCOM.EditText).Value = oRs.Fields.Item("U_EXO_DESLARGA").Value.ToString()

            End If

            sSql = "Select T0.""DocEntry"", T0.""U_EXO_CODFAM"" ""Familia"" , T0.""U_EXO_CODSUBFAM"" ""Subfamilia"", T0.""U_EXO_DESSUBFAM"" ""Descripcion"", T0.""U_EXO_FABDES"" ""Fabricante"" " _
            & " FROM ""@EXO_FAMSUBFAM""  T0" _
            & " WHERE T0.""U_EXO_CODFAM"" ='" & CType(oForm.Items.Item("cmbFam").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' ORDER BY T0.""U_EXO_CODSUBFAM"""
            oForm.DataSources.DataTables.Item("dtDatos").ExecuteQuery(sSql)

            'oculto columnas docentry y codfam
            CType(oForm.Items.Item("grdSubfam").Specific, SAPbouiCOM.Grid).Columns.Item("DocEntry").Visible = False
            CType(oForm.Items.Item("grdSubfam").Specific, SAPbouiCOM.Grid).Columns.Item("Familia").Visible = False

            LimpiarDatos(oForm)

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)


        End Try
    End Sub

    Private Sub AddDatos(ByRef oForm As Form)
        Dim sSql As String = ""
        Dim oDI_SUBFAM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sDocEntry As String = ""
        Dim sFab As String = ""
        Dim oOMRC As SAPbobsCOM.Manufacturers = Nothing

        Try
            oDI_SUBFAM = New EXO_DIAPI.EXO_UDOEntity(objGlobal.refDi.comunes, "EXO_FAMSUBFAM") 'UDO 
            'oDI_SUBFAM = New EXO_DIAPI.EXO_UDOEntity(objGlobal.refDi.comunes, "EXO_FAMSUBFAM") 'UDO 
            If CType(oForm.Items.Item("txtCodSF").Specific, SAPbouiCOM.EditText).Value.ToString <> "" Then


                sSql = "SELECT * FROM ""@EXO_FAMSUBFAM"" WHERE ""U_EXO_CODFAM""='" & CType(oForm.Items.Item("cmbFam").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' and ""U_EXO_CODSUBFAM"" ='" & CType(oForm.Items.Item("txtCodSF").Specific, SAPbouiCOM.EditText).Value.ToString & "' "
                oRs.DoQuery(sSql)
                If oRs.RecordCount > 0 Then
                    oDI_SUBFAM.GetByKey(oRs.Fields.Item("DocEntry").Value.ToString)
                    oDI_SUBFAM.SetValue("U_EXO_DESSUBFAM") = CType(oForm.Items.Item("txtDesSF").Specific, SAPbouiCOM.EditText).Value.ToString
                    oDI_SUBFAM.SetValue("U_EXO_CODSUBFAM") = CType(oForm.Items.Item("txtCodSF").Specific, SAPbouiCOM.EditText).Value.ToString
                    oDI_SUBFAM.SetValue("U_EXO_DESFAM") = CType(oForm.Items.Item("txtDesL").Specific, SAPbouiCOM.EditText).Value.ToString
                    oDI_SUBFAM.SetValue("U_EXO_FABDES") = CType(oForm.Items.Item("txtFab").Specific, SAPbouiCOM.EditText).Value.ToString
                    If oDI_SUBFAM.UDO_Update = False Then
                        Throw New Exception("(EXO) - Error al acutalizar registro Fam: " & CType(oForm.Items.Item("cmbFam").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & " Sufam: " & CType(oForm.Items.Item("txtCodSF").Specific, SAPbouiCOM.EditText).Value.ToString & " Error: " & oDI_SUBFAM.GetLastError)
                    End If

                Else
                    oDI_SUBFAM.GetNew()
                    oDI_SUBFAM.SetValue("U_EXO_CODFAM") = CType(oForm.Items.Item("cmbFam").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString
                    oDI_SUBFAM.SetValue("U_EXO_CODSUBFAM") = CType(oForm.Items.Item("txtCodSF").Specific, SAPbouiCOM.EditText).Value.ToString
                    oDI_SUBFAM.SetValue("U_EXO_DESSUBFAM") = CType(oForm.Items.Item("txtDesSF").Specific, SAPbouiCOM.EditText).Value.ToString
                    oDI_SUBFAM.SetValue("U_EXO_DESFAM") = CType(oForm.Items.Item("txtDesL").Specific, SAPbouiCOM.EditText).Value.ToString
                    oDI_SUBFAM.SetValue("U_EXO_FABDES") = CType(oForm.Items.Item("txtFab").Specific, SAPbouiCOM.EditText).Value.ToString
                    If oDI_SUBFAM.UDO_Add = False Then
                        Throw New Exception("(EXO) - Error al añadir registro Fam: " & CType(oForm.Items.Item("cmbFam").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & " Sufam: " & CType(oForm.Items.Item("txtCodSF").Specific, SAPbouiCOM.EditText).Value.ToString & " Error: " & oDI_SUBFAM.GetLastError)
                    End If
                End If
                'comprobar si existe el fabricante en SAP, y sino crearlo.
                sSql = "SELECT T0.""FirmCode"", T0.""FirmName"" FROM OMRC T0 WHERE T0.""FirmName"" ='" & CType(oForm.Items.Item("txtFab").Specific, SAPbouiCOM.EditText).Value.ToString & "'  "
                oRs.DoQuery(sSql)
                If oRs.RecordCount = 0 Then
                    'lo creo
                    oOMRC = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oManufacturers), SAPbobsCOM.Manufacturers)
                    oOMRC.ManufacturerName = CType(oForm.Items.Item("txtFab").Specific, SAPbouiCOM.EditText).Value.ToString
                    If oOMRC.Add() <> 0 Then
                        Throw New Exception("(EXO) - Error al añadir registro Fabricante: " & objGlobal.compañia.GetLastErrorCode & " / " & objGlobal.compañia.GetLastErrorDescription)
                    End If

                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            If oOMRC IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOMRC)
            'If oDI_SUBFAM IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oDI_SUBFAM)
        End Try
    End Sub

    Private Sub DelDatos(ByRef oForm As Form)

        Dim oDI_SUBFAM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos
        Dim oGrid As SAPbouiCOM.Grid = Nothing
        Dim iRespuesta As Integer = 0
        Try
            oDI_SUBFAM = New EXO_DIAPI.EXO_UDOEntity(objGlobal.refDi.comunes, "EXO_FAMSUBFAM") 'UDO 
            oGrid = CType(oForm.Items.Item("grdSubfam").Specific, Grid)
            If oGrid.Rows.SelectedRows.Count > 0 Then
                Dim intSelRow As Integer = oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)
                Dim sValorGrid As String = CType(oGrid.DataTable.GetValue("DocEntry", intSelRow), String)
                'no dejar borrar si esta vinculado oitm
                iRespuesta = objGlobal.SBOApp.MessageBox("Se va a proceder a borrar la relación familia-subfamilia. ¿Desea continuar?", 2, "Ok", "Cancelar")
                If iRespuesta = 2 Then
                    Exit Sub
                Else
                    If oDI_SUBFAM.UDO_Delete(sValorGrid) = True Then
                        CargarGrid(oForm)
                    End If
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub LimpiarDatos(ByRef oForm As Form)
        Try

            CType(oForm.Items.Item("txtCodSF").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("txtDesSF").Specific, SAPbouiCOM.EditText).Value = ""
            CType(oForm.Items.Item("txtFab").Specific, SAPbouiCOM.EditText).Value = ""
        Catch ex As Exception

        End Try
    End Sub
End Class

Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System.IO

Public Class EXOACTDTOS
    Inherits EXO_UIAPI.EXO_DLLBase

    Private oCompanyService As CompanyService

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then

            'cargaAutorizaciones()
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
                    Case "EXO-MnActDtos"
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
                        Case "EXOACTDTOS"
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

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                Case SAPbouiCOM.BoEventTypes.et_CLICK


                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXOACTDTOS"
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
                        Case "EXOACTDTOS"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXOACTDTOS"
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
                Case "btRev"
                    RevisionDatos(oForm)
                Case "btAct"
                    AddDatos(oForm)
                    'Arreglar(oForm)
                Case "btSalir"
                    oForm.Close()
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
                        'cargar subfamlia
                        CargaComboSubFam(oForm)

                    End If
                Case "cmbSubFam"
                    If pVal.ItemChanged = True Then
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
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXOACTDTOS.srf")

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

    Private Function CargaComboSubFam(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaComboSubFam = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sCodArt As String = ""

        Try
            'cargo en la caja de texto la descripcion larga de la familia.
            sSQL = "SELECT  T0.""U_EXO_DESLARGA"" FROM  ""OITB"" T0 where T0.""ItmsGrpCod"" = '" & CType(oForm.Items.Item("cmbFam").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                CType(oForm.Items.Item("txtDesL").Specific, SAPbouiCOM.EditText).Value = oRs.Fields.Item("U_EXO_DESLARGA").Value.ToString()
            End If

            sSQL = ""

            sSQL = "SELECT t1.""FirmCode"", T0.""U_EXO_FABDES"" || '-' || T0.""U_EXO_DESSUBFAM"" ""Subfamlia"",T0.""U_EXO_FABDES"" FROM ""@EXO_FAMSUBFAM""  T0 " _
                & " INNER JOIN ""OMRC"" T1 ON T0.""U_EXO_FABDES"" = T1.""FirmName""   " _
                & " WHERE ""U_EXO_CODFAM""='" & CType(oForm.Items.Item("cmbFam").Specific, SAPbouiCOM.ComboBox).Selected.Value & "'" _
                & " UNION ALL SELECT '-1' , '','' FROM ""DUMMY"" "
            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                If CType(oForm.Items.Item("cmbFam").Specific, SAPbouiCOM.ComboBox).Selected.Value <> "" Then
                    CType(oForm.Items.Item("cmbSubFam").Specific, SAPbouiCOM.ComboBox).Select("", BoSearchKey.psk_ByValue)
                    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cmbSubFam").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

                End If
            End If
            'cargar fabircante
            CargaComboSubFam = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Sub CargarGrid(ByRef oForm As SAPbouiCOM.Form)
       Dim sSql As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try

            oForm.Freeze(True)
            ' --& " And T1.""ObjType"" =43 " _
            'cargo en la caja de texto la descripcion larga de la familia.
            sSql = "SELECT T0.""AbsEntry"",T1.""ObjType"", T1.""ObjKey"" ,T0.""ObjCode"" ""Cliente"" ,T2.""FirmName"" ""Fabricante"", T1.""Discount"" ""Descuento"" " _
                & "  FROM ""OEDG"" T0  " _
                & " INNER JOIN ""EDG1"" T1 On T0.""AbsEntry"" = T1.""AbsEntry"" " _
                & " INNER Join ""OMRC"" T2 ON T1.""ObjKey"" = T2.""FirmCode"" " _
                & " INNER JOIN ""OCRD"" T3 ON T0.""ObjCode"" = T3.""CardCode""" _
                & " WHERE T0.""ObjType"" = 2 " _
                & " And T1.""ObjType"" =43 " _
                & " and T2.""FirmCode"" = '" & CType(oForm.Items.Item("cmbSubFam").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' " _
                & " and T3.""CardType"" = 'C' " _
                & " order by T0.""ObjCode"",T2.""FirmName"" "

            oForm.DataSources.DataTables.Item("dtDatos").ExecuteQuery(sSql)
            'oculto columnas docentry y codfam
            CType(oForm.Items.Item("grdDatos").Specific, SAPbouiCOM.Grid).Columns.Item("AbsEntry").Visible = False
            CType(oForm.Items.Item("grdDatos").Specific, SAPbouiCOM.Grid).Columns.Item("ObjType").Visible = False
            CType(oForm.Items.Item("grdDatos").Specific, SAPbouiCOM.Grid).Columns.Item("ObjKey").Visible = False



            'LimpiarDatos(oForm)

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)


        End Try
    End Sub
    Private Sub Arreglar(ByRef oForm As Form)

        Dim dblPorcenAnt As Double
        Dim dblPorcenN As Double
        Dim sError As String = ""
        Dim sCodCli As String = ""
        Dim sFabri As String = "517"
        Try


            'RECORRER

            'objGlobal.SBOApp.StatusBar.SetText("Actualizando registro " & i + 1 & " de " & oForm.DataSources.DataTables.Item("dtDatos").Rows.Count & " Código cliente: " & oForm.DataSources.DataTables.Item("dtDatos").GetValue("Cliente", i).ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Dim oBP As SAPbobsCOM.BusinessPartners = Nothing

            oBP = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), BusinessPartners)
            sCodCli = "C0368"
            If (oBP.GetByKey(sCodCli)) = True Then
                If oBP.Frozen = BoYesNoEnum.tNO Then


                    For J As Integer = oBP.DiscountGroups.Count - 1 To 0 Step -1
                        'oDiscountGroup = oBP.DiscountGroups
                        oBP.DiscountGroups.SetCurrentLine(J)
                        If oBP.DiscountGroups.BPCode = sCodCli Then
                            'BORRAR
                            If oBP.DiscountGroups.ObjectEntry = sFabri Then
                                dblPorcenAnt = oBP.DiscountGroups.DiscountPercentage
                                oBP.DiscountGroups.Delete()
                                Exit For
                            End If
                        End If
                    Next
                    'añadir
                    dblPorcenN = 25
                    If (dblPorcenN + dblPorcenAnt) > 0 Then
                        'oDiscountGroup = oBP.DiscountGroups
                        oBP.DiscountBaseObject = DiscountGroupBaseObjectEnum.dgboManufacturer
                        oBP.DiscountGroups.ObjectEntry = sFabri
                        oBP.DiscountGroups.DiscountPercentage = dblPorcenN + dblPorcenAnt

                        oBP.DiscountGroups.Add()
                    End If
                End If

            End If
            'ACTAULIZADO REGISTRO 
            'If oBP.Frozen = BoYesNoEnum.tYES Then
            '    oBP.Frozen = BoYesNoEnum.tNO
            'End If

            If oBP.Update() <> 0 Then
                    sError = sError & " Código cliente " & oBP.CardCode & vbCrLf
                    objGlobal.SBOApp.StatusBar.SetText("Error: " & compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
                'oDiscountGroup = Nothing
                oBP = Nothing

            If sError <> "" Then
                objGlobal.SBOApp.MessageBox("Proceso realizado" & " Clientes sin poder actualizar: " & sError)
            Else
                objGlobal.SBOApp.MessageBox("Proceso realizado correctamente")
            End If

            CargarGrid(oForm)
        Catch ex As Exception
            objGlobal.SBOApp.StatusBar.SetText("Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Throw ex
        Finally
            ' If oDiscountGroup IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oDiscountGroup)
            'If oBP IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBP)

        End Try
    End Sub
    Private Sub AddDatos(ByRef oForm As Form)
        Dim oBP As SAPbobsCOM.BusinessPartners = Nothing
        Dim dblPorcenAnt As Double
        Dim dblPorcenN As Double
        Dim sError As String = ""
        Dim sCodCli As String = ""
        Dim sErrorSAP As String = ""
        Try


            'RECORRER
            dblPorcenN = CDbl(oForm.DataSources.UserDataSources.Item("dsPorcen").Value)
            For i As Integer = 0 To oForm.DataSources.DataTables.Item("dtDatos").Rows.Count - 1
                objGlobal.SBOApp.StatusBar.SetText("Actualizando registro " & i + 1 & " de " & oForm.DataSources.DataTables.Item("dtDatos").Rows.Count & " Código cliente: " & oForm.DataSources.DataTables.Item("dtDatos").GetValue("Cliente", i).ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                oBP = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), BusinessPartners)
                sCodCli = oForm.DataSources.DataTables.Item("dtDatos").GetValue("Cliente", i).ToString
                If (oBP.GetByKey(sCodCli)) = True Then
                    For J As Integer = oBP.DiscountGroups.Count - 1 To 0 Step -1
                        'oDiscountGroup = oBP.DiscountGroups
                        oBP.DiscountGroups.SetCurrentLine(J)
                        'BORRAR
                        If oBP.DiscountGroups.ObjectEntry = oForm.DataSources.DataTables.Item("dtDatos").GetValue("ObjKey", i).ToString.Trim And oBP.DiscountGroups.BaseObjectType = DiscountGroupBaseObjectEnum.dgboManufacturer Then
                            oBP.DiscountGroups.SetCurrentLine(J)
                            dblPorcenAnt = oBP.DiscountGroups.DiscountPercentage
                            oBP.DiscountBaseObject = DiscountGroupBaseObjectEnum.dgboManufacturer
                            oBP.DiscountGroups.Delete()
                            If oBP.Update() <> 0 Then
                                sError = sError & " Código cliente " & oBP.CardCode & vbCrLf
                                sErrorSAP = compañia.GetLastErrorDescription
                                objGlobal.SBOApp.StatusBar.SetText("Error: " & compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            End If
                            Exit For
                        End If
                    Next
                    'añadir

                    If (dblPorcenN + dblPorcenAnt) > 0 Then
                        oBP.DiscountGroups.Add()
                        'oDiscountGroup = oBP.DiscountGroups
                        oBP.DiscountBaseObject = DiscountGroupBaseObjectEnum.dgboManufacturer
                        oBP.DiscountGroups.ObjectEntry = oForm.DataSources.DataTables.Item("dtDatos").GetValue("ObjKey", i).ToString.Trim
                        oBP.DiscountGroups.DiscountPercentage = dblPorcenN + dblPorcenAnt
                        oBP.DiscountGroups.Add()
                        If oBP.Update() <> 0 Then
                            sError = sError & " Código cliente " & oBP.CardCode & vbCrLf
                            sErrorSAP = compañia.GetLastErrorDescription
                            objGlobal.SBOApp.StatusBar.SetText("Error: " & compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    End If

                End If
                'ACTAULIZADO REGISTRO 
                'If oBP.Frozen = BoYesNoEnum.tYES Then
                '    oBP.Frozen = BoYesNoEnum.tNO
                'End If
                If oBP.Update() <> 0 Then
                    sError = sError & " Código cliente " & oBP.CardCode & vbCrLf
                    sErrorSAP = compañia.GetLastErrorDescription
                    objGlobal.SBOApp.StatusBar.SetText("Error: " & compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
                'oDiscountGroup = Nothing
                oBP = Nothing
            Next
            If sError <> "" Then
                objGlobal.SBOApp.MessageBox("Proceso realizado" & " Clientes sin poder actualizar: " & sError)
            Else
                objGlobal.SBOApp.MessageBox("Proceso realizado correctamente")
            End If

            CargarGrid(oForm)
        Catch ex As Exception
            objGlobal.SBOApp.StatusBar.SetText("Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Throw ex
        Finally
            'If oDiscountGroup IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oDiscountGroup)
            If oBP IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBP)

        End Try
    End Sub

    Private Sub RevisionDatos(ByRef oForm As Form)

        Dim dblPorcenAnt As Double
        Dim dblPorcenN As Double = 5
        Dim sError As String = ""
        Dim sCodCli As String = ""
        Dim sSql As String = ""
        Dim oBP As SAPbobsCOM.BusinessPartners = Nothing
        Dim sFabri As String = "524"
        Dim sFabriEO As String = "517"
        Dim ruta As String = "C:\Tiara\"
        ':::Nombre del archivo
        Dim archivo As String = "Clientes error.txt"
        Dim fs As FileStream
        Dim sErrorTran As String = ""
        Try
            If File.Exists(ruta) Then

                ':::Si la carpeta existe creamos o sobreescribios el archivo txt
                fs = File.Create(ruta & archivo)
                fs.Close()

            Else

                ':::Si la carpeta no existe la creamos
                Directory.CreateDirectory(ruta)

                ':::Una vez creada la carpeta creamos o sobreescribios el archivo txt
                fs = File.Create(ruta & archivo)
                fs.Close()

            End If
            Dim escribir As New StreamWriter(ruta & archivo)


            'RECORRER
            dblPorcenN = 5
            'If objGlobal.compañia.InTransaction = True Then
            '    objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            'End If

            'objGlobal.compañia.StartTransaction()
            oBP = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), BusinessPartners)

            ' And T0.""ObjCode""='C0004'
            'and T1.""ObjKey""=517

            sSql = "SELECT distinct  T0.""ObjCode"" FROM ""OEDG""  T0  INNER JOIN ""EDG1"" T1 ON T0.""AbsEntry"" = T1.""AbsEntry""  " _
                & " INNER JOIN ""OCRD"" T2 ON T0.""ObjCode"" = T2.""CardCode""" _
            & " WHERE T0.""ObjType""= 2 AND  T1.""ObjType"" = 43 and T2.""CardType"" = 'C'  and T1.""ObjKey""<> " & sFabri & " " _
            & " ORDER BY T0.""ObjCode"" "

            oForm.DataSources.DataTables.Item("dtDatos2").ExecuteQuery(sSql)


            For i As Integer = 0 To oForm.DataSources.DataTables.Item("dtDatos2").Rows.Count - 1
                objGlobal.SBOApp.StatusBar.SetText("Actualizando registro " & i + 1 & " de " & oForm.DataSources.DataTables.Item("dtDatos2").Rows.Count & " Código cliente: " & oForm.DataSources.DataTables.Item("dtDatos2").GetValue("ObjCode", i).ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


                sCodCli = oForm.DataSources.DataTables.Item("dtDatos2").GetValue("ObjCode", i).ToString
                If (oBP.GetByKey(sCodCli)) = True Then

                    'añadir
                    If (dblPorcenN + dblPorcenAnt) > 0 Then
                        'oDiscountGroup = oBP.DiscountGroups
                        oBP.DiscountBaseObject = DiscountGroupBaseObjectEnum.dgboManufacturer
                        oBP.DiscountGroups.ObjectEntry = sFabri 'codigo silvia
                        oBP.DiscountGroups.DiscountPercentage = dblPorcenN
                        oBP.DiscountGroups.Add()
                    End If



                    If oBP.Update() <> 0 Then
                        escribir.WriteLine(oBP.CardCode)
                        sError = sError & " Código cliente " & oBP.CardCode & vbCrLf

                        objGlobal.SBOApp.StatusBar.SetText("Error: " & compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        'compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If

                End If

                'oDiscountGroup = Nothing
                'oBP = Nothing
            Next

            If sError = "" Then

                'If objGlobal.compañia.InTransaction = True Then
                '    objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                'End If
            Else
                'If objGlobal.compañia.InTransaction = True Then
                '    compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                'End If
            End If

            escribir.Close()

            If sError <> "" Then
                objGlobal.SBOApp.MessageBox("Proceso realizado con errores. Consulte fichero detalle error clientes")
            Else
                objGlobal.SBOApp.MessageBox("Proceso realizado correctamente")
            End If

            CargarGrid(oForm)
        Catch exCOM As System.Runtime.InteropServices.COMException
            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objGlobal.SBOApp.StatusBar.SetText("Error: " & exCOM.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Throw exCOM
            End If
        Catch ex As Exception
            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                'End If
                objGlobal.SBOApp.StatusBar.SetText("Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If
            Throw ex
        Finally

            If oBP IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBP)

        End Try
    End Sub

    'ARREGLAR


    'Private Sub AddDatos(ByRef oForm As Form)

    '    Dim dblPorcenAnt As Double
    '    Dim dblPorcenN As Double
    '    Dim sError As String = ""
    '    Dim sCodCli As String = ""
    '    Try


    '        'RECORRER
    '        dblPorcenN = CDbl(oForm.DataSources.UserDataSources.Item("dsPorcen").Value)
    '        For i As Integer = 0 To oForm.DataSources.DataTables.Item("dtDatos").Rows.Count - 1
    '            objGlobal.SBOApp.StatusBar.SetText("Actualizando registro " & i + 1 & " de " & oForm.DataSources.DataTables.Item("dtDatos").Rows.Count & " Código cliente: " & oForm.DataSources.DataTables.Item("dtDatos").GetValue("Cliente", i).ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '            Dim oBP As SAPbobsCOM.BusinessPartners = Nothing
    '            Dim oDiscountGroup As SAPbobsCOM.DiscountGroups = Nothing

    '            oBP = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), BusinessPartners)
    '            sCodCli = oForm.DataSources.DataTables.Item("dtDatos").GetValue("Cliente", i).ToString
    '            If (oBP.GetByKey(sCodCli)) = True Then
    '                For J As Integer = oBP.DiscountGroups.Count - 1 To 0 Step -1
    '                    'oDiscountGroup = oBP.DiscountGroups
    '                    oDiscountGroup.SetCurrentLine(J)
    '                    If oDiscountGroup.BPCode = sCodCli Then
    '                        'BORRAR
    '                        If oDiscountGroup.ObjectEntry = oForm.DataSources.DataTables.Item("dtDatos").GetValue("ObjKey", i).ToString Then
    '                            dblPorcenAnt = oDiscountGroup.DiscountPercentage
    '                            oDiscountGroup.Delete()
    '                            Exit For
    '                        End If
    '                    End If
    '                Next
    '                'añadir
    '                If (dblPorcenN + dblPorcenAnt) > 0 Then
    '                    oDiscountGroup = oBP.DiscountGroups
    '                    oBP.DiscountBaseObject = DiscountGroupBaseObjectEnum.dgboManufacturer
    '                    oDiscountGroup.ObjectEntry = oForm.DataSources.DataTables.Item("dtDatos").GetValue("ObjKey", i).ToString
    '                    oDiscountGroup.DiscountPercentage = dblPorcenN + dblPorcenAnt

    '                    oDiscountGroup.Add()
    '                End If

    '            End If
    '            'ACTAULIZADO REGISTRO 
    '            'If oBP.Frozen = BoYesNoEnum.tYES Then
    '            '    oBP.Frozen = BoYesNoEnum.tNO
    '            'End If
    '            If oBP.Update() <> 0 Then
    '                sError = sError & " Código cliente " & oBP.CardCode & vbCrLf
    '                objGlobal.SBOApp.StatusBar.SetText("Error: " & compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '            End If
    '            oDiscountGroup = Nothing
    '            oBP = Nothing
    '        Next
    '        If sError <> "" Then
    '            objGlobal.SBOApp.MessageBox("Proceso realizado" & " Clientes sin poder actualizar: " & sError)
    '        Else
    '            objGlobal.SBOApp.MessageBox("Proceso realizado correctamente")
    '        End If

    '        CargarGrid(oForm)
    '    Catch ex As Exception
    '        objGlobal.SBOApp.StatusBar.SetText("Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '        Throw ex
    '    Finally
    '        ' If oDiscountGroup IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oDiscountGroup)
    '        'If oBP IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBP)

    '    End Try
    'End Sub
End Class

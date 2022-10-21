Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OITM
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaDatos()

        End If
    End Sub
    Private Sub cargaDatos()
        Dim sXML As String = ""
        Dim res As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OITM.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OITM", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_OADMINTER.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_OADMINTER", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            res = objGlobal.SBOApp.GetLastBatchResults


        End If
    End Sub
    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
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
                        Case "150"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_ComboSelect_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

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
                        Case "150"

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
                        Case "150"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "150"

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
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function

    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument
        Dim sItemCode As String = ""

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "150"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "150"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If infoEvento.ActionSuccess Then

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess Then


                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If infoEvento.ActionSuccess Then
                                    oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

                                    'If oForm.Visible = True Then
                                    If CargaComboSubFam(oForm) = False Then
                                        Return False
                                    End If

                                    ' End If
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
    Private Function EventHandler_Form_Load(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oItem As SAPbouiCOM.Item

        Dim Path As String = ""
        Dim XmlDoc As New System.Xml.XmlDocument


        EventHandler_Form_Load = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            'oForm.Visible = True
            If pVal.ActionSuccess = False Then
                objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                '
                oItem = oForm.Items.Add("cmbSubFam", BoFormItemTypes.it_COMBO_BOX)
                oItem.Top = oForm.Items.Item("39").Top
                oItem.Left = oForm.Items.Item("107").Left
                oItem.Height = oForm.Items.Item("39").Height
                oItem.Width = oForm.Items.Item("39").Width + 50
                oItem.LinkTo = "107"
                oItem.FromPane = 0
                oItem.ToPane = 0


                CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "OITM", "U_EXO_SUBFAM")
                CType(oItem.Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_DescriptionOnly
                CType(oItem.Specific, SAPbouiCOM.ComboBox).Item.DisplayDesc = True

                oItem = oForm.Items.Add("lblSubFam", BoFormItemTypes.it_STATIC)
                oItem.Top = oForm.Items.Item("40").Top
                oItem.Left = oForm.Items.Item("106").Left
                oItem.Height = oForm.Items.Item("106").Height
                oItem.Width = oForm.Items.Item("106").Width
                oItem.LinkTo = "106"
                oItem.FromPane = 0
                oItem.ToPane = 0
                CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Subfamilia"

            End If

            EventHandler_Form_Load = True


        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            'oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function


    Private Function EventHandler_ComboSelect_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ComboSelect_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "39"
                    If pVal.ItemChanged = True Then
                        'cargarSUBFAMLIAS
                        CargaComboSubFam(oForm)
                        'AsignarPropiedad(oForm)
                    End If
                Case "cmbSubFam"
                    If pVal.ItemChanged = True Then
                        'cargarFABRICANTE
                        If oForm.Mode <> BoFormMode.fm_OK_MODE Then
                            CargaComboFab(oForm)
                        End If

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
    Private Function AsignarPropiedad(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim sCodArt As String = ""
        Dim sGrupoArt As String = ""
        Dim sSql As String = ""
        Dim sPropiedad As String = ""
        Dim oChk As CheckBox
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        AsignarPropiedad = False

        Try
            sGrupoArt = CType(oForm.Items.Item("39").Specific, SAPbouiCOM.ComboBox).Selected.Value
            sSql = "SELECT T0.""ItmsGrpCod"", T0.""ItmsGrpNam"",T1.""U_EXO_PROPIEDAD"" 
            FROM OITB T0 
            LEFT OUTER JOIN ""@EXO_TIPOFAM""  T1 ON T0.""U_EXO_TIPFAM"" = T1.""Code""
            WHERE T0.""ItmsGrpCod""='" & sGrupoArt & "'
            ORDER BY T0.""ItmsGrpNam"""
            oRs.DoQuery(sSql)
            If oRs.RecordCount > 0 Then
                sPropiedad = oRs.Fields.Item("U_EXO_PROPIEDAD").Value.ToString
                If sPropiedad <> "" Then
                    'marcar propiead
                    'RECORRER MATRIX
                    CType(oForm.Items.Item("11").Specific, SAPbouiCOM.Folder).Select()
                    For i As Integer = 1 To CType(oForm.Items.Item("129").Specific, SAPbouiCOM.Matrix).RowCount
                        If CInt(sPropiedad) = i Then
                            oChk = CType(CType(oForm.Items.Item("129").Specific, SAPbouiCOM.Matrix).Columns.Item("2").Cells.Item(i).Specific, CheckBox)
                            oChk.Checked = True
                        End If
                    Next
                End If
            End If


            AsignarPropiedad = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function CargaComboSubFam(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaComboSubFam = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sCodArt As String = ""

        Try

            sSQL = "SELECT t0.""DocEntry"", T0.""U_EXO_FABDES"" || '-' || T0.""U_EXO_DESSUBFAM"" ""Subfamlia"",T0.""U_EXO_FABDES"" FROM ""@EXO_FAMSUBFAM""  T0" _
                & " WHERE ""U_EXO_CODFAM""='" & CType(oForm.Items.Item("39").Specific, SAPbouiCOM.ComboBox).Selected.Value & "'" _
                & " UNION ALL SELECT '-1' , '','' FROM ""DUMMY"" "
            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                If CType(oForm.Items.Item("39").Specific, SAPbouiCOM.ComboBox).Selected.Value <> "" Then


                    ' CType(oForm.Items.Item("cmbSubFam").Specific, SAPbouiCOM.ComboBox).Select("", BoSearchKey.psk_ByValue)
                    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cmbSubFam").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)

                    sCodArt = CType(oForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value
                    'If CType(oForm.Items.Item("cmbSubFam").Specific, SAPbouiCOM.ComboBox).Selected.Value <> "" Or CType(oForm.Items.Item("cmbSubFam").Specific, SAPbouiCOM.ComboBox).Selected.Value <> "-1" Then
                    '    sSQL = "Select T0.""FirmCode"", T0.""FirmName"" FROM OMRC T0 WHERE T0.""FirmName""='" & oRs.Fields.Item("U_EXO_FABDES").Value.ToString & "'"
                    '    oRs.DoQuery(sSQL)
                    '    If oRs.RecordCount > 0 Then
                    '        If oRs.Fields.Item("FirmCode").Value.ToString <> "" Then
                    '            CType(oForm.Items.Item("114").Specific, SAPbouiCOM.ComboBox).Select(oRs.Fields.Item("FirmCode").Value.ToString, BoSearchKey.psk_ByValue)
                    '        End If
                    '    Else

                    '        CType(oForm.Items.Item("114").Specific, SAPbouiCOM.ComboBox).Select("-1", BoSearchKey.psk_ByValue)

                    '    End If
                    'Else
                    '    CType(oForm.Items.Item("114").Specific, SAPbouiCOM.ComboBox).Select("-1", BoSearchKey.psk_ByValue)
                    'End If
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

    Private Function CargaComboFab(ByRef oForm As SAPbouiCOM.Form) As Boolean
        CargaComboFab = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try

            sSQL = "SELECT t0.""DocEntry"", T0.""U_EXO_FABDES"" || '-' || T0.""U_EXO_DESSUBFAM"" ""Subfamlia"",T0.""U_EXO_FABDES"" FROM ""@EXO_FAMSUBFAM""  T0" _
                & " WHERE  T0.""DocEntry"" ='" & CType(oForm.Items.Item("cmbSubFam").Specific, SAPbouiCOM.ComboBox).Selected.Value & "'"
            oRs.DoQuery(sSQL)

            If oRs.RecordCount >= 0 Then

                If oRs.Fields.Item("U_EXO_FABDES").Value.ToString <> "" Then
                    CType(oForm.Items.Item("114").Specific, SAPbouiCOM.ComboBox).Select(oRs.Fields.Item("U_EXO_FABDES").Value.ToString, BoSearchKey.psk_ByValue)
                End If

            End If

            'cargar fabircante
            CargaComboFab = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
End Class

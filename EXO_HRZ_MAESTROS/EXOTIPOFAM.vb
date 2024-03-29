﻿Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class EXOTIPOFAM
    Inherits EXO_UIAPI.EXO_DLLBase
    Dim bolCargarCombo As Boolean = False
    Private oCompanyService As CompanyService

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            'cargaDatos()
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

    'Private Sub cargaDatos()
    '    Dim sXML As String = ""
    '    Dim res As String = ""

    '    If objGlobal.refDi.comunes.esAdministrador Then

    '        sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_TIPOFAM.xml")
    '        objGlobal.refDi.comunes.LoadBDFromXML(sXML)
    '        objGlobal.SBOApp.StatusBar.SetText("Validando: UDO UDO_EXO_TIPOFAM", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '        res = objGlobal.SBOApp.GetLastBatchResults
    '    End If
    'End Sub

    'Private Sub cargaAutorizaciones()
    '    Dim autorizacionXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_AUTIPOFAM.xml")
    '    objGlobal.refDi.comunes.LoadBDFromXML(autorizacionXML)
    '    Dim res As String = objGlobal.SBOApp.GetLastBatchResults
    'End Sub

#End Region
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then

            Else

                Select Case infoEvento.MenuUID
                    Case "EXO-MnTipoFam"
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
            Dim oForm As SAPbouiCOM.Form = Nothing

            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_TIPOFAM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

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
                        Case "UDO_FT_EXO_TIPOFAM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_TIPOFAM"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                    If oForm.Visible = True Then
                                        bolCargarCombo = False
                                        CargaComboTipoFamilia(oForm)
                                        bolCargarCombo = True
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    'If infoEvento.ActionSuccess Then
                                    '    oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                    '    If oForm.Visible = True Then

                                    '        CargaComboTipoFamilia(oForm)

                                    '    End If
                                    'End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_TIPOFAM"
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
    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument
        Dim sItemCode As String = ""

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_TIPOFAM"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_TIPOFAM"
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
                                    CargaComboTipoFamilia(oForm)

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
    Public Function CargarForm() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)
        Dim bt As SAPbouiCOM.Item = Nothing
        CargarForm = False

        Try

            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "UDO_EXO_TIPOFAM.srf")

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
            CargaComboTipoFamilia(oForm)

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

    Private Function CargaComboTipoFamilia(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaComboTipoFamilia = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oItem As SAPbouiCOM.Item
        Dim oEdit As SAPbouiCOM.ComboBox

        Try
            If oForm.ActiveItem <> "" And bolCargarCombo = False Then

                bolCargarCombo = True
                oForm.Freeze(True)
                'crear el combo por desarrollo

                'propiedad
                Try


                    oItem = oForm.Items.Add("13_U_S", BoFormItemTypes.it_STATIC)
                    oItem.Top = oForm.Items.Item("0_U_S").Top + oForm.Items.Item("0_U_S").Height + 1
                    oItem.Left = oForm.Items.Item("0_U_S").Left
                    oItem.Height = oForm.Items.Item("0_U_S").Height
                    oItem.Width = oForm.Items.Item("0_U_S").Width
                    CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Propiedad"
                Catch ex As Exception

                End Try

                oItem = oForm.Items.Add("cmbProp", BoFormItemTypes.it_COMBO_BOX)
                oItem.Top = oForm.Items.Item("13_U_S").Top
                oItem.Left = oForm.Items.Item("0_U_E").Left
                oItem.Height = oForm.Items.Item("0_U_E").Height
                oItem.Width = oForm.Items.Item("0_U_E").Width * 2

                oItem.FromPane = 0
                oItem.ToPane = 0
                oItem.Enabled = True
                oItem.DisplayDesc = True
                CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "@EXO_TIPOFAM", "U_EXO_PROPIEDAD")


                ' oItem = oForm.Items.Item(1)

                oEdit = CType(oItem.Specific, ComboBox)
                oEdit.TabOrder = 2

                oForm.Items.Item("13_U_S").LinkTo = "cmbProp"


                sSQL = "SELECT T0.""ItmsTypCod"", T0.""ItmsGrpNam"" FROM OITG T0 ORDER BY T0.""ItmsTypCod"""
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cmbProp").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                    oCombo = CType(oForm.Items.Item("cmbProp").Specific, ComboBox)
                    oCombo.ExpandType = BoExpandType.et_DescriptionOnly

                Else
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If

                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_View, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                CargaComboTipoFamilia = True

            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.FormCombo(oCombo)
        End Try
    End Function


End Class

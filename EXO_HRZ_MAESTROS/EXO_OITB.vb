Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OITB
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaDatos()

        End If
    End Sub

#Region "SAP"

#End Region
    Private Sub cargaDatos()
        Dim sXML As String = ""
        Dim res As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_TIPOFAM.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO UDO_EXO_TIPOFAM", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OITB.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OITB", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
                        Case "63"
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
                        Case "63"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK
                                    If EventHandler_CLICK_BEFORE(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    Dim S As String = ""

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "63"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "63"

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
    Private Function EventHandler_Form_Load(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oItem As SAPbouiCOM.Item
        Dim Path As String = ""
        Dim XmlDoc As New System.Xml.XmlDocument


        EventHandler_Form_Load = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oForm.Freeze(True)
            If pVal.ActionSuccess = False Then
                objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                '
                oItem = oForm.Items.Add("txtDesL", BoFormItemTypes.it_EDIT)
                oItem.Top = oForm.Items.Item("10002024").Top + 15
                oItem.Left = oForm.Items.Item("10002024").Left
                oItem.Height = oForm.Items.Item("10002024").Height
                oItem.Width = oForm.Items.Item("10002024").Width
                oItem.FromPane = 1
                oItem.ToPane = 1


                CType(oItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "OITB", "U_EXO_DESLARGA")

                oItem = oForm.Items.Add("lblDesL", BoFormItemTypes.it_STATIC)
                oItem.Top = oForm.Items.Item("txtDesL").Top
                oItem.Left = oForm.Items.Item("10002023").Left
                oItem.Height = oForm.Items.Item("10002023").Height
                oItem.Width = oForm.Items.Item("10002023").Width
                oItem.LinkTo = "txtDesL"
                oItem.FromPane = 1
                oItem.ToPane = 1
                CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Descripción larga"

                'campo tipo familia
                oItem = oForm.Items.Add("cmbTipFam", BoFormItemTypes.it_COMBO_BOX)
                oItem.Top = oForm.Items.Item("txtDesL").Top + 15
                oItem.Left = oForm.Items.Item("txtDesL").Left
                oItem.Height = oForm.Items.Item("txtDesL").Height
                oItem.Width = oForm.Items.Item("txtDesL").Width
                'oItem.LinkTo = "107"
                oItem.FromPane = 1
                oItem.ToPane = 1


                CType(oItem.Specific, SAPbouiCOM.ComboBox).DataBind.SetBound(True, "OITB", "U_EXO_TIPFAM")
                CType(oItem.Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_DescriptionOnly
                CType(oItem.Specific, SAPbouiCOM.ComboBox).Item.DisplayDesc = True

                oItem = oForm.Items.Add("lblTipFam", BoFormItemTypes.it_STATIC)
                oItem.Top = oForm.Items.Item("cmbTipFam").Top
                oItem.Left = oForm.Items.Item("lblDesL").Left
                oItem.Height = oForm.Items.Item("lblDesL").Height
                oItem.Width = oForm.Items.Item("lblDesL").Width
                oItem.LinkTo = "cmbTipFam"
                oItem.FromPane = 1
                oItem.ToPane = 1
                CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Tipo Familia"

                oItem = oForm.Items.Add("lkTipFam", BoFormItemTypes.it_LINKED_BUTTON)
                oItem.Top = oForm.Items.Item("cmbTipFam").Top
                oItem.Left = oForm.Items.Item("136").Left
                oItem.FromPane = 1
                oItem.ToPane = 1

                'CType(oItem.Specific, SAPbouiCOM.LinkedButton).LinkedObject = BoLinkedObject.lf_ServiceCall
                CType(oItem.Specific, SAPbouiCOM.LinkedButton).LinkedObjectType = "EXO_TIPOFAM"
                CType(oItem.Specific, SAPbouiCOM.LinkedButton).Item.LinkTo = "cmbTipFam"

                'cargar datos combo
                CargaComboTipoFamilia(oForm)


            End If

            EventHandler_Form_Load = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))

        End Try
    End Function
    Sub CargaComboTipoFamilia(ByRef oForm As Form)
        Dim sSQL As String = ""
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        Try
            oForm.Freeze(True)
            sSQL = "SELECT T0.""Code"", T0.""Name"" FROM ""@EXO_TIPOFAM""  T0"
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cmbTipFam").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                oCombo = CType(oForm.Items.Item("cmbTipFam").Specific, ComboBox)
                oCombo.ExpandType = BoExpandType.et_DescriptionOnly
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            ' EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.FormCombo(oCombo)

        End Try
    End Sub
    Private Function EventHandler_CLICK_BEFORE(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim STipo As String = ""
        Dim sUser As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oFormA As SAPbouiCOM.Form = Nothing

        EventHandler_CLICK_BEFORE = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "cmbTipFam" Then
                If pVal.ActionSuccess = False Then
                    Try
                        'CARGARCOMBO
                        CargaComboTipoFamilia(oForm)

                    Catch ex As Exception

                    End Try

                End If

            End If

            EventHandler_CLICK_BEFORE = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            '_sCodeALMDIVEMP = ""
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
End Class

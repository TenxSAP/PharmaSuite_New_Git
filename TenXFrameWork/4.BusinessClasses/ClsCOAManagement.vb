Imports SAPbouiCOM

Public Class ClsCOAManagement

#Region "       Declaration             "
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details, oDBs_Details1, oDBs_Details2, oDBs_Details3, oDBs_Details4, oDBs_Details5 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, objMatrix2, objMatrix3, objMatrix4, objMatrix5, objMatrix6 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim str, str1 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim objutilities As Utilities
    Dim MATRIXS As String
    Dim ChkMatrix As String
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("COAManagement.xml", "10X_COA", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("10X_COA", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_COA_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_COA_T")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_COA_A")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_COA_APP")

            objMatrix1 = objForm.Items.Item("MXT_1").Specific
            objMatrix2 = objForm.Items.Item("MXT_2").Specific
            objMatrix3 = objForm.Items.Item("MXT_3").Specific


            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNXCOAUDO", "Primary"))
            oDBs_Head.SetValue("U_DDS", 0, DateTime.Now.ToString("yyyyMMdd"))
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DDS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DDS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Me.SetDefault(objForm.UniqueID)


            'objForm.DataBrowser.BrowseBy = "DocNum"
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("DDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("DDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("Status").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("Status").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            'objForm.Items.Item("BatchSize").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("BatchSize").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Freeze(False)
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)

            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "10X_COA" And pVal.BeforeAction = False Then
                Me.CreateForm()

            ElseIf pVal.MenuUID = "1281" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                If objForm.TypeEx = "10X_COA" Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                End If

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'objForm = objMain.objApplication.Forms.ActiveForm
                Me.SetDefault(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx = "10X_COA" Then

                    objMatrix1 = objForm.Items.Item("MXT_1").Specific

                    objMatrix1.AddRow()
                    Me.SetNewLine(objForm.UniqueID)

                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                    objMatrix1.AutoResizeColumns()

                End If


            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm
                objMatrix1 = objForm.Items.Item("MXT_1").Specific
                'Me.SetNewLine(objForm.UniqueID)
                Me.SetNewLine1(objForm.UniqueID, MATRIXS)
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                End If

            End If

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = False AndAlso pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        Me.SetDefault(objForm.UniqueID)

                    End If


                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    If pVal.ItemUID = "MXT_2" And pVal.ColUID = "TPA" And pVal.BeforeAction = False Then
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("MXT_2").Specific


                        If pVal.Row > 0 AndAlso pVal.Row <= objMatrix.VisualRowCount Then
                            Dim fullPath As String = objMatrix.Columns.Item("TPA").Cells.Item(pVal.Row).Specific.Value
                            If Not String.IsNullOrEmpty(fullPath) AndAlso fullPath.Contains("\") Then
                                Dim indexLoc As Integer = fullPath.LastIndexOf("\")
                                Dim filename As String = fullPath.Substring(indexLoc + 1)
                                objMatrix.Columns.Item("FN").Cells.Item(pVal.Row).Specific.Value = filename
                                objMatrix.Columns.Item("ATD").Cells.Item(pVal.Row).Specific.Value = DateTime.Now.ToString("yyyyMMdd")
                                ' objForm.Items.Item("btn_Del").Enabled = True
                            End If
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then

                        If (pVal.ItemUID = "MXT_1" Or
                            pVal.ItemUID = "MXT_2" Or
                            pVal.ItemUID = "MXT_3" Or
                            pVal.ItemUID = "MXT_4" Or
                            pVal.ItemUID = "MXT_5" Or
                            pVal.ItemUID = "MXT_6") _
                            And pVal.BeforeAction = True Then

                            MATRIXS = pVal.ItemUID

                        End If

                    End If

            End Select

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub SetDefault(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            'objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_COA_H")

            ' oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNXCOAUDO", "Primary"))
            ' oDBs_Head.SetValue("U_DDate", oDBs_Details.Offset, DateTime.Now.ToString("yyyyMMdd"))
            ' objForm.Items.Item("Item_7").Click(BoCellClickType.ct_Regular)
            oDBs_Head.SetValue("U_DS", 0, "Open")

            ' Dim objComboBox1 = objForm.Items.Item("Status").Specific
            'objComboBox1.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNXCOAUDO"))
            oDBs_Head.SetValue("U_DDS", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            'oDS.SetValue("U_DOCDATE", oDS.Offset, DateTime.Now.ToString("yyyyMMdd"))
            Me.SetNewLine(FormUID)
            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

    Sub SetNewLine(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_COA_T")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_COA_A")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_COA_APP")


            objMatrix1 = objForm.Items.Item("MXT_1").Specific
            objMatrix2 = objForm.Items.Item("MXT_2").Specific
            objMatrix3 = objForm.Items.Item("MXT_3").Specific



            '======================== MATRIX 1 ========================

            objMatrix1 = objForm.Items.Item("MXT_1").Specific

            If objMatrix1.VisualRowCount = 0 Then

                objMatrix1.AddRow()

                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix1.VisualRowCount)
                oDBs_Details.SetValue("U_TestCode", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_TestName", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_TestM", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Unit", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_SpecMin", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_SpecMax", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_SpecText", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ResultV", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ResultT", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ResultS", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Analyst", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_TestD", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Instrument", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Remarks", oDBs_Details.Offset, "")

                objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                objMatrix1.AutoResizeColumns()

            End If

            '======================== MATRIX 2 ========================

            objMatrix2 = objForm.Items.Item("MXT_2").Specific

            If objMatrix2.VisualRowCount = 0 Then

                objMatrix2.AddRow()

                oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix2.VisualRowCount)
                oDBs_Details1.SetValue("U_FileN", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_FileT", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_FileP", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_AttachE", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_UPU", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_UPD", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_RM", oDBs_Details1.Offset, "")

                objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                objMatrix2.AutoResizeColumns()

            End If

            '======================== MATRIX 3 ========================

            objMatrix3 = objForm.Items.Item("MXT_3").Specific

            If objMatrix3.VisualRowCount = 0 Then

                objMatrix3.AddRow()

                oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix3.VisualRowCount)
                oDBs_Details2.SetValue("U_ALevel", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_AppR", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_AppU", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_Status", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_ActionA", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_Esign", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_Com", oDBs_Details2.Offset, "")

                objMatrix3.SetLineData(objMatrix3.VisualRowCount)

                objMatrix3.AutoResizeColumns()

            End If



        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

    Private Sub SetNewLine1(ByVal FormUID As String,
                        ByVal MatrixUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            Select Case MatrixUID

            '====================================================
            ' MATRIX 1
            '====================================================

                Case "MXT_1"

                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_COA_T")

                    objMatrix1 = objForm.Items.Item("MXT_1").Specific

                    objMatrix1.AddRow()

                    oDBs_Details.SetValue("LineId",
                                          oDBs_Details.Offset,
                                          objMatrix1.VisualRowCount.ToString())

                    oDBs_Details.SetValue("U_TestCode", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_TestName", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_TestM", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Unit", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_SpecMin", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_SpecMax", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_SpecText", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ResultV", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ResultT", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ResultS", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Analyst", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_TestD", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Instrument", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Remarks", oDBs_Details.Offset, "")

                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                    objMatrix1.AutoResizeColumns()

            '====================================================
            ' MATRIX 2
            '====================================================

                Case "MXT_2"

                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_COA_A")

                    objMatrix2 = objForm.Items.Item("MXT_2").Specific

                    objMatrix2.AddRow()

                    oDBs_Details1.SetValue("LineId",
                                           oDBs_Details1.Offset,
                                           objMatrix2.VisualRowCount.ToString())

                    oDBs_Details1.SetValue("U_FileN", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_FileT", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_FileP", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_AttachE", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_UPU", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_UPD", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_RM", oDBs_Details1.Offset, "")

                    objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                    objMatrix2.AutoResizeColumns()

            '====================================================
            ' MATRIX 3
            '====================================================

                Case "MXT_3"

                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_COA_APP")

                    objMatrix3 = objForm.Items.Item("MXT_3").Specific

                    objMatrix3.AddRow()

                    oDBs_Details2.SetValue("LineId",
                                           oDBs_Details2.Offset,
                                           objMatrix3.VisualRowCount.ToString())

                    oDBs_Details2.SetValue("U_ALevel", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_AppR", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_AppU", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_Status", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_ActionA", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_Esign", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_Com", oDBs_Details2.Offset, "")

                    objMatrix3.SetLineData(objMatrix3.VisualRowCount)

                    objMatrix3.AutoResizeColumns()

            End Select

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "SetNewLine1 Error : " & ex.Message)

        End Try

    End Sub
End Class

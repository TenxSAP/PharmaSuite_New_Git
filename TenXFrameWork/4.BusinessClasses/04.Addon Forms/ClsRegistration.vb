Imports SAPbouiCOM

Public Class ClsRegistration

#Region "       Declaration             "
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details, oDBs_Details1, oDBs_Details2, oDBs_Details3, oDBs_Details4, oDBs_Details5 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, objMatrix2, objMatrix3, objMatrix4, objMatrix5, objMatrix6, objMatrix7 As SAPbouiCOM.Matrix
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
            objMain.objUtilities.LoadForm("RegistratioRenewal.xml", "10X_REG_RENEW", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("10X_REG_RENEW", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_RENH")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_REG_RENL")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_REG_REND")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_REG_RENQ")

            objMatrix1 = objForm.Items.Item("0_U_G").Specific
            objMatrix2 = objForm.Items.Item("1_U_G").Specific
            objMatrix3 = objForm.Items.Item("2_U_G").Specific

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_REG_RENEW", "Primary"))
            oDBs_Head.SetValue("U_DDR", 0, DateTime.Now.ToString("yyyyMMdd"))
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DDR").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DDR").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Me.SetDefault(objForm.UniqueID)


            'objForm.DataBrowser.BrowseBy = "DocNum"
            'objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("DDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("DDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("Status").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("Status").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            'objForm.Items.Item("BatchSize").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("BatchSize").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Freeze(False)
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)
            objForm.EnableMenu("1282", True)

            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "10X_REG_RENEW" And pVal.BeforeAction = False Then
                Me.CreateForm()

            ElseIf pVal.MenuUID = "1281" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                If objForm.TypeEx = "10X_REG_RENEW" Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                End If

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'objForm = objMain.objApplication.Forms.ActiveForm
                Me.SetDefault(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx = "10X_REG_RENEW" Then

                    objMatrix1 = objForm.Items.Item("0_U_G").Specific

                    objMatrix1.AddRow()
                    Me.SetNewLine(objForm.UniqueID)

                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                    objMatrix1.AutoResizeColumns()

                End If


            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm
                objMatrix1 = objForm.Items.Item("0_U_G").Specific
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
                        If (pVal.ItemUID = "0_U_G" Or
    pVal.ItemUID = "1_U_G" Or
    pVal.ItemUID = "2_U_G") _
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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_RENH")

            ' oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNXCOAUDO", "Primary"))
            ' oDBs_Head.SetValue("U_DDate", oDBs_Details.Offset, DateTime.Now.ToString("yyyyMMdd"))
            ' objForm.Items.Item("Item_7").Click(BoCellClickType.ct_Regular)
            ' oDBs_Head.SetValue("U_DS", 0, "Open")

            ' Dim objComboBox1 = objForm.Items.Item("Status").Specific
            'objComboBox1.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_REG_RENEW"))
            oDBs_Head.SetValue("U_DDR", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
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
            '  oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_RENH")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_REG_RENL")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_REG_REND")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_REG_RENQ")

            objMatrix1 = objForm.Items.Item("0_U_G").Specific
            objMatrix2 = objForm.Items.Item("1_U_G").Specific
            objMatrix3 = objForm.Items.Item("2_U_G").Specific



            '======================== MATRIX 1 ========================

            objMatrix1 = objForm.Items.Item("0_U_G").Specific

            If objMatrix1.VisualRowCount = 0 Then

                objMatrix1.AddRow()

                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix1.VisualRowCount)
                oDBs_Details.SetValue("U_ActivityDate", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ActivityType", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ActivityOwner", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ActivityStatus", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_DueDate", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_CompletionDate", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Remarks", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_AttachEntry", oDBs_Details.Offset, "")



                objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                objMatrix1.AutoResizeColumns()

            End If

            '======================== MATRIX 2 ========================

            objMatrix2 = objForm.Items.Item("1_U_G").Specific

            If objMatrix2.VisualRowCount = 0 Then

                objMatrix2.AddRow()

                oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix2.VisualRowCount)
                oDBs_Details1.SetValue("U_DocType", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_DocName", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_Required", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_Available", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_DocVersion", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_ValidFrom", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_ValidTo", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_AttachEntry", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_VerifiedBy", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_VerifiedDate", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_Remarks", oDBs_Details1.Offset, "")

                objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                objMatrix2.AutoResizeColumns()

            End If

            '======================== MATRIX 3 ========================

            objMatrix3 = objForm.Items.Item("2_U_G").Specific

            If objMatrix3.VisualRowCount = 0 Then

                objMatrix3.AddRow()
                oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix3.VisualRowCount)
                oDBs_Details2.SetValue("U_QueryNo", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_QueryDate", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_QueryType", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_QueryDescription", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_ResponseDueDate", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_ResponseDate", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_ResponseBy", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_ResponseAttach", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_QueryStatus", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_EscalationRequired", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_Remarks", oDBs_Details2.Offset, "")

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

                Case "0_U_G"

                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_REG_RENL")

                    objMatrix1 = objForm.Items.Item("0_U_G").Specific

                    objMatrix1.AddRow()

                    oDBs_Details.SetValue("LineId",
                                          oDBs_Details.Offset,
                                          objMatrix1.VisualRowCount.ToString())
                    oDBs_Details.SetValue("U_ActivityDate", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ActivityType", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ActivityOwner", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ActivityStatus", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_DueDate", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_CompletionDate", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Remarks", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_AttachEntry", oDBs_Details.Offset, "")

                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                    objMatrix1.AutoResizeColumns()

            '====================================================
            ' MATRIX 2
            '====================================================

                Case "1_U_G"

                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_REG_REND")

                    objMatrix2 = objForm.Items.Item("1_U_G").Specific

                    objMatrix2.AddRow()

                    oDBs_Details1.SetValue("LineId",
                                           oDBs_Details1.Offset,
                                           objMatrix2.VisualRowCount.ToString())

                    oDBs_Details1.SetValue("U_DocType", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_DocName", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_Required", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_Available", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_DocVersion", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_ValidFrom", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_ValidTo", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_AttachEntry", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VerifiedBy", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VerifiedDate", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_Remarks", oDBs_Details1.Offset, "")
                    objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                    objMatrix2.AutoResizeColumns()

            '====================================================
            ' MATRIX 3
            '====================================================

                Case "2_U_G"

                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_REG_RENQ")

                    objMatrix3 = objForm.Items.Item("2_U_G").Specific

                    objMatrix3.AddRow()

                    oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix3.VisualRowCount)
                    oDBs_Details2.SetValue("U_QueryNo", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_QueryDate", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_QueryType", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_QueryDescription", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_ResponseDueDate", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_ResponseDate", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_ResponseBy", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_ResponseAttach", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_QueryStatus", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_EscalationRequired", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_Remarks", oDBs_Details2.Offset, "")
                    objMatrix3.SetLineData(objMatrix3.VisualRowCount)

                    objMatrix3.AutoResizeColumns()

            End Select

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "SetNewLine1 Error : " & ex.Message)

        End Try

    End Sub
End Class

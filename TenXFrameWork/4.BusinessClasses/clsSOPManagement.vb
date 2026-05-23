Imports SAPbouiCOM

Public Class clsSOPManagement
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
    Dim objComboBox1, objComboBox2
#End Region
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("SOPManagement.xml", "frm_SOPMGT", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("frm_SOPMGT", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities

            Me.SetDefault(objForm.UniqueID)

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

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
            If pVal.MenuUID = "10X_COMP_SOP" And pVal.BeforeAction = False Then
                Me.CreateForm()

            ElseIf pVal.MenuUID = "1281" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                If objForm.TypeEx = "frm_SOPMGT" Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                End If

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm
                ' objMatrix1 = objForm.Items.Item("MXT_1").Specific
                Me.SetNewLine1(objForm.UniqueID, MATRIXS)
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                End If

            ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx = "frm_SOPMGT" Then

                    objMatrix1 = objForm.Items.Item("MXT_1").Specific
                    objMatrix1.AddRow()
                    Me.SetNewLine(objForm.UniqueID)
                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                    objMatrix1.AutoResizeColumns()

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

                'Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                '    If pVal.ItemUID = "MXT_2" And pVal.ColUID = "TPA" And pVal.BeforeAction = False Then
                '        objForm = objMain.objApplication.Forms.Item(FormUID)
                '        Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("MXT_2").Specific


                '        If pVal.Row > 0 AndAlso pVal.Row <= objMatrix.VisualRowCount Then
                '            Dim fullPath As String = objMatrix.Columns.Item("TPA").Cells.Item(pVal.Row).Specific.Value
                '            If Not String.IsNullOrEmpty(fullPath) AndAlso fullPath.Contains("\") Then
                '                Dim indexLoc As Integer = fullPath.LastIndexOf("\")
                '                Dim filename As String = fullPath.Substring(indexLoc + 1)
                '                objMatrix.Columns.Item("FN").Cells.Item(pVal.Row).Specific.Value = filename
                '                objMatrix.Columns.Item("ATD").Cells.Item(pVal.Row).Specific.Value = DateTime.Now.ToString("yyyyMMdd")
                '                ' objForm.Items.Item("btn_Del").Enabled = True
                '            End If
                '        End If
                '    End If

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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_SOPH")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_SOP_REV")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_APR")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_TRN")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_DIST")
            oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_CAT")
            oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_TMP")

            objMatrix1 = objForm.Items.Item("MTX_1").Specific
            objMatrix2 = objForm.Items.Item("MTX_2").Specific
            objMatrix3 = objForm.Items.Item("MTX_3").Specific
            objMatrix4 = objForm.Items.Item("MTX_4").Specific
            objMatrix5 = objForm.Items.Item("MTX_5").Specific
            objMatrix6 = objForm.Items.Item("MTX_6").Specific
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_SOP"))
            oDBs_Head.SetValue("U_EffDate", 0, DateTime.Now.ToString("yyyyMMdd"))

            objForm.Items.Item("Item_75").Click(BoCellClickType.ct_Regular)

            'objComboBox1 = objForm.Items.Item("Status").Specific
            'objComboBox1.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByValue)
            'objComboBox2 = objForm.Items.Item("QAStatus").Specific
            'objComboBox2.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByValue)

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

            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_SOP_REV")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_APR")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_TRN")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_DIST")
            oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_CAT")
            oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_TMP")


            objMatrix1 = objForm.Items.Item("MTX_1").Specific
            objMatrix2 = objForm.Items.Item("MTX_2").Specific
            objMatrix3 = objForm.Items.Item("MTX_3").Specific
            objMatrix4 = objForm.Items.Item("MTX_4").Specific
            objMatrix5 = objForm.Items.Item("MTX_5").Specific
            objMatrix6 = objForm.Items.Item("MTX_6").Specific



            ' =========================================================
            ' MATRIX 1 - SOP REVISION
            ' =========================================================

            If objMatrix1.VisualRowCount = 0 Then

                objMatrix1.AddRow()

                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix1.VisualRowCount)
                oDBs_Details.SetValue("U_SecNo", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_SecTitle", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ChgType", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ChgDesc", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Reason", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ImpArea", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ImpLvl", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ChgBy", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ChgDate", oDBs_Details.Offset, "")

                objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                objMatrix1.AutoResizeColumns()

            End If



            ' =========================================================
            ' MATRIX 2 - SOP APPROVAL
            ' =========================================================

            If objMatrix2.VisualRowCount = 0 Then

                objMatrix2.AddRow()

                oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix2.VisualRowCount)
                oDBs_Details1.SetValue("U_Level", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_ApprRole", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_ApprUser", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_ActDate", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_ActTime", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_Comments", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_ESign", oDBs_Details1.Offset, "")

                objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                objMatrix2.AutoResizeColumns()

            End If



            ' =========================================================
            ' MATRIX 3 - SOP TRAINING
            ' =========================================================

            If objMatrix3.VisualRowCount = 0 Then

                objMatrix3.AddRow()

                oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix3.VisualRowCount)
                oDBs_Details2.SetValue("U_Dept", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_Positn", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_EmpID", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_EmpName", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_TrainReq", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_DueDate", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_TrainDoc", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_Status", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_CompDate", oDBs_Details2.Offset, "")

                objMatrix3.SetLineData(objMatrix3.VisualRowCount)

                objMatrix3.AutoResizeColumns()

            End If



            ' =========================================================
            ' MATRIX 4 - SOP DISTRIBUTION
            ' =========================================================

            If objMatrix4.VisualRowCount = 0 Then

                objMatrix4.AddRow()

                oDBs_Details3.SetValue("LineId", oDBs_Details3.Offset, objMatrix4.VisualRowCount)
                oDBs_Details3.SetValue("U_Dept", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_UserCode", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_DistDate", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_Ack", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_AckDate", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_Remarks", oDBs_Details3.Offset, "")

                objMatrix4.SetLineData(objMatrix4.VisualRowCount)

                objMatrix4.AutoResizeColumns()

            End If



            ' =========================================================
            ' MATRIX 5 - SOP CATEGORY
            ' =========================================================

            If objMatrix5.VisualRowCount = 0 Then

                objMatrix5.AddRow()

                oDBs_Details4.SetValue("LineId", oDBs_Details4.Offset, objMatrix5.VisualRowCount)
                oDBs_Details4.SetValue("U_Code", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_Name", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_Dept", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_RevCycle", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_TrainReq", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_ApprRoute", oDBs_Details4.Offset, "")

                objMatrix5.SetLineData(objMatrix5.VisualRowCount)

                objMatrix5.AutoResizeColumns()

            End If



            ' =========================================================
            ' MATRIX 6 - SOP TEMPLATE
            ' =========================================================

            If objMatrix6.VisualRowCount = 0 Then

                objMatrix6.AddRow()

                oDBs_Details5.SetValue("LineId", oDBs_Details5.Offset, objMatrix6.VisualRowCount)
                oDBs_Details5.SetValue("U_Code", oDBs_Details5.Offset, "")
                oDBs_Details5.SetValue("U_Name", oDBs_Details5.Offset, "")
                oDBs_Details5.SetValue("U_SOPType", oDBs_Details5.Offset, "")
                oDBs_Details5.SetValue("U_DefSect", oDBs_Details5.Offset, "")
                oDBs_Details5.SetValue("U_AttEntry", oDBs_Details5.Offset, "")

                objMatrix6.SetLineData(objMatrix6.VisualRowCount)

                objMatrix6.AutoResizeColumns()

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
        ' MATRIX 1 - SOP REVISION
        '====================================================

                Case "MTX_1"

                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_SOP_REV")

                    objMatrix1 = objForm.Items.Item("MTX_1").Specific

                    objMatrix1.AddRow()

                    oDBs_Details.SetValue("LineId",
                                          oDBs_Details.Offset,
                                          objMatrix1.VisualRowCount.ToString())

                    oDBs_Details.SetValue("U_SecNo", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_SecTitle", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ChgType", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ChgDesc", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Reason", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ImpArea", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ImpLvl", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ChgBy", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ChgDate", oDBs_Details.Offset, "")

                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                    objMatrix1.AutoResizeColumns()



        '====================================================
        ' MATRIX 2 - SOP APPROVAL
        '====================================================

                Case "MTX_2"

                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_APR")

                    objMatrix2 = objForm.Items.Item("MTX_2").Specific

                    objMatrix2.AddRow()

                    oDBs_Details1.SetValue("LineId",
                                           oDBs_Details1.Offset,
                                           objMatrix2.VisualRowCount.ToString())

                    oDBs_Details1.SetValue("U_Level", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_ApprRole", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_ApprUser", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_ActDate", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_ActTime", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_Comments", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_ESign", oDBs_Details1.Offset, "")

                    objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                    objMatrix2.AutoResizeColumns()



        '====================================================
        ' MATRIX 3 - SOP TRAINING
        '====================================================

                Case "MTX_3"

                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_TRN")

                    objMatrix3 = objForm.Items.Item("MTX_3").Specific

                    objMatrix3.AddRow()

                    oDBs_Details2.SetValue("LineId",
                                           oDBs_Details2.Offset,
                                           objMatrix3.VisualRowCount.ToString())

                    oDBs_Details2.SetValue("U_Dept", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_Positn", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_EmpID", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_EmpName", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_TrainReq", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_DueDate", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_TrainDoc", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_Status", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_CompDate", oDBs_Details2.Offset, "")

                    objMatrix3.SetLineData(objMatrix3.VisualRowCount)

                    objMatrix3.AutoResizeColumns()



        '====================================================
        ' MATRIX 4 - SOP DISTRIBUTION
        '====================================================

                Case "MTX_4"

                    oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_DIST")

                    objMatrix4 = objForm.Items.Item("MTX_4").Specific

                    objMatrix4.AddRow()

                    oDBs_Details3.SetValue("LineId",
                                           oDBs_Details3.Offset,
                                           objMatrix4.VisualRowCount.ToString())

                    oDBs_Details3.SetValue("U_Dept", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_UserCode", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_DistDate", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_Ack", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_AckDate", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_Remarks", oDBs_Details3.Offset, "")

                    objMatrix4.SetLineData(objMatrix4.VisualRowCount)

                    objMatrix4.AutoResizeColumns()



        '====================================================
        ' MATRIX 5 - SOP CATEGORY
        '====================================================

                Case "MTX_5"

                    oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_CAT")

                    objMatrix5 = objForm.Items.Item("MTX_5").Specific

                    objMatrix5.AddRow()

                    oDBs_Details4.SetValue("LineId",
                                           oDBs_Details4.Offset,
                                           objMatrix5.VisualRowCount.ToString())

                    oDBs_Details4.SetValue("U_Code", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_Name", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_Dept", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_RevCycle", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_TrainReq", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_ApprRoute", oDBs_Details4.Offset, "")

                    objMatrix5.SetLineData(objMatrix5.VisualRowCount)

                    objMatrix5.AutoResizeColumns()



        '====================================================
        ' MATRIX 6 - SOP TEMPLATE
        '====================================================

                Case "MTX_6"

                    oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_SOP_TMP")

                    objMatrix6 = objForm.Items.Item("MTX_6").Specific

                    objMatrix6.AddRow()

                    oDBs_Details5.SetValue("LineId",
                                           oDBs_Details5.Offset,
                                           objMatrix6.VisualRowCount.ToString())

                    oDBs_Details5.SetValue("U_Code", oDBs_Details5.Offset, "")
                    oDBs_Details5.SetValue("U_Name", oDBs_Details5.Offset, "")
                    oDBs_Details5.SetValue("U_SOPType", oDBs_Details5.Offset, "")
                    oDBs_Details5.SetValue("U_DefSect", oDBs_Details5.Offset, "")
                    oDBs_Details5.SetValue("U_AttEntry", oDBs_Details5.Offset, "")

                    objMatrix6.SetLineData(objMatrix6.VisualRowCount)

                    objMatrix6.AutoResizeColumns()

            End Select

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "SetNewLine1 Error : " & ex.Message)

        End Try

    End Sub
End Class

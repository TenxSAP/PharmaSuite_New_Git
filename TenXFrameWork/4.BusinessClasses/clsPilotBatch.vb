Imports SAPbouiCOM

Public Class clsPilotBatch

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
            objMain.objUtilities.LoadForm("PilotBatch.xml", "PHPILOT", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("PHPILOT", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities

            Me.SetDefault(objForm.UniqueID)


            'objForm.DataBrowser.BrowseBy = "DocNum"
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("CrtDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("CrtDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

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
            If pVal.MenuUID = "10X_PILOT" And pVal.BeforeAction = False Then
                Me.CreateForm()

            ElseIf pVal.MenuUID = "1281" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                If objForm.TypeEx = "PHPILOT" Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                End If

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'objForm = objMain.objApplication.Forms.ActiveForm
                Me.SetDefault(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx = "PHPILOT" Then

                    objMatrix1 = objForm.Items.Item("MXT_1").Specific

                    objMatrix1.AddRow()
                    Me.SetNewLine(objForm.UniqueID)

                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                    objMatrix1.AutoResizeColumns()

                End If


            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm
                objMatrix1 = objForm.Items.Item("MXT_1").Specific
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



                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then

                        If (pVal.ItemUID = "MXT_1" Or
                            pVal.ItemUID = "MXT_3" Or
                            pVal.ItemUID = "Item_13" Or
                            pVal.ItemUID = "Item_23" Or
                            pVal.ItemUID = "MXT_5") _
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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PB_HDR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PB_PROC")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PB_MAT")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PB_ISS")
            oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PB_QC")
            oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_PB_YIELD")

            objMatrix1 = objForm.Items.Item("MXT_1").Specific
            objMatrix2 = objForm.Items.Item("MXT_3").Specific
            objMatrix3 = objForm.Items.Item("Item_13").Specific
            objMatrix4 = objForm.Items.Item("Item_23").Specific
            objMatrix5 = objForm.Items.Item("MXT_5").Specific

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_PILOT"))
            oDBs_Head.SetValue("U_CrtDate", 0, DateTime.Now.ToString("yyyyMMdd"))


            objForm.Items.Item("Item_7").Click(BoCellClickType.ct_Regular)


            'objComboBox1 = objForm.Items.Item("Status").Specific
            'objComboBox1.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByValue)
            'objComboBox2 = objForm.Items.Item("QAStatus").Specific
            'objComboBox2.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByValue)

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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PB_HDR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PB_PROC")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PB_MAT")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PB_ISS")
            oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PB_QC")
            oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_PB_YIELD")

            objMatrix1 = objForm.Items.Item("MXT_1").Specific
            objMatrix2 = objForm.Items.Item("MXT_3").Specific
            objMatrix3 = objForm.Items.Item("Item_13").Specific
            objMatrix4 = objForm.Items.Item("Item_23").Specific
            objMatrix5 = objForm.Items.Item("MXT_5").Specific



            '======================== MATRIX 1 ========================

            objMatrix1 = objForm.Items.Item("MXT_1").Specific

            If objMatrix1.VisualRowCount = 0 Then

                objMatrix1.AddRow()

                oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)

                oDBs_Details1.SetValue("U_StepNo", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_ProcStage", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_Instruc", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_MacCode", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_MacName", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_PlanTime", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_ActStart", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_ActEnd", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_Oper", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_Remarks", oDBs_Details1.Offset, "")

                objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                objMatrix1.AutoResizeColumns()

            End If

            '======================== MATRIX 2 ========================

            objMatrix2 = objForm.Items.Item("MXT_3").Specific

            If objMatrix2.VisualRowCount = 0 Then

                objMatrix2.AddRow()

                oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)
                oDBs_Details2.SetValue("U_ItemCode", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_ItemName", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_IngType", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_ForQty", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_ForUOM", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_ScaleFac", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_ReqQty", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_IssWhs", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_BatMng", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_TolPer", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_Remarks", oDBs_Details2.Offset, "")

                objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                objMatrix2.AutoResizeColumns()

            End If

            '======================== MATRIX 3 ========================

            objMatrix3 = objForm.Items.Item("Item_13").Specific

            If objMatrix3.VisualRowCount = 0 Then

                objMatrix3.AddRow()

                oDBs_Details3.SetValue("LineId", oDBs_Details3.Offset, objMatrix3.VisualRowCount)

                oDBs_Details3.SetValue("U_ItemCode", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_ReqQty", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_IssQty", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_BatchNo", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_ExpDate", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_WhsCode", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_DiffQty", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_TolStat", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_IssBy", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_IssDate", oDBs_Details3.Offset, "")
                oDBs_Details3.SetValue("U_SAPIssNo", oDBs_Details3.Offset, "")

                objMatrix3.SetLineData(objMatrix3.VisualRowCount)

                objMatrix3.AutoResizeColumns()

            End If

            '======================== MATRIX 4 ========================

            objMatrix4 = objForm.Items.Item("Item_23").Specific

            If objMatrix4.VisualRowCount = 0 Then

                objMatrix4.AddRow()

                oDBs_Details4.SetValue("LineId", oDBs_Details4.Offset, objMatrix4.VisualRowCount)

                oDBs_Details4.SetValue("U_TestCode", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_TestName", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_StdVal", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_MinVal", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_MaxVal", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_ActVal", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_Result", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_ChkBy", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_ChkDate", oDBs_Details4.Offset, "")
                oDBs_Details4.SetValue("U_Remarks", oDBs_Details4.Offset, "")

                objMatrix4.SetLineData(objMatrix4.VisualRowCount)

                objMatrix4.AutoResizeColumns()

            End If

            '======================== MATRIX 5 ========================

            objMatrix5 = objForm.Items.Item("MXT_5").Specific

            If objMatrix5.VisualRowCount = 0 Then

                objMatrix5.AddRow()

                oDBs_Details5.SetValue("LineId", oDBs_Details5.Offset, objMatrix5.VisualRowCount)

                oDBs_Details5.SetValue("U_PlanOut", oDBs_Details5.Offset, "")
                oDBs_Details5.SetValue("U_ActOut", oDBs_Details5.Offset, "")
                oDBs_Details5.SetValue("U_RejQty", oDBs_Details5.Offset, "")
                oDBs_Details5.SetValue("U_SampQty", oDBs_Details5.Offset, "")
                oDBs_Details5.SetValue("U_LossQty", oDBs_Details5.Offset, "")
                oDBs_Details5.SetValue("U_YieldPer", oDBs_Details5.Offset, "")
                oDBs_Details5.SetValue("U_LossPer", oDBs_Details5.Offset, "")
                oDBs_Details5.SetValue("U_YieldStat", oDBs_Details5.Offset, "")
                oDBs_Details5.SetValue("U_Remarks", oDBs_Details5.Offset, "")

                objMatrix5.SetLineData(objMatrix5.VisualRowCount)

                objMatrix5.AutoResizeColumns()

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
' MATRIX 1  |  UID : MXT_1  |  TABLE : @TNX_PB_PROC
'====================================================

                Case "MXT_1"

                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PB_PROC")

                    objMatrix1 = objForm.Items.Item("MXT_1").Specific

                    objMatrix1.AddRow()

                    oDBs_Details.SetValue("LineId",
                          oDBs_Details.Offset,
                          objMatrix1.VisualRowCount.ToString())

                    oDBs_Details.SetValue("U_StepNo", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ProcStage", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Instruc", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_MacCode", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_MacName", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_PlanTime", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ActStart", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_ActEnd", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Oper", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Status", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Remarks", oDBs_Details.Offset, "")

                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                    objMatrix1.AutoResizeColumns()

'====================================================
' MATRIX 2  |  UID : MXT_3  |  TABLE : @TNX_PB_MAT
'====================================================

                Case "MXT_3"

                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PB_MAT")

                    objMatrix2 = objForm.Items.Item("MXT_3").Specific

                    objMatrix2.AddRow()

                    oDBs_Details1.SetValue("LineId",
                           oDBs_Details1.Offset,
                           objMatrix2.VisualRowCount.ToString())

                    oDBs_Details1.SetValue("U_ItemCode", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_ItemName", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_IngType", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_ForQty", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_ForUOM", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_ScaleFac", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_ReqQty", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_IssWhs", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_BatMng", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_TolPer", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_Remarks", oDBs_Details1.Offset, "")

                    objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                    objMatrix2.AutoResizeColumns()

'====================================================
' MATRIX 3  |  UID : Item_13  |  TABLE : @TNX_PB_ISS
'====================================================

                Case "Item_13"

                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PB_ISS")

                    objMatrix3 = objForm.Items.Item("Item_13").Specific

                    objMatrix3.AddRow()

                    oDBs_Details2.SetValue("LineId",
                           oDBs_Details2.Offset,
                           objMatrix3.VisualRowCount.ToString())

                    oDBs_Details2.SetValue("U_ItemCode", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_ReqQty", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_IssQty", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_BatchNo", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_ExpDate", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_WhsCode", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_DiffQty", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_TolStat", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_IssBy", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_IssDate", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_SAPIssNo", oDBs_Details2.Offset, "")

                    objMatrix3.SetLineData(objMatrix3.VisualRowCount)

                    objMatrix3.AutoResizeColumns()

'====================================================
' MATRIX 4  |  UID : Item_23  |  TABLE : @TNX_PB_QC
'====================================================

                Case "Item_23"

                    oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PB_QC")

                    objMatrix4 = objForm.Items.Item("Item_23").Specific

                    objMatrix4.AddRow()

                    oDBs_Details3.SetValue("LineId",
                           oDBs_Details3.Offset,
                           objMatrix4.VisualRowCount.ToString())

                    oDBs_Details3.SetValue("U_TestCode", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_TestName", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_StdVal", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_MinVal", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_MaxVal", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_ActVal", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_Result", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_ChkBy", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_ChkDate", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_Remarks", oDBs_Details3.Offset, "")

                    objMatrix4.SetLineData(objMatrix4.VisualRowCount)

                    objMatrix4.AutoResizeColumns()

'====================================================
' MATRIX 5  |  UID : MXT_5  |  TABLE : @TNX_PB_YIELD
'====================================================

                Case "MXT_5"

                    oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PB_YIELD")

                    objMatrix5 = objForm.Items.Item("MXT_5").Specific

                    objMatrix5.AddRow()

                    oDBs_Details4.SetValue("LineId",
                           oDBs_Details4.Offset,
                           objMatrix5.VisualRowCount.ToString())

                    oDBs_Details4.SetValue("U_PlanOut", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_ActOut", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_RejQty", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_SampQty", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_LossQty", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_YieldPer", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_LossPer", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_YieldStat", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_Remarks", oDBs_Details4.Offset, "")

                    objMatrix5.SetLineData(objMatrix5.VisualRowCount)

                    objMatrix5.AutoResizeColumns()


            End Select

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "SetNewLine1 Error : " & ex.Message)

        End Try

    End Sub
End Class

Imports System
Imports System.Configuration
    Imports System.Net

    Public Class Cfrm_BMR

#Region "       Declaration             "
        Public objForm As SAPbouiCOM.Form
        Dim oDBs_Head, oDBs_Details1, oDBs_Details2, oDBs_Details3, oDBs_Details4, oDBs_Details5, oDBs_Details6 As SAPbouiCOM.DBDataSource
        Dim objMatrix1, objMatrix2, objMatrix3, objMatrix4, objMatrix5, objMatrix6 As SAPbouiCOM.Matrix
        Dim objComboBox As SAPbouiCOM.ComboBox
        Dim str, str1 As String
        Public rs, RsNum As SAPbobsCOM.Recordset
        Dim LostFocusFlag As Boolean = False
        Dim oGrid As SAPbouiCOM.Grid
        Dim oDt As SAPbouiCOM.DataTable
        Dim objutilities As Utilities

#End Region

        Sub CreateForm()
            Try
            objMain.objUtilities.LoadForm("BRMExecution.xml", "BMRR", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("BMRR", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objForm.Freeze(True)
                objutilities = New Utilities
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_H")
                oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_STAGE")
                oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_MAT")
                oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_EQP")
                oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_IPQC")
                oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_DEV")
                oDBs_Details6 = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_APP")

                objMatrix1 = objForm.Items.Item("Mtx1").Specific
                objMatrix2 = objForm.Items.Item("Mtx2").Specific
                objMatrix3 = objForm.Items.Item("Mtx3").Specific
            objMatrix4 = objForm.Items.Item("3_U_G").Specific
            objMatrix5 = objForm.Items.Item("Mtx5").Specific
                objMatrix6 = objForm.Items.Item("Mtx6").Specific

                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_PBMR"))
                oDBs_Head.SetValue("U_MfgDate", 0, DateTime.Now.ToString("yyyyMMdd"))
                objForm.EnableMenu("1292", True)
                objForm.EnableMenu("1293", True)
                objForm.Freeze(False)
                Me.SetDefault(objForm.UniqueID)
                objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Catch ex As Exception
                objForm.Freeze(False)
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
            If pVal.MenuUID = "10X_BMR" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                objMatrix1 = objForm.Items.Item("Mtx1").Specific
                objMatrix2 = objForm.Items.Item("Mtx2").Specific
                objMatrix3 = objForm.Items.Item("Mtx3").Specific
                objMatrix4 = objForm.Items.Item("3_U_G").Specific
                objMatrix5 = objForm.Items.Item("Mtx5").Specific
                objMatrix6 = objForm.Items.Item("Mtx6").Specific

                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                objMatrix1 = objForm.Items.Item("Mtx1").Specific
                objMatrix2 = objForm.Items.Item("Mtx2").Specific
                objMatrix3 = objForm.Items.Item("Mtx3").Specific
                objMatrix4 = objForm.Items.Item("3_U_G").Specific
                objMatrix5 = objForm.Items.Item("Mtx5").Specific
                objMatrix6 = objForm.Items.Item("Mtx6").Specific
            ElseIf pVal.MenuUID = "1281" And pVal.BeforeAction = False Then

            End If
            Catch ex As Exception
                objForm.Freeze(False)
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub
        Public Function Validate() As Boolean
            Dim CustomerCode As SAPbouiCOM.Matrix
            Try
                If oDBs_Head.GetValue("U_FormulaCode", 0) = "" Then
                    objMain.objApplication.SetStatusBarMessage("Formula Code is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                    Return False
                    Exit Function
                End If
                Return True
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message & "Errors in Validation Function", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Function

        Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try

        Catch ex As Exception
                objForm.Freeze(False)
                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

    Private Sub SetNewLine1(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oDBDS As SAPbouiCOM.DBDataSource
            Dim oMatrix As SAPbouiCOM.Matrix

            oDBDS = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_STAGE")
            oMatrix = objForm.Items.Item("Mtx1").Specific

            'oMatrix.FlushToDataSource()

            'oDBDS.InsertRecord(oDBDS.Size)
            'oDBDS.Offset = oDBDS.Size - 1


            'oMatrix = objForm.Items.Item("Mtx").Specific

            oMatrix.AddRow()

            oDBDS.SetValue("LineId",
                                   oDBDS.Offset,
                                   oMatrix.VisualRowCount.ToString())

            ' oDBDS.SetValue("LineId", oDBDS.Offset, (oDBDS.Offset + 1).ToString())
            oDBDS.SetValue("U_StageCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_StageName", oDBDS.Offset, "")
            oDBDS.SetValue("U_SeqNo", oDBDS.Offset, "")
            oDBDS.SetValue("U_StartTime", oDBDS.Offset, "")
            oDBDS.SetValue("U_EndTime", oDBDS.Offset, "")
            oDBDS.SetValue("U_DurationMin", oDBDS.Offset, "")
            oDBDS.SetValue("U_Operator", oDBDS.Offset, "")
            oDBDS.SetValue("U_Supervisor", oDBDS.Offset, "")
            oDBDS.SetValue("U_EquipCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_Temp", oDBDS.Offset, "")
            oDBDS.SetValue("U_Humidity", oDBDS.Offset, "")
            oDBDS.SetValue("U_RPM", oDBDS.Offset, "")
            oDBDS.SetValue("U_Status", oDBDS.Offset, "")
            oDBDS.SetValue("U_Remarks", oDBDS.Offset, "")

            oMatrix.SetLineData(oMatrix.VisualRowCount)

            oMatrix.LoadFromDataSource()
            oMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Private Sub SetNewLine2(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oDBDS As SAPbouiCOM.DBDataSource
            Dim oMatrix As SAPbouiCOM.Matrix

            oDBDS = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_MAT")
            oMatrix = objForm.Items.Item("Mtx2").Specific

            oMatrix.AddRow()

            oDBDS.SetValue("LineId",
                                   oDBDS.Offset,
                                   oMatrix.VisualRowCount.ToString())

            oDBDS.SetValue("U_BaseLine", oDBDS.Offset, "")
            oDBDS.SetValue("U_ItemCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_ItemName", oDBDS.Offset, "")
            oDBDS.SetValue("U_PlannedQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_IssuedQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_ConsumedQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_ReturnQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_WasteQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_BatchNo", oDBDS.Offset, "")
            oDBDS.SetValue("U_WhsCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_UOM", oDBDS.Offset, "")
            oDBDS.SetValue("U_VarianceQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_VariancePct", oDBDS.Offset, "")
            oDBDS.SetValue("U_Status", oDBDS.Offset, "")
            oDBDS.SetValue("U_Remarks", oDBDS.Offset, "")
            oMatrix.SetLineData(oMatrix.VisualRowCount)

            oMatrix.LoadFromDataSource()
            oMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Private Sub SetNewLine3(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oDBDS As SAPbouiCOM.DBDataSource
            Dim oMatrix As SAPbouiCOM.Matrix

            oDBDS = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_EQP")
            oMatrix = objForm.Items.Item("Mtx3").Specific


            oMatrix.AddRow()

            oDBDS.SetValue("LineId",
                                   oDBDS.Offset,
                                   oMatrix.VisualRowCount.ToString())

            oDBDS.SetValue("U_EquipCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_EquipName", oDBDS.Offset, "")
            oDBDS.SetValue("U_StageCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_CleaningStatus", oDBDS.Offset, "")
            oDBDS.SetValue("U_CalibrationStatus", oDBDS.Offset, "")
            oDBDS.SetValue("U_UsedFrom", oDBDS.Offset, "")
            oDBDS.SetValue("U_UsedTo", oDBDS.Offset, "")
            oDBDS.SetValue("U_Operator", oDBDS.Offset, "")
            oDBDS.SetValue("U_Remarks", oDBDS.Offset, "")
            oMatrix.SetLineData(oMatrix.VisualRowCount)

            oMatrix.LoadFromDataSource()
            oMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub SetNewLine4(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oDBDS As SAPbouiCOM.DBDataSource
            Dim oMatrix As SAPbouiCOM.Matrix

            oDBDS = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_IPQC")
            oMatrix = objForm.Items.Item("3_U_G").Specific

            oMatrix.AddRow()

            oDBDS.SetValue("LineId",
                                   oDBDS.Offset,
                                   oMatrix.VisualRowCount.ToString())


            oDBDS.SetValue("U_StageCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_TestCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_TestName", oDBDS.Offset, "")
            oDBDS.SetValue("U_Specification", oDBDS.Offset, "")
            oDBDS.SetValue("U_ResultValue", oDBDS.Offset, "")
            oDBDS.SetValue("U_ResultStatus", oDBDS.Offset, "")
            oDBDS.SetValue("U_CheckedBy", oDBDS.Offset, "")
            oDBDS.SetValue("U_CheckedDate", oDBDS.Offset, "")
            oDBDS.SetValue("U_Remarks", oDBDS.Offset, "")
            oMatrix.SetLineData(oMatrix.VisualRowCount)

            oMatrix.LoadFromDataSource()
            oMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub SetNewLine5(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oDBDS As SAPbouiCOM.DBDataSource
            Dim oMatrix As SAPbouiCOM.Matrix

            oDBDS = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_DEV")
            oMatrix = objForm.Items.Item("Mtx5").Specific
            oMatrix.AddRow()

            oDBDS.SetValue("LineId",
                                   oDBDS.Offset,
                                   oMatrix.VisualRowCount.ToString())


            oDBDS.SetValue("U_DeviationNo", oDBDS.Offset, "")
            oDBDS.SetValue("U_StageCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_DeviationType", oDBDS.Offset, "")
            oDBDS.SetValue("U_Description", oDBDS.Offset, "")
            oDBDS.SetValue("U_Severity", oDBDS.Offset, "")
            oDBDS.SetValue("U_ActionTaken", oDBDS.Offset, "")
            oDBDS.SetValue("U_CAPARequired", oDBDS.Offset, "")
            oDBDS.SetValue("U_Status", oDBDS.Offset, "")
            oMatrix.SetLineData(oMatrix.VisualRowCount)

            oMatrix.LoadFromDataSource()
            oMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub SetNewLine6(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oDBDS As SAPbouiCOM.DBDataSource
            Dim oMatrix As SAPbouiCOM.Matrix

            oDBDS = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_APP")
            oMatrix = objForm.Items.Item("Mtx6").Specific
            oMatrix.AddRow()

            oDBDS.SetValue("LineId",
                                   oDBDS.Offset,
                                   oMatrix.VisualRowCount.ToString())


            oDBDS.SetValue("U_ApprovalLevel", oDBDS.Offset, "")
            oDBDS.SetValue("U_ApproverRole", oDBDS.Offset, "")
            oDBDS.SetValue("U_ApproverUser", oDBDS.Offset, "")
            oDBDS.SetValue("U_Status", oDBDS.Offset, "")
            oDBDS.SetValue("U_ApprovedDate", oDBDS.Offset, "")
            oDBDS.SetValue("U_Remarks", oDBDS.Offset, "")
            oMatrix.SetLineData(oMatrix.VisualRowCount)

            oMatrix.LoadFromDataSource()
            oMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
            Try
                objForm = objMain.objApplication.Forms.Item(FormUID)
                objForm.Freeze(True)
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_H")
                oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_STAGE")
                oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_MAT")
                oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_EQP")
                oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_IPQC")
                oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_DEV")
                oDBs_Details6 = objForm.DataSources.DBDataSources.Item("@TNX_PBMR_APP")

                objMatrix1 = objForm.Items.Item("Mtx1").Specific
                objMatrix2 = objForm.Items.Item("Mtx2").Specific
                objMatrix3 = objForm.Items.Item("Mtx3").Specific
            objMatrix4 = objForm.Items.Item("3_U_G").Specific
            objMatrix5 = objForm.Items.Item("Mtx5").Specific
                objMatrix6 = objForm.Items.Item("Mtx6").Specific

                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_PBMR", "Primary"))
                oDBs_Head.SetValue("U_MfgDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
                objMatrix1.Clear()
                objMatrix2.Clear()
                objMatrix3.Clear()
                objMatrix4.Clear()
                objMatrix5.Clear()
                objMatrix6.Clear()

                oDBs_Details1.Clear()
                oDBs_Details2.Clear()
                oDBs_Details3.Clear()
                oDBs_Details4.Clear()
                oDBs_Details5.Clear()
                oDBs_Details6.Clear()
                objMatrix1.FlushToDataSource()
                objMatrix2.FlushToDataSource()
                objMatrix3.FlushToDataSource()
                objMatrix4.FlushToDataSource()
                objMatrix5.FlushToDataSource()
                objMatrix6.FlushToDataSource()

                Me.SetNewLine1(objForm.UniqueID)
                Me.SetNewLine2(objForm.UniqueID)
                Me.SetNewLine3(objForm.UniqueID)
                Me.SetNewLine4(objForm.UniqueID)
                Me.SetNewLine5(objForm.UniqueID)
            Me.SetNewLine6(objForm.UniqueID)
            objForm.Freeze(False)
            Catch ex As Exception
                objForm.Freeze(False)
                objMain.objApplication.StatusBar.SetText(ex.Message)
            End Try
        End Sub

End Class
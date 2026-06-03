Public Class VendorPerformanceReview
    Public objForm As SAPbouiCOM.Form

    Public oMatrixKPI As SAPbouiCOM.Matrix
    Public oMatrixAction As SAPbouiCOM.Matrix

    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_KPI As SAPbouiCOM.DBDataSource
    Dim oDBs_Action As SAPbouiCOM.DBDataSource
#Region "Create Form"

    Public Sub CreateForm()

        Try

            objMain.objUtilities.LoadForm("VendorPerformanceReview.xml", "UDO_TNX_VPR", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("UDO_TNX_VPR", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VPR_H")
            oDBs_KPI = objForm.DataSources.DBDataSources.Item("@TNX_VPR_KPI")
            oDBs_Action = objForm.DataSources.DBDataSources.Item("@TNX_VPR_ACT")

            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_VPR", "Primary"))
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_VPR", "Primary"))
            oDBs_Head.SetValue("U_Status", oDBs_Head.Offset, "Draft")
            oDBs_Head.SetValue("U_FinalScore", oDBs_Head.Offset, "0")
            oDBs_Head.SetValue("U_Rating", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_TotalPO", oDBs_Head.Offset, "0")
            oDBs_Head.SetValue("U_TotalGRPO", oDBs_Head.Offset, "0")
            oDBs_Head.SetValue("U_RCount", oDBs_Head.Offset, "0")
            oDBs_Head.SetValue("U_DelayCount", oDBs_Head.Offset, "0")

            '====================================================
            ' MATRIX SETTINGS
            '====================================================

            oMatrixKPI = CType(objForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)
            oMatrixAction = CType(objForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix)

            oMatrixKPI.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oMatrixAction.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            '====================================================
            ' DEFAULT TAB
            '====================================================

            objForm.PaneLevel = 1

            objForm.Items.Item("0_U_FD").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            '====================================================
            ' ENABLE MENUS
            '====================================================

            objForm.EnableMenu("1282", True) 'Add
            objForm.EnableMenu("1281", True) 'find
            objForm.EnableMenu("1288", True) 'First
            objForm.EnableMenu("1289", True) 'Prev
            objForm.EnableMenu("1290", True) 'Next
            objForm.EnableMenu("1291", True) 'Last
            objForm.EnableMenu("1292", True) 'Add Row
            objForm.EnableMenu("1293", True) 'Delete Row

            '====================================================
            ' DEFAULT MATRIX ROWS
            '====================================================

            SetNewLine_KPI(objForm.UniqueID)
            SetNewLine_Action(objForm.UniqueID)

            '====================================================
            ' FIND MODE SETTINGS
            '====================================================
            objForm.Items.Item("DocEntry").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("2").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText("Vendor Performance Review Form Loaded Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            Try
                objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

#End Region
    Public Sub SetNewLine_KPI(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VPR_H")
            oDBs_KPI = objForm.DataSources.DBDataSources.Item("@TNX_VPR_KPI")

            oMatrixKPI = CType(objForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)
            oMatrixKPI.AddRow()

            oDBs_KPI.SetValue("LineId", oDBs_KPI.Offset, oMatrixKPI.VisualRowCount)
            oDBs_KPI.SetValue("U_KPI", oDBs_KPI.Offset, "")
            oDBs_KPI.SetValue("U_Target", oDBs_KPI.Offset, "")
            oDBs_KPI.SetValue("U_Actual", oDBs_KPI.Offset, "")
            oDBs_KPI.SetValue("U_Score", oDBs_KPI.Offset, "0")
            oDBs_KPI.SetValue("U_Remarks", oDBs_KPI.Offset, "")

            oMatrixKPI.SetLineData(oMatrixKPI.VisualRowCount)
            oMatrixKPI.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Public Sub SetNewLine_Action(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VPR_H")
            oDBs_Action = objForm.DataSources.DBDataSources.Item("@TNX_VPR_ACT")

            oMatrixAction = CType(objForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix)

            oMatrixAction.AddRow()

            oDBs_Action.SetValue("LineId", oDBs_Action.Offset, oMatrixAction.VisualRowCount)
            oDBs_Action.SetValue("U_Action", oDBs_Action.Offset, "")
            oDBs_Action.SetValue("U_Owner", oDBs_Action.Offset, "")
            oDBs_Action.SetValue("U_DueDate", oDBs_Action.Offset, "")
            oDBs_Action.SetValue("U_Status", oDBs_Action.Offset, "Open")

            oMatrixAction.SetLineData(oMatrixAction.VisualRowCount)
            oMatrixAction.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
    Public Sub SetDefault(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VPR_H")
            oDBs_KPI = objForm.DataSources.DBDataSources.Item("@TNX_VPR_KPI")
            oDBs_Action = objForm.DataSources.DBDataSources.Item("@TNX_VPR_ACT")

            objForm.Freeze(True)

            'Header Defaults
            oDBs_Head.SetValue("DocEntry", 0, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_VPR", "Primary"))
            oDBs_Head.SetValue("DocNum", 0, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_VPR", "Primary"))

            oDBs_Head.SetValue("U_Status", 0, "Draft")
            oDBs_Head.SetValue("U_FinalScore", 0, "0")
            oDBs_Head.SetValue("U_Rating", 0, "")
            oDBs_Head.SetValue("U_TotalPO", 0, "0")
            oDBs_Head.SetValue("U_TotalGRPO", 0, "0")
            oDBs_Head.SetValue("U_RCount", 0, "0")
            oDBs_Head.SetValue("U_DelayCount", 0, "0")

            'Default Tab
            objForm.PaneLevel = 1
            objForm.Items.Item("0_U_FD").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            'Matrix References
            oMatrixKPI = CType(objForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)
            oMatrixAction = CType(objForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix)

            'Clear Existing Rows
            While oDBs_KPI.Size > 0
                oDBs_KPI.RemoveRecord(0)
            End While

            While oDBs_Action.Size > 0
                oDBs_Action.RemoveRecord(0)
            End While

            'Add First Rows
            SetNewLine_KPI(objForm.UniqueID)
            SetNewLine_Action(objForm.UniqueID)

            objForm.Freeze(False)

        Catch ex As Exception

            Try
                objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(
            "SetDefault Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
#Region "Menu Event"

    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

        Try

            If pVal.MenuUID = "10X_VEN_PERF" AndAlso pVal.BeforeAction = False Then

                Me.CreateForm()

            ElseIf pVal.MenuUID = "1282" AndAlso pVal.BeforeAction = False Then
                Me.CreateForm()

            ElseIf pVal.MenuUID = "1292" AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "UDO_TNX_VPR" Then Exit Sub

                If objForm.PaneLevel = 1 Then

                    SetNewLine_KPI(objForm.UniqueID)

                ElseIf objForm.PaneLevel = 2 Then

                    SetNewLine_Action(objForm.UniqueID)

                End If

            ElseIf pVal.MenuUID = "1293" AndAlso pVal.BeforeAction = True Then
                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "UDO_TNX_VPR" Then Exit Sub

                    BubbleEvent = False

                    objForm.Freeze(True)

                    '=========================================================
                    ' DELETE ROW - KPI MATRIX
                    '=========================================================
                    If objForm.PaneLevel = 1 Then

                        oMatrixKPI = CType(objForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)
                        oDBs_KPI = objForm.DataSources.DBDataSources.Item("@TNX_VPR_KPI")

                        Dim selectedRow As Integer = oMatrixKPI.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder)

                        If selectedRow <= 0 Then

                            objMain.objApplication.StatusBar.SetText("Please select KPI row to delete", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            Exit Try

                        End If

                        oMatrixKPI.DeleteRow(selectedRow)
                        oMatrixKPI.FlushToDataSource()

                        While oDBs_KPI.Size > oMatrixKPI.VisualRowCount
                            oDBs_KPI.RemoveRecord(oDBs_KPI.Size - 1)
                        End While

                        If oDBs_KPI.Size = 0 Then

                            oDBs_KPI.InsertRecord(0)

                            oDBs_KPI.SetValue("LineId", 0, "1")
                            oDBs_KPI.SetValue("U_KPI", 0, "")
                            oDBs_KPI.SetValue("U_Target", 0, "")
                            oDBs_KPI.SetValue("U_Actual", 0, "")
                            oDBs_KPI.SetValue("U_Score", 0, "0")
                            oDBs_KPI.SetValue("U_Remarks", 0, "")

                        End If

                        For i As Integer = 0 To oDBs_KPI.Size - 1

                            oDBs_KPI.SetValue("LineId", i, (i + 1).ToString())

                        Next

                        oMatrixKPI.LoadFromDataSource()
                        oMatrixKPI.AutoResizeColumns()

                    End If

                    '=========================================================
                    ' DELETE ROW - ACTION MATRIX
                    '=========================================================
                    If objForm.PaneLevel = 2 Then

                        oMatrixAction = CType(objForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix)
                        oDBs_Action = objForm.DataSources.DBDataSources.Item("@TNX_VPR_ACT")

                        Dim selectedRow As Integer = oMatrixAction.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder)

                        If selectedRow <= 0 Then

                            objMain.objApplication.StatusBar.SetText("Please select Action row to delete", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            Exit Try

                        End If

                        oMatrixAction.DeleteRow(selectedRow)

                        oMatrixAction.FlushToDataSource()

                        While oDBs_Action.Size > oMatrixAction.VisualRowCount
                            oDBs_Action.RemoveRecord(oDBs_Action.Size - 1)
                        End While

                        If oDBs_Action.Size = 0 Then

                            oDBs_Action.InsertRecord(0)

                            oDBs_Action.SetValue("LineId", 0, "1")
                            oDBs_Action.SetValue("U_Action", 0, "")
                            oDBs_Action.SetValue("U_Owner", 0, "")
                            oDBs_Action.SetValue("U_DueDate", 0, "")
                            oDBs_Action.SetValue("U_Status", 0, "Open")

                        End If

                        For i As Integer = 0 To oDBs_Action.Size - 1

                            oDBs_Action.SetValue("LineId", i, (i + 1).ToString())

                        Next

                        oMatrixAction.LoadFromDataSource()
                        oMatrixAction.AutoResizeColumns()

                    End If

                    '=========================================================
                    ' UPDATE MODE
                    '=========================================================
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                    End If

                Catch ex As Exception

                    objMain.objApplication.StatusBar.SetText("Delete Row Error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Finally

                    Try
                        objForm.Freeze(False)
                    Catch
                    End Try

                End Try

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

#End Region
#Region "Item Event"

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

        Try
            Try
                objForm = objMain.objApplication.Forms.Item(FormUID)
            Catch
                Exit Sub
            End Try

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    Try
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                    Catch
                        Exit Sub
                    End Try
                    If pVal.ItemUID = "1" Then
                        Me.CreateForm()
                    End If


                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.BeforeAction = False Then

                    End If



            End Select

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "ItemEvent Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

#End Region
End Class


Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports SAPbouiCOM

Public Class LabTesting




#Region "Declaration"

    Public objForm As SAPbouiCOM.Form
    Public oMatrixResult, oMatrixAttach, oMatrixApprove, objMatrix3 As SAPbouiCOM.Matrix
    Dim oDBs_Head, oDBs_Approve, oDBs_Result As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim oDBs_Attach As SAPbouiCOM.DBDataSource
    Dim oDS As SAPbouiCOM.DBDataSource


#End Region

#Region "Create Form"

    Public Sub CreateForm()

        Try

            objMain.objUtilities.LoadForm("Labtesting.xml", "TNXPH_QCLAB", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("TNXPH_QCLAB",
                  objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            '================================================================
            ' DATASOURCES
            '================================================================
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXPH_QCLABH")
            oDBs_Result = objForm.DataSources.DBDataSources.Item("@TNXPH_QCLABL")
            oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNXPH_QCLABATT")
            oDBs_Approve = objForm.DataSources.DBDataSources.Item("@TNXPH_QCLABAPP")

            '================================================================
            ' DEFAULT VALUES
            '================================================================
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset,
                           objMain.objUtilities.GetNextDocNum(objForm, "TNXPH_QCLAB", "Primary"))

            oDBs_Head.SetValue("U_TestDate", 0, Date.Now.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_Status", 0, "Draft")

            '================================================================
            ' DOCNUM CONTROL
            '================================================================
            'objForm.Items.Item("DocNum").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            '-1,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'objForm.Items.Item("DocNum").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Find,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            '================================================================
            ' HEADER FIELD CONTROL
            '================================================================

            'Test No
            'objForm.Items.Item("TestNo").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Add,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            ''Sample No
            'objForm.Items.Item("SampleNo").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Add,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            ''Item Code
            'objForm.Items.Item("ItemCode").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Add,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            ''Item Name - Auto Fill / Non Editable
            'objForm.Items.Item("ItemName").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            '-1,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            ''Batch No
            'objForm.Items.Item("BatchNo").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Add,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            ''Spec Code
            'objForm.Items.Item("SpecCode").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Add,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            ''Spec Version
            'objForm.Items.Item("SpecVer").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            '-1,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            ''Sample Type
            'objForm.Items.Item("SType").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Add,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            ''Source Type
            'objForm.Items.Item("SrcType").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Add,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            ''Test Date
            'objForm.Items.Item("TestDate").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Add,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            ''Analyst
            'objForm.Items.Item("Analyst").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Add,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            ''Reviewer
            'objForm.Items.Item("Reviewer").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Add,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            ''Overall Result
            'objForm.Items.Item("Overall").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Add,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            ''Status - System Controlled
            'objForm.Items.Item("Status").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            '-1,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            ''Remarks
            'objForm.Items.Item("Item_1").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Add,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            ''================================================================
            '' FIND MODE CONTROLS
            ''================================================================

            'objForm.Items.Item("TestDate").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Find,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'objForm.Items.Item("Overall").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Find,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'objForm.Items.Item("Status").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            'SAPbouiCOM.BoAutoFormMode.afm_Find,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            '================================================================
            ' MATRIX SETTINGS
            '================================================================

            oMatrixResult = objForm.Items.Item("MtxResult").Specific
            oMatrixAttach = objForm.Items.Item("MtxAttach").Specific
            oMatrixApprove = objForm.Items.Item("MtxAppr").Specific

            oMatrixResult.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oMatrixAttach.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oMatrixApprove.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            '================================================================
            ' BUTTON SETTINGS
            '================================================================

            objForm.Items.Item("BtnSub").Enabled = True
            objForm.Items.Item("BtnDev").Enabled = True
            objForm.Items.Item("Btn_dlt").Enabled = True

            '================================================================
            ' TAB DEFAULT
            '================================================================
            objForm.PaneLevel = 1
            objForm.Items.Item("Tab1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            '================================================================
            ' MENU ENABLE
            '================================================================
            Me.objForm.EnableMenu("1282", True)
            Me.objForm.EnableMenu("1288", True)
            Me.objForm.EnableMenu("1289", True)
            Me.objForm.EnableMenu("1290", True)
            Me.objForm.EnableMenu("1291", True)
            Me.objForm.EnableMenu("1292", True)
            Me.objForm.EnableMenu("1293", True)

            '================================================================
            ' DEFAULT LINES
            '================================================================
            SetNewLine_Result(objForm.UniqueID)
            SetNewLine_Attachment(objForm.UniqueID)
            SetNewLine_Approval(objForm.UniqueID)
            objForm.Items.Item("1").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("2").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)


            objForm.Items.Item("BtnSub").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("BtnDev").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
            "QC Lab Testing Form Loaded Successfully",
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            Try
                objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

#End Region


#Region "Add New Line"

    Public Sub SetNewLine_Result(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXPH_QCLABH")
            oDBs_Result = objForm.DataSources.DBDataSources.Item("@TNXPH_QCLABL")

            oMatrixResult = objForm.Items.Item("MtxResult").Specific

            oMatrixResult.AddRow()

            oDBs_Result.SetValue("LineId", oDBs_Result.Offset, oMatrixResult.VisualRowCount)

            oDBs_Result.SetValue("U_TestCode", oDBs_Result.Offset, "")
            oDBs_Result.SetValue("U_TestName", oDBs_Result.Offset, "")
            oDBs_Result.SetValue("U_Parameter", oDBs_Result.Offset, "")
            oDBs_Result.SetValue("U_TestMethod", oDBs_Result.Offset, "")
            oDBs_Result.SetValue("U_Unit", oDBs_Result.Offset, "")

            oDBs_Result.SetValue("U_MinValue", oDBs_Result.Offset, "0")
            oDBs_Result.SetValue("U_MaxValue", oDBs_Result.Offset, "0")

            oDBs_Result.SetValue("U_ActualValue", oDBs_Result.Offset, "")
            oDBs_Result.SetValue("U_TextResult", oDBs_Result.Offset, "")

            oDBs_Result.SetValue("U_ResultStatus", oDBs_Result.Offset, "Pending")

            oDBs_Result.SetValue("U_InstrumentName", oDBs_Result.Offset, "")
            oDBs_Result.SetValue("U_TestedBy", oDBs_Result.Offset, "")

            oMatrixResult.SetLineData(oMatrixResult.VisualRowCount)

            oMatrixResult.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

    Public Sub SetNewLine_Attachment(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXPH_QCLABH")
            oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNXPH_QCLABATT")

            oMatrixAttach = objForm.Items.Item("MtxAttach").Specific

            oMatrixAttach.AddRow()

            oDBs_Attach.SetValue("LineId", oDBs_Attach.Offset, oMatrixAttach.VisualRowCount)
            oDBs_Attach.SetValue("U_AttachType", oDBs_Attach.Offset, "")
            oDBs_Attach.SetValue("U_FileName", oDBs_Attach.Offset, "")
            oDBs_Attach.SetValue("U_FilePath", oDBs_Attach.Offset, "")
            oDBs_Attach.SetValue("U_FileExt", oDBs_Attach.Offset, "")

            oDBs_Attach.SetValue("U_UploadedBy", oDBs_Attach.Offset,
                             objMain.objCompany.UserName)

            oDBs_Attach.SetValue("U_UploadedDate", oDBs_Attach.Offset,
                             Date.Now.ToString("yyyyMMdd"))

            oMatrixAttach.SetLineData(oMatrixAttach.VisualRowCount)

            oMatrixAttach.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

    Public Sub SetNewLine_Approval(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNXPH_QCLABH")
            oDBs_Approve = objForm.DataSources.DBDataSources.Item("@TNXPH_QCLABAPP")

            oMatrixApprove = objForm.Items.Item("MtxAppr").Specific

            oMatrixApprove.AddRow()

            oDBs_Approve.SetValue("LineId", oDBs_Approve.Offset, oMatrixApprove.VisualRowCount)
            oDBs_Approve.SetValue("U_ApprovalLevel", oDBs_Approve.Offset,
                              oMatrixApprove.VisualRowCount)

            oDBs_Approve.SetValue("U_ApproverRole", oDBs_Approve.Offset, "")
            oDBs_Approve.SetValue("U_ApproverUser", oDBs_Approve.Offset, "")

            oDBs_Approve.SetValue("U_ApprovalStatus", oDBs_Approve.Offset, "Pending")

            oDBs_Approve.SetValue("U_ApprovalDate", oDBs_Approve.Offset, "")
            oDBs_Approve.SetValue("U_ApprovalTime", oDBs_Approve.Offset, "")

            oDBs_Approve.SetValue("U_Comments", oDBs_Approve.Offset, "")

            oMatrixApprove.SetLineData(oMatrixApprove.VisualRowCount)

            oMatrixApprove.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

#End Region



#Region "Menu Event"
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "10X_LABTEST" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'oMatrixAttach = objForm.Items.Item("MXT_3").Specific
                'SetNewLine_Result(objForm.UniqueID)
                'SetNewLine_Attachment(objForm.UniqueID)
                'SetNewLine_Approval(objForm.UniqueID)
                ' objForm.Items.Item("APPI").Enabled = False
                Me.CreateForm()
                ' objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                '  Me.SetDefault(objForm.UniqueID)
                SetNewLine_Result(objForm.UniqueID)
                SetNewLine_Attachment(objForm.UniqueID)
                SetNewLine_Approval(objForm.UniqueID)
                '        oMatrixAttach = objForm.Items.Item("MXT_3").Specific
                ' objForm.Items.Item("APPI").Enabled = False
                ' Me.SetNewLine(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1281" And pVal.BeforeAction = False Then

            ElseIf pVal.MenuUID = "1293" AndAlso pVal.BeforeAction = True Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "TNXPH_QCLAB" Then Exit Sub

                    BubbleEvent = False

                    objForm.Freeze(True)

                    '=========================================================
                    ' DELETE ROW - TEST RESULT MATRIX
                    '=========================================================
                    If objForm.PaneLevel = 1 Then

                        oMatrixResult = CType(objForm.Items.Item("MtxResult").Specific,
                                  SAPbouiCOM.Matrix)

                        oDBs_Result = objForm.DataSources.DBDataSources.Item("@TNXPH_QCLABL")

                        Dim selectedRow As Integer =
            oMatrixResult.GetNextSelectedRow(0,
            SAPbouiCOM.BoOrderType.ot_RowOrder)

                        If selectedRow <= 0 Then

                            objMain.objApplication.StatusBar.SetText(
                "Please select Result row to delete",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            Exit Try

                        End If

                        'Delete selected row
                        oMatrixResult.DeleteRow(selectedRow)

                        'Sync Matrix to DB
                        oMatrixResult.FlushToDataSource()

                        'Remove extra datasource rows
                        While oDBs_Result.Size > oMatrixResult.VisualRowCount
                            oDBs_Result.RemoveRecord(oDBs_Result.Size - 1)
                        End While

                        'Keep minimum one row
                        If oDBs_Result.Size = 0 Then

                            oDBs_Result.InsertRecord(0)

                            oDBs_Result.SetValue("LineId", 0, "1")
                            oDBs_Result.SetValue("U_TestCode", 0, "")
                            oDBs_Result.SetValue("U_TestName", 0, "")
                            oDBs_Result.SetValue("U_Parameter", 0, "")
                            oDBs_Result.SetValue("U_TestMethod", 0, "")
                            oDBs_Result.SetValue("U_Unit", 0, "")
                            oDBs_Result.SetValue("U_MinValue", 0, "0")
                            oDBs_Result.SetValue("U_MaxValue", 0, "0")
                            oDBs_Result.SetValue("U_ActualValue", 0, "")
                            oDBs_Result.SetValue("U_TextResult", 0, "")
                            oDBs_Result.SetValue("U_ResultStatus", 0, "Pending")
                            oDBs_Result.SetValue("U_InstrumentName", 0, "")
                            oDBs_Result.SetValue("U_TestedBy", 0, "")

                        End If

                        'Re-sequence line numbers
                        For i As Integer = 0 To oDBs_Result.Size - 1

                            oDBs_Result.SetValue("LineId",
                                     i,
                                     (i + 1).ToString())

                        Next

                        oMatrixResult.LoadFromDataSource()
                        oMatrixResult.AutoResizeColumns()

                    End If

                    '=========================================================
                    ' DELETE ROW - APPROVAL MATRIX
                    '=========================================================
                    If objForm.PaneLevel = 2 Then

                        oMatrixApprove = CType(objForm.Items.Item("MtxAppr").Specific,
                                   SAPbouiCOM.Matrix)

                        oDBs_Approve = objForm.DataSources.DBDataSources.Item("@TNXPH_QCLABAPP")

                        Dim selectedRow As Integer =
            oMatrixApprove.GetNextSelectedRow(0,
            SAPbouiCOM.BoOrderType.ot_RowOrder)

                        If selectedRow <= 0 Then

                            objMain.objApplication.StatusBar.SetText(
                "Please select Approval row to delete",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            Exit Try

                        End If

                        'Delete selected row
                        oMatrixApprove.DeleteRow(selectedRow)

                        'Sync Matrix to DB
                        oMatrixApprove.FlushToDataSource()

                        'Remove extra datasource rows
                        While oDBs_Approve.Size > oMatrixApprove.VisualRowCount
                            oDBs_Approve.RemoveRecord(oDBs_Approve.Size - 1)
                        End While

                        'Keep minimum one row
                        If oDBs_Approve.Size = 0 Then

                            oDBs_Approve.InsertRecord(0)

                            oDBs_Approve.SetValue("LineId", 0, "1")
                            oDBs_Approve.SetValue("U_ApprovalLevel", 0, "1")
                            oDBs_Approve.SetValue("U_ApproverRole", 0, "")
                            oDBs_Approve.SetValue("U_ApproverUser", 0, "")
                            oDBs_Approve.SetValue("U_ApprovalStatus", 0, "Pending")
                            oDBs_Approve.SetValue("U_ApprovalDate", 0, "")
                            oDBs_Approve.SetValue("U_ApprovalTime", 0, "")
                            oDBs_Approve.SetValue("U_Comments", 0, "")

                        End If

                        'Re-sequence line numbers
                        For i As Integer = 0 To oDBs_Approve.Size - 1

                            oDBs_Approve.SetValue("LineId",
                                      i,
                                      (i + 1).ToString())

                        Next

                        oMatrixApprove.LoadFromDataSource()
                        oMatrixApprove.AutoResizeColumns()

                    End If

                    '=========================================================
                    ' DELETE ROW - ATTACHMENT MATRIX
                    '=========================================================
                    If objForm.PaneLevel = 3 Then

                        oMatrixAttach = CType(objForm.Items.Item("MtxAttach").Specific,
                                  SAPbouiCOM.Matrix)

                        oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNXPH_QCLABATT")

                        Dim selectedRow As Integer =
            oMatrixAttach.GetNextSelectedRow(0,
            SAPbouiCOM.BoOrderType.ot_RowOrder)

                        If selectedRow <= 0 Then

                            objMain.objApplication.StatusBar.SetText(
                "Please select Attachment row to delete",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            Exit Try

                        End If

                        'Delete selected row
                        oMatrixAttach.DeleteRow(selectedRow)

                        'Sync Matrix to DB
                        oMatrixAttach.FlushToDataSource()

                        'Remove extra datasource rows
                        While oDBs_Attach.Size > oMatrixAttach.VisualRowCount
                            oDBs_Attach.RemoveRecord(oDBs_Attach.Size - 1)
                        End While

                        'Keep minimum one row
                        If oDBs_Attach.Size = 0 Then

                            oDBs_Attach.InsertRecord(0)

                            oDBs_Attach.SetValue("LineId", 0, "1")
                            oDBs_Attach.SetValue("U_AttachType", 0, "")
                            oDBs_Attach.SetValue("U_FileName", 0, "")
                            oDBs_Attach.SetValue("U_FilePath", 0, "")
                            oDBs_Attach.SetValue("U_FileExt", 0, "")
                            oDBs_Attach.SetValue("U_UploadedBy", 0, "")
                            oDBs_Attach.SetValue("U_UploadedDate", 0, "")

                        End If

                        'Re-sequence line numbers
                        For i As Integer = 0 To oDBs_Attach.Size - 1

                            oDBs_Attach.SetValue("LineId",
                                     i,
                                     (i + 1).ToString())

                        Next

                        oMatrixAttach.LoadFromDataSource()
                        oMatrixAttach.AutoResizeColumns()

                    End If

                    '=========================================================
                    ' UPDATE MODE
                    '=========================================================
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If

                Catch ex As Exception

                    objMain.objApplication.StatusBar.SetText(
        "Delete Row Error : " & ex.Message,
        SAPbouiCOM.BoMessageTime.bmt_Short,
        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Finally

                    Try
                        objForm.Freeze(False)
                    Catch
                    End Try

                End Try
            ElseIf pVal.MenuUID = "519" AndAlso pVal.BeforeAction = True Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "VATR" Then Exit Sub

                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VATRP")

                    Dim frs As SAPbobsCOM.Recordset =
           objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'VATREPORT' AND ""Type"" = 'C'")

                    If frs.RecordCount = 0 Then

                        objMain.objApplication.MessageBox("VATREPORT Layout not found. Please import Crystal Layout with same name.", 0, "OK")

                    Else

                        objMain.objApplication.ActivateMenuItem(
               frs.Fields.Item(0).Value.ToString())

                        Dim CrForm As SAPbouiCOM.Form
                        Dim oedt As SAPbouiCOM.EditText

                        CrForm = objMain.objApplication.Forms.ActiveForm

                        oedt = CrForm.Items.Item("1000003").Specific

                        oedt.Value = oDBs_Head.GetValue("DocEntry", 0)

                        CrForm.Items.Item("1").Click(
               SAPbouiCOM.BoCellClickType.ct_Regular)

                    End If

                Catch ex As Exception

                    objMain.objApplication.StatusBar.SetText("Preview Error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                End Try



            ElseIf pVal.MenuUID = "520" AndAlso pVal.BeforeAction = True Then

                Try

                    Dim LayoutSelection As Integer =
           objMain.objApplication.MessageBox("Please select layout for printing", 1, "VAT Report", "")

                    If LayoutSelection = 1 Then

                        objForm = objMain.objApplication.Forms.ActiveForm

                        If objForm.TypeEx <> "VATR" Then Exit Sub

                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VATRP")

                        Dim frs As SAPbobsCOM.Recordset =
               objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        frs.DoQuery("SELECT ""MenuUID"" FROM ""OCMN"" WHERE ""Name"" = 'VATREPORT' AND ""Type"" = 'C'")

                        If frs.RecordCount = 0 Then

                            objMain.objApplication.MessageBox("VATREPORT Layout not found. Please import Crystal Layout with same name.", 0, "OK")

                        Else

                            objMain.objApplication.ActivateMenuItem(frs.Fields.Item(0).Value.ToString())

                            Dim CrForm As SAPbouiCOM.Form
                            Dim oedt As SAPbouiCOM.EditText

                            CrForm = objMain.objApplication.Forms.ActiveForm

                            oedt = CrForm.Items.Item("1000003").Specific

                            oedt.Value = oDBs_Head.GetValue("DocEntry", 0)

                            CrForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        End If

                    End If

                Catch ex As Exception

                    objMain.objApplication.StatusBar.SetText("Print Error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                End Try

            End If

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


#End Region



#Region "Item Event"
    Public Sub ItemEvent(ByVal FormUID As String,
                 ByRef pVal As SAPbouiCOM.ItemEvent,
                 ByRef BubbleEvent As Boolean)

        Try



            If pVal.EventType = BoEventTypes.et_ITEM_PRESSED _
                    AndAlso pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.Item(FormUID)
                If pVal.ItemUID = "1" Then
                    Me.CreateForm()
                End If

                '    If pVal.ItemUID = "Item_8" Then

                '        Dim FromDate As String = objForm.Items.Item("VRPY").Specific.Value
                '        Dim ToDate As String = objForm.Items.Item("VATTO").Specific.Value

                '        If FromDate = "" OrElse ToDate = "" Then
                '            objMain.objApplication.StatusBar.SetText(
                '          "Please enter From Date and To Date",
                '         SAPbouiCOM.BoMessageTime.bmt_Short,
                'SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '            Exit Sub
                '        End If

                '        Try
                '            'Open Tax Report
                '            objMain.objApplication.ActivateMenuItem("13068")

                '            Dim oTaxForm As SAPbouiCOM.Form
                '            oTaxForm = objMain.objApplication.Forms.ActiveForm

                '            'Set From Date and To Date
                '            CType(oTaxForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value = FromDate
                '            CType(oTaxForm.Items.Item("7").Specific, SAPbouiCOM.EditText).Value = ToDate
                '            oTaxForm.Items.Item("9").Click(BoCellClickType.ct_Regular)
                '        Catch ex As Exception
                '            objMain.objApplication.StatusBar.SetText(
                '"Tax Report Open Error : " & ex.Message,
                'SAPbouiCOM.BoMessageTime.bmt_Short,
                'SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '        End Try

                '    End If

                '    Select Case pVal.ItemUID


                '        Case "Item_7"      'Sales
                '            objForm.PaneLevel = 1

                '        Case "Item_9"      'Purchase
                '            objForm.PaneLevel = 2

                '        Case "Item_13"     'Attachment
                '            objForm.PaneLevel = 3
                '        Case "Item_11"
                '            Dim FromDate As String = GetEditValue(objForm, "VRPY")
                '            Dim ToDate As String = GetEditValue(objForm, "VATTO")



                '            FromDate = ConvertSAPDateToSQL(FromDate)
                '            ToDate = ConvertSAPDateToSQL(ToDate)

                '            ExportVATToExcel(FromDate, ToDate)


                '        Case "Item_10"     'Get Data Button

                '            Dim FromDate As String = GetEditValue(objForm, "VRPY")
                '            Dim ToDate As String = GetEditValue(objForm, "VATTO")

                '            If FromDate = "" OrElse ToDate = "" Then
                '                objMain.objApplication.StatusBar.SetText(
                '                    "Please enter VAT Return Period From and To Date",
                '                    SAPbouiCOM.BoMessageTime.bmt_Short,
                '                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '                Exit Sub
                '            End If


                '            If Not ValidateVATDates(objForm) Then
                '                Exit Sub
                '            End If


                '            FromDate = ConvertSAPDateToSQL(FromDate)
                '            ToDate = ConvertSAPDateToSQL(ToDate)

                '            LoadVATReportLines(FormUID, FromDate, ToDate)

                '            '  objForm.Items.Item("Item_10").Enabled = False


                '    End Select

            End If

            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK AndAlso pVal.BeforeAction = False Then

                If pVal.ItemUID = "btn_Del" Then

                    Try
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        objForm.Freeze(True)

                        oMatrixAttach = CType(objForm.Items.Item("MXT_3").Specific, SAPbouiCOM.Matrix)
                        oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_ATTACH_C3")

                        Dim selectedRow As Integer = 0

                        For i As Integer = 1 To oMatrixAttach.VisualRowCount
                            If oMatrixAttach.IsRowSelected(i) = True Then
                                selectedRow = i
                                Exit For
                            End If
                        Next

                        If selectedRow = 0 Then
                            objMain.objApplication.StatusBar.SetText("Please select attachment row.")
                            Exit Try
                        End If

                        oMatrixAttach.FlushToDataSource()

                        oDBs_Attach.RemoveRecord(selectedRow - 1)

                        If oDBs_Attach.Size = 0 Then
                            oDBs_Attach.InsertRecord(0)
                            oDBs_Attach.SetValue("LineId", 0, "1")
                            oDBs_Attach.SetValue("U_TPA", 0, "")
                            oDBs_Attach.SetValue("U_FN", 0, "")
                            oDBs_Attach.SetValue("U_ATD", 0, "")
                            oDBs_Attach.SetValue("U_FTT", 0, "")
                        Else

                            For i As Integer = 0 To oDBs_Attach.Size - 1
                                oDBs_Attach.SetValue("LineId", i, (i + 1).ToString())
                            Next
                        End If

                        oMatrixAttach.LoadFromDataSource()
                        oMatrixAttach.AutoResizeColumns()

                        ' Ensure form goes into update mode after deleting a row when it was in OK mode.
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If

                        objForm.Items.Item("btn_Del").Enabled = False

                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(
                "Delete Attachment Error : " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Finally
                        Try
                            objForm.Freeze(False)
                        Catch
                        End Try
                    End Try

                End If

            End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK AndAlso pVal.BeforeAction = False Then
                If pVal.ItemUID = "oMatrixAttach" And pVal.ColUID = "FPATH" And pVal.BeforeAction = False Then
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("oMatrixAttach").Specific


                    If pVal.Row > 0 AndAlso pVal.Row <= objMatrix.VisualRowCount Then
                        Dim fullPath As String = objMatrix.Columns.Item("FPATH").Cells.Item(pVal.Row).Specific.Value
                        If Not String.IsNullOrEmpty(fullPath) AndAlso fullPath.Contains("\") Then
                            Dim indexLoc As Integer = fullPath.LastIndexOf("\")
                            Dim filename As String = fullPath.Substring(indexLoc + 1)
                            objMatrix.Columns.Item("FNAME").Cells.Item(pVal.Row).Specific.Value = filename

                            objMatrix.Columns.Item("ATD").Cells.Item(pVal.Row).Specific.Value = DateTime.Now.ToString("yyyyMMdd")
                            objForm.Items.Item("btn_Del").Enabled = True
                        End If
                    End If
                End If
            End If
            '            'sreeja
            '            Select Case pVal.EventType
            '                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
            '                    objForm = objMain.objApplication.Forms.Item(FormUID)
            '                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = False AndAlso pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '                        Me.SetDefault(objForm.UniqueID)
            '                        objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            '                    End If
            '                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = True AndAlso pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '                        Try
            '                            If Not ValidateVATDates(objForm) Then
            '                                BubbleEvent = False
            '                                Exit Sub
            '                            End If

            '                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VATRP")

            '                            Dim dblTVRP As Double = 0

            '                            Try

            '                                Dim TVRPValue As String = oDBs_Head.GetValue("U_NVAT", 0).ToString().Trim()

            '                                If TVRPValue = "" Then

            '                                    objMain.objApplication.StatusBar.SetText("Net VAT Due value should not be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            '                                    BubbleEvent = False
            '                                    Exit Sub

            '                                End If

            '                                Double.TryParse(TVRPValue, dblTVRP)

            '                                If dblTVRP <= 0 Then

            '                                    objMain.objApplication.StatusBar.SetText("Net VAT Due value should be greater than Zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            '                                    BubbleEvent = False
            '                                    Exit Sub

            '                                End If

            '                            Catch ex As Exception

            '                                objMain.objApplication.StatusBar.SetText("NVAT Validation Error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            '                                BubbleEvent = False
            '                                Exit Sub

            '                            End Try

            '                            If oDBs_Head Is Nothing Then

            '                                Exit Try
            '                            End If

            '                            Dim currentAppId As String = ""
            '                            Try
            '                                currentAppId = oDBs_Head.GetValue("U_APPI", oDBs_Head.Offset).Trim()
            '                            Catch
            '                                currentAppId = ""
            '                            End Try

            '                            If String.IsNullOrEmpty(currentAppId) Then
            '                                AutoDocentryNumber(objForm.UniqueID)
            '                                Try
            '                                    currentAppId = oDBs_Head.GetValue("U_APPI", oDBs_Head.Offset).Trim()
            '                                Catch
            '                                    currentAppId = ""
            '                                End Try
            '                            End If

            '                            Dim DocType As String = "VPA"
            '                            Dim TableName As String = "@TNX_VATRP"
            '                            Dim AppIDField As String = "U_APPI"
            '                            Dim AppStatField As String = "U_FTY"

            '                            Dim rsapp As SAPbobsCOM.Recordset = Nothing
            '                            Dim approvalExists As Boolean = False

            '                            Try
            '                                rsapp = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            '                                Dim q As String =
            '                                "SELECT T0.""Code"" As ""TemplateID"",T1.""U_Name"" As ""Originator"", " &
            '                                "S1.""Code"" As ""Stage"",S2.""U_UKey"" As ""Authorizers"", " &
            '                                "S1.""U_NAP"" As ""NoOfAuth"",S1.""U_NRJ"" As ""NoOfRej"", " &
            '                                "T0.""CreateDate"", " &
            '                                "(Case When T3.""LineId""='1' Then 'S' Else 'O' End) As ""Status"" " &
            '                                "FROM ""@SBO_APPHDR"" T0 " &
            '                                "INNER JOIN ""@SBO_APPREQ"" T1 ON T1.""Code""=T0.""Code"" " &
            '                                "INNER JOIN ""@SBO_APPDOC"" T2 ON T2.""Code"" = T0.""Code"" " &
            '                                "INNER JOIN ""@SBO_APPAUT"" T3 ON T3.""Code"" = T0.""Code"" " &
            '                                "INNER JOIN ""@SBO_AST"" S1 ON T3.""U_M3_1""=S1.""Code"" " &
            '                                "INNER JOIN ""@SBO_AST_C0"" S2 ON S1.""Code""=S2.""Code"" " &
            '                                "WHERE ""U_Active"" = 'Y' " &
            '                                "AND T1.""U_Name"" = '" & objMain.objCompany.UserName & "' " &
            '                                "AND T2.""U_" & DocType & """ = 'Y' " &
            '                                "AND IFNULL(S2.""U_UKey"",'')<>'' " &
            '                                "ORDER BY T3.""LineId"""

            '                                rsapp.DoQuery(q)

            '                                If rsapp.RecordCount > 0 Then
            '                                    approvalExists = True
            '                                End If

            '                            Catch ex As Exception
            '                                objMain.objApplication.StatusBar.SetText("Approval check error: " & ex.Message)
            '                            Finally
            '                                If rsapp IsNot Nothing Then
            '                                    Marshal.ReleaseComObject(rsapp)
            '                                    rsapp = Nothing
            '                                End If
            '                            End Try

            '                            If approvalExists Then
            '                                ' Me.ApprovalTrigger(DocType, currentAppId, TableName, AppIDField, AppStatField, "VAT Report")
            '                            Else
            '                                Try
            '                                    oDBs_Head.SetValue("U_FTY", oDBs_Head.Offset, "A")

            '                                Catch ex As Exception

            '                                End Try

            '                                objMain.objApplication.StatusBar.SetText("Document will be auto-approved (no approval template).", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
            '                            End If

            '                        Catch ex As Exception
            '                            objMain.objApplication.StatusBar.SetText("ItemEvent (Before Add) error: " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            '                        End Try


            '                        Try
            '                            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VATRP")
            '                        Catch ex As Exception
            '                            oDBs_Head = Nothing
            '                        End Try

            '                    End If
            '                    If pVal.ItemUID = "Item_5" And pVal.BeforeAction = False And SAPbouiCOM.BoFormMode.fm_OK_MODE Then
            '                        Try
            '                            Dim AppID As String = objForm.Items.Item("APPI").Specific.value
            '                            Me.ApprovalTrigger("VPA", AppID, "@TNX_VATRP", "U_APPI", "U_FTY", "VAT Report")
            '                        Catch ex As Exception
            '                        End Try

            '                    End If
            '            End Select
            '            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED _
            'AndAlso pVal.BeforeAction = False Then

            '                If pVal.ItemUID = "MXT_1" AndAlso pVal.ColUID = "ACN" Then

            '                    Try

            '                        objForm = objMain.objApplication.Forms.Item(FormUID)

            '                        objMatrix = CType(objForm.Items.Item("MXT_1").Specific,
            '                              SAPbouiCOM.Matrix)

            '                        Dim AccountCode As String =
            '            CType(objMatrix.Columns.Item("ACN").Cells.Item(pVal.Row).Specific,
            '                  SAPbouiCOM.EditText).Value

            '                        If AccountCode <> "" Then

            '                            objMain.objApplication.OpenForm(
            '                    SAPbouiCOM.BoFormObjectEnum.fo_GLAccounts,
            '                    "",
            '                    AccountCode)

            '                        End If

            '                    Catch ex As Exception

            '                        objMain.objApplication.StatusBar.SetText(
            '            "Linked Button Error : " & ex.Message,
            '            SAPbouiCOM.BoMessageTime.bmt_Short,
            '            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            '                    End Try

            '                End If

            '            End If
            '            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED _
            'AndAlso pVal.BeforeAction = False Then

            '                If pVal.ItemUID = "MXT_2" AndAlso pVal.ColUID = "ARN" Then

            '                    Try

            '                        objForm = objMain.objApplication.Forms.Item(FormUID)

            '                        objMatrix = CType(objForm.Items.Item("MXT_2").Specific,
            '                              SAPbouiCOM.Matrix)

            '                        Dim AccountCode As String =
            '            CType(objMatrix.Columns.Item("ARN").Cells.Item(pVal.Row).Specific,
            '                  SAPbouiCOM.EditText).Value

            '                        If AccountCode <> "" Then

            '                            objMain.objApplication.OpenForm(
            '                    SAPbouiCOM.BoFormObjectEnum.fo_GLAccounts,
            '                    "",
            '                    AccountCode)

            '                        End If

            '                    Catch ex As Exception

            '                        objMain.objApplication.StatusBar.SetText(
            '            "Linked Button Error : " & ex.Message,
            '            SAPbouiCOM.BoMessageTime.bmt_Short,
            '            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            '                    End Try

            '                End If

            '            End If


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("ItemEvent Error: " & ex.Message)
        End Try

    End Sub


    Private Function GetEditValue(ByVal oForm As SAPbouiCOM.Form,
                              ByVal ItemUID As String) As String
        Try
            Return CType(oForm.Items.Item(ItemUID).Specific,
                     SAPbouiCOM.EditText).Value.Trim()
        Catch
            Return ""
        End Try
    End Function
    'Private Function ConvertSAPDateToSQL(ByVal sapDate As String) As String
    '    Try
    '        If sapDate = "" Then Return ""

    '        'SAP B1 date format: yyyyMMdd
    '        If sapDate.Length = 8 Then
    '            Return sapDate.Substring(0, 4) & "-" &
    '               sapDate.Substring(4, 2) & "-" &
    '               sapDate.Substring(6, 2)
    '        End If

    '        Return sapDate

    '    Catch
    '        Return sapDate
    '    End Try
    'End Function
    'Private Sub LoadVATReportLines(ByVal FormUID As String,
    '                           ByVal FromDate As String,
    '                           ByVal ToDate As String)

    '    Dim oForm As SAPbouiCOM.Form = Nothing

    '    Try
    '        oForm = objMain.objApplication.Forms.Item(FormUID)
    '        oForm.Freeze(True)

    '        Dim oMatrixSales As SAPbouiCOM.Matrix =
    '        CType(oForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)

    '        Dim oMatrixPurchase As SAPbouiCOM.Matrix =
    '        CType(oForm.Items.Item("MXT_2").Specific, SAPbouiCOM.Matrix)

    '        Dim oDBH As SAPbouiCOM.DBDataSource =
    '        oForm.DataSources.DBDataSources.Item("@TNX_VATRP")

    '        Dim oDBSales As SAPbouiCOM.DBDataSource =
    '        oForm.DataSources.DBDataSources.Item("@TNX_VATCTM_C1")

    '        Dim oDBPurchase As SAPbouiCOM.DBDataSource =
    '        oForm.DataSources.DBDataSources.Item("@TNX_VATRP_C0")

    '        ClearMatrix(oMatrixSales, oDBSales)
    '        ClearMatrix(oMatrixPurchase, oDBPurchase)

    '        Dim rs As SAPbobsCOM.Recordset =
    '        objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        Dim sql As String =
    '        "CALL ""TNX_VAT_RETURN_RPT"" " &
    '        "(TO_DATE('" & FromDate & "','YYYY-MM-DD'), " &
    '        "TO_DATE('" & ToDate & "','YYYY-MM-DD'))"

    '        rs.DoQuery(sql)

    '        If rs.RecordCount = 0 Then
    '            objMain.objApplication.StatusBar.SetText("No VAT report data found")
    '            Exit Sub
    '        End If

    '        Dim rowSales As Integer = 0
    '        Dim rowPurchase As Integer = 0

    '        While Not rs.EoF

    '            Dim BoxNo As String = GetRSValue(rs, "BoxNo")
    '            Dim AccNo As String = GetRSValue(rs, "AccountNo")
    '            Dim Desc As String = GetRSValue(rs, "Description")
    '            Dim Amt As String = ToNumber(GetRSValue(rs, "Amount_AED"))
    '            Dim Vat As String = ToNumber(GetRSValue(rs, "VAT_AED"))
    '            Dim Adj As String = ToNumber(GetRSValue(rs, "Adjustment_AED"))

    '            If BoxNo = "12" Then oDBH.SetValue("U_TVD", 0, Vat)
    '            If BoxNo = "13" Then oDBH.SetValue("U_TVRP", 0, Vat)
    '            If BoxNo = "14" Then oDBH.SetValue("U_NVAT", 0, Vat)

    '            If IsOutputVATBox(BoxNo) Then

    '                oDBSales.InsertRecord(rowSales)
    '                oDBSales.SetValue("LineId", rowSales, (rowSales + 1).ToString())
    '                oDBSales.SetValue("U_ACN", rowSales, AccNo)
    '                oDBSales.SetValue("U_TRN", rowSales, Desc)
    '                oDBSales.SetValue("U_AMT", rowSales, Amt)
    '                oDBSales.SetValue("U_VATA", rowSales, Vat)
    '                oDBSales.SetValue("U_AST", rowSales, Adj)

    '                rowSales += 1

    '            ElseIf IsInputVATBox(BoxNo) Then

    '                oDBPurchase.InsertRecord(rowPurchase)
    '                oDBPurchase.SetValue("LineId", rowPurchase, (rowPurchase + 1).ToString())
    '                oDBPurchase.SetValue("U_ARN", rowPurchase, AccNo)
    '                oDBPurchase.SetValue("U_VATE", rowPurchase, Desc)
    '                oDBPurchase.SetValue("U_AUT", rowPurchase, Amt)
    '                oDBPurchase.SetValue("U_RVAT", rowPurchase, Vat)
    '                oDBPurchase.SetValue("U_AVAT", rowPurchase, Adj)

    '                rowPurchase += 1

    '            End If

    '            rs.MoveNext()

    '        End While

    '        oForm.PaneLevel = 1
    '        oMatrixSales.LoadFromDataSource()
    '        oMatrixSales.AutoResizeColumns()

    '        oForm.PaneLevel = 2
    '        oMatrixPurchase.LoadFromDataSource()
    '        oMatrixPurchase.AutoResizeColumns()

    '        oForm.PaneLevel = 1

    '        objMain.objApplication.StatusBar.SetText(
    '        "VAT matrix data loaded successfully",
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(
    '        "LoadVATReportLines Error: " & ex.Message,
    '        SAPbouiCOM.BoMessageTime.bmt_Short,
    '        SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    Finally
    '        If oForm IsNot Nothing Then oForm.Freeze(False)
    '    End Try

    'End Sub
    'Private Sub ClearMatrix(ByVal oMatrix As SAPbouiCOM.Matrix,
    '                    ByVal oDB As SAPbouiCOM.DBDataSource)

    '    oMatrix.FlushToDataSource()

    '    For i As Integer = oDB.Size - 1 To 0 Step -1
    '        oDB.RemoveRecord(i)
    '    Next

    '    oMatrix.Clear()

    'End Sub
    'Private Sub LoadVATHeaderOnly(ByVal FormUID As String)

    '    Dim oForm As SAPbouiCOM.Form = Nothing

    '    Try

    '        oForm = objMain.objApplication.Forms.Item(FormUID)

    '        oForm.Freeze(True)

    '        Dim oDBH As SAPbouiCOM.DBDataSource =
    '    oForm.DataSources.DBDataSources.Item("@TNX_VATRP")

    '        If oDBH.Size = 0 Then
    '            oDBH.InsertRecord(0)
    '        End If

    '        oDBH.Offset = 0

    '        Dim rs As SAPbobsCOM.Recordset =
    '    objMain.objCompany.GetBusinessObject(
    '    SAPbobsCOM.BoObjectTypes.BoRecordset)

    '        Dim sql As String =
    '    "SELECT " &
    '    "CURRENT_DATE AS ""SubmissionDate"", " &
    '    "IFNULL(""TaxIdNum"",'') AS ""TRN"", " &
    '    "IFNULL(""CompnyName"",'') AS ""TaxablePersonNameEnglish"", " &
    '    "IFNULL(""PrintHeadr"",'') AS ""TaxablePersonNameArabic"", " &
    '    "IFNULL(""CompnyAddr"",'') AS ""TaxablePersonAddress"" " &
    '    "FROM OADM"

    '        rs.DoQuery(sql)

    '        If rs.RecordCount > 0 Then

    '            Try
    '                oDBH.SetValue(
    '            "U_SD",
    '            0,
    '            ToSAPDate(GetRSValue(rs, "SubmissionDate")))
    '            Catch
    '            End Try

    '            Try
    '                oDBH.SetValue(
    '            "U_TRNM",
    '            0,
    '            GetRSValue(rs, "TRN"))
    '            Catch
    '            End Try

    '            Try
    '                oDBH.SetValue(
    '            "U_TPNE",
    '            0,
    '            GetRSValue(rs, "TaxablePersonNameEnglish"))
    '            Catch
    '            End Try

    '            Try
    '                oDBH.SetValue(
    '            "U_TPNA",
    '            0,
    '            GetRSValue(rs, "TaxablePersonNameArabic"))
    '            Catch
    '            End Try

    '            Try
    '                oDBH.SetValue(
    '            "U_TPA",
    '            0,
    '            GetRSValue(rs, "TaxablePersonAddress"))
    '            Catch
    '            End Try

    '        End If

    '    Catch ex As Exception

    '        objMain.objApplication.StatusBar.SetText(
    '    "LoadVATHeaderOnly Error: " & ex.Message,
    '    SAPbouiCOM.BoMessageTime.bmt_Short,
    '    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    Finally

    '        Try
    '            If oForm IsNot Nothing Then
    '                oForm.Freeze(False)
    '            End If
    '        Catch
    '        End Try

    '    End Try

    'End Sub

#End Region

End Class

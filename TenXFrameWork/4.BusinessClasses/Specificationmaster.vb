Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.AccessControl
Imports System.Text

Public Class Specificationmaster



#Region "Declaration"

    Public objForm As SAPbouiCOM.Form
    Public objMatrix, objMatrix1, objMatrix2, objMatrix3 As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim oDBs_Attach As SAPbouiCOM.DBDataSource
    Dim oDS As SAPbouiCOM.DBDataSource


#End Region

#Region "Create Form"

    Public Sub CreateForm()

        Try

            objMain.objUtilities.LoadForm(
            "SpecificationMaster.xml",
            "XPH_QSPEC",
            ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm(
            "XPH_QSPEC",
            objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)


            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PH_QSPECH")

            objMatrix = objForm.Items.Item("MtxParam").Specific
            objMatrix1 = objForm.Items.Item("MtxMap").Specific
            oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_QSPECM_ATT")

            Me.SetNewLine(objForm.UniqueID)
            Me.SetNewLine1(objForm.UniqueID)
            Me.SetNewLine2(objForm.UniqueID)
            oDBs_Head.SetValue("U_EffDate", 0, Date.Now.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_Status", 0, "Draft")
            oDBs_Head.SetValue("Docnum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "XPH_QSPEC", "Primary"))
            'oDBs_Head.SetValue("U_DA", 0, DateTime.Now.ToString("yyyyMMdd"))
            objForm.Items.Item("Docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("Docnum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("SpecCode").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("SpecCode").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("SpecName").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("SpecName").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ItemCode").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ItemCode").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ItemName").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("ItemName").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("ItemType").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ItemType").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Category").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Category").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Version").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Version").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("EffDate").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ValidFrom").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ValidTo").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Status").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("AppBy").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("AppBy").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("AppDate").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("AppDate").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("Remarks").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("MtxParam").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("MtxMap").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("1").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("2").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("SpecCode").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("SpecCode").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("SpecName").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("SpecName").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ItemCode").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ItemCode").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ItemName").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("ItemName").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("ItemType").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ItemType").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Category").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Category").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("Version").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Version").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("EffDate").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ValidFrom").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("ValidTo").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("Status").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("AppBy").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("AppBy").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("AppDate").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("AppDate").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("Remarks").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("MtxParam").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("MtxMap").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("1").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("2").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Me.objForm.EnableMenu("1282", True)
            Me.objForm.EnableMenu("1288", True)
            Me.objForm.EnableMenu("1289", True)
            Me.objForm.EnableMenu("1290", True)
            Me.objForm.EnableMenu("1291", True)
            Me.objForm.EnableMenu("1292", True)
            Me.objForm.EnableMenu("1293", True)

            SetDefault(objForm.UniqueID)

            objForm.Items.Item("Item_7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)


        End Try

    End Sub

#End Region

#Region "Default Values"

    Public Sub SetDefault(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PH_QSPECH")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PH_QSPECL")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PH_QSPECM")
            oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_QSPECM_ATT")
            objForm.Freeze(True)



            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                oDBs_Head.SetValue("Docnum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "XPH_QSPEC", "Primary"))
                oDBs_Head.SetValue("U_EffDate", 0, Date.Now.ToString("yyyyMMdd"))
                oDBs_Head.SetValue("U_Status", 0, "Draft")
                ' oDBs_Head.SetValue("U_FTY", 0, "Open")
                '  AddAttachmentNewLine(FormUID)
            End If
            '  Me.CreateForm()

            objForm.PaneLevel = 1

            Me.SetNewLine(objForm.UniqueID)
            Me.SetNewLine1(objForm.UniqueID)
            Me.SetNewLine2(objForm.UniqueID)


            objForm.Freeze(False)

        Catch ex As Exception

            Try
                objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub


#End Region

#Region "Add New Line"

    Public Sub SetNewLine(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PH_QSPECH")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PH_QSPECL")

            objMatrix = objForm.Items.Item("MtxParam").Specific

            objMatrix.AddRow()

            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)

            oDBs_Details.SetValue("U_TestCode", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_TestName", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_TestCat", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_TestMethod", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Unit", oDBs_Details.Offset, "")

            oDBs_Details.SetValue("U_MinValue", oDBs_Details.Offset, "0")
            oDBs_Details.SetValue("U_MaxValue", oDBs_Details.Offset, "0")
            oDBs_Details.SetValue("U_TargetVal", oDBs_Details.Offset, "0")

            oDBs_Details.SetValue("U_TextLimit", oDBs_Details.Offset, "")

            oDBs_Details.SetValue("U_ResultType", oDBs_Details.Offset, "")

            oDBs_Details.SetValue("U_Mandatory", oDBs_Details.Offset, "Y")

            oDBs_Details.SetValue("U_SeqNo", oDBs_Details.Offset, objMatrix.VisualRowCount)

            oDBs_Details.SetValue("U_Remarks", oDBs_Details.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)

            objMatrix.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub
    Public Sub SetNewLine1(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PH_QSPECH")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PH_QSPECM")

            objMatrix = objForm.Items.Item("MtxMap").Specific

            objMatrix.AddRow()

            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)

            oDBs_Details.SetValue("U_TestCode", oDBs_Details.Offset, "")

            oDBs_Details.SetValue("U_MethodCode", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_MethodName", oDBs_Details.Offset, "")

            oDBs_Details.SetValue("U_SOPNo", oDBs_Details.Offset, "")

            oDBs_Details.SetValue("U_Instrument", oDBs_Details.Offset, "")

            oDBs_Details.SetValue("U_CalibReq", oDBs_Details.Offset, "N")

            oDBs_Details.SetValue("U_Frequency", oDBs_Details.Offset, "")

            oDBs_Details.SetValue("U_Remarks", oDBs_Details.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)

            objMatrix.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub
    Public Sub SetNewLine2(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PH_QSPECH")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_QSPECM_ATT")

            objMatrix = objForm.Items.Item("Item_4").Specific


            objMatrix.AddRow()
            oDBs_Attach.SetValue("LineId", oDBs_Attach.Offset, objMatrix.VisualRowCount)
            oDBs_Attach.SetValue("U_TPH", oDBs_Attach.Offset, "")   'Target Path
            oDBs_Attach.SetValue("U_FNM", oDBs_Attach.Offset, "")    'File Name
            oDBs_Attach.SetValue("U_FTR", oDBs_Attach.Offset, "")   'Free Text
            oDBs_Attach.SetValue("U_ATCD", oDBs_Attach.Offset, "") 'Attachment Date
            objMatrix.SetLineData(objMatrix.VisualRowCount)

            objMatrix.SetLineData(objMatrix.VisualRowCount)

            objMatrix.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub
#End Region



#Region "Menu Event"
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent,
              ByRef BubbleEvent As Boolean)

        Try

            If pVal.MenuUID = "10X_SPEC" And pVal.BeforeAction = False Then

                Me.CreateForm()

            ElseIf pVal.MenuUID = "1282" _
        And pVal.BeforeAction = False Then


                SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1292" _
        And pVal.BeforeAction = False Then

                objMatrix = objForm.Items.Item("MtxParam").Specific
                objMatrix1 = objForm.Items.Item("MtxMap").Specific

                Me.SetNewLine(objForm.UniqueID)
                Me.SetNewLine1(objForm.UniqueID)
                Me.SetNewLine2(objForm.UniqueID)
                ' Me.SetNewLine(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1281" And pVal.BeforeAction = False Then

            ElseIf pVal.MenuUID = "1293" AndAlso pVal.BeforeAction = True Then
                objForm = objMain.objApplication.Forms.Item("XPH_QSPEC")

                If objForm.TypeEx <> "XPH_QSPEC" Then Exit Sub

                objForm.Freeze(True)

                objMatrix = CType(objForm.Items.Item("MtxParam").Specific,
                                  SAPbouiCOM.Matrix)

                Select Case objMatrix.UniqueID

                    Case "MtxParam"

                        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PH_QSPECL")

                    Case "MtxMap"

                        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PH_QSPECM")

                    Case "Item_3"

                        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PH_QSPECM")

                End Select

                Dim selectedRow As Integer =
                    objMatrix.GetNextSelectedRow(0,
                    SAPbouiCOM.BoOrderType.ot_RowOrder)

                If selectedRow <= 0 Then

                    objMain.objApplication.StatusBar.SetText(
                        "Please select row to delete",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    Exit Try

                End If


                '=====================================================
                ' Delete Selected Row
                '=====================================================
                objMatrix.DeleteRow(selectedRow)

                objMatrix.FlushToDataSource()


                '=====================================================
                ' Remove Extra Records
                '=====================================================
                While oDBs_Details.Size > objMatrix.VisualRowCount

                    oDBs_Details.RemoveRecord(oDBs_Details.Size - 1)

                End While


                If oDBs_Details.Size = 0 Then

                    oDBs_Details.InsertRecord(0)

                    oDBs_Details.SetValue("LineId", 0, "1")

                    If objMatrix.UniqueID = "MtxParam" Then

                        oDBs_Details.SetValue("U_TestCode", 0, "")
                        oDBs_Details.SetValue("U_TestName", 0, "")
                        oDBs_Details.SetValue("U_TestCategory", 0, "")
                        oDBs_Details.SetValue("U_TestMethod", 0, "")
                        oDBs_Details.SetValue("U_Unit", 0, "")

                        oDBs_Details.SetValue("U_MinValue", 0, "0")
                        oDBs_Details.SetValue("U_MaxValue", 0, "0")
                        oDBs_Details.SetValue("U_TargetValue", 0, "0")

                        oDBs_Details.SetValue("U_TextLimit", 0, "")

                        oDBs_Details.SetValue("U_ResultType", 0, "Numeric")

                        oDBs_Details.SetValue("U_Mandatory", 0, "Y")

                        oDBs_Details.SetValue("U_SeqNo", 0, "1")

                        oDBs_Details.SetValue("U_Remarks", 0, "")

                    End If

                    If objMatrix1.UniqueID = "MtxMap" Or objMatrix3.UniqueID = "Item_3" Then

                        oDBs_Details.SetValue("U_TestCode", 0, "")

                        oDBs_Details.SetValue("U_MethodCode", 0, "")
                        oDBs_Details.SetValue("U_MethodName", 0, "")

                        oDBs_Details.SetValue("U_SOPNo", 0, "")

                        oDBs_Details.SetValue("U_Instrument", 0, "")

                        oDBs_Details.SetValue("U_CalibReq", 0, "N")

                        oDBs_Details.SetValue("U_Frequency", 0, "")

                        oDBs_Details.SetValue("U_Remarks", 0, "")

                    End If
                    If oDBs_Attach.Size = 0 Then

                        oDBs_Attach.InsertRecord(0)

                        oDBs_Attach.SetValue("LineId", 0, "1")
                        oDBs_Attach.SetValue("U_TPA", 0, "")
                        oDBs_Attach.SetValue("U_FN", 0, "")
                        oDBs_Attach.SetValue("U_ATD", 0, "")
                        oDBs_Attach.SetValue("U_FTT", 0, "")

                    End If

                End If

                For i As Integer = 0 To oDBs_Details.Size - 1

                    oDBs_Details.SetValue("LineId",
                                          i,
                                          (i + 1).ToString())

                Next

                objMatrix.LoadFromDataSource()

                objMatrix.AutoResizeColumns()


                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                End If



            End If



        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText("Print Error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub


#End Region

    Public Sub AutoDocentryNumber(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VATRP")

            If oDBs_Head.Size = 0 Then
                oDBs_Head.InsertRecord(0)
            End If

            Dim oRsDocNum As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            Dim Query1 As String = "SELECT IFNULL(MAX(""DocNum""),0)+1 AS ""DocNum"" FROM ""@TNX_VATRP"""

            oRsDocNum.DoQuery(Query1)

            oDBs_Head.SetValue("DocNum", 0, oRsDocNum.Fields.Item("DocNum").Value.ToString())


            Dim rsAppId As SAPbobsCOM.Recordset = CType(objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            Dim str As String = "SELECT 'V' || LPAD(" & "TO_NVARCHAR(IFNULL(MAX(TO_INTEGER(REPLACE(""U_APPI"", 'V', ''))),0)+1), 6, '0') AS ""AppId"" " & "FROM ""@TNX_VATRP"""

            rsAppId.DoQuery(str)

            oDBs_Head.SetValue("U_APPI", 0, rsAppId.Fields.Item("AppId").Value.ToString())

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("AutoDocentryNumber Error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub


    Public Sub ItemEvent(ByVal FormUID As String,
             ByRef pVal As SAPbouiCOM.ItemEvent,
             ByRef BubbleEvent As Boolean)

        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK AndAlso pVal.BeforeAction = False Then
            If pVal.ItemUID = "1" Then
                Me.CreateForm()
            End If
            If pVal.ItemUID = "Item_4" Then
                Try
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    Dim objMatrix As SAPbouiCOM.Matrix = CType(objForm.Items.Item("Item_4").Specific, SAPbouiCOM.Matrix)

                    If pVal.Row > 0 AndAlso pVal.Row <= objMatrix.VisualRowCount Then

                        Dim fullPath As String = ""
                        Dim colId As String = pVal.ColUID

                        ' Try to read value from the clicked column; fallback to Col_0
                        Try
                            fullPath = Convert.ToString(objMatrix.Columns.Item(colId).Cells.Item(pVal.Row).Specific.Value)
                        Catch
                            Try
                                fullPath = Convert.ToString(objMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.Value)
                            Catch
                                fullPath = String.Empty
                            End Try
                        End Try

                        If Not String.IsNullOrEmpty(fullPath) Then
                            ' Support both backslash and forward slash, use Path.GetFileName for reliability
                            Dim filename As String = Path.GetFileName(fullPath)
                            If Not String.IsNullOrEmpty(filename) Then
                                objMatrix.Columns.Item("FN").Cells.Item(pVal.Row).Specific.Value = filename
                                objMatrix.Columns.Item("ATD").Cells.Item(pVal.Row).Specific.Value = DateTime.Now.ToString("yyyyMMdd")
                                objForm.Items.Item("btn_Del").Enabled = True
                                objMatrix.FlushToDataSource()
                            End If
                        End If
                    End If

                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText("Attachment double-click error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try
            End If
        End If


    End Sub






End Class

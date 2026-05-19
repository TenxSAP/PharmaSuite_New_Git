

Imports SAPbouiCOM
Imports SAPbobsCOM
Imports System.IO
Imports System.Text
Imports System.Diagnostics
Imports System.Runtime.InteropServices

Public Class ClsCOAManagement

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
            objMain.objUtilities.LoadForm("COAManagement.xml", "10X_COA", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("10X_COA",
                      objMain.objApplication.Forms.ActiveForm.TypeCount)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_COA_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_COA_T")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_COA_A")
            oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_COA_A")
            objMatrix = objForm.Items.Item("MXT_1").Specific
            objMatrix = objForm.Items.Item("MXT_2").Specific
            objMatrix = objForm.Items.Item("MXT_3").Specific

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNXCOAUDO", "Primary"))
            'oDBs_Head.SetValue("U_DA", 0, DateTime.Now.ToString("yyyyMMdd"))
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("DS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            Me.objForm.EnableMenu("1282", True)

            Me.objForm.EnableMenu("1292", True)
            SetDefault(objForm.UniqueID)
            ' Me.SetNewLine(objForm.UniqueID)
            Me.objForm.EnableMenu("1293", True)
            'objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("APPI").Enabled = False
            'objForm.Items.Item("Item_7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
            "Coa Management Form Loaded Successfully",
            BoMessageTime.bmt_Short,
            BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            Try
                objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            BoMessageTime.bmt_Short,
            BoStatusBarMessageType.smt_Warning)

        End Try

    End Sub

#End Region

#Region "Default Values"

    Public Sub SetDefault(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_COA_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_COA_T")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_COA_A")
            oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_COA_APP")
            objForm.Freeze(True)


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

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            objMatrix = objForm.Items.Item(MatrixUID).Specific

            If MatrixUID = "MXT_1" Then

                oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_COA_T")

                objMatrix.AddRow()
                oDBs_Details.Offset = objMatrix.VisualRowCount - 1

                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount.ToString())
                oDBs_Details.SetValue("U_TestCode", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_TestName", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_TestM", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Unit", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_SpecMin", oDBs_Details.Offset, "0")
                oDBs_Details.SetValue("U_SpecMax", oDBs_Details.Offset, "0")
                oDBs_Details.SetValue("U_SpecText", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ResultV", oDBs_Details.Offset, "0")
                oDBs_Details.SetValue("U_ResultT", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_ResultS", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Analyst", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_TestD", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Instrument", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Remarks", oDBs_Details.Offset, "")

                objMatrix.SetLineData(objMatrix.VisualRowCount)

            ElseIf MatrixUID = "MXT_2" Then


                oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_COA_A")

                objMatrix.AddRow()
                oDBs_Details1.Offset = objMatrix.VisualRowCount - 1

                oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix.VisualRowCount.ToString())
                oDBs_Details1.SetValue("U_FileN", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_FileT", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_FileP", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_AttachE", oDBs_Details1.Offset, "0")
                oDBs_Details1.SetValue("U_UPU", oDBs_Details1.Offset, objMain.objCompany.UserName)
                oDBs_Details1.SetValue("U_UPD", oDBs_Details1.Offset, DateTime.Now.ToString("yyyyMMdd"))
                oDBs_Details1.SetValue("U_RM", oDBs_Details1.Offset, "")

                objMatrix.SetLineData(objMatrix.VisualRowCount)

            ElseIf MatrixUID = "MXT_3" Then


                oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_COA_APP")

                objMatrix.AddRow()
                oDBs_Attach.Offset = objMatrix.VisualRowCount - 1

                oDBs_Attach.SetValue("LineId", oDBs_Attach.Offset, objMatrix.VisualRowCount.ToString())
                oDBs_Attach.SetValue("U_ALevel", oDBs_Attach.Offset, "0")
                oDBs_Attach.SetValue("U_AppR", oDBs_Attach.Offset, "")
                oDBs_Attach.SetValue("U_AppU", oDBs_Attach.Offset, "")
                oDBs_Attach.SetValue("U_Status", oDBs_Attach.Offset, "Pending")
                oDBs_Attach.SetValue("U_ActionA", oDBs_Attach.Offset, "")
                oDBs_Attach.SetValue("U_Esign", oDBs_Attach.Offset, "")
                oDBs_Attach.SetValue("U_Com", oDBs_Attach.Offset, "")

                objMatrix.SetLineData(objMatrix.VisualRowCount)

            End If

            objMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "SetNewLine Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            Try
                objForm.Freeze(False)
            Catch
            End Try
        End Try

    End Sub


#End Region




    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "10X_COA" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

                '
                ' objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                '  Me.SetDefault(objForm.UniqueID)
                ' Me.SetNewLine(objForm.UniqueID)

                objMatrix = objForm.Items.Item("MXT_3").Specific
                objForm.Items.Item("APPI").Enabled = False
                ' Me.SetNewLine(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1281" And pVal.BeforeAction = False Then

            ElseIf pVal.MenuUID = "1293" AndAlso pVal.BeforeAction = True Then



            End If


        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST


                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED


                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED



                Case SAPbouiCOM.BoEventTypes.et_VALIDATE


            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub








End Class

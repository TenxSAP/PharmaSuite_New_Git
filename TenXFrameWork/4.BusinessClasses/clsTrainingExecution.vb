Imports Org.BouncyCastle.Security
Imports SAPbouiCOM

Public Class clsTrainingExecution

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
            objMain.objUtilities.LoadForm("TrainingExecution.xml", "frm_TRNEXE", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("frm_TRNEXE", objMain.objApplication.Forms.ActiveForm.TypeCount)
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
            If pVal.MenuUID = "TRAINE" And pVal.BeforeAction = False Then
                Me.CreateForm()

            ElseIf pVal.MenuUID = "1281" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                If objForm.TypeEx = "frm_TRNEXE" Then
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

                If objForm.TypeEx = "frm_TRNEXE" Then

                    objMatrix1 = objForm.Items.Item("0_U_G").Specific
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

                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then

                        If (pVal.ItemUID = "0_U_G" Or
                            pVal.ItemUID = "1_U_G") And pVal.BeforeAction = True Then

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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_TRNEH")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_TRNEL")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_TRNASM")


            objMatrix1 = objForm.Items.Item("0_U_G").Specific
            objMatrix2 = objForm.Items.Item("1_U_G").Specific

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_TRNEXE"))
            oDBs_Head.SetValue("U_SessDate", 0, DateTime.Now.ToString("yyyyMMdd"))

            objForm.Items.Item("0_U_FD").Click(BoCellClickType.ct_Regular)

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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_TRNEH")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_TRNEL")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_TRNASM")

            objMatrix1 = objForm.Items.Item("0_U_G").Specific
            objMatrix2 = objForm.Items.Item("1_U_G").Specific


            If objMatrix1.VisualRowCount = 0 Then

                objMatrix1.AddRow()

                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix1.VisualRowCount)

                oDBs_Details.SetValue("U_EmpID", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_EmpName", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Attended", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_AttTime", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Sign", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Status", oDBs_Details.Offset, "")
                oDBs_Details.SetValue("U_Remarks", oDBs_Details.Offset, "")

                objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                objMatrix1.AutoResizeColumns()

            End If


            If objMatrix2.VisualRowCount = 0 Then

                objMatrix2.AddRow()

                oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix2.VisualRowCount)
                oDBs_Details1.SetValue("U_EmpID", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_AssType", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_TotMarks", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_MarksOb", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_PassPer", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_Result", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_AssBy", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_AssDate", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_Remarks", oDBs_Details1.Offset, "")

                objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                objMatrix2.AutoResizeColumns()

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
                Case "0_U_G"

                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_TRNPL")

                    objMatrix1 = objForm.Items.Item("0_U_G").Specific

                    objMatrix1.AddRow()

                    oDBs_Details.SetValue("LineId",
                                          oDBs_Details.Offset,
                                          objMatrix1.VisualRowCount.ToString())

                    oDBs_Details.SetValue("U_EmpID", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_EmpName", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Position", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Dept", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Required", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_DueDate", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Status", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_Remarks", oDBs_Details.Offset, "")

                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                    objMatrix1.AutoResizeColumns()



                Case "1_U_G"

                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_TRNP_ATT")

                    objMatrix2 = objForm.Items.Item("1_U_G").Specific

                    objMatrix2.AddRow()

                    oDBs_Details1.SetValue("LineId",
                                           oDBs_Details1.Offset,
                                           objMatrix2.VisualRowCount.ToString())

                    oDBs_Details1.SetValue("U_TPA", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_FN", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_ATD", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_FTT", oDBs_Details1.Offset, "")

                    objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                    objMatrix2.AutoResizeColumns()



            End Select

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "SetNewLine1 Error : " & ex.Message)

        End Try

    End Sub
End Class

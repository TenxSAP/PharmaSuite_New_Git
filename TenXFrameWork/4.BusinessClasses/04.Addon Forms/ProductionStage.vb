Imports Org.BouncyCastle.Security
Imports SAPbouiCOM

Public Class ProductionStage
#Region "Decliration"

    Public objForm As SAPbouiCOM.Form
    Public oDBs_Content As SAPbouiCOM.DBDataSource
    Public Prgbar As SAPbouiCOM.ProgressBar
    Dim objutilities As Utilities

    Public objMatrix, oMatrixAttach, oMatrixApprove, objMatrix3 As SAPbouiCOM.Matrix
    Dim oDBs_Head, oDBs_Approve, oDBs_Result As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim oDBs_Attach As SAPbouiCOM.DBDataSource
    Dim oDS As SAPbouiCOM.DBDataSource

#End Region
    Public Sub CreateForm()

        Try
            objMain.objUtilities.LoadForm("ProductionStageMaster.xml", "TNX_PSTG", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("TNX_PSTG",
                      objMain.objApplication.Forms.ActiveForm.TypeCount)



            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PSTG_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PSTG_L")

            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_PSTG", "Primary"))
            oDBs_Head.SetValue("U_StageSeq", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_PSTG", "Primary"))
            'oDBs_Head.SetValue("U_DA", 0, DateTime.Now.ToString("yyyyMMdd"))
            'objForm.DataBrowser.BrowseBy = "DocEntry"
            ' Production Stage Master - Auto Managed Attributes
            '========================================================================

            objForm.PaneLevel = 1

            '-----------------------------------------------------------------------
            ' System Controlled Fields
            '-----------------------------------------------------------------------

            'DocEntry
            objForm.Items.Item("0_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            '-----------------------------------------------------------------------
            ' Header Fields - Editable in Add/OK, Disabled in Find
            '-----------------------------------------------------------------------

            Dim HeaderItems As String() = {
    "23_U_E",
    "24_U_E",
    "25_U_E",
    "26_U_E",
    "27_U_E",
    "28_U_E",
    "29_U_E",
    "30_U_E",
    "31_U_E"
}

            For Each itemId As String In HeaderItems

                'Enable in all normal modes
                objForm.Items.Item(itemId).SetAutoManagedAttribute(
        SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
        -1,
        SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                'Disable in Find Mode
                objForm.Items.Item(itemId).SetAutoManagedAttribute(
        SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
        SAPbouiCOM.BoAutoFormMode.afm_Find,
        SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Next

            '-----------------------------------------------------------------------
            ' Matrix/Grid
            '-----------------------------------------------------------------------

            objForm.Items.Item("0_U_G").SetAutoManagedAttribute(
    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
    -1,
    SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("0_U_G").SetAutoManagedAttribute(
    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
    SAPbouiCOM.BoAutoFormMode.afm_Find,
    SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            '-----------------------------------------------------------------------
            ' Folder
            '-----------------------------------------------------------------------

            '        objForm.Items.Item("0_U_FD").SetAutoManagedAttribute(
            'SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            '-1,
            'SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            '-----------------------------------------------------------------------
            ' OK Button
            '-----------------------------------------------------------------------

            objForm.Items.Item("1").SetAutoManagedAttribute(
    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
    SAPbouiCOM.BoAutoFormMode.afm_Find,
    SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("0_U_E").SetAutoManagedAttribute(
    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
    SAPbouiCOM.BoAutoFormMode.afm_Find,
    SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("24_U_E").SetAutoManagedAttribute(
    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
    SAPbouiCOM.BoAutoFormMode.afm_Add,
    SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            '-----------------------------------------------------------------------
            ' Cancel Button
            '-----------------------------------------------------------------------

            objForm.Items.Item("2").SetAutoManagedAttribute(
    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
    SAPbouiCOM.BoAutoFormMode.afm_Find,
    SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            '========================================================================
            ' Optional:
            ' Lock fields after document is added (OK Mode)
            '========================================================================

            'Stage Sequence
            objForm.Items.Item("23_U_E").SetAutoManagedAttribute(
    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
    SAPbouiCOM.BoAutoFormMode.afm_Ok,
    SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            'Dosage Form
            objForm.Items.Item("24_U_E").SetAutoManagedAttribute(
    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
    SAPbouiCOM.BoAutoFormMode.afm_Ok,
    SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            '========================================================================
            ' Matrix Columns Editable Control
            '========================================================================

            objMatrix = objForm.Items.Item("0_U_G").Specific

            objMatrix.Columns.Item("C_0_1").Editable = True
            objMatrix.Columns.Item("C_0_2").Editable = True
            objMatrix.Columns.Item("C_0_3").Editable = True
            objMatrix.Columns.Item("C_0_4").Editable = True
            objMatrix.Columns.Item("C_0_5").Editable = True
            objMatrix.Columns.Item("C_0_6").Editable = True
            objMatrix.Columns.Item("C_0_7").Editable = True

            'Row Number Column
            objMatrix.Columns.Item("LineId").Editable = False

            Me.objForm.EnableMenu("1282", True)

            Me.objForm.EnableMenu("1281", True)
            Me.objForm.EnableMenu("519", True)
            Me.objForm.EnableMenu("520", True)
            'SetDefault(objForm.UniqueID)
            Me.objForm.EnableMenu("1292", True)
            ' Me.SetNewLine(objForm.UniqueID)
            Me.objForm.EnableMenu("1293", True)


            Me.SetNewLine(objForm.UniqueID)
            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
            "Productionstage Form Loaded Successfully",
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
    Public Sub SetNewLine(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PSTG_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PSTG_L")

            objMatrix = objForm.Items.Item("0_U_G").Specific


            objMatrix.AddRow()

            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount.ToString())

            oDBs_Details.SetValue("U_ParamCode", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_ParamName", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_MinValue", oDBs_Details.Offset, "0")
            oDBs_Details.SetValue("U_MaxValue", oDBs_Details.Offset, "0")
            oDBs_Details.SetValue("U_UOM", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_IsCrit", oDBs_Details.Offset, "N")
            oDBs_Details.SetValue("U_DevReq", oDBs_Details.Offset, "N")


            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
#Region "Menu Event"

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent,
              ByRef BubbleEvent As Boolean)

        Try

            '========================================================================
            ' OPEN UDO FORM
            '========================================================================
            If pVal.MenuUID = "10X_PMS_STAGE" _
            AndAlso pVal.BeforeAction = False Then

                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" _
      AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm
                Me.SetDefault(objForm.UniqueID)
                '========================================================================
                ' ADD ROW
                '========================================================================
            ElseIf pVal.MenuUID = "1292" _
            AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "TNX_PSTG" Then Exit Sub

                SetNewLine(objForm.UniqueID)

                '========================================================================
                ' DELETE ROW
                '========================================================================
            ElseIf pVal.MenuUID = "1293" _
            AndAlso pVal.BeforeAction = True Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "TNX_PSTG" Then Exit Sub

                    BubbleEvent = False

                    objForm.Freeze(True)

                    objMatrix = CType(objForm.Items.Item("0_U_G").Specific,
                                  SAPbouiCOM.Matrix)

                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PSTG_L")

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

                    '------------------------------------------------------------
                    ' DELETE MATRIX ROW
                    '------------------------------------------------------------
                    objMatrix.DeleteRow(selectedRow)

                    '------------------------------------------------------------
                    ' REMOVE EXTRA DATASOURCE ROWS
                    '------------------------------------------------------------
                    While oDBs_Details.Size > objMatrix.VisualRowCount

                        oDBs_Details.RemoveRecord(oDBs_Details.Size - 1)

                    End While

                    '------------------------------------------------------------
                    ' KEEP MINIMUM ONE ROW
                    '------------------------------------------------------------
                    If oDBs_Details.Size = 0 Then

                        oDBs_Details.InsertRecord(0)

                        oDBs_Details.SetValue("LineId", 0, "1")

                        oDBs_Details.SetValue("U_ParamCode", 0, "")
                        oDBs_Details.SetValue("U_ParamName", 0, "")
                        oDBs_Details.SetValue("U_MinValue", 0, "0")
                        oDBs_Details.SetValue("U_MaxValue", 0, "0")
                        oDBs_Details.SetValue("U_UOM", 0, "")
                        oDBs_Details.SetValue("U_IsCrit", 0, "N")
                        oDBs_Details.SetValue("U_DevReq", 0, "N")

                    End If

                    '------------------------------------------------------------
                    ' RE-SEQUENCE LINE NUMBERS
                    '------------------------------------------------------------
                    For i As Integer = 0 To oDBs_Details.Size - 1

                        oDBs_Details.SetValue("LineId",
                                       i,
                                       (i + 1).ToString())

                    Next

                    '------------------------------------------------------------
                    ' RELOAD MATRIX
                    '------------------------------------------------------------
                    Try

                        objMatrix.LoadFromDataSource()

                    Catch
                    End Try

                    objMatrix.AutoResizeColumns()

                    '------------------------------------------------------------
                    ' UPDATE MODE
                    '------------------------------------------------------------
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

                '========================================================================
                ' FIND MODE
                '========================================================================
            ElseIf pVal.MenuUID = "1281" _
            AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "TNX_PSTG" Then Exit Sub

                objForm.PaneLevel = 1

            End If

        Catch ex As Exception

            Try
                If objForm IsNot Nothing Then
                    objForm.Freeze(False)
                End If
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Public Sub SetDefault(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            '========================================================================
            ' DATASOURCES
            '========================================================================

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PSTG_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PSTG_L")

            objForm.Freeze(True)

            '========================================================================
            ' AUTO DOCUMENT NUMBER
            '========================================================================

            oDBs_Head.SetValue("DocNum",
                           0,
                           objMain.objUtilities.GetNextDocNum(objForm,
                                                              "TNX_PSTG",
                                                              "Primary"))
            oDBs_Head.SetValue("DocEntry",
                           0,
                           objMain.objUtilities.GetNextDocNum(objForm,
                                                              "TNX_PSTG",
                                                              "Primary"))

            '========================================================================
            ' DEFAULT VALUES
            '========================================================================

            oDBs_Head.SetValue("U_Status", 0, "Draft")

            ' oDBs_Head.SetValue("U_QAReq", 0, "N")

            oDBs_Head.SetValue("U_ApprReq", 0, "N")

            '========================================================================
            ' DEFAULT PANE
            '========================================================================

            objForm.PaneLevel = 1

            '========================================================================
            ' MATRIX DEFAULT ROW
            '========================================================================

            SetNewLine(FormUID)

            '========================================================================
            ' BUTTON SETTINGS
            '========================================================================

            objForm.Items.Item("1").Enabled = True
            objForm.Items.Item("2").Enabled = True

            objForm.Freeze(False)

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
    Public Sub ItemEvent(ByVal FormUID As String,
                 ByRef pVal As SAPbouiCOM.ItemEvent,
                 ByRef BubbleEvent As Boolean)
        Try

            If pVal.EventType = BoEventTypes.et_ITEM_PRESSED _
                    AndAlso pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.Item(FormUID)
                If pVal.ItemUID = "1" Then
                    Me.SetDefault(FormUID)
                End If
            End If
        Catch ex As Exception
            Try
                If objForm IsNot Nothing Then objForm.Freeze(False)
            Catch
            End Try
            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            BoMessageTime.bmt_Short,
            BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
End Class

Imports SAPbouiCOM

Public Class LineClerenceChecklist
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
            objMain.objUtilities.LoadForm("LineClerencechecklist.xml", "TNX_PLCC", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("TNX_PLCC",
                      objMain.objApplication.Forms.ActiveForm.TypeCount)



            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PLCC_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PLCC_H")

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_PLCC", "Primary"))
            '========================================================================
            ' DEFAULT VALUES
            '========================================================================

            objForm.PaneLevel = 1

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset,
                   objMain.objUtilities.GetNextDocNum(objForm, "TNX_PLCC", "Primary"))
            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset,
                   objMain.objUtilities.GetNextDocNum(objForm, "TNX_PLCC", "Primary"))

            oDBs_Head.SetValue("U_Status", 0, "Draft")
            oDBs_Head.SetValue("U_ApprReq", 0, "N")
            oDBs_Head.SetValue("U_QAReq", 0, "N")

            Me.SetNewLine(objForm.UniqueID)
            '========================================================================
            ' DOCENTRY / DOCNUM CONTROL
            '========================================================================

            objForm.Items.Item("0_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("0_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("1_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            '========================================================================
            ' HEADER FIELD CONTROL
            '========================================================================

            'Clearance Type
            objForm.Items.Item("23_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Add,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'Dosage Form
            objForm.Items.Item("24_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Add,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'Production Stage
            objForm.Items.Item("25_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Add,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'Production Area
            objForm.Items.Item("26_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Add,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'QA Required
            objForm.Items.Item("27_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Add,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'Approval Required
            objForm.Items.Item("28_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Add,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'Status - System Controlled
            objForm.Items.Item("29_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            '========================================================================
            ' FIND MODE CONTROLS
            '========================================================================

            objForm.Items.Item("24_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("25_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("26_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("27_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("28_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("29_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            '========================================================================
            ' MATRIX SETTINGS
            '========================================================================

            objMatrix = CType(objForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)

            objMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            '========================================================================
            ' MATRIX COLUMN SETTINGS
            '========================================================================

            objMatrix.Columns.Item("C_0_1").Editable = True   'Checklist Check Point
            objMatrix.Columns.Item("C_0_2").Editable = True   'Expected Result
            objMatrix.Columns.Item("C_0_3").Editable = True   'Result Type
            objMatrix.Columns.Item("C_0_4").Editable = True   'Mandatory
            objMatrix.Columns.Item("C_0_5").Editable = True   'Critical Checkpoint
            objMatrix.Columns.Item("C_0_6").Editable = True   'Deviation Required
            objMatrix.Columns.Item("C_0_7").Editable = True   'Remarks

            'objMatrix.Columns.Item("#").Editable = False

            '========================================================================
            ' BUTTON SETTINGS
            '========================================================================

            objForm.Items.Item("1").Enabled = True
            objForm.Items.Item("2").Enabled = True

            '========================================================================
            ' MENU SETTINGS
            '========================================================================

            Me.objForm.EnableMenu("1282", True)
            Me.objForm.EnableMenu("519", True)
            Me.objForm.EnableMenu("520", True)
            Me.objForm.EnableMenu("1292", True)
            Me.objForm.EnableMenu("1293", True)

            '========================================================================
            ' DEFAULT PANE
            '========================================================================

            objForm.PaneLevel = 1

            '========================================================================
            ' ADD DEFAULT ROW
            '========================================================================


            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
            "Line Clearance Checklist Form Loaded Successfully",
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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PLCC_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PLCC_L")

            objMatrix = CType(objForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)

            objMatrix.AddRow()

            oDBs_Details.SetValue("LineId",
                                  oDBs_Details.Offset,
                                  objMatrix.VisualRowCount.ToString())

            oDBs_Details.SetValue("U_ChkPoint",
                                  oDBs_Details.Offset,
                                  "")

            oDBs_Details.SetValue("U_ExpResult",
                                  oDBs_Details.Offset,
                                  "")

            oDBs_Details.SetValue("U_ResType",
                                  oDBs_Details.Offset,
                                  "")

            oDBs_Details.SetValue("U_IsMand",
                                  oDBs_Details.Offset,
                                  "Y")

            oDBs_Details.SetValue("U_IsCrit",
                                  oDBs_Details.Offset,
                                  "N")

            oDBs_Details.SetValue("U_DevReq",
                                  oDBs_Details.Offset,
                                  "N")

            oDBs_Details.SetValue("U_Remarks",
                                  oDBs_Details.Offset,
                                  "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)

            objMatrix.AutoResizeColumns()

        Catch ex As Exception

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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PLCC_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PLCC_L")

            objForm.Freeze(True)

            '========================================================================
            ' AUTO DOCUMENT NUMBER
            '========================================================================

            oDBs_Head.SetValue("DocNum",
                       0,
                       objMain.objUtilities.GetNextDocNum(objForm,
                                                          "TNX_PLCC",
                                                          "Primary"))
            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset,
                   objMain.objUtilities.GetNextDocNum(objForm, "TNX_PLCC", "Primary"))

            '========================================================================
            ' DEFAULT VALUES
            '========================================================================

            oDBs_Head.SetValue("U_Status", 0, "Draft")

            oDBs_Head.SetValue("U_QAReq", 0, "N")

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
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent,
          ByRef BubbleEvent As Boolean)

        Try

            '========================================================================
            ' OPEN UDO FORM
            '========================================================================
            If pVal.MenuUID = "10X_PMS_LCC" _
        AndAlso pVal.BeforeAction = False Then

                Me.CreateForm()

                '========================================================================
                ' ADD ROW
                '========================================================================
            ElseIf pVal.MenuUID = "1282" _
        AndAlso pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1292" _
        AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "TNX_PLCC" Then Exit Sub

                SetNewLine(objForm.UniqueID)

                '========================================================================
                ' DELETE ROW
                '========================================================================
            ElseIf pVal.MenuUID = "1293" _
        AndAlso pVal.BeforeAction = True Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "TNX_PLCC" Then Exit Sub

                    BubbleEvent = False

                    objForm.Freeze(True)

                    objMatrix = CType(objForm.Items.Item("0_U_G").Specific,
                              SAPbouiCOM.Matrix)

                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PLCC_L")

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

                        oDBs_Details.SetValue("U_ChkPoint", 0, "")
                        oDBs_Details.SetValue("U_ExpResult", 0, "")
                        oDBs_Details.SetValue("U_ResType", 0, "")
                        oDBs_Details.SetValue("U_IsMand", 0, "N")
                        oDBs_Details.SetValue("U_IsCrit", 0, "N")
                        oDBs_Details.SetValue("U_DevReq", 0, "N")
                        oDBs_Details.SetValue("U_Remarks", 0, "")

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

                If objForm.TypeEx <> "TNX_PLCC" Then Exit Sub

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

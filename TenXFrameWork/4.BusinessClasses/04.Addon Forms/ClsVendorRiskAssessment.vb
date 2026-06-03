Imports SAPbouiCOM

Public Class ClsVendorRiskAssessment

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
    Public Sub CreateForm()

        Try

            objMain.objUtilities.LoadForm("VendorRiskAssessment.xml", "frm_VRA", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("frm_VRA", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VRA_H")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_VRA_SCORE")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_VRA_REC")

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_VRA", "Primary"))
            oDBs_Head.SetValue("U_Status", oDBs_Head.Offset, "Open")
            oDBs_Head.SetValue("U_TotalScore", oDBs_Head.Offset, "0")
            oDBs_Head.SetValue("U_RiskLevel", oDBs_Head.Offset, "")

            '====================================================
            ' MATRIX SETTINGS
            '====================================================

            objMatrix1 = CType(objForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)
            objMatrix2 = CType(objForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix)

            objMatrix1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            objMatrix2.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            '====================================================
            ' DEFAULT TAB
            '====================================================

            objForm.PaneLevel = 1

            objForm.Items.Item("0_U_FD").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            '====================================================
            ' ENABLE MENUS
            '====================================================

            objForm.EnableMenu("1282", True)
            objForm.EnableMenu("1281", True)
            objForm.EnableMenu("1288", True)
            objForm.EnableMenu("1289", True)
            objForm.EnableMenu("1290", True)
            objForm.EnableMenu("1291", True)
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)

            '====================================================
            ' DEFAULT MATRIX ROWS
            '====================================================

            SetNewLine(objForm.UniqueID)
            SetNewLine1(objForm.UniqueID)

            '====================================================
            ' FIND MODE SETTINGS
            '====================================================

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("2").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
            "Vendor Risk Assessment Form Loaded Successfully",
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
    Public Sub SetNewLine(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VRA_H")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_VRA_SCORE")

            objMatrix1 = CType(objForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)

            objMatrix1.AddRow()

            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
            oDBs_Details1.SetValue("U_Factor", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_Weight", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_Score", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_WScore", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_Remarks", oDBs_Details1.Offset, "")

            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
            objMatrix1.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Public Sub SetNewLine1(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VRA_H")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_VRA_REC")

            objMatrix2 = CType(objForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix)

            objMatrix2.AddRow()

            oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)
            oDBs_Details2.SetValue("U_Recommendation", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_ActionType", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_TargetDate", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_Status", oDBs_Details2.Offset, "")

            objMatrix2.SetLineData(objMatrix2.VisualRowCount)
            objMatrix2.AutoResizeColumns()

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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VRA_H")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_VRA_SCORE")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_VRA_REC")

            objForm.Freeze(True)

            'Header Defaults
            oDBs_Head.SetValue("DocNum", 0, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_VRA", "Primary"))

            oDBs_Head.SetValue("U_Status", 0, "Open")
            oDBs_Head.SetValue("U_TotalScore", 0, "0")
            oDBs_Head.SetValue("U_RiskLevel", 0, "")

            'Default Tab
            objForm.PaneLevel = 1
            objForm.Items.Item("0_U_FD").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            'Matrix References
            objMatrix1 = CType(objForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)
            objMatrix2 = CType(objForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix)

            'Clear Existing Rows
            While oDBs_Details1.Size > 0
                oDBs_Details1.RemoveRecord(0)
            End While

            While oDBs_Details2.Size > 0
                oDBs_Details2.RemoveRecord(0)
            End While

            'Add First Rows
            SetNewLine(objForm.UniqueID)
            SetNewLine1(objForm.UniqueID)

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
    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

        Try

            If pVal.MenuUID = "10X_VEN_RISK" AndAlso pVal.BeforeAction = False Then

                Me.CreateForm()

            ElseIf pVal.MenuUID = "1282" AndAlso pVal.BeforeAction = False Then

                Me.CreateForm()

            ElseIf pVal.MenuUID = "1292" AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "frm_VRA" Then Exit Sub

                If objForm.PaneLevel = 1 Then

                    SetNewLine(objForm.UniqueID)

                ElseIf objForm.PaneLevel = 2 Then

                    SetNewLine1(objForm.UniqueID)

                End If

            ElseIf pVal.MenuUID = "1293" AndAlso pVal.BeforeAction = True Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "frm_VRA" Then Exit Sub

                    BubbleEvent = False

                    objForm.Freeze(True)

                    '=========================================================
                    ' DELETE ROW - SCORE MATRIX
                    '=========================================================
                    If objForm.PaneLevel = 1 Then

                        objMatrix1 = CType(objForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)
                        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_VRA_SCORE")

                        Dim selectedRow As Integer = objMatrix1.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder)

                        If selectedRow <= 0 Then

                            objMain.objApplication.StatusBar.SetText(
                        "Please select row to delete",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            Exit Try

                        End If

                        objMatrix1.DeleteRow(selectedRow)

                        objMatrix1.FlushToDataSource()

                        While oDBs_Details1.Size > objMatrix1.VisualRowCount

                            oDBs_Details1.RemoveRecord(oDBs_Details1.Size - 1)

                        End While

                        If oDBs_Details1.Size = 0 Then

                            oDBs_Details1.InsertRecord(0)

                            oDBs_Details1.SetValue("LineId", 0, "1")
                            oDBs_Details1.SetValue("U_Factor", 0, "")
                            oDBs_Details1.SetValue("U_Weight", 0, "")
                            oDBs_Details1.SetValue("U_Score", 0, "")
                            oDBs_Details1.SetValue("U_WScore", 0, "")
                            oDBs_Details1.SetValue("U_Remarks", 0, "")

                        End If

                        For i As Integer = 0 To oDBs_Details1.Size - 1

                            oDBs_Details1.SetValue("LineId", i, (i + 1).ToString())

                        Next

                        objMatrix1.LoadFromDataSource()
                        objMatrix1.AutoResizeColumns()

                    End If

                    '=========================================================
                    ' DELETE ROW - RECOMMENDATION MATRIX
                    '=========================================================
                    If objForm.PaneLevel = 2 Then

                        objMatrix2 = CType(objForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix)
                        oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_VRA_REC")

                        Dim selectedRow As Integer = objMatrix2.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder)

                        If selectedRow <= 0 Then

                            objMain.objApplication.StatusBar.SetText(
                        "Please select row to delete",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            Exit Try

                        End If

                        objMatrix2.DeleteRow(selectedRow)

                        objMatrix2.FlushToDataSource()

                        While oDBs_Details2.Size > objMatrix2.VisualRowCount

                            oDBs_Details2.RemoveRecord(oDBs_Details2.Size - 1)

                        End While

                        If oDBs_Details2.Size = 0 Then

                            oDBs_Details2.InsertRecord(0)

                            oDBs_Details2.SetValue("LineId", 0, "1")
                            oDBs_Details2.SetValue("U_Recommendation", 0, "")
                            oDBs_Details2.SetValue("U_ActionType", 0, "")
                            oDBs_Details2.SetValue("U_TargetDate", 0, "")
                            oDBs_Details2.SetValue("U_Status", 0, "")

                        End If

                        For i As Integer = 0 To oDBs_Details2.Size - 1

                            oDBs_Details2.SetValue("LineId", i, (i + 1).ToString())

                        Next

                        objMatrix2.LoadFromDataSource()
                        objMatrix2.AutoResizeColumns()

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

            End If

        Catch ex As Exception

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


    'Sub SetDefault(ByVal FormUID As String)

    '    Try

    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_VRA_H")
    '        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_VRA_SCORE")
    '        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_VRA_REC")


    '        objMatrix1 = objForm.Items.Item("0_U_G").Specific
    '        objMatrix2 = objForm.Items.Item("1_U_G").Specific

    '        oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_VRA"))
    '        'oDBs_Head.SetValue("U_CreatedOn", 0, DateTime.Now.ToString("yyyyMMdd"))

    '        objForm.Items.Item("0_U_FD").Click(BoCellClickType.ct_Regular)

    '        Me.SetNewLine(FormUID)
    '        objForm.Freeze(False)

    '    Catch ex As Exception
    '        objForm.Freeze(False)
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try

    'End Sub
    'Sub SetNewLine(ByVal FormUID As String)

    '    Try

    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_VRA_H")
    '        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_VRA_SCORE")
    '        oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_VRA_REC")

    '        objMatrix1 = objForm.Items.Item("0_U_G").Specific
    '        objMatrix2 = objForm.Items.Item("1_U_G").Specific


    '        If objMatrix1.VisualRowCount = 0 Then

    '            objMatrix1.AddRow()

    '            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)

    '            oDBs_Details1.SetValue("U_Factor", oDBs_Details1.Offset, "")
    '            oDBs_Details1.SetValue("U_Weight", oDBs_Details1.Offset, "")
    '            oDBs_Details1.SetValue("U_Score", oDBs_Details1.Offset, "")
    '            oDBs_Details1.SetValue("U_WScore", oDBs_Details1.Offset, "")
    '            oDBs_Details1.SetValue("U_Remarks", oDBs_Details1.Offset, "")

    '            objMatrix1.SetLineData(objMatrix1.VisualRowCount)

    '            objMatrix1.AutoResizeColumns()

    '        End If


    '        If objMatrix2.VisualRowCount = 0 Then

    '            objMatrix2.AddRow()

    '            oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)
    '            oDBs_Details2.SetValue("U_Recommendation", oDBs_Details2.Offset, "")
    '            oDBs_Details2.SetValue("U_ActionType", oDBs_Details2.Offset, "")
    '            oDBs_Details2.SetValue("U_TargetDate", oDBs_Details2.Offset, "")
    '            oDBs_Details2.SetValue("U_Status", oDBs_Details2.Offset, "")

    '            objMatrix2.SetLineData(objMatrix2.VisualRowCount)

    '            objMatrix2.AutoResizeColumns()

    '        End If


    '    Catch ex As Exception

    '        objMain.objApplication.StatusBar.SetText(ex.Message)

    '    End Try

    'End Sub
    'Private Sub SetNewLine1(ByVal FormUID As String,
    '                    ByVal MatrixUID As String)

    '    Try

    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        Select Case MatrixUID

    '    '====================================================
    '    ' MATRIX 1
    '    '====================================================

    '            Case "0_U_G"

    '                oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_VRA_SCORE")

    '                objMatrix1 = objForm.Items.Item("0_U_G").Specific

    '                objMatrix1.AddRow()

    '                oDBs_Details1.SetValue("LineId",
    '                              oDBs_Details1.Offset,
    '                              objMatrix1.VisualRowCount.ToString())

    '                oDBs_Details1.SetValue("U_Factor", oDBs_Details1.Offset, "")
    '                oDBs_Details1.SetValue("U_Weight", oDBs_Details1.Offset, "")
    '                oDBs_Details1.SetValue("U_Score", oDBs_Details1.Offset, "")
    '                oDBs_Details1.SetValue("U_WScore", oDBs_Details1.Offset, "")
    '                oDBs_Details1.SetValue("U_Remarks", oDBs_Details1.Offset, "")

    '                objMatrix1.SetLineData(objMatrix1.VisualRowCount)

    '                objMatrix1.AutoResizeColumns()


    '    '====================================================
    '    ' MATRIX 2
    '    '====================================================

    '            Case "1_U_G"

    '                oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_VRA_REC")

    '                objMatrix2 = objForm.Items.Item("1_U_G").Specific

    '                objMatrix2.AddRow()

    '                oDBs_Details2.SetValue("LineId",
    '                               oDBs_Details2.Offset,
    '                               objMatrix2.VisualRowCount.ToString())

    '                oDBs_Details2.SetValue("U_Recommendation", oDBs_Details2.Offset, "")
    '                oDBs_Details2.SetValue("U_ActionType", oDBs_Details2.Offset, "")
    '                oDBs_Details2.SetValue("U_TargetDate", oDBs_Details2.Offset, "")
    '                oDBs_Details2.SetValue("U_Status", oDBs_Details2.Offset, "")

    '                objMatrix2.SetLineData(objMatrix2.VisualRowCount)

    '                objMatrix2.AutoResizeColumns()

    '        End Select

    '    Catch ex As Exception

    '        objMain.objApplication.StatusBar.SetText(
    '"SetNewLine1 Error : " & ex.Message)

    '    End Try

    'End Sub
End Class

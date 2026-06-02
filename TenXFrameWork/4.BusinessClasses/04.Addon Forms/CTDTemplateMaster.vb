Public Class CTDTemplateMaster
    Public objForm As SAPbouiCOM.Form
    Public objMatrix As SAPbouiCOM.Matrix

    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource

    Public Sub CreateForm()

        Try

            objMain.objUtilities.LoadForm("CTDTemplateMaster.xml",
                                      "REG_CTDTMP",
                                      ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm(
                    "REG_CTDTMP",
                    objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            '=================================================
            ' DATASOURCES
            '=================================================
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_CTDTMP")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_REG_CTDL")

            'oDBs_Head.SetValue("DocEntry",
            '               oDBs_Head.Offset,
            '               objMain.objUtilities.GetNextDocNum(
            '               objForm,
            '               "REG_CTDTMP",
            '               "Primary"))
            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset,
                   objMain.objUtilities.GetNextDocNum(objForm, "UDO_REG_CTDTMP", "Primary"))

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset,
                   objMain.objUtilities.GetNextDocNum(objForm, "UDO_REG_CTDTMP", "Primary"))

            'oDBs_Head.SetValue("DocNum",
            '               oDBs_Head.Offset,
            '               objMain.objUtilities.GetNextDocNum(
            '               objForm,
            '               "REG_CTDTMP",
            '               "Primary"))

            oDBs_Head.SetValue("U_Status", 0, "Active")

            oDBs_Head.SetValue("U_EffDate",
                           0,
                           DateTime.Now.ToString("yyyyMMdd"))

            If Trim(oDBs_Head.GetValue("U_VersionNo", 0)) = "" Then
                oDBs_Head.SetValue("U_VersionNo", 0, "1.0")
            End If

            objForm.PaneLevel = 1

            '=================================================
            ' DOCENTRY
            '=================================================
            objForm.Items.Item("DocEntry").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("DocEntry").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            '=================================================
            ' DOCNUM
            '=================================================
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("DocNum").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            '=================================================
            ' STATUS
            '=================================================
            objForm.Items.Item("29_U_E").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            -1,
            SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            '=================================================
            ' MATRIX
            '=================================================
            objMatrix = CType(
            objForm.Items.Item("0_U_G").Specific,
            SAPbouiCOM.Matrix)

            objMatrix.SelectionMode =
            SAPbouiCOM.BoMatrixSelect.ms_Single

            objMatrix.Columns.Item("LineId").Editable = False

            objMatrix.Columns.Item("C_0_1").Editable = True
            objMatrix.Columns.Item("C_0_2").Editable = True
            objMatrix.Columns.Item("C_0_3").Editable = True
            objMatrix.Columns.Item("C_0_4").Editable = True
            objMatrix.Columns.Item("C_0_5").Editable = True
            objMatrix.Columns.Item("C_0_6").Editable = True
            objMatrix.Columns.Item("C_0_7").Editable = True
            objMatrix.Columns.Item("C_0_8").Editable = True
            objMatrix.Columns.Item("C_0_9").Editable = True

            '=================================================
            ' BUTTONS
            '=================================================
            objForm.Items.Item("1").Enabled = True
            objForm.Items.Item("2").Enabled = True

            '=================================================
            ' MENU SETTINGS
            '=================================================
            objForm.EnableMenu("1281", True)
            objForm.EnableMenu("1282", True)
            objForm.EnableMenu("1288", True)
            objForm.EnableMenu("1289", True)
            objForm.EnableMenu("1290", True)
            objForm.EnableMenu("1291", True)
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)

            '=================================================
            ' DEFAULT PANE
            '=================================================
            objForm.PaneLevel = 1

            '=================================================
            ' FIRST MATRIX ROW
            '=================================================
            Me.SetNewLine(objForm.UniqueID)

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
            "CTD eCTD Template Master Loaded Successfully",
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
    Private Function ItemExists(ByVal frm As SAPbouiCOM.Form, ByVal itemId As String) As Boolean
        Try
            If frm Is Nothing Then Return False
            Dim dummy = frm.Items.Item(itemId)
            Return True
        Catch
            Return False
        End Try
    End Function
    Public Sub SetDefault(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_CTDTMP")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_REG_CTDL")

            objForm.Freeze(True)


            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset,
                   objMain.objUtilities.GetNextDocNum(objForm, "UDO_REG_CTDTMP", "Primary"))
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset,
                   objMain.objUtilities.GetNextDocNum(objForm, "UDO_REG_CTDTMP", "Primary"))
            If oDBs_Head.Size = 0 Then oDBs_Head.InsertRecord(0)
            oDBs_Head.SetValue("U_Status", 0, "Active")

            oDBs_Head.SetValue("U_EffDate", 0,
                               DateTime.Now.ToString("yyyyMMdd"))

            If String.IsNullOrEmpty(oDBs_Head.GetValue("U_VersionNo", 0).Trim()) Then
                oDBs_Head.SetValue("U_VersionNo", 0, "1.0")
            End If

            objForm.PaneLevel = 1


            SetNewLine(objForm.UniqueID)

            If ItemExists(objForm, "DocEntry") Then
                objForm.Items.Item("DocEntry").Enabled = False
            End If

            If ItemExists(objForm, "1_U_E") Then
                objForm.Items.Item("1_U_E").Enabled = False
            End If

            If ItemExists(objForm, "29_U_E") Then
                objForm.Items.Item("29_U_E").Enabled = False
            End If

            '=========================================
            ' OPEN FOLDER
            '=========================================
            If ItemExists(objForm, "0_U_FD") Then
                objForm.Items.Item("0_U_FD").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            End If

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
    'Sub SetNewLine(ByVal FormUID As String)

    '    Try

    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_REG_CTDL")

    '        objMatrix = objForm.Items.Item("0_U_G").Specific

    '        objMatrix.AddRow()

    '        oDBs_Details.SetValue("LineId",
    '                                  oDBs_Details.Offset,
    '                                  objMatrix.VisualRowCount.ToString())

    '        oDBs_Details.SetValue("U_ModuleNo", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_SecCode", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_SecName", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_DocTypCod", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_Mandatory", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_SeqNo", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_AllowMulti", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_AttachReq", oDBs_Details.Offset, "")
    '        oDBs_Details.SetValue("U_ApprReq", oDBs_Details.Offset, "")
    '        objMatrix.SetLineData(objMatrix.VisualRowCount)

    '        objMatrix.AutoResizeColumns()

    '    Catch ex As Exception

    '        objMain.objApplication.StatusBar.SetText(
    '            ex.Message,
    '            SAPbouiCOM.BoMessageTime.bmt_Short,
    '            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '    End Try

    'End Sub
    Sub SetNewLine(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_REG_CTDL")

            objMatrix = objForm.Items.Item("0_U_G").Specific

            If objMatrix.VisualRowCount > 0 Then
                objMatrix.FlushToDataSource()
            End If

            Dim rowIndex As Integer

            If oDBs_Details.Size = 1 _
        AndAlso oDBs_Details.GetValue("LineId", 0).Trim = "" Then

                rowIndex = 0

            Else

                rowIndex = oDBs_Details.Size
                oDBs_Details.InsertRecord(rowIndex)

            End If

            oDBs_Details.Offset = rowIndex

            oDBs_Details.SetValue("LineId", rowIndex, (rowIndex + 1).ToString())

            oDBs_Details.SetValue("U_ModuleNo", rowIndex, "")
            oDBs_Details.SetValue("U_SecCode", rowIndex, "")
            oDBs_Details.SetValue("U_SecName", rowIndex, "")
            oDBs_Details.SetValue("U_DocTypCod", rowIndex, "")
            oDBs_Details.SetValue("U_Mandatory", rowIndex, "N")
            oDBs_Details.SetValue("U_SeqNo", rowIndex, "")
            oDBs_Details.SetValue("U_AllowMulti", rowIndex, "N")
            oDBs_Details.SetValue("U_AttachReq", rowIndex, "N")
            oDBs_Details.SetValue("U_ApprReq", rowIndex, "N")

            objMatrix.LoadFromDataSource()

            For i As Integer = 1 To objMatrix.VisualRowCount

                objMatrix.Columns.Item("LineId") _
                     .Cells.Item(i).Specific.Value = i.ToString()

            Next

            objMatrix.AutoResizeColumns()

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

            If objForm Is Nothing _
            OrElse objForm.TypeEx <> "REG_CTDTMP" Then
                Exit Sub
            End If

            Select Case pVal.EventType

            '=========================================
            ' BUTTON EVENTS
            '=========================================
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "1" Then

                        If pVal.BeforeAction = True Then

                            objMatrix = CType(objForm.Items.Item("0_U_G").Specific,
                              SAPbouiCOM.Matrix)

                            objMatrix.FlushToDataSource()

                        Else

                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                SetDefault(FormUID)
                            End If

                        End If

                    End If
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                    If pVal.BeforeAction = True Then

                        objMatrix = CType(objForm.Items.Item("0_U_G").Specific,
                                          SAPbouiCOM.Matrix)

                        objMatrix.FlushToDataSource()

                    Else

                        objMain.objApplication.StatusBar.SetText(
                            "Document Added Successfully",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    End If
            '=========================================
            ' MATRIX CLICK
            '=========================================
                Case SAPbouiCOM.BoEventTypes.et_CLICK

                    If pVal.BeforeAction = False Then

                        If pVal.ItemUID = "0_U_G" Then

                            objMatrix = CType(objForm.Items.Item("0_U_G").Specific,
                                              SAPbouiCOM.Matrix)

                        End If

                    End If

            '=========================================
            ' MATRIX VALIDATION
            '=========================================
                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                    If pVal.BeforeAction = False Then

                        If pVal.ItemUID = "0_U_G" Then

                            If pVal.Row > 0 Then

                                objMatrix = CType(objForm.Items.Item("0_U_G").Specific,
                                                  SAPbouiCOM.Matrix)

                                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If

                            End If

                        End If

                    End If

            '=========================================
            ' CFL EVENTS (ADD LATER IF REQUIRED)
            '=========================================
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                    If pVal.BeforeAction = False Then

                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent =
                            CType(pVal, SAPbouiCOM.IChooseFromListEvent)

                        Dim oDataTable As SAPbouiCOM.DataTable =
                            oCFLEvento.SelectedObjects

                        If oDataTable Is Nothing Then Exit Select

                        ' Add CFL handling here if Country,
                        ' Authority Code, Section Code etc.
                        ' are linked to CFLs.

                    End If

            End Select

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
                "REG_CTDTMP Item Event Error : " & ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent,
                     ByRef BubbleEvent As Boolean)

        Try

            '=========================================
            ' OPEN FORM
            '=========================================
            If pVal.MenuUID = "10X_REG_CTDTMP" _
            AndAlso pVal.BeforeAction = False Then

                Me.CreateForm()

                '=========================================
                ' ADD MODE
                '=========================================
            ElseIf pVal.MenuUID = "1282" _
            AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "REG_CTDTMP" Then Exit Sub

                Me.SetDefault(objForm.UniqueID)

                '=========================================
                ' FIND MODE
                '=========================================
            ElseIf pVal.MenuUID = "1281" _
            AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "REG_CTDTMP" Then Exit Sub

                '=========================================
                ' ADD ROW
                '=========================================
            ElseIf pVal.MenuUID = "1292" _
            AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "REG_CTDTMP" Then Exit Sub

                SetNewLine(objForm.UniqueID)

                '=========================================
                ' DELETE ROW
                '=========================================
            ElseIf pVal.MenuUID = "1293" _
            AndAlso pVal.BeforeAction = True Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "REG_CTDTMP" Then Exit Sub

                    BubbleEvent = False

                    objForm.Freeze(True)

                    objMatrix = CType(
                        objForm.Items.Item("0_U_G").Specific,
                        SAPbouiCOM.Matrix)

                    oDBs_Details =
                        objForm.DataSources.DBDataSources.Item("@TNX_REG_CTDL")

                    Dim selectedRow As Integer =
                        objMatrix.GetNextSelectedRow(
                            0,
                            SAPbouiCOM.BoOrderType.ot_RowOrder)

                    If selectedRow <= 0 Then

                        objMain.objApplication.StatusBar.SetText(
                            "Please select a row.",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        Exit Try

                    End If

                    objMatrix.FlushToDataSource()

                    objMatrix.DeleteRow(selectedRow)

                    While oDBs_Details.Size > objMatrix.VisualRowCount

                        oDBs_Details.RemoveRecord(
                            oDBs_Details.Size - 1)

                    End While

                    If oDBs_Details.Size = 0 Then

                        oDBs_Details.InsertRecord(0)

                        oDBs_Details.SetValue("LineId", 0, "1")
                        oDBs_Details.SetValue("U_ModuleNo", 0, "")
                        oDBs_Details.SetValue("U_SecCode", 0, "")
                        oDBs_Details.SetValue("U_SecName", 0, "")
                        oDBs_Details.SetValue("U_DocTypCod", 0, "")
                        oDBs_Details.SetValue("U_Mandatory", 0, "")
                        oDBs_Details.SetValue("U_SeqNo", 0, "")
                        oDBs_Details.SetValue("U_AllowMulti", 0, "")
                        oDBs_Details.SetValue("U_AttachReq", 0, "")
                        oDBs_Details.SetValue("U_ApprReq", 0, "")

                    End If

                    For i As Integer = 0 To oDBs_Details.Size - 1

                        oDBs_Details.SetValue(
                            "LineId",
                            i,
                            (i + 1).ToString())

                    Next

                    objMatrix.LoadFromDataSource()
                    objMatrix.AutoResizeColumns()

                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                        objForm.Mode =
                            SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

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
End Class

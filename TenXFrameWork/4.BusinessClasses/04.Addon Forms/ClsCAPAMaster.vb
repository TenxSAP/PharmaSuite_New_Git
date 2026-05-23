Public Class ClsCAPAMaster

#Region "Declaration"
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
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
            objMain.objUtilities.LoadForm("CAPACategoryMaster.xml", "10X_CAPACAT", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("10X_CAPACAT", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            'CheckCorporateTaxConfigSetup()
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_QC_CAPACAT")
            objMatrix = objForm.Items.Item("MXT_1").Specific
            'oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "TNX_GLAUDO"))

            Me.MatrixLoad()

            objForm.Freeze(False)
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)
            '  Me.SetNewLine(objForm.UniqueID)
            ' Me.GlAccountCode()
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

        Try

            If pVal.MenuUID = "10X_QC_CAPACAT" AndAlso pVal.BeforeAction = False Then

                Me.CreateForm()

            ElseIf pVal.MenuUID = "1282" AndAlso pVal.BeforeAction = False Then

                'Add mode logic if required

            ElseIf pVal.MenuUID = "1292" AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "10X_CAPACAT" Then Exit Sub

                objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)
                Me.SetNewLine(objForm.UniqueID)


            ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then
                objMatrix = objForm.Items.Item("MXT_1").Specific
                Dim row As Integer = objMatrix.VisualRowCount
                If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then
                    objMatrix.AddRow()
                    oDBs_Details.SetValue("DocEntry", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    oDBs_Details.SetValue("Code", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                End If
                If objMatrix.IsRowSelected(row) = True Then
                    objMatrix.DeleteRow(row)
                Else
                    For i As Integer = 1 To objMatrix.VisualRowCount - 1

                        If objMatrix.IsRowSelected(i) = True Then
                            objMatrix.DeleteRow(i)
                        End If
                    Next
                End If
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("DocEntry").Cells.Item(i).Specific.Value = i
                    objMatrix.Columns.Item("Code").Cells.Item(i).Specific.Value = i
                Next

            End If

        Catch ex As Exception

            Try
                objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(
            "MenuEvent Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub


    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_QC_CAPACAT")
                    ' oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PALLET_C0")
                    objMatrix = objForm.Items.Item("MXT_1").Specific
                    '    Dim oCFL As SAPbouiCOM.ChooseFromList
                    '    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    '    Dim CFL_Id As String
                    '    CFL_Id = CFLEvent.ChooseFromListUID
                    '    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    '    Dim oDT As SAPbouiCOM.DataTable
                    '    oDT = CFLEvent.SelectedObjects

                    '    If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then
                    '        CFLFilterGLAccounts(FormUID, "CFL_1")
                    '    ElseIf oCFL.UniqueID = "CFL_0" And pVal.BeforeAction = True Then
                    '        CFLFilterGLAccounts(FormUID, "CFL_0")

                    '    End If

                    '    If Not (oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                    '        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    '        If oCFL.UniqueID = "CFL_1" Then
                    '            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value)
                    '            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.Columns.Item("Code").Cells.Item(pVal.Row).Specific.Value)
                    '            oDBs_Head.SetValue("U_MnProfit", oDBs_Head.Offset, objMatrix.Columns.Item("MnProfit").Cells.Item(pVal.Row).Specific.Value)
                    '            oDBs_Head.SetValue("U_MxProfit", oDBs_Head.Offset, objMatrix.Columns.Item("MxProfit").Cells.Item(pVal.Row).Specific.Value)
                    '            oDBs_Head.SetValue("U_TaxPrc", oDBs_Head.Offset, objMatrix.Columns.Item("TaxPrc").Cells.Item(pVal.Row).Specific.Value)
                    '            oDBs_Head.SetValue("U_EAccount", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                    '            oDBs_Head.SetValue("U_LAccount", oDBs_Head.Offset, objMatrix.Columns.Item("LAccount").Cells.Item(pVal.Row).Specific.Value)
                    '            oDBs_Head.SetValue("U_FINA", oDBs_Head.Offset, objMatrix.Columns.Item("FINA").Cells.Item(pVal.Row).Specific.Value)
                    '            'oDBs_Head.SetValue("U_TNXSTUS", oDBs_Head.Offset, objMatrix.Columns.Item("TNXSTUS").Cells.Item(pVal.Row).Specific.Value)

                    '            objMatrix.SetLineData(pVal.Row)
                    '        End If
                    '        If oCFL.UniqueID = "CFL_0" Then
                    '            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.Value)
                    '            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.Columns.Item("Code").Cells.Item(pVal.Row).Specific.Value)
                    '            oDBs_Head.SetValue("U_MnProfit", oDBs_Head.Offset, objMatrix.Columns.Item("MnProfit").Cells.Item(pVal.Row).Specific.Value)
                    '            oDBs_Head.SetValue("U_MxProfit", oDBs_Head.Offset, objMatrix.Columns.Item("MxProfit").Cells.Item(pVal.Row).Specific.Value)
                    '            oDBs_Head.SetValue("U_TaxPrc", oDBs_Head.Offset, objMatrix.Columns.Item("TaxPrc").Cells.Item(pVal.Row).Specific.Value)
                    '            oDBs_Head.SetValue("U_EAccount", oDBs_Head.Offset, objMatrix.Columns.Item("EAccount").Cells.Item(pVal.Row).Specific.Value)
                    '            oDBs_Head.SetValue("U_LAccount", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                    '            oDBs_Head.SetValue("U_FINA", oDBs_Head.Offset, objMatrix.Columns.Item("FINA").Cells.Item(pVal.Row).Specific.Value)
                    '            objMatrix.SetLineData(pVal.Row)
                    '            ' CFLFilterGLAccount(FormUID, "CFL_0")
                    '        End If
                    '    End If

                    'Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                    '    If pVal.BeforeAction = False Then

                    '        objForm = objMain.objApplication.Forms.Item(FormUID)
                    '        objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)

                    '        If pVal.ItemUID = "MXT_1" AndAlso
                    '           (pVal.ColUID = "LAccount" OrElse pVal.ColUID = "EAccount") Then

                    '            Dim AcctCode As String =
                    '                CType(objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific,
                    '                SAPbouiCOM.EditText).Value.Trim()

                    '            If AcctCode <> "" Then
                    '                objMain.objApplication.OpenForm(
                    '                    SAPbouiCOM.BoFormObjectEnum.fo_GLAccounts,
                    '                    "",
                    '                    AcctCode)
                    '            End If

                    '        End If

                    '    End If
                    'Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    '    If pVal.BeforeAction = True Then

                    '        If pVal.ItemUID = "1" Then   'OK / Update button

                    '            objForm = objMain.objApplication.Forms.Item(FormUID)
                    '            objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)

                    '            For i As Integer = 1 To objMatrix.VisualRowCount
                    '                If ValidateProfitRange(FormUID, i) = False Then
                    '                    BubbleEvent = False
                    '                    Exit Sub
                    '                End If
                    '            Next

                    '            If ValidateGLAccounts(FormUID) = False Then
                    '                BubbleEvent = False
                    '                Exit Sub
                    '            End If

                    '        End If

                    '    End If
                    '    'Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                    '    '    If pVal.BeforeAction = False Then

                    '    '        If pVal.ItemUID = "MXT_1" AndAlso
                    '    '           (pVal.ColUID = "MnProfit" OrElse pVal.ColUID = "MxProfit") Then

                    '    '            If ValidateProfitRange(FormUID, pVal.Row) = False Then
                    '    '                BubbleEvent = False
                    '    '            End If

                    '    '        End If

                    '    '    End If
                    'Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                    '    If pVal.BeforeAction = False Then

                    '        If pVal.ItemUID = "MXT_1" AndAlso
                    '           (pVal.ColUID = "MnProfit" OrElse
                    '            pVal.ColUID = "MxProfit") Then

                    '            objForm = objMain.objApplication.Forms.Item(FormUID)
                    '            objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)

                    '            Dim MinValue As String =
                    '                CType(objMatrix.Columns.Item("MnProfit").Cells.Item(pVal.Row).Specific,
                    '                SAPbouiCOM.EditText).Value.Trim()

                    '            Dim MaxValue As String =
                    '                CType(objMatrix.Columns.Item("MxProfit").Cells.Item(pVal.Row).Specific,
                    '                SAPbouiCOM.EditText).Value.Trim()

                    '            'Run validation only when both values entered
                    '            If MinValue <> "" AndAlso MaxValue <> "" Then

                    '                If ValidateProfitRange(FormUID, pVal.Row) = False Then
                    '                    BubbleEvent = False
                    '                    Exit Sub
                    '                End If

                    '            End If

                    '        End If

                    '   End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function ToDecimalValue(ByVal value As String) As Decimal
        If value Is Nothing Then Return 0D

        value = value.Trim().Replace(",", "")

        Dim result As Decimal = 0D
        Decimal.TryParse(value, result)

        Return result
    End Function

    Private Function ValidateProfitRange(ByVal FormUID As String, ByVal CurrentRow As Integer) As Boolean

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = CType(objForm.Items.Item("MXT_1").Specific, SAPbouiCOM.Matrix)

            If CurrentRow <= 0 Then Return True

            Dim CurMin As Decimal =
            ToDecimalValue(CType(objMatrix.Columns.Item("MnProfit").Cells.Item(CurrentRow).Specific, SAPbouiCOM.EditText).Value)

            Dim CurMax As Decimal =
            ToDecimalValue(CType(objMatrix.Columns.Item("MxProfit").Cells.Item(CurrentRow).Specific, SAPbouiCOM.EditText).Value)

            If CurMin = 0 OrElse CurMax = 0 Then Return True

            If CurMin >= CurMax Then
                objMain.objApplication.StatusBar.SetText(
                "Minimum Profit should be less than Maximum Profit.",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False
            End If

            For i As Integer = 1 To objMatrix.VisualRowCount

                If i = CurrentRow Then Continue For

                Dim MinVal As Decimal =
                ToDecimalValue(CType(objMatrix.Columns.Item("MnProfit").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value)

                Dim MaxVal As Decimal =
                ToDecimalValue(CType(objMatrix.Columns.Item("MxProfit").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value)

                If MinVal = 0 OrElse MaxVal = 0 Then Continue For

                If CurMin <= MaxVal AndAlso CurMax >= MinVal Then

                    objMain.objApplication.StatusBar.SetText(
                    "Profit range overlap found in row " & i & ". Please enter different Minimum and Maximum Profit.",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    Return False
                End If

            Next

            Return True

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(
            "Validate Profit Range Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            Return False
        End Try

    End Function

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_QC_CAPACAT")
            objMatrix = objForm.Items.Item("MXT_1").Specific
            objMatrix.AddRow()
            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.VisualRowCount)
            oDBs_Head.SetValue("Name", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_ActionTyp", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_DefOwn", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_EffCheck", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_TarDays", oDBs_Head.Offset, "")
            'oDBs_Head.SetValue("U_Active", oDBs_Head.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub CFLFilterGLAccounts(ByVal FormUID As String, ByVal CFL_ID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition

            oChooseFromList = objForm.ChooseFromLists.Item(CFL_ID)

            'Clear existing conditions
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList.SetConditions(emptyCon)

            'Get conditions object
            oConditions = oChooseFromList.GetConditions()

            'Filter only Postable Accounts
            oCondition = oConditions.Add()
            oCondition.Alias = "Postable"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"

            oChooseFromList.SetConditions(oConditions)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message,
                                                 SAPbouiCOM.BoMessageTime.bmt_Short,
                                                 SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub




    Public Sub MatrixLoad()
        Try

            oDBs_Head = objForm.DataSources.DBDataSources.Add("@TNX_QC_CAPACAT")
            Dim objsectMat As SAPbouiCOM.Matrix = objForm.Items.Item("MXT_1").Specific

            Dim rs1 As String = "SELECT ""DocEntry"",""Code"", ""Name"", ""U_ActionTyp"", ""U_DefOwn"", ""U_EffCheck"", ""U_TarDays"" " &
"FROM ""@TNX_QC_CAPACAT"" ORDER BY ""Code"" "
            Dim ors1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ors1.DoQuery(rs1)
            If ors1.RecordCount > 0 Then
                objsectMat.Clear()
                For j As Integer = 1 To ors1.RecordCount
                    objsectMat.AddRow()
                    oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
                    oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMatrix.VisualRowCount)
                    oDBs_Head.SetValue("Name", oDBs_Head.Offset, ors1.Fields.Item("Name").Value)
                    oDBs_Head.SetValue("U_ActionTyp", oDBs_Head.Offset, ors1.Fields.Item("U_ActionTyp").Value)
                    oDBs_Head.SetValue("U_DefOwn", oDBs_Head.Offset, ors1.Fields.Item("U_DefOwn").Value)
                    oDBs_Head.SetValue("U_EffCheck", oDBs_Head.Offset, ors1.Fields.Item("U_EffCheck").Value)
                    oDBs_Head.SetValue("U_TarDays", oDBs_Head.Offset, ors1.Fields.Item("U_TarDays").Value)
                    'oDBs_Head.SetValue("U_Active", oDBs_Head.Offset, ors1.Fields.Item("U_Active").Value)
                    objsectMat.SetLineData(objsectMat.VisualRowCount)
                    ors1.MoveNext()
                Next
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            Else
                Me.SetNewLine(objForm.UniqueID)
            End If
            objMatrix.FlushToDataSource()
            ' objsectMat.AutoResizeColumns()
            ' objsectMat.FlushToDataSource()
            'objMatrix.LoadFromDataSource()

        Catch ex As Exception
            oDBs_Head.Freeze(False)
        End Try
    End Sub

End Class

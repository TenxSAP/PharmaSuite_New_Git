Public Class RegulatoryAuthority

#Region "Declaration"

    Public objForm As SAPbouiCOM.Form

    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objUtilities As Utilities

    Public oDBs_Head As SAPbouiCOM.DBDataSource

#End Region

#Region "Create Form"

    Sub CreateForm()

        Try

            objMain.objUtilities.LoadForm(
                "RegulatoryAuthority.xml",
                "REG_AUTH",
                ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm(
                "REG_AUTH",
                objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            objUtilities = New Utilities

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_AUTH")

            objMatrix = CType(objForm.Items.Item("CMtx").Specific,
                              SAPbouiCOM.Matrix)

            Dim rs As SAPbobsCOM.Recordset

            rs = objMain.objCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery("SELECT TOP 1 ""DocEntry"" FROM ""@TNX_REG_AUTH""")

            If rs.RecordCount = 0 Then

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

                Me.SetNewLine(objForm.UniqueID)

            Else

                Me.MatrixLoad()

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

            End If

            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)

            objMatrix.Columns.Item("Code").Editable = False
            objMatrix.Columns.Item("Code").Visible = False

            objMatrix.AutoResizeColumns()

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
                "Successfully initialized, Please proceed...",
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

#Region "Menu Event"

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent,
                  ByRef BubbleEvent As Boolean)

        Try

            '========================================================================
            ' OPEN FORM
            '========================================================================
            If pVal.MenuUID = "10X_RMS_AUTH" _
                AndAlso pVal.BeforeAction = False Then

                Me.CreateForm()

                '========================================================================
                ' ADD ROW
                '========================================================================
            ElseIf pVal.MenuUID = "1292" _
                AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "REG_AUTH" Then Exit Sub

                objMatrix = CType(objForm.Items.Item("CMtx").Specific,
                                  SAPbouiCOM.Matrix)

                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                End If

                Me.SetNewLine(objForm.UniqueID)

                '========================================================================
                ' DELETE ROW
                '========================================================================
            ElseIf pVal.MenuUID = "1293" _
                AndAlso pVal.BeforeAction = True Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "REG_AUTH" Then Exit Sub

                    BubbleEvent = False

                    objForm.Freeze(True)

                    objMatrix = CType(objForm.Items.Item("CMtx").Specific,
                                      SAPbouiCOM.Matrix)

                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_AUTH")

                    Dim selectedRow As Integer

                    selectedRow = objMatrix.GetNextSelectedRow(
                        0,
                        SAPbouiCOM.BoOrderType.ot_RowOrder)

                    If selectedRow <= 0 Then

                        objMain.objApplication.StatusBar.SetText(
                            "Please select row to delete",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        Exit Try

                    End If

                    objMatrix.DeleteRow(selectedRow)

                    While oDBs_Head.Size > objMatrix.VisualRowCount

                        oDBs_Head.RemoveRecord(oDBs_Head.Size - 1)

                    End While

                    If oDBs_Head.Size = 0 Then

                        oDBs_Head.InsertRecord(0)

                        oDBs_Head.SetValue("DocEntry", 0, "1")

                        oDBs_Head.SetValue("U_Country", 0, "")
                        oDBs_Head.SetValue("U_AuthType", 0, "")
                        oDBs_Head.SetValue("U_Website", 0, "")
                        oDBs_Head.SetValue("U_Email", 0, "")
                        oDBs_Head.SetValue("U_Phone", 0, "")
                        oDBs_Head.SetValue("U_SubMode", 0, "")
                        oDBs_Head.SetValue("U_DefTime", 0, "")
                        oDBs_Head.SetValue("U_AgentReq", 0, "Y")
                        oDBs_Head.SetValue("U_Status", 0, "Y")
                        oDBs_Head.SetValue("U_Remarks", 0, "")

                    End If

                    objMatrix.LoadFromDataSource()

                    For i As Integer = 1 To objMatrix.VisualRowCount

                        objMatrix.Columns.Item("DocEntry") _
                            .Cells.Item(i).Specific.Value = i.ToString()

                    Next

                    objMatrix.AutoResizeColumns()

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

                If objForm.TypeEx <> "REG_AUTH" Then Exit Sub

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

#End Region

#Region "Item Event"

    Sub ItemEvent(ByVal FormUID As String,
                  ByRef pVal As SAPbouiCOM.ItemEvent,
                  ByRef BubbleEvent As Boolean)

        Try

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "1" _
                        AndAlso pVal.BeforeAction = False _
                        AndAlso pVal.ActionSuccess = True Then

                        Me.MatrixLoad()

                        objForm.Mode =
                            SAPbouiCOM.BoFormMode.fm_OK_MODE

                    End If

            End Select

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
                ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

#End Region

#Region "Set New Line"

    Sub SetNewLine(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_AUTH")

            objMatrix = CType(objForm.Items.Item("CMtx").Specific,
                              SAPbouiCOM.Matrix)

            If objMatrix.VisualRowCount > 0 Then
                objMatrix.FlushToDataSource()
            End If

            Dim rowIndex As Integer

            If oDBs_Head.Size = 1 _
                AndAlso oDBs_Head.GetValue("DocEntry", 0).Trim = "" Then

                rowIndex = 0

            Else

                rowIndex = oDBs_Head.Size

                oDBs_Head.InsertRecord(rowIndex)

            End If

            oDBs_Head.Offset = rowIndex

            oDBs_Head.SetValue("DocEntry",
                               rowIndex,
                               (rowIndex + 1).ToString())

            oDBs_Head.SetValue("U_Country", rowIndex, "")
            oDBs_Head.SetValue("U_AuthType", rowIndex, "")
            oDBs_Head.SetValue("U_Website", rowIndex, "")
            oDBs_Head.SetValue("U_Email", rowIndex, "")
            oDBs_Head.SetValue("U_Phone", rowIndex, "")
            oDBs_Head.SetValue("U_SubMode", rowIndex, "")
            oDBs_Head.SetValue("U_DefTime", rowIndex, "")
            oDBs_Head.SetValue("U_AgentReq", rowIndex, "Y")
            oDBs_Head.SetValue("U_Status", rowIndex, "Y")
            oDBs_Head.SetValue("U_Remarks", rowIndex, "")

            objMatrix.LoadFromDataSource()

            For i As Integer = 1 To objMatrix.VisualRowCount

                objMatrix.Columns.Item("DocEntry") _
                    .Cells.Item(i).Specific.Value = i.ToString()

            Next

            objMatrix.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

#End Region

#Region "Matrix Load"

    Public Sub MatrixLoad()

        Try

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_AUTH")

            objMatrix = CType(objForm.Items.Item("CMtx").Specific,
                              SAPbouiCOM.Matrix)

            Dim strQuery As String

            strQuery = "SELECT " &
                       """DocEntry"", " &
                       """U_Country"", " &
                       """U_AuthType"", " &
                       """U_Website"", " &
                       """U_Email"", " &
                       """U_Phone"", " &
                       """U_SubMode"", " &
                       """U_DefTime"", " &
                       """U_AgentReq"", " &
                       """U_Status"", " &
                       """U_Remarks"" " &
                       "FROM ""@TNX_REG_AUTH"" " &
                       "ORDER BY ""DocEntry"""

            Dim ors1 As SAPbobsCOM.Recordset

            ors1 = objMain.objCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.BoRecordset)

            ors1.DoQuery(strQuery)

            oDBs_Head.Clear()

            If ors1.RecordCount > 0 Then

                For j As Integer = 0 To ors1.RecordCount - 1

                    oDBs_Head.InsertRecord(j)

                    oDBs_Head.Offset = j

                    oDBs_Head.SetValue(
                        "DocEntry",
                        j,
                        ors1.Fields.Item("DocEntry").Value.ToString())

                    oDBs_Head.SetValue(
                        "U_Country",
                        j,
                        ors1.Fields.Item("U_Country").Value.ToString())

                    oDBs_Head.SetValue(
                        "U_AuthType",
                        j,
                        ors1.Fields.Item("U_AuthType").Value.ToString())

                    oDBs_Head.SetValue(
                        "U_Website",
                        j,
                        ors1.Fields.Item("U_Website").Value.ToString())

                    oDBs_Head.SetValue(
                        "U_Email",
                        j,
                        ors1.Fields.Item("U_Email").Value.ToString())

                    oDBs_Head.SetValue(
                        "U_Phone",
                        j,
                        ors1.Fields.Item("U_Phone").Value.ToString())

                    oDBs_Head.SetValue(
                        "U_SubMode",
                        j,
                        ors1.Fields.Item("U_SubMode").Value.ToString())

                    oDBs_Head.SetValue(
                        "U_DefTime",
                        j,
                        ors1.Fields.Item("U_DefTime").Value.ToString())

                    oDBs_Head.SetValue(
                        "U_AgentReq",
                        j,
                        ors1.Fields.Item("U_AgentReq").Value.ToString())

                    oDBs_Head.SetValue(
                        "U_Status",
                        j,
                        ors1.Fields.Item("U_Status").Value.ToString())

                    oDBs_Head.SetValue(
                        "U_Remarks",
                        j,
                        ors1.Fields.Item("U_Remarks").Value.ToString())

                    ors1.MoveNext()

                Next

                objMatrix.LoadFromDataSource()

                For i As Integer = 1 To objMatrix.VisualRowCount

                    objMatrix.Columns.Item("DocEntry") _
                        .Cells.Item(i).Specific.Value = i.ToString()

                Next

                objForm.Mode =
                    SAPbouiCOM.BoFormMode.fm_OK_MODE

            Else

                Me.SetNewLine(objForm.UniqueID)

            End If

            objMatrix.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

#End Region

End Class
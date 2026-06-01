Public Class RegulatoryStatus
#Region "Declaration"

    Public objForm As SAPbouiCOM.Form

    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objUtilities As Utilities

    Public oDBs_Head As SAPbouiCOM.DBDataSource

#End Region

    Sub CreateForm()

        Try

            objMain.objUtilities.LoadForm(
                    "RegulatoryStatusMaster.xml",
                    "REG_STAT",
                    ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm(
                    "REG_STAT",
                    objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            objUtilities = New Utilities

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_STAT")

            objMatrix = objForm.Items.Item("Mtx").Specific

            Dim rs As SAPbobsCOM.Recordset

            rs = objMain.objCompany.GetBusinessObject(
                    SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(
                    "SELECT TOP 1 ""Code"" FROM ""@TNX_REG_STAT""")

            If rs.RecordCount = 0 Then

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

                Me.SetNewLine(objForm.UniqueID)

            Else

                Me.MatrixLoad()

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

            End If

            objMatrix.Columns.Item("Code").Editable = True

            objForm.Freeze(False)

            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)

            objMain.objApplication.StatusBar.SetText(
                    "Successfully initialized, Please proceed...",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
                    ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent,
                      ByRef BubbleEvent As Boolean)

        Try

            If pVal.MenuUID = "10X_RMS_STAT" _
                    And pVal.BeforeAction = False Then

                Me.CreateForm()

            ElseIf pVal.MenuUID = "1292" _
                    And pVal.BeforeAction = False Then

                objMatrix = objForm.Items.Item("Mtx").Specific

                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If

                Me.SetNewLine(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1293" _
                    And pVal.BeforeAction = False Then

                objMatrix = objForm.Items.Item("Mtx").Specific

                Dim row As Integer = objMatrix.VisualRowCount

                If objMatrix.IsRowSelected(1) <> True _
                        And objMatrix.VisualRowCount < 1 Then

                    objMatrix.AddRow()

                    oDBs_Head.SetValue(
                            "DocEntry",
                            oDBs_Head.Offset,
                            objMatrix.VisualRowCount)

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

                    objMatrix.Columns.Item("DocEntry") _
                            .Cells.Item(i).Specific.Value = i

                Next

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
                    ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Sub ItemEvent(ByVal FormUID As String,
                      ByRef pVal As SAPbouiCOM.ItemEvent,
                      ByRef BubbleEvent As Boolean)

        Try

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "1" _
                            And pVal.BeforeAction = False _
                            And pVal.ActionSuccess = True Then

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

    Sub SetNewLine(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_STAT")

            objMatrix = objForm.Items.Item("Mtx").Specific

            If objMatrix.VisualRowCount > 0 Then
                objMatrix.FlushToDataSource()
            End If

            Dim rowIndex As Integer

            If oDBs_Head.Size = 1 _
                    And oDBs_Head.GetValue("Code", 0).Trim = "" Then

                rowIndex = 0

            Else

                rowIndex = oDBs_Head.Size
                oDBs_Head.InsertRecord(rowIndex)

            End If

            oDBs_Head.Offset = rowIndex

            oDBs_Head.SetValue("Code", rowIndex, (rowIndex + 1).ToString())
            oDBs_Head.SetValue("Name", rowIndex, "")
            oDBs_Head.SetValue("U_StatGroup", rowIndex, "")
            oDBs_Head.SetValue("U_AppForm", rowIndex, "")
            oDBs_Head.SetValue("U_Editable", rowIndex, "Y")
            oDBs_Head.SetValue("U_Final", rowIndex, "N")
            oDBs_Head.SetValue("U_AllowCopy", rowIndex, "N")
            oDBs_Head.SetValue("U_ReqAppr", rowIndex, "N")
            oDBs_Head.SetValue("U_SeqNo", rowIndex, "")
            oDBs_Head.SetValue("U_Status", rowIndex, "Y")
            ' oDBs_Head.SetValue("U_FromStatus", rowIndex, "")
            oDBs_Head.SetValue("U_changby", rowIndex, "")
            oDBs_Head.SetValue("U_StatusDate", rowIndex, "")
            oDBs_Head.SetValue("U_Reason", rowIndex, "")
            'oDBs_Head.SetValue("U_Remarks", rowIndex, "")

            objMatrix.LoadFromDataSource()

            For i As Integer = 1 To objMatrix.VisualRowCount

                objMatrix.Columns.Item("DocEntry") _
                        .Cells.Item(i).Specific.Value = i

            Next

            objMatrix.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

    Public Sub MatrixLoad()

        Try

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_STAT")

            objMatrix = objForm.Items.Item("Mtx").Specific

            Dim rs1 As String

            rs1 = "SELECT " &
              """Code"", " &
              """Name"", " &
              """U_StatGroup"", " &
              """U_AppForm"", " &
              """U_Editable"", " &
              """U_Final"", " &
              """U_AllowCopy"", " &
              """U_ReqAppr"", " &
              """U_SeqNo"", " &
              """U_Status"", " &
              """U_StatDate"", " &
              """U_FromStat"", " &
              """U_ChangBy"", " &
              """U_Reason"", " &
              """U_StatusDate"", " &
              """U_FromStatus"" " &
              "FROM ""@TNX_REG_STAT"" ORDER BY ""Code"""

            Dim ors1 As SAPbobsCOM.Recordset

            ors1 = objMain.objCompany.GetBusinessObject(
            SAPbobsCOM.BoObjectTypes.BoRecordset)

            ors1.DoQuery(rs1)

            oDBs_Head.Clear()

            If ors1.RecordCount > 0 Then

                For j As Integer = 0 To ors1.RecordCount - 1

                    oDBs_Head.InsertRecord(j)

                    oDBs_Head.Offset = j

                    oDBs_Head.SetValue("Code", j,
                    ors1.Fields.Item("Code").Value.ToString())

                    oDBs_Head.SetValue("Name", j,
                    ors1.Fields.Item("Name").Value.ToString())

                    oDBs_Head.SetValue("U_StatGroup", j,
                    ors1.Fields.Item("U_StatGroup").Value.ToString())

                    oDBs_Head.SetValue("U_AppForm", j,
                    ors1.Fields.Item("U_AppForm").Value.ToString())

                    oDBs_Head.SetValue("U_Editable", j,
                    ors1.Fields.Item("U_Editable").Value.ToString())

                    oDBs_Head.SetValue("U_Final", j,
                    ors1.Fields.Item("U_Final").Value.ToString())

                    oDBs_Head.SetValue("U_AllowCopy", j,
                    ors1.Fields.Item("U_AllowCopy").Value.ToString())

                    oDBs_Head.SetValue("U_ReqAppr", j,
                    ors1.Fields.Item("U_ReqAppr").Value.ToString())

                    oDBs_Head.SetValue("U_SeqNo", j,
                    ors1.Fields.Item("U_SeqNo").Value.ToString())

                    oDBs_Head.SetValue("U_Status", j,
                    ors1.Fields.Item("U_Status").Value.ToString())

                    'oDBs_Head.SetValue("U_StatDate", j,
                    'ors1.Fields.Item("U_StatDate").Value.ToString())

                    oDBs_Head.SetValue("U_FromStat", j,
                    ors1.Fields.Item("U_FromStat").Value.ToString())

                    oDBs_Head.SetValue("U_ChangBy", j,
                    ors1.Fields.Item("U_ChangBy").Value.ToString())

                    oDBs_Head.SetValue("U_Reason", j,
                    ors1.Fields.Item("U_Reason").Value.ToString())

                    'oDBs_Head.SetValue("U_StatusDate", j,
                    'ors1.Fields.Item("U_StatusDate").Value.ToString())

                    oDBs_Head.SetValue("U_FromStatus", j,
                    ors1.Fields.Item("U_FromStatus").Value.ToString())

                    ors1.MoveNext()

                Next

                objMatrix.LoadFromDataSource()

                For i As Integer = 1 To objMatrix.VisualRowCount

                    objMatrix.Columns.Item("LineId") _
                        .Cells.Item(i).Specific.Value = i

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

End Class

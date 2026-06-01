Public Class RegulatorySubmission


#Region "Declaration"

    Public objForm As SAPbouiCOM.Form

        Dim objMatrix As SAPbouiCOM.Matrix
        Dim objUtilities As Utilities

        Public oDBs_Head As SAPbouiCOM.DBDataSource

#End Region

        Sub CreateForm()

            Try

                objMain.objUtilities.LoadForm(
                "RegulatorySubmission.xml",
                "TNX_REG_SUBTYP",
                ResourceType.Embeded)

                objForm = objMain.objApplication.Forms.GetForm(
                "TNX_REG_SUBTYP",
                objMain.objApplication.Forms.ActiveForm.TypeCount)

                objForm.Freeze(True)

                objUtilities = New Utilities

                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_SUBTYP")

                objMatrix = objForm.Items.Item("Mtx").Specific

                Dim rs As SAPbobsCOM.Recordset

                rs = objMain.objCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.BoRecordset)

                rs.DoQuery(
                "SELECT TOP 1 ""Code"" FROM ""@TNX_REG_SUBTYP""")

                If rs.RecordCount = 0 Then

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

                    Me.SetNewLine(objForm.UniqueID)

                Else

                    Me.MatrixLoad()

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                End If

                objMatrix = objForm.Items.Item("Mtx").Specific

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

            If pVal.MenuUID = "10X_RMS_SUBTYP" _
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
                        "LineId",
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

                        objMatrix.Columns.Item("LineId") _
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

                        objMatrix = objForm.Items.Item("Mtx").Specific

                        objMatrix.FlushToDataSource()

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

                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_SUBTYP")

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
                oDBs_Head.SetValue("U_Category", rowIndex, "")
                oDBs_Head.SetValue("U_DossReq", rowIndex, "N")
                oDBs_Head.SetValue("U_ArtReq", rowIndex, "N")
                oDBs_Head.SetValue("U_ApprReq", rowIndex, "N")
                oDBs_Head.SetValue("U_ExpDays", rowIndex, "")
                oDBs_Head.SetValue("U_QueryAll", rowIndex, "N")
                oDBs_Head.SetValue("U_CCReq", rowIndex, "N")
                oDBs_Head.SetValue("U_Status", rowIndex, "Y")

                objMatrix.LoadFromDataSource()

                For i As Integer = 1 To objMatrix.VisualRowCount

                    objMatrix.Columns.Item("LineId") _
                    .Cells.Item(i).Specific.Value = i

                Next

                objMatrix.AutoResizeColumns()

            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try

        End Sub

        Public Sub MatrixLoad()

            Try

                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_SUBTYP")

                objMatrix = objForm.Items.Item("Mtx").Specific

                Dim rs1 As String

                rs1 =
                "SELECT " &
                """Code"", " &
                """Name"", " &
                """U_Category"", " &
                """U_DossReq"", " &
                """U_ArtReq"", " &
                """U_ApprReq"", " &
                """U_ExpDays"", " &
                """U_QueryAll"", " &
                """U_CCReq"", " &
                """U_Status"" " &
                "FROM ""@TNX_REG_SUBTYP"" " &
                "ORDER BY ""Code"""

                Dim ors1 As SAPbobsCOM.Recordset

                ors1 = objMain.objCompany.GetBusinessObject(
                SAPbobsCOM.BoObjectTypes.BoRecordset)

                ors1.DoQuery(rs1)

                oDBs_Head.Clear()

                If ors1.RecordCount > 0 Then

                    For j As Integer = 0 To ors1.RecordCount - 1

                        oDBs_Head.InsertRecord(j)

                        oDBs_Head.Offset = j

                        oDBs_Head.SetValue(
                        "Code",
                        j,
                        ors1.Fields.Item("Code").Value.ToString())

                        oDBs_Head.SetValue(
                        "Name",
                        j,
                        ors1.Fields.Item("Name").Value.ToString())

                        oDBs_Head.SetValue(
                        "U_Category",
                        j,
                        ors1.Fields.Item("U_Category").Value.ToString())

                        oDBs_Head.SetValue(
                        "U_DossReq",
                        j,
                        ors1.Fields.Item("U_DossReq").Value.ToString())

                        oDBs_Head.SetValue(
                        "U_ArtReq",
                        j,
                        ors1.Fields.Item("U_ArtReq").Value.ToString())

                        oDBs_Head.SetValue(
                        "U_ApprReq",
                        j,
                        ors1.Fields.Item("U_ApprReq").Value.ToString())

                        oDBs_Head.SetValue(
                        "U_ExpDays",
                        j,
                        ors1.Fields.Item("U_ExpDays").Value.ToString())

                        oDBs_Head.SetValue(
                        "U_QueryAll",
                        j,
                        ors1.Fields.Item("U_QueryAll").Value.ToString())

                        oDBs_Head.SetValue(
                        "U_CCReq",
                        j,
                        ors1.Fields.Item("U_CCReq").Value.ToString())

                        oDBs_Head.SetValue(
                        "U_Status",
                        j,
                        ors1.Fields.Item("U_Status").Value.ToString())

                        ors1.MoveNext()

                    Next

                    objMatrix.LoadFromDataSource()

                    For i As Integer = 1 To objMatrix.VisualRowCount

                    objMatrix.Columns.Item("DocEntry") _
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

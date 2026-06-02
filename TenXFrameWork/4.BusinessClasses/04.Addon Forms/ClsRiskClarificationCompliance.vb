Public Class ClsRiskClarificationCompliance


#Region "Decliration"

    Public objForm As SAPbouiCOM.Form

        Dim objMatrix As SAPbouiCOM.Matrix
        Dim objutilities As Utilities

        Public oDBs_Head As SAPbouiCOM.DBDataSource

#End Region

        Sub CreateForm()

            Try

                objMain.objUtilities.LoadForm("RiskClarificationMasterCompliance.xml", "RISKCLMSTR", ResourceType.Embeded)

                objForm = objMain.objApplication.Forms.GetForm("RISKCLMSTR", objMain.objApplication.Forms.ActiveForm.TypeCount)

                objForm.Freeze(True)

                objutilities = New Utilities

                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_RISK")

                objMatrix = objForm.Items.Item("CMtx").Specific

                Dim rs As SAPbobsCOM.Recordset

                rs = objMain.objCompany.GetBusinessObject(
            SAPbobsCOM.BoObjectTypes.BoRecordset)

                rs.DoQuery(
            "SELECT TOP 1 ""Code"" FROM ""@TNX_RISK""")

                If rs.RecordCount = 0 Then

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

                    Me.SetNewLine(objForm.UniqueID)

                Else

                    Me.MatrixLoad()

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                End If

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

                If pVal.MenuUID = "10X_CMS_RISK" _
        And pVal.BeforeAction = False Then

                    Me.CreateForm()

                ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                    objMatrix = objForm.Items.Item("CMtx").Specific
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    Me.SetNewLine(objForm.UniqueID)

                ElseIf pVal.MenuUID = "1293" And pVal.BeforeAction = False Then
                    objMatrix = objForm.Items.Item("CMtx").Specific
                    Dim row As Integer = objMatrix.VisualRowCount
                    'If objMatrix.VisualRowCount <= 1 Then
                    If objMatrix.IsRowSelected(1) <> True And objMatrix.VisualRowCount < 1 Then
                        objMatrix.AddRow()
                        oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMatrix.VisualRowCount)
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

                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_RISK")

                objMatrix = objForm.Items.Item("CMtx").Specific

                If objMatrix.VisualRowCount > 0 Then
                    objMatrix.FlushToDataSource()
                End If
                Dim rowIndex As Integer

                ' ===== FIX EMPTY FIRST ROW ISSUE =====
                If oDBs_Head.Size = 1 _
        And oDBs_Head.GetValue("Code", 0).Trim = "" Then

                    rowIndex = 0

                Else

                    rowIndex = oDBs_Head.Size
                    oDBs_Head.InsertRecord(rowIndex)

                End If

                oDBs_Head.Offset = rowIndex

                oDBs_Head.SetValue("Code", rowIndex, (rowIndex + 1).ToString())
            oDBs_Head.SetValue("U_RiskLevel", rowIndex, "")
            oDBs_Head.SetValue("U_MinScore", rowIndex, "")
            oDBs_Head.SetValue("U_MaxScore", rowIndex, "")
            oDBs_Head.SetValue("U_QAApprvl", rowIndex, "")
            oDBs_Head.SetValue("U_MgmtAppr", rowIndex, "")
            oDBs_Head.SetValue("U_CAPAReq", rowIndex, "")

            oDBs_Head.SetValue("U_ValidReq", rowIndex, "")
            oDBs_Head.SetValue("U_EscDays", rowIndex, "")

            oDBs_Head.SetValue("U_Active", rowIndex, "Y")

                objMatrix.LoadFromDataSource()


                For i As Integer = 1 To objMatrix.VisualRowCount

                    objMatrix.Columns.Item("DocEntry") _
            .Cells.Item(i).Specific.Value = i

                Next

                objMatrix.AutoResizeColumns()
                'objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE


            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try

        End Sub
        Public Sub MatrixLoad()

            Try

                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_RISK")

                objMatrix = objForm.Items.Item("CMtx").Specific

                Dim rs1 As String

            rs1 = "SELECT ""Code"", ""U_RiskLevel"", ""U_MinScore"", ""U_MaxScore"", " &
              """U_QAApprvl"", ""U_MgmtAppr"", ""U_CAPAReq"",""U_ValidReq"",""U_EscDays"", ""U_Active"" " &
              "FROM ""@TNX_RISK"" ORDER BY ""Code"""

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

                    oDBs_Head.SetValue("U_RiskLevel", j,
                                   ors1.Fields.Item("U_RiskLevel").Value.ToString())

                    oDBs_Head.SetValue("U_MinScore", j,
                                   ors1.Fields.Item("U_MinScore").Value.ToString())

                    oDBs_Head.SetValue("U_MaxScore", j,
                                   ors1.Fields.Item("U_MaxScore").Value.ToString())

                    oDBs_Head.SetValue("U_QAApprvl", j,
                                   ors1.Fields.Item("U_QAApprvl").Value.ToString())

                    oDBs_Head.SetValue("U_MgmtAppr", j,
                                   ors1.Fields.Item("U_MgmtAppr").Value.ToString())

                    oDBs_Head.SetValue("U_CAPAReq", j,
                                   ors1.Fields.Item("U_CAPAReq").Value.ToString())

                    oDBs_Head.SetValue("U_ValidReq", j,
                                   ors1.Fields.Item("U_ValidReq").Value.ToString())

                    oDBs_Head.SetValue("U_EscDays", j,
                                   ors1.Fields.Item("U_EscDays").Value.ToString())

                    oDBs_Head.SetValue("U_Active", j,
                                   ors1.Fields.Item("U_Active").Value.ToString())

                        ors1.MoveNext()

                    Next

                    objMatrix.LoadFromDataSource()

                    For i As Integer = 1 To objMatrix.VisualRowCount

                        objMatrix.Columns.Item("DocEntry").Cells.Item(i).Specific.Value = i

                    Next

                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                Else

                    Me.SetNewLine(objForm.UniqueID)

                End If

                objMatrix.AutoResizeColumns()

            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try

        End Sub






    End Class



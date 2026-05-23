Public Class clsMstrComplianceSett

#Region "Decliration"

    Public objForm As SAPbouiCOM.Form

    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objutilities As Utilities

    Public oDBs_Head As SAPbouiCOM.DBDataSource

#End Region

    Sub CreateForm()

        Try

            objMain.objUtilities.LoadForm(
            "MstrComplianceSett.xml",
            "frm_COMPSET",
            ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm(
            "frm_COMPSET",
            objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            objutilities = New Utilities

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_COMPSET")

            objMatrix = objForm.Items.Item("CMtx").Specific

            Dim rs As SAPbobsCOM.Recordset

            rs = objMain.objCompany.GetBusinessObject(
            SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(
            "SELECT TOP 1 ""Code"" FROM ""@TNX_COMPSET""")

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

            If pVal.MenuUID = "10X_CMS_COMPSET" _
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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_COMPSET")

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
            oDBs_Head.SetValue("Name", rowIndex, "")
            oDBs_Head.SetValue("U_AuditTrail", rowIndex, "Y")
            oDBs_Head.SetValue("U_ESign", rowIndex, "Y")
            oDBs_Head.SetValue("U_AutoNumber", rowIndex, "Y")
            oDBs_Head.SetValue("U_EmailAlert", rowIndex, "Y")
            oDBs_Head.SetValue("U_Active", rowIndex, "Y")

            oDBs_Head.SetValue("U_AttachM", rowIndex, "Y")
            oDBs_Head.SetValue("U_AutoTr", rowIndex, "Y")
            oDBs_Head.SetValue("U_CAPAAuto", rowIndex, "Y")
            oDBs_Head.SetValue("U_WhatsApp", rowIndex, "Y")
            oDBs_Head.SetValue("U_LockC", rowIndex, "Y")
            oDBs_Head.SetValue("U_AllowB", rowIndex, "Y")
            oDBs_Head.SetValue("U_MaxBD", rowIndex, "Y")

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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_COMPSET")

            objMatrix =
        objForm.Items.Item("CMtx").Specific

            Dim rs1 As String

            rs1 = "SELECT ""Code"",""Name"", ""U_AuditTrail"", ""U_ESign"", ""U_AutoNumber"", ""U_EmailAlert"", ""U_Active"",""U_AttachM"",""U_AutoTr"",""U_CAPAAuto"",""U_WhatsApp"",""U_LockC"",""U_AllowB"",""U_MaxBD"" " &
        "FROM ""@TNX_COMPSET"" ORDER BY ""Code"""

            Dim ors1 As SAPbobsCOM.Recordset

            ors1 =
        objMain.objCompany.GetBusinessObject(
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
                "U_AuditTrail",
                j,
                ors1.Fields.Item("U_AuditTrail").Value.ToString())

                    oDBs_Head.SetValue(
                "U_ESign",
                j,
                ors1.Fields.Item("U_ESign").Value.ToString())

                    oDBs_Head.SetValue(
                "U_AutoNumber",
                j,
                ors1.Fields.Item("U_AutoNumber").Value.ToString())

                    oDBs_Head.SetValue(
                "U_EmailAlert",
                j,
                ors1.Fields.Item("U_EmailAlert").Value.ToString())

                    oDBs_Head.SetValue(
                "U_AttachM",
                j,
                ors1.Fields.Item("U_AttachM").Value.ToString())
                    oDBs_Head.SetValue(
                "U_AutoTr",
                j,
                ors1.Fields.Item("U_AutoTr").Value.ToString())
                    oDBs_Head.SetValue(
                "U_CAPAAuto",
                j,
                ors1.Fields.Item("U_CAPAAuto").Value.ToString())
                    oDBs_Head.SetValue(
                "U_WhatsApp",
                j,
                ors1.Fields.Item("U_WhatsApp").Value.ToString())

                    oDBs_Head.SetValue(
                "U_LockC",
                j,
                ors1.Fields.Item("U_LockC").Value.ToString())

                    oDBs_Head.SetValue(
                "U_AllowB",
                j,
                ors1.Fields.Item("U_AllowB").Value.ToString())

                    oDBs_Head.SetValue(
                "U_MaxBD",
                j,
                ors1.Fields.Item("U_MaxBD").Value.ToString())

                    oDBs_Head.SetValue(
                "U_Active",
                j,
                ors1.Fields.Item("U_Active").Value.ToString())

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

            objMain.objApplication.StatusBar.SetText(
        ex.Message)

        End Try

    End Sub
End Class

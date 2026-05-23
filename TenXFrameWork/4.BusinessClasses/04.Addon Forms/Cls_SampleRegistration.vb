Imports System.Configuration
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports Newtonsoft.Json.Linq

Public Class Cls_SampleRegistration

#Region "Declaration"

    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBS1, oDBS2 As SAPbouiCOM.DBDataSource
    Dim oDBs_Details1, oDBs_Details2 As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim str, str1 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim objutilities As Utilities
    Dim MATRIXS As String

#End Region

    Sub CreateForm()

        Try

            objMain.objUtilities.LoadForm("SampleRegistration.xml", "SAMPLE", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("SAMPLE",
                                                           objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            objutilities = New Utilities

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_QASMPH")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_QASMPL")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_QASMPA")

            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("774", True)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Me.SetDefault(objForm.UniqueID)

            Dim item As SAPbouiCOM.Item

            item = objForm.Items.Item("DocNum")

            item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                                         SAPbouiCOM.BoFormMode.fm_ADD_MODE,
                                         SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                                         SAPbouiCOM.BoFormMode.fm_UPDATE_MODE,
                                         SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                                         SAPbouiCOM.BoFormMode.fm_OK_MODE,
                                         SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                                         SAPbouiCOM.BoFormMode.fm_FIND_MODE,
                                         SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix1 = objForm.Items.Item("Mtx1").Specific

            objForm.Freeze(False)


            Me.SetNewLine(objForm.UniqueID)
            Me.SetNewLine2(objForm.UniqueID)

            'Me.AutoDocentryNumber()

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

            If pVal.MenuUID = "10X_SAMPLE" And pVal.BeforeAction = False Then

                Me.CreateForm()

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

                objMatrix = objForm.Items.Item("Mtx").Specific
                objMatrix1 = objForm.Items.Item("Mtx1").Specific
                Me.SetDefault(objForm.UniqueID)

                Me.SetNewLine(objForm.UniqueID)
                Me.SetNewLine2(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

                objMatrix = objForm.Items.Item("Mtx").Specific
                objMatrix1 = objForm.Items.Item("Mtx1").Specific

                Me.SetNewLine(objForm.UniqueID)
                Me.SetNewLine2(objForm.UniqueID)

            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.GetForm("SAMPLE",
                                                               objMain.objApplication.Forms.ActiveForm.TypeCount)

                objMatrix = objForm.Items.Item("Mtx").Specific
                objMatrix1 = objForm.Items.Item("Mtx1").Specific

                Dim row As Integer = objMatrix.VisualRowCount

                If MATRIXS.Equals("Mtx") = True Then

                    If objMatrix.IsRowSelected(1) <> True And
                       objMatrix.VisualRowCount < 1 Then

                        objMatrix.AddRow()

                        oDBs_Details1.SetValue("LineId",
                                       oDBs_Details1.Offset,
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

                        objMatrix.Columns.Item("LineId").Cells.Item(i).Specific.Value = i

                    Next

                ElseIf MATRIXS.Equals("Mtx1") = True Then

                    Dim row1 As Integer = objMatrix1.VisualRowCount

                    If objMatrix1.IsRowSelected(1) <> True And
                       objMatrix1.VisualRowCount < 1 Then

                        objMatrix1.AddRow()

                        oDBs_Details2.SetValue("LineId",
                                       oDBs_Details2.Offset,
                                       objMatrix1.VisualRowCount)

                        objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                    End If

                    If objMatrix1.IsRowSelected(row1) = True Then

                        objMatrix1.DeleteRow(row1)

                    Else

                        For i As Integer = 1 To objMatrix1.VisualRowCount - 1

                            If objMatrix1.IsRowSelected(i) = True Then
                                objMatrix1.DeleteRow(i)
                            End If

                        Next

                    End If

                    For i As Integer = 1 To objMatrix1.VisualRowCount

                        objMatrix1.Columns.Item("LineId").Cells.Item(i).Specific.Value = i

                    Next

                End If

                'If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or
                '   objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

                '    'objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                'End If
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
            End If

        Catch ex As Exception

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
                ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Private Sub DisableMatrix(ByVal FormUID As String)

        Try

            Dim objForm As SAPbouiCOM.Form
            Dim objMatrix1 As SAPbouiCOM.Matrix

            objForm = objMain.objApplication.Forms.Item(FormUID)

            objMatrix1 = CType(objForm.Items.Item("Mtx").Specific,
                               SAPbouiCOM.Matrix)

            objForm.Freeze(True)

            objMatrix1.Columns.Item("CREAT").Visible = False
            objMatrix1.Columns.Item("CREBY").Visible = False
            objMatrix1.Columns.Item("UPDAT").Visible = False
            objMatrix1.Columns.Item("UPDBY").Visible = False

            objForm.Freeze(False)

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

    Sub ItemEvent(ByVal FormUID As String,
                  ByRef pVal As SAPbouiCOM.ItemEvent,
                  ByRef BubbleEvent As Boolean)

        Try

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_CLICK

                    If (pVal.ItemUID = "Mtx") And
                       pVal.BeforeAction = False Then

                        MATRIXS = pVal.ItemUID

                    End If

                    If (pVal.ItemUID = "Mtx1") And
                       pVal.BeforeAction = False Then

                        MATRIXS = pVal.ItemUID

                    End If


                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = False Then
                        Me.SetDefault(objForm.UniqueID)
                        Me.SetNewLine(objForm.UniqueID)

                        Me.SetNewLine2(objForm.UniqueID)
                    End If

            End Select

        Catch ex As Exception

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
                ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Sub SetNewLine(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_QASMPL")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_QASMPA")

            objMatrix = objForm.Items.Item("Mtx").Specific

            objMatrix.AddRow()

            oDBs_Details1.SetValue("LineId",
                                   oDBs_Details1.Offset,
                                   objMatrix.VisualRowCount.ToString())

            oDBs_Details1.SetValue("U_TESTCD", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_TENM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_TESTCA", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MEDCD", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MTDNM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_MINVAL", oDBs_Details1.Offset, 0)
            oDBs_Details1.SetValue("U_MAXVAL", oDBs_Details1.Offset, 0)
            oDBs_Details1.SetValue("U_STDVAL", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_UOM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_ISMNDT", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_TESTS", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_ASSIGN", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_TADATE", oDBs_Details1.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)

            objMatrix.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

    Sub SetNewLine2(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_QASMPA")

            objMatrix1 = CType(objForm.Items.Item("Mtx1").Specific,
                               SAPbouiCOM.Matrix)

            objMatrix1.AddRow()

            oDBs_Details2.SetValue("LineId",
                                   oDBs_Details2.Offset,
                                   objMatrix1.VisualRowCount.ToString())

            oDBs_Details2.SetValue("U_FIPTH", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_FINM", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_UPBY", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_AType", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_UPDATE", oDBs_Details2.Offset, "")
            'oDBs_Details2.SetValue("U_REMK", oDBs_Details2.Offset, "")

            ' FIXED ISSUE HERE
            objMatrix1.SetLineData(objMatrix1.VisualRowCount)

            objMatrix1.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

    Sub SetDefault(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_QASMPH")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_QASMPL")

            'oDBs_Head.SetValue("DocNum",
            '                   0,
            '                   objMain.objUtilities.GetNextDocNum(objForm,
            '                                                      "TNX_QA_SAMPLE",
            '                                                      "Primary"))

            objForm.DataBrowser.BrowseBy = "DocNum"

            'aa

            Dim oRsDocNum As SAPbobsCOM.Recordset

            oRsDocNum = objMain.objCompany.GetBusinessObject(
            SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim Query1 As String =
            "SELECT IFNULL(MAX(""DocNum""),0)+1 AS ""DocNum"" FROM ""@TNX_QASMPH"""

            oRsDocNum.DoQuery(Query1)

            Dim nextDocNum As String =
            oRsDocNum.Fields.Item("DocNum").Value.ToString().Trim()

            oDBs_Head.SetValue("DocNum", 0, nextDocNum)

            objForm.Items.Item("DocNum").Specific.Value = nextDocNum
            'a







            oDBs_Head.SetValue("U_DATE", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))

            'Me.SetNewLine(objForm.UniqueID)
            'Me.SetNewLine2(objForm.UniqueID)
            'objMatrix.AutoResizeColumns()

            objForm.Freeze(False)

        Catch ex As Exception

            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

    Public Function AutoDocentryNumber()

        objForm = objMain.objApplication.Forms.Item(objForm.UniqueID)

        oDBs_Head = objForm.DataSources.DBDataSources.Item("@SBO_ITEMREVIEWS")

        Dim oRsDocNum As SAPbobsCOM.Recordset

        oRsDocNum = objMain.objCompany.GetBusinessObject(
            SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim Query1 As String =
            "SELECT IFNULL(MAX(""DocNum""),0)+1 AS ""DocNum"" FROM ""@SBO_ITEMREVIEWS"""

        oRsDocNum.DoQuery(Query1)

        Dim nextDocNum As String =
            oRsDocNum.Fields.Item("DocNum").Value.ToString().Trim()

        oDBs_Head.SetValue("DocNum", 0, nextDocNum)

        objForm.Items.Item("DocNum").Specific.Value = nextDocNum

        Dim rsAppId As SAPbobsCOM.Recordset

        rsAppId = objMain.objCompany.GetBusinessObject(
            SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim str As String =
            "SELECT 'SP' || LPAD('',6 - LENGTH(LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_APPID""), '0'), 'SP', '') + 1))),'0') || LTRIM(RTRIM(REPLACE(IFNULL(MAX(""U_APPID""), '0'), 'SP', '') + 1)) FROM ""@SBO_ITEMREVIEWS"";"

        rsAppId.DoQuery(str)

        Dim appp As String =
            rsAppId.Fields.Item(0).Value.ToString().Trim()

        oDBs_Head.SetValue("U_APPID", oDBs_Head.Offset, appp)

        Return True

    End Function

End Class
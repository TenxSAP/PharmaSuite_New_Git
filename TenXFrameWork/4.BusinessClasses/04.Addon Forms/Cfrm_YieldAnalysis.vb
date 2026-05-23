Imports System
Imports System.Configuration
    Imports System.Net
    Imports System.Net.Mail
    Imports System.Net.Security
    Imports System.Security.Cryptography.X509Certificates
    Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
    Imports Newtonsoft.Json.Linq

    Public Class Cfrm_YieldAnalysis

#Region "Declaration"

        Public objForm As SAPbouiCOM.Form
        Dim oDBs_Head, oDBS1, oDBS2 As SAPbouiCOM.DBDataSource
        Dim oDBs_Details1, oDBs_Details2, oDBs_Details3, oDBs_Details4, oDBs_Details5 As SAPbouiCOM.DBDataSource
        Dim oDBs_Details As SAPbouiCOM.DBDataSource
        Dim objMatrix, objMatrix1, objMatrix2, objMatrix3, objMatrix4, objMatrix5 As SAPbouiCOM.Matrix
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

                objMain.objUtilities.LoadForm("YieldAnalysis.xml", "TNXPYLD", ResourceType.Embeded)

                objForm = objMain.objApplication.Forms.GetForm("TNXPYLD",
                                                               objMain.objApplication.Forms.ActiveForm.TypeCount)

                objForm.Freeze(True)

                objutilities = New Utilities

                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_H")
                oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_MAT")
                oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_OUT")
                oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_VAR")
                oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_APR")
                oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_ATT")

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
            'objMatrix1 = objForm.Items.Item("Mtx1").Specific

            objForm.Freeze(False)


                Me.SetNewLine(objForm.UniqueID)
                Me.SetNewLine2(objForm.UniqueID)
            Me.SetNewLine3(objForm.UniqueID)
            Me.SetNewLine4(objForm.UniqueID)
            Me.SetNewLine5(objForm.UniqueID)

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

                If pVal.MenuUID = "10X_YIELD" And pVal.BeforeAction = False Then

                    Me.CreateForm()

                ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

                objMatrix = objForm.Items.Item("Mtx").Specific

                Me.SetDefault(objForm.UniqueID)

                    Me.SetNewLine(objForm.UniqueID)
                    Me.SetNewLine2(objForm.UniqueID)
                    Me.SetNewLine3(objForm.UniqueID)
                    Me.SetNewLine4(objForm.UniqueID)
                    Me.SetNewLine5(objForm.UniqueID)



                ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                objMatrix = objForm.Items.Item("Mtx").Specific
                objMatrix1 = CType(objForm.Items.Item("MOTX").Specific, SAPbouiCOM.Matrix)
                objMatrix2 = CType(objForm.Items.Item("VMTX").Specific,
                       SAPbouiCOM.Matrix)
                objMatrix3 = CType(objForm.Items.Item("AMTX").Specific, SAPbouiCOM.Matrix)
                objMatrix5 = CType(objForm.Items.Item("ATTX").Specific, SAPbouiCOM.Matrix)

                Me.SetNewLine(objForm.UniqueID)
                    Me.SetNewLine2(objForm.UniqueID)
                    Me.SetNewLine3(objForm.UniqueID)
                    Me.SetNewLine4(objForm.UniqueID)
                    Me.SetNewLine5(objForm.UniqueID)



                ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then

                    objForm = objMain.objApplication.Forms.GetForm("TNXPYLD",
                                                                   objMain.objApplication.Forms.ActiveForm.TypeCount)

                    objMatrix = objForm.Items.Item("Mtx").Specific

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

                    ElseIf MATRIXS.Equals("MOTX") = True Then

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
                        Me.SetNewLine3(objForm.UniqueID)
                        Me.SetNewLine4(objForm.UniqueID)
                        Me.SetNewLine5(objForm.UniqueID)

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
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_H")
                oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_MAT")
                oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_OUT")
                oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_VAR")
                oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_APR")
                oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_ATT")
                objMatrix = objForm.Items.Item("Mtx").Specific

                objMatrix.AddRow()
                oDBs_Details1.SetValue("LineId",
                                       oDBs_Details1.Offset,
                                       objMatrix.VisualRowCount.ToString())
            'oDBs_Details1.SetValue("U_TESTCD", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_CMPCD", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_COMNM", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_ITMTY", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_BATNO", oDBs_Details1.Offset, "")

                oDBs_Details1.SetValue("U_BASEQTY", oDBs_Details1.Offset, 0)
                oDBs_Details1.SetValue("U_PLANQT", oDBs_Details1.Offset, 0)
                oDBs_Details1.SetValue("U_ISSQT", oDBs_Details1.Offset, 0)
                oDBs_Details1.SetValue("U_RETQT", oDBs_Details1.Offset, 0)
                oDBs_Details1.SetValue("U_CONQT", oDBs_Details1.Offset, 0)
                oDBs_Details1.SetValue("U_ACLQT", oDBs_Details1.Offset, 0)
                oDBs_Details1.SetValue("U_VAQT", oDBs_Details1.Offset, 0)

                oDBs_Details1.SetValue("U_STLOSP", oDBs_Details1.Offset, 0)
                oDBs_Details1.SetValue("U_VARP", oDBs_Details1.Offset, 0)
                oDBs_Details1.SetValue("U_ALLTO", oDBs_Details1.Offset, 0)

                oDBs_Details1.SetValue("U_Status", oDBs_Details1.Offset, "")
                oDBs_Details1.SetValue("U_RECOD", oDBs_Details1.Offset, "")
                objMatrix.SetLineData(objMatrix.VisualRowCount)
                objMatrix.AutoResizeColumns()
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try

        End Sub


        Sub SetNewLine2(ByVal FormUID As String)

            Try

                objForm = objMain.objApplication.Forms.Item(FormUID)

                oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_OUT")

                objMatrix1 = CType(objForm.Items.Item("MOTX").Specific,
                                   SAPbouiCOM.Matrix)

                objMatrix1.AddRow()

            oDBs_Details2.SetValue("LineId",
                                       oDBs_Details2.Offset,
                                       objMatrix1.VisualRowCount.ToString())
            oDBs_Details2.SetValue("U_REFRM", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_REPN", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_RELI", oDBs_Details2.Offset, "")

                oDBs_Details2.SetValue("U_REPDT", oDBs_Details2.Offset, "")

                oDBs_Details2.SetValue("U_ItemCode", oDBs_Details2.Offset, "")
                oDBs_Details2.SetValue("U_BatchNo", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_WhsCode", oDBs_Details2.Offset, "")

            oDBs_Details2.SetValue("U_REQTY", oDBs_Details2.Offset, 0)
                oDBs_Details2.SetValue("U_QCQty", oDBs_Details2.Offset, 0)
                oDBs_Details2.SetValue("U_ACCPT", oDBs_Details2.Offset, 0)

                oDBs_Details2.SetValue("U_UOM", oDBs_Details2.Offset, "")

                objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                objMatrix1.AutoResizeColumns()

            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try

        End Sub


        Sub SetNewLine3(ByVal FormUID As String)

            Try

                objForm = objMain.objApplication.Forms.Item(FormUID)

                oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_VAR")

                objMatrix2 = CType(objForm.Items.Item("VMTX").Specific,
                                   SAPbouiCOM.Matrix)

                objMatrix2.AddRow()

            oDBs_Details3.SetValue("LineId",
                                       oDBs_Details3.Offset,
                                       objMatrix2.VisualRowCount.ToString())
            oDBs_Details3.SetValue("U_VAR_TYPE", oDBs_Details1.Offset, "")
                oDBs_Details3.SetValue("U_EXPQTY", oDBs_Details1.Offset, 0)
                oDBs_Details3.SetValue("U_ACTQTY", oDBs_Details1.Offset, 0)
                oDBs_Details3.SetValue("U_VARQTY", oDBs_Details1.Offset, 0)
                oDBs_Details3.SetValue("U_VAPCT", oDBs_Details1.Offset, 0)
                oDBs_Details3.SetValue("U_TOPCT", oDBs_Details1.Offset, 0)

                oDBs_Details3.SetValue("U_IMPACT", oDBs_Details1.Offset, "")
                oDBs_Details3.SetValue("U_REAE", oDBs_Details1.Offset, "")
                oDBs_Details3.SetValue("U_REESC", oDBs_Details1.Offset, "")

                oDBs_Details3.SetValue("U_ACED", oDBs_Details1.Offset, "")
                oDBs_Details3.SetValue("U_CARED", oDBs_Details1.Offset, "")
                oDBs_Details3.SetValue("U_DIRED", oDBs_Details1.Offset, "")

                objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                objMatrix2.AutoResizeColumns()

            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try

        End Sub


        Sub SetNewLine4(ByVal FormUID As String)

            Try

                objForm = objMain.objApplication.Forms.Item(FormUID)
                oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_APR")
                objMatrix3 = CType(objForm.Items.Item("AMTX").Specific,
                                   SAPbouiCOM.Matrix)
                objMatrix3.AddRow()
            oDBs_Details4.SetValue("LineId",
                                       oDBs_Details4.Offset,
                                       objMatrix3.VisualRowCount.ToString())
            oDBs_Details4.SetValue("U_STAGE", oDBs_Details1.Offset, "")
                oDBs_Details4.SetValue("U_APER", oDBs_Details1.Offset, "")
                oDBs_Details4.SetValue("U_APPS", oDBs_Details1.Offset, "")

                oDBs_Details4.SetValue("U_ATPE", oDBs_Details1.Offset, "")
                oDBs_Details4.SetValue("U_AIME", oDBs_Details1.Offset, "")

                oDBs_Details4.SetValue("U_RMKS", oDBs_Details1.Offset, "")
                objMatrix3.SetLineData(objMatrix3.VisualRowCount)

                objMatrix3.AutoResizeColumns()

            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try

        End Sub
        Sub SetNewLine5(ByVal FormUID As String)

            Try

                objForm = objMain.objApplication.Forms.Item(FormUID)
                oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_ATT")
                objMatrix5 = CType(objForm.Items.Item("ATTX").Specific,
                                   SAPbouiCOM.Matrix)
                objMatrix5.AddRow()
            oDBs_Details5.SetValue("LineId",
                                       oDBs_Details5.Offset,
                                       objMatrix5.VisualRowCount.ToString())
            oDBs_Details5.SetValue("U_FILEE", oDBs_Details1.Offset, "")
                oDBs_Details5.SetValue("U_FITH", oDBs_Details1.Offset, "")
                oDBs_Details5.SetValue("U_FIYPE", oDBs_Details1.Offset, "")
                oDBs_Details5.SetValue("U_ATTBY", oDBs_Details1.Offset, "")

                oDBs_Details5.SetValue("U_AE", oDBs_Details1.Offset, "")

                oDBs_Details5.SetValue("U_REMS", oDBs_Details1.Offset, "")
                objMatrix5.SetLineData(objMatrix5.VisualRowCount)

                objMatrix5.AutoResizeColumns()

            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(ex.Message)

            End Try

        End Sub

        Sub SetDefault(ByVal FormUID As String)

            Try

                objForm = objMain.objApplication.Forms.Item(FormUID)

                objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PYLD_H")

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
                "SELECT IFNULL(MAX(""DocNum""),0)+1 AS ""DocNum"" FROM ""@TNX_PYLD_H"""

            oRsDocNum.DoQuery(Query1)

                Dim nextDocNum As String =
                oRsDocNum.Fields.Item("DocNum").Value.ToString().Trim()

                oDBs_Head.SetValue("DocNum", 0, nextDocNum)

                objForm.Items.Item("DocNum").Specific.Value = nextDocNum
            'a







            oDBs_Head.SetValue("U_DocDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))

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

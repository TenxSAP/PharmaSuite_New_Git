Imports System.Configuration
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports Newtonsoft.Json.Linq

Public Class Cls_ShelfLifeAnalysis

#Region "Declaration"

    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBS1, oDBS2 As SAPbouiCOM.DBDataSource
    Dim oDBs_Details1, oDBs_Details2, oDBs_Details3 As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1, objMatrix2, objMatrix3 As SAPbouiCOM.Matrix
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

            objMain.objUtilities.LoadForm("ShelfLifeAnalysis.xml", "STAB_SHELF", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("STAB_SHELF",
                                                               objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_STAB_SHELF")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_STAB_SHELF_L")
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

            objMatrix = objForm.Items.Item("0_U_G").Specific
            'objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objForm.Freeze(False)


            Me.SetNewLineShelf(objForm.UniqueID)
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

            If pVal.MenuUID = "10X_STB_SHELFL" And pVal.BeforeAction = False Then
                Me.CreateForm()

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                objMatrix = objForm.Items.Item("0_U_G").Specific
                'objMatrix1 = objForm.Items.Item("Mtx1").Specific
                Me.SetDefault(objForm.UniqueID)

                Me.SetNewLineShelf(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

                objMatrix2 = CType(objForm.Items.Item("Item_3").Specific,
                                   SAPbouiCOM.Matrix)
                objMatrix1 = objForm.Items.Item("0_U_G").Specific

                Me.SetNewLineShelf(objForm.UniqueID)

            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.GetForm("STAB_SHELF",
                                                                   objMain.objApplication.Forms.ActiveForm.TypeCount)

                ' objMatrix = objForm.Items.Item("Mtx").Specific
                '  objMatrix1 = objForm.Items.Item("Mtx1").Specific

                Dim row As Integer = objMatrix.VisualRowCount

                If MATRIXS.Equals("0_U_G") = True Then

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



    Sub ItemEvent(ByVal FormUID As String,
                      ByRef pVal As SAPbouiCOM.ItemEvent,
                      ByRef BubbleEvent As Boolean)

        Try

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_CLICK

                    If (pVal.ItemUID = "0_U_G") And
                           pVal.BeforeAction = False Then

                        MATRIXS = pVal.ItemUID

                    End If

                    If (pVal.ItemUID = "Item_3") And
                           pVal.BeforeAction = False Then

                        MATRIXS = pVal.ItemUID

                    End If


                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = False Then
                        Me.SetDefault(objForm.UniqueID)
                        Me.SetNewLineShelf(objForm.UniqueID)

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



    Sub SetNewLineShelf(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_STAB_SHELF_L")

            objMatrix1 = CType(objForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)

            objMatrix1.AddRow()

            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount.ToString())
            oDBs_Details1.SetValue("U_TestCode", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_TimePoint", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_RESLT", oDBs_Details1.Offset, "0")
            oDBs_Details1.SetValue("U_TRENDS", oDBs_Details1.Offset, "0")
            oDBs_Details1.SetValue("U_RISK", oDBs_Details1.Offset, "")

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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_STAB_SHELF")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_STAB_SHELF_L")
            objForm.EnableMenu("1292", True)
            objForm.DataBrowser.BrowseBy = "DocNum"
            Dim oRsDocNum As SAPbobsCOM.Recordset

            oRsDocNum = objMain.objCompany.GetBusinessObject(
            SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim Query1 As String =
            "SELECT IFNULL(MAX(""DocNum""),0)+1 AS ""DocNum"" FROM ""@TNX_STAB_SHELF"""
            oRsDocNum.DoQuery(Query1)
            Dim nextDocNum As String =
            oRsDocNum.Fields.Item("DocNum").Value.ToString().Trim()
            oDBs_Head.SetValue("DocNum", 0, nextDocNum)
            objForm.Items.Item("DocNum").Specific.Value = nextDocNum
            ' oDBs_Head.SetValue("U_StartDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))

            'Me.SetNewLine(objForm.UniqueID)
            'Me.SetNewLine2(objForm.UniqueID)
            'objMatrix.AutoResizeColumns()

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub



End Class



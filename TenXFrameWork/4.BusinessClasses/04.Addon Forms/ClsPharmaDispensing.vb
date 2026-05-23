
Imports System

Public Class ClsPharmaDispensing

#Region "       Declaration             "
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1, oDBs_Details2, oDBs_Details3, oDBs_Details4, oDBs_Details5, oDBs_Details6 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, objMatrix2, objMatrix3, objMatrix4, objMatrix5, objMatrix6 As SAPbouiCOM.Matrix
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
            objMain.objUtilities.LoadForm("PharmaDispensing.xml", "10X_DISP", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("10X_DISP", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)
            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PDSP_H")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PDSP_L")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@@TNX_PDSP_BT")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@@TNX_PDSP_APR")
            oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@@TNX_PDSP_ATT")

            objMatrix1 = objForm.Items.Item("MTX_1").Specific
            objMatrix2 = objForm.Items.Item("MTX_2").Specific
            objMatrix3 = objForm.Items.Item("MTX_3").Specific
            objMatrix4 = objForm.Items.Item("MTX_4").Specific
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_PDSP"))
            oDBs_Head.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"))
            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("1293", True)
            objForm.Freeze(False)
            Me.SetDefault(objForm.UniqueID)
            objMain.objApplication.StatusBar.SetText("Successfully initialized, Please proceed...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "10X_DISP" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                objMatrix1 = objForm.Items.Item("MTX_1").Specific
                objMatrix2 = objForm.Items.Item("MTX_2").Specific
                objMatrix3 = objForm.Items.Item("MTX_3").Specific
                objMatrix4 = objForm.Items.Item("MTX_4").Specific
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                objMatrix1 = objForm.Items.Item("MTX_1").Specific
                objMatrix2 = objForm.Items.Item("MTX_2").Specific
                objMatrix3 = objForm.Items.Item("MTX_3").Specific
                objMatrix4 = objForm.Items.Item("MTX_4").Specific
            ElseIf pVal.MenuUID = "1281" And pVal.BeforeAction = False Then

            End If
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Function Validate() As Boolean
        Dim CustomerCode As SAPbouiCOM.Matrix
        Try
            If oDBs_Head.GetValue("U_FormulaCode", 0) = "" Then
                objMain.objApplication.SetStatusBarMessage("Formula Code is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)                'Me.FormText(enControlName.Financeyear).Active = True
                Return False
                Exit Function
            End If
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message & "Errors in Validation Function", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Function

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            'Select Case pVal.EventType

            '    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
            '        objForm = objMain.objApplication.Forms.Item(FormUID)
            '        'objForm.Freeze(True)
            '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_ACNGW")
            '        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_ACNGW_C0")
            '        objMatrix = objForm.Items.Item("Matrix").Specific
            '        Dim oCFL As SAPbouiCOM.ChooseFromList
            '        Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal        'CHOOSE FORM LIST'
            '        Dim CFL_Id As String
            '        CFL_Id = CFLEvent.ChooseFromListUID
            '        oCFL = objForm.ChooseFromLists.Item(CFL_Id)
            '        Dim oDT As SAPbouiCOM.DataTable
            '        oDT = CFLEvent.SelectedObjects
            '        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
            '        If oCFL.UniqueID = "CFL_1" And pVal.BeforeAction = True Then 'CHOOSE FORM TABLE UNIQUNE ID
            '            CFLFilter(FormUID, "CFL_1")
            '        End If

            '        If oCFL.UniqueID = "CFL_4" And pVal.BeforeAction = True Then 'CHOOSE FORM TABLE UNIQUNE ID
            '            Dim StrCFLQry As String = ""
            '            StrCFLQry = "Select ""OcrCode"", ""OcrName"" FROM ""OOCR"" WHERE ""DimCode"" = 1 And ""Active"" = 'Y'"
            '            CFLFilter1(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "OcrCode")
            '        End If
            '        If oCFL.UniqueID = "CFL_5" And pVal.BeforeAction = True Then 'CHOOSE FORM TABLE UNIQUNE ID
            '            Dim StrCFLQry As String = ""
            '            StrCFLQry = "Select ""OcrCode"", ""OcrName"" FROM ""OOCR"" WHERE ""DimCode"" = 2 And ""Active"" = 'Y'"
            '            CFLFilter2(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "OcrCode")
            '        End If
            '        If oCFL.UniqueID = "CFL_6" And pVal.BeforeAction = True Then 'CHOOSE FORM TABLE UNIQUNE ID
            '            Dim StrCFLQry As String = ""
            '            StrCFLQry = "Select ""OcrCode"", ""OcrName"" FROM ""OOCR"" WHERE ""DimCode"" = 3 And ""Active"" = 'Y'"
            '            CFLFilter3(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "OcrCode")
            '        End If
            '        If oCFL.UniqueID = "CFL_7" And pVal.BeforeAction = True Then 'CHOOSE FORM TABLE UNIQUNE ID
            '            Dim StrCFLQry As String = ""
            '            StrCFLQry = "Select ""OcrCode"", ""OcrName"" FROM ""OOCR"" WHERE ""DimCode"" = 4 And ""Active"" = 'Y'"
            '            CFLFilter4(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "OcrCode")
            '        End If
            '        If oCFL.UniqueID = "CFL_8" And pVal.BeforeAction = True Then 'CHOOSE FORM TABLE UNIQUNE ID
            '            Dim StrCFLQry As String = ""
            '            StrCFLQry = "Select ""OcrCode"", ""OcrName"" FROM ""OOCR"" WHERE ""DimCode"" = 5 And ""Active"" = 'Y'"
            '            CFLFilter5(pVal, objMain.objApplication, objMain.objCompany, StrCFLQry, "OcrCode")
            '        End If
            '        If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
            '            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            '            objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
            '            If oCFL.UniqueID = "CFL_1" Then
            '                oDBs_Head.SetValue("U_TNXCODE", oDBs_Head.Offset, oDT.GetValue("CardCode", 0))
            '                oDBs_Head.SetValue("U_TNXCNM", oDBs_Head.Offset, oDT.GetValue("CardName", 0))   '(U_CN DESGIN FORM CUSTOMER NAME ALIS NAME)
            '            End If
            '            If oCFL.UniqueID = "CFL_2" Then
            '                oDBs_Head.SetValue("U_TNXIC", oDBs_Head.Offset, oDT.GetValue("ItemCode", 0))
            '            End If
            '            If oCFL.UniqueID = "CFL_3" Then
            '                oDBs_Details.SetValue("LineID", oDBs_Details.Offset, pVal.Row)
            '                oDBs_Details.SetValue("U_TNXICD", oDBs_Details.Offset, oDT.GetValue("ItemCode", 0))
            '                oDBs_Details.SetValue("U_TNXINM", oDBs_Details.Offset, oDT.GetValue("ItemName", 0))
            '                oDBs_Details.SetValue("U_TNXINVN", oDBs_Details.Offset, objMatrix.Columns.Item("TNXINVN").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXIENT", oDBs_Details.Offset, objMatrix.Columns.Item("TNXIENT").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXQTY", oDBs_Details.Offset, objMatrix.Columns.Item("TNXQTY").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXUPRI", oDBs_Details.Offset, objMatrix.Columns.Item("TNXUPRI").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXGT", oDBs_Details.Offset, objMatrix.Columns.Item("TNXGT").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXDIS", oDBs_Details.Offset, objMatrix.Columns.Item("TNXDIS").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXNTL", oDBs_Details.Offset, objMatrix.Columns.Item("TNXNTL").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXLNID", oDBs_Details.Offset, objMatrix.Columns.Item("TNXLNID").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXIND", oDBs_Details.Offset, objMatrix.Columns.Item("TNXIND").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXDISA", oDBs_Details.Offset, objMatrix.Columns.Item("TNXDISA").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode2", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode2").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode3", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode3").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode4", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode4").Cells.Item(pVal.Row).Specific.Value)
            '                ' oDBs_Details.SetValue("U_OcrCode5", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode5").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_LTSNO", oDBs_Details.Offset, objMatrix.Columns.Item("LTSNO").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_CostCenter", oDBs_Details.Offset, objMatrix.Columns.Item("CostCenter").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_GMANo", oDBs_Details.Offset, objMatrix.Columns.Item("GMANo").Cells.Item(pVal.Row).Specific.Value)
            '                objMatrix.SetLineData(pVal.Row)
            '                ' Me.SetNewLine(objForm.UniqueID)
            '            End If
            '            If oCFL.UniqueID = "CFL_4" Then
            '                oDBs_Details.SetValue("LineID", oDBs_Details.Offset, pVal.Row)
            '                oDBs_Details.SetValue("U_TNXICD", oDBs_Details.Offset, objMatrix.Columns.Item("TNXICD").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXINM", oDBs_Details.Offset, objMatrix.Columns.Item("TNXINM").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXINVN", oDBs_Details.Offset, objMatrix.Columns.Item("TNXINVN").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXIENT", oDBs_Details.Offset, objMatrix.Columns.Item("TNXIENT").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXQTY", oDBs_Details.Offset, objMatrix.Columns.Item("TNXQTY").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXUPRI", oDBs_Details.Offset, objMatrix.Columns.Item("TNXUPRI").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXGT", oDBs_Details.Offset, objMatrix.Columns.Item("TNXGT").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXDIS", oDBs_Details.Offset, objMatrix.Columns.Item("TNXDIS").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXNTL", oDBs_Details.Offset, objMatrix.Columns.Item("TNXNTL").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXLNID", oDBs_Details.Offset, objMatrix.Columns.Item("TNXLNID").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXIND", oDBs_Details.Offset, objMatrix.Columns.Item("TNXIND").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXDISA", oDBs_Details.Offset, objMatrix.Columns.Item("TNXDISA").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode", oDBs_Details.Offset, oDT.GetValue("OcrCode", 0))
            '                oDBs_Details.SetValue("U_OcrCode2", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode2").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode3", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode3").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode4", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode4").Cells.Item(pVal.Row).Specific.Value)
            '                '  oDBs_Details.SetValue("U_OcrCode5", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode5").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_LTSNO", oDBs_Details.Offset, objMatrix.Columns.Item("LTSNO").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_CostCenter", oDBs_Details.Offset, objMatrix.Columns.Item("CostCenter").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_GMANo", oDBs_Details.Offset, objMatrix.Columns.Item("GMANo").Cells.Item(pVal.Row).Specific.Value)
            '                objMatrix.SetLineData(pVal.Row)
            '                ' Me.SetNewLine(objForm.UniqueID)
            '            End If
            '            If oCFL.UniqueID = "CFL_5" Then
            '                oDBs_Details.SetValue("LineID", oDBs_Details.Offset, pVal.Row)
            '                oDBs_Details.SetValue("U_TNXICD", oDBs_Details.Offset, objMatrix.Columns.Item("TNXICD").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXINM", oDBs_Details.Offset, objMatrix.Columns.Item("TNXINM").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXINVN", oDBs_Details.Offset, objMatrix.Columns.Item("TNXINVN").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXIENT", oDBs_Details.Offset, objMatrix.Columns.Item("TNXIENT").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXQTY", oDBs_Details.Offset, objMatrix.Columns.Item("TNXQTY").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXUPRI", oDBs_Details.Offset, objMatrix.Columns.Item("TNXUPRI").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXGT", oDBs_Details.Offset, objMatrix.Columns.Item("TNXGT").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXDIS", oDBs_Details.Offset, objMatrix.Columns.Item("TNXDIS").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXNTL", oDBs_Details.Offset, objMatrix.Columns.Item("TNXNTL").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXLNID", oDBs_Details.Offset, objMatrix.Columns.Item("TNXLNID").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXIND", oDBs_Details.Offset, objMatrix.Columns.Item("TNXIND").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXDISA", oDBs_Details.Offset, objMatrix.Columns.Item("TNXDISA").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode2", oDBs_Details.Offset, oDT.GetValue("OcrCode", 0))
            '                oDBs_Details.SetValue("U_OcrCode3", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode3").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode4", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode4").Cells.Item(pVal.Row).Specific.Value)
            '                ' oDBs_Details.SetValue("U_OcrCode5", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode5").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_LTSNO", oDBs_Details.Offset, objMatrix.Columns.Item("LTSNO").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_CostCenter", oDBs_Details.Offset, objMatrix.Columns.Item("CostCenter").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_GMANo", oDBs_Details.Offset, objMatrix.Columns.Item("GMANo").Cells.Item(pVal.Row).Specific.Value)
            '                objMatrix.SetLineData(pVal.Row)
            '                ' Me.SetNewLine(objForm.UniqueID)
            '            End If
            '            If oCFL.UniqueID = "CFL_6" Then
            '                oDBs_Details.SetValue("LineID", oDBs_Details.Offset, pVal.Row)
            '                oDBs_Details.SetValue("U_TNXICD", oDBs_Details.Offset, objMatrix.Columns.Item("TNXICD").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXINM", oDBs_Details.Offset, objMatrix.Columns.Item("TNXINM").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXINVN", oDBs_Details.Offset, objMatrix.Columns.Item("TNXINVN").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXIENT", oDBs_Details.Offset, objMatrix.Columns.Item("TNXIENT").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXQTY", oDBs_Details.Offset, objMatrix.Columns.Item("TNXQTY").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXUPRI", oDBs_Details.Offset, objMatrix.Columns.Item("TNXUPRI").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXGT", oDBs_Details.Offset, objMatrix.Columns.Item("TNXGT").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXDIS", oDBs_Details.Offset, objMatrix.Columns.Item("TNXDIS").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXNTL", oDBs_Details.Offset, objMatrix.Columns.Item("TNXNTL").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXLNID", oDBs_Details.Offset, objMatrix.Columns.Item("TNXLNID").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXIND", oDBs_Details.Offset, objMatrix.Columns.Item("TNXIND").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXDISA", oDBs_Details.Offset, objMatrix.Columns.Item("TNXDISA").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode2", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode2").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode3", oDBs_Details.Offset, oDT.GetValue("OcrCode", 0))
            '                oDBs_Details.SetValue("U_OcrCode4", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode4").Cells.Item(pVal.Row).Specific.Value)
            '                ' oDBs_Details.SetValue("U_OcrCode5", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode5").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_LTSNO", oDBs_Details.Offset, objMatrix.Columns.Item("LTSNO").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_CostCenter", oDBs_Details.Offset, objMatrix.Columns.Item("CostCenter").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_GMANo", oDBs_Details.Offset, objMatrix.Columns.Item("GMANo").Cells.Item(pVal.Row).Specific.Value)
            '                objMatrix.SetLineData(pVal.Row)
            '                ' Me.SetNewLine(objForm.UniqueID)
            '            End If
            '            If oCFL.UniqueID = "CFL_7" Then
            '                oDBs_Details.SetValue("LineID", oDBs_Details.Offset, pVal.Row)
            '                oDBs_Details.SetValue("U_TNXICD", oDBs_Details.Offset, objMatrix.Columns.Item("TNXICD").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXINM", oDBs_Details.Offset, objMatrix.Columns.Item("TNXINM").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXINVN", oDBs_Details.Offset, objMatrix.Columns.Item("TNXINVN").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXIENT", oDBs_Details.Offset, objMatrix.Columns.Item("TNXIENT").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXQTY", oDBs_Details.Offset, objMatrix.Columns.Item("TNXQTY").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXUPRI", oDBs_Details.Offset, objMatrix.Columns.Item("TNXUPRI").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXGT", oDBs_Details.Offset, objMatrix.Columns.Item("TNXGT").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXDIS", oDBs_Details.Offset, objMatrix.Columns.Item("TNXDIS").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXNTL", oDBs_Details.Offset, objMatrix.Columns.Item("TNXNTL").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXLNID", oDBs_Details.Offset, objMatrix.Columns.Item("TNXLNID").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXIND", oDBs_Details.Offset, objMatrix.Columns.Item("TNXIND").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_TNXDISA", oDBs_Details.Offset, objMatrix.Columns.Item("TNXDISA").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode2", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode2").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode3", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode3").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_OcrCode4", oDBs_Details.Offset, oDT.GetValue("OcrCode", 0))
            '                ' oDBs_Details.SetValue("U_OcrCode5", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode5").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_LTSNO", oDBs_Details.Offset, objMatrix.Columns.Item("LTSNO").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_CostCenter", oDBs_Details.Offset, objMatrix.Columns.Item("CostCenter").Cells.Item(pVal.Row).Specific.Value)
            '                oDBs_Details.SetValue("U_GMANo", oDBs_Details.Offset, objMatrix.Columns.Item("GMANo").Cells.Item(pVal.Row).Specific.Value)
            '                objMatrix.SetLineData(pVal.Row)
            '                ' Me.SetNewLine(objForm.UniqueID)
            '            End If
            '            'If oCFL.UniqueID = "CFL_8" Then
            '            '    oDBs_Details.SetValue("LineID", oDBs_Details.Offset, pVal.Row)
            '            '    oDBs_Details.SetValue("U_TNXICD", oDBs_Details.Offset, objMatrix.Columns.Item("TNXICD").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_TNXINM", oDBs_Details.Offset, objMatrix.Columns.Item("TNXINM").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_TNXINVN", oDBs_Details.Offset, objMatrix.Columns.Item("TNXINVN").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_TNXIENT", oDBs_Details.Offset, objMatrix.Columns.Item("TNXIENT").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_TNXQTY", oDBs_Details.Offset, objMatrix.Columns.Item("TNXQTY").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_TNXUPRI", oDBs_Details.Offset, objMatrix.Columns.Item("TNXUPRI").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_TNXGT", oDBs_Details.Offset, objMatrix.Columns.Item("TNXGT").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_TNXDIS", oDBs_Details.Offset, objMatrix.Columns.Item("TNXDIS").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_TNXNTL", oDBs_Details.Offset, objMatrix.Columns.Item("TNXNTL").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_TNXLNID", oDBs_Details.Offset, objMatrix.Columns.Item("TNXLNID").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_TNXIND", oDBs_Details.Offset, objMatrix.Columns.Item("TNXIND").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_TNXDISA", oDBs_Details.Offset, objMatrix.Columns.Item("TNXDISA").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_OcrCode", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_OcrCode2", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode2").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_OcrCode3", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode3").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_OcrCode4", oDBs_Details.Offset, objMatrix.Columns.Item("OcrCode4").Cells.Item(pVal.Row).Specific.Value)
            '            '    '  oDBs_Details.SetValue("U_OcrCode5", oDBs_Details.Offset, oDT.GetValue("OcrCode", 0))
            '            '    oDBs_Details.SetValue("U_LTSNO", oDBs_Details.Offset, objMatrix.Columns.Item("U_LTSNO").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_CostCenter", oDBs_Details.Offset, objMatrix.Columns.Item("U_CostCenter").Cells.Item(pVal.Row).Specific.Value)
            '            '    oDBs_Details.SetValue("U_GMANo", oDBs_Details.Offset, objMatrix.Columns.Item("U_GMANo").Cells.Item(pVal.Row).Specific.Value)
            '            '    ' objMatrix.SetLineData(pVal.Row)
            '            '    ' Me.SetNewLine(objForm.UniqueID)
            '            'End If
            '        End If
            '    'Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
            '    '    If pVal.FormTypeEx = "ACNGW" Then
            '    '        ' Release references
            '    '        objForm = Nothing
            '    '    End If
            '    '    If objMain.objApplication.Forms.Count > 0 Then
            '    '        Try
            '    '            Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
            '    '            If objForm IsNot Nothing AndAlso objForm.TypeEx = "ACNGW" Then
            '    '                ' Now safely access form objects
            '    '            End If
            '    '        Catch ex As Exception
            '    '            objMain.objApplication.StatusBar.SetText("Form no longer available: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error.smt_Warning)
            '    '        End Try
            '    '    End If


            '    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
            '        objForm = objMain.objApplication.Forms.Item(FormUID)
            '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_ACNGW")
            '        oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_ACNGW_C0")
            '        objMatrix = objForm.Items.Item("Matrix").Specific
            '        If pVal.ItemUID = "Matrix" And pVal.ColUID = "TNXINVN" Then
            '            If objMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() <> "" Then
            '                If pVal.Row = objMatrix.VisualRowCount Then
            '                End If
            '            End If
            '        End If

            '        'If pVal.ItemUID = "Matrix" And pVal.ColUID = "TNXQTY" Or pVal.ColUID = "TNXUPRI" Then
            '        '    Dim qtyText As String = objMatrix.Columns.Item("TNXQTY").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()
            '        '    Dim priceText As String = objMatrix.Columns.Item("TNXUPRI").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()

            '        '    Dim Quantity As Double = 0
            '        '    Dim UnitPrice As Double = 0

            '        '    If Double.TryParse(qtyText, Quantity) AndAlso Double.TryParse(priceText, UnitPrice) Then
            '        '        Dim GrossTotal As Double = Quantity * UnitPrice
            '        '        objMatrix.Columns.Item("TNXGT").Cells.Item(pVal.Row).Specific.Value = GrossTotal.ToString("N2")
            '        '    Else
            '        '        ' Optionally clear if input is invalid
            '        '        objMatrix.Columns.Item("TNXGT").Cells.Item(pVal.Row).Specific.Value = "0.00"
            '        '    End If
            '        'End If

            '        If pVal.ItemUID = "Matrix" And pVal.ColUID = "TNXNTL" Or pVal.ColUID = "TNXIND" Then
            '            objForm.Freeze(True)
            '            If objMatrix.Columns.Item("TNXNTL").Cells.Item(pVal.Row).Specific.Value.ToString().Trim().Trim() <> "" And objMatrix.Columns.Item("TNXNTL").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() <> "" Then
            '                Dim NetTotal As Double = CDbl(objMatrix.Columns.Item("TNXNTL").Cells.Item(pVal.Row).Specific.Value)
            '                Dim InvoiceDiscount As Double = CDbl(objMatrix.Columns.Item("TNXIND").Cells.Item(pVal.Row).Specific.Value)
            '                Dim DiscountAmount As Double = (NetTotal * (InvoiceDiscount / 100))
            '                objMatrix.Columns.Item("TNXDISA").Cells.Item(pVal.Row).Specific.Value = DiscountAmount
            '                objMatrix.AutoResizeColumns()
            '            End If
            '            objForm.Freeze(False)
            '        End If

            '    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
            '        objForm = objMain.objApplication.Forms.Item(FormUID)
            '        'Dim ff As Integer
            '        'objForm = objMain.objApplication.Forms.Item(FormUID)
            '        If pVal.ItemUID = "AV" And pVal.BeforeAction = False And objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '            'objForm.Items.Item("1").Visible = True
            '            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '            'objForm.Items.Item("1").Visible = False
            '        End If
            '        If pVal.ItemUID = "Item_23" And pVal.BeforeAction = False Then
            '            Dim ff As Integer
            '            ff = objMain.objApplication.MessageBox("Do You want To Confirm!", 1, "Ok", "Cancel")

            '            If ff = 1 Then ' 1 = Ok, 2 = Cancel
            '                Me.PostARCreditMemo(objForm.UniqueID)
            '            End If
            '        End If

            '        If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And
            '        (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
            '            Dim oComboBox As SAPbouiCOM.ButtonCombo = objForm.Items.Item("AV").Specific

            '            Dim selectedOption As String = oComboBox.Selected.Description

            '            If selectedOption = "Add & New" Then
            '                Me.SetDefault(objForm.UniqueID)
            '                'objForm.Items.Item("AV").Visible = True
            '                'objForm.Items.Item("1").Visible = False
            '            Else
            '                Dim strQry As String = "Select Max(""DocNum"") From ""@TNX_ACNGW"" "
            '                Dim rs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            '                rs.DoQuery(strQry)
            '                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            '                objForm.Items.Item("DocNum").Specific.Value = rs.Fields.Item(0).Value
            '                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '                'objForm.Items.Item("AV").Visible = False
            '                'objForm.Items.Item("1").Visible = True
            '            End If
            '        End If
            '        If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
            '            If Me.Validate() = False Then
            '                BubbleEvent = False
            '            End If
            '        End If
            '        If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            '            objForm.Freeze(True)
            '            Try
            '                Dim rowCount As Integer = objMatrix.RowCount
            '                For i As Integer = 1 To rowCount
            '                    Dim netTotalStr As String = objMatrix.Columns.Item("TNXNTL").Cells.Item(i).Specific.Value.ToString().Trim()
            '                    Dim discountStr As String = objMatrix.Columns.Item("TNXIND").Cells.Item(i).Specific.Value.ToString().Trim()

            '                    If netTotalStr <> "" AndAlso discountStr <> "" Then
            '                        Dim NetTotal As Double = CDbl(netTotalStr)
            '                        Dim InvoiceDiscount As Double = CDbl(discountStr)
            '                        Dim DiscountAmount As Double = NetTotal * (InvoiceDiscount / 100)
            '                        objMatrix.Columns.Item("TNXDISA").Cells.Item(i).Specific.Value = DiscountAmount
            '                    End If
            '                Next
            '                objMatrix.AutoResizeColumns()
            '            Catch ex As Exception
            '                objMain.objApplication.MessageBox("Error during discount calculation: " & ex.Message)
            '            Finally
            '                objForm.Freeze(False)
            '            End Try
            '        End If

            '        'If pVal.ItemUID = "1" And pVal.BeforeAction = False And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
            '        '    pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
            '        '    LostFocusFlag = True
            '        '    If Me.Validate() = False Then
            '        '        LostFocusFlag = False
            '        '        BubbleEvent = False
            '        '    End If
            '        'End If

            '        If pVal.ItemUID = "LoadGC" And pVal.BeforeAction = False Then
            '            Me.LoadMatrixDetails(objForm.UniqueID)
            '            Dim discountStr As String = objForm.Items.Item("TNXDISC").Specific.Value
            '            ' objForm.Freeze(True)
            '            If discountStr <> "" Then
            '                Dim discountPercent As Double = CDbl(discountStr)
            '                Dim objMatrix As SAPbouiCOM.Matrix = objForm.Items.Item("Matrix").Specific
            '                For i As Integer = 1 To objMatrix.VisualRowCount
            '                    objMatrix.Columns.Item("TNXIND").Cells.Item(i).Specific.Value = discountPercent
            '                Next
            '                objMatrix.AutoResizeColumns()
            '            Else
            '                objMain.objApplication.MessageBox("Please enter a discount % in the header field before loading.")
            '            End If
            '        End If

            'End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub SetNewLine1(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oDBDS As SAPbouiCOM.DBDataSource
            Dim oMatrix As SAPbouiCOM.Matrix

            oDBDS = objForm.DataSources.DBDataSources.Item("@TNX_PDSP_L")
            oMatrix = objForm.Items.Item("0_U_G").Specific

            oMatrix.FlushToDataSource()

            oDBDS.InsertRecord(oDBDS.Size)
            oDBDS.Offset = oDBDS.Size - 1

            oDBDS.SetValue("LineId", oDBDS.Offset, (oDBDS.Offset + 1).ToString())
            oDBDS.SetValue("U_RMItemCd", oDBDS.Offset, "")
            oDBDS.SetValue("U_RMItemNm", oDBDS.Offset, "")
            oDBDS.SetValue("U_ItemGroup", oDBDS.Offset, "")
            oDBDS.SetValue("U_ReqQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_PrevDispQ", oDBDS.Offset, "")
            oDBDS.SetValue("U_BalQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_PlanDispQ", oDBDS.Offset, "")
            oDBDS.SetValue("U_WeightQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_UOM", oDBDS.Offset, "")
            oDBDS.SetValue("U_TolPct", oDBDS.Offset, "")
            oDBDS.SetValue("U_MinQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_MaxQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_WhsCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_BatchMng", oDBDS.Offset, "")
            oDBDS.SetValue("U_BatchNo", oDBDS.Offset, "")
            oDBDS.SetValue("U_QCStatus", oDBDS.Offset, "")
            oDBDS.SetValue("U_ExpDate", oDBDS.Offset, "")
            oDBDS.SetValue("U_ScaleID", oDBDS.Offset, "")
            oDBDS.SetValue("U_WeightBy", oDBDS.Offset, "")
            oDBDS.SetValue("U_WeightDT", oDBDS.Offset, "")
            oDBDS.SetValue("U_VerifyBy", oDBDS.Offset, "")
            oDBDS.SetValue("U_VerifyDT", oDBDS.Offset, "")
            oDBDS.SetValue("U_LineSt", oDBDS.Offset, "")
            oDBDS.SetValue("U_Remarks", oDBDS.Offset, "")
            oDBDS.SetValue("U_BaseLine", oDBDS.Offset, "")
            oDBDS.SetValue("U_RMItemCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_RMItemName", oDBDS.Offset, "")
            oDBDS.SetValue("U_RequiredQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_PreviouslyDispQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_BalanceQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_PlannedDispQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_WeighedQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_TolerancePct", oDBDS.Offset, "")
            oDBDS.SetValue("U_BatchManaged", oDBDS.Offset, "")
            oDBDS.SetValue("U_ExpiryDate", oDBDS.Offset, "")
            oDBDS.SetValue("U_WeighedBy", oDBDS.Offset, "")
            oDBDS.SetValue("U_WeighedDateTime", oDBDS.Offset, "")
            oDBDS.SetValue("U_VerifiedBy", oDBDS.Offset, "")
            oDBDS.SetValue("U_VerifiedDateTime", oDBDS.Offset, "")
            oDBDS.SetValue("U_LineStatus", oDBDS.Offset, "")

            oMatrix.LoadFromDataSource()
            oMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Private Sub SetNewLine2(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oDBDS As SAPbouiCOM.DBDataSource
            Dim oMatrix As SAPbouiCOM.Matrix

            oDBDS = objForm.DataSources.DBDataSources.Item("@TNX_PDSP_BT")
            oMatrix = objForm.Items.Item("MTX_2").Specific

            oMatrix.FlushToDataSource()

            oDBDS.InsertRecord(oDBDS.Size)
            oDBDS.Offset = oDBDS.Size - 1

            oDBDS.SetValue("LineId", oDBDS.Offset, (oDBDS.Offset + 1).ToString())
            oDBDS.SetValue("U_RMItemCd", oDBDS.Offset, "")
            oDBDS.SetValue("U_BatchNo", oDBDS.Offset, "")
            oDBDS.SetValue("U_WhsCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_AvailQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_AllocQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_WeightQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_ExpDate", oDBDS.Offset, "")
            oDBDS.SetValue("U_MfgDate", oDBDS.Offset, "")
            oDBDS.SetValue("U_QCStatus", oDBDS.Offset, "")
            oDBDS.SetValue("U_RetestDt", oDBDS.Offset, "")
            oDBDS.SetValue("U_Barcode", oDBDS.Offset, "")
            oDBDS.SetValue("U_Status", oDBDS.Offset, "")
            oDBDS.SetValue("U_SAPBatAbs", oDBDS.Offset, "")
            oDBDS.SetValue("U_ParentLineId", oDBDS.Offset, "")
            oDBDS.SetValue("U_RMItemCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_SAPBatchAbs", oDBDS.Offset, "")
            oDBDS.SetValue("U_AvailableQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_AllocatedQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_WeighedQty", oDBDS.Offset, "")
            oDBDS.SetValue("U_ExpiryDate", oDBDS.Offset, "")
            oDBDS.SetValue("U_RetestDate", oDBDS.Offset, "")

            oMatrix.LoadFromDataSource()
            oMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub SetNewLine3(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oDBDS As SAPbouiCOM.DBDataSource
            Dim oMatrix As SAPbouiCOM.Matrix

            oDBDS = objForm.DataSources.DBDataSources.Item("@TNX_PDSP_APR")
            oMatrix = objForm.Items.Item("MTX_3").Specific

            oMatrix.FlushToDataSource()

            oDBDS.InsertRecord(oDBDS.Size)
            oDBDS.Offset = oDBDS.Size - 1

            oDBDS.SetValue("LineId", oDBDS.Offset, (oDBDS.Offset + 1).ToString())
            oDBDS.SetValue("U_Role", oDBDS.Offset, "")
            oDBDS.SetValue("U_UserCode", oDBDS.Offset, "")
            oDBDS.SetValue("U_Status", oDBDS.Offset, "")
            oDBDS.SetValue("U_ActionDT", oDBDS.Offset, "")
            oDBDS.SetValue("U_Remarks", oDBDS.Offset, "")
            oDBDS.SetValue("U_Level", oDBDS.Offset, "")
            oDBDS.SetValue("U_ActionDate", oDBDS.Offset, "")

            oMatrix.LoadFromDataSource()
            oMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub SetNewLine4(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oDBDS As SAPbouiCOM.DBDataSource
            Dim oMatrix As SAPbouiCOM.Matrix

            oDBDS = objForm.DataSources.DBDataSources.Item("@TNX_PDSP_ATT")
            oMatrix = objForm.Items.Item("MTX_4").Specific

            oMatrix.FlushToDataSource()

            oDBDS.InsertRecord(oDBDS.Size)
            oDBDS.Offset = oDBDS.Size - 1

            oDBDS.SetValue("LineId", oDBDS.Offset, (oDBDS.Offset + 1).ToString())
            oDBDS.SetValue("U_FileName", oDBDS.Offset, "")
            oDBDS.SetValue("U_FilePath", oDBDS.Offset, "")
            oDBDS.SetValue("U_FileType", oDBDS.Offset, "")
            oDBDS.SetValue("U_UploadBy", oDBDS.Offset, "")
            oDBDS.SetValue("U_UploadDt", oDBDS.Offset, "")
            oDBDS.SetValue("U_UploadedBy", oDBDS.Offset, "")
            oDBDS.SetValue("U_UploadedDate", oDBDS.Offset, "")

            oMatrix.LoadFromDataSource()
            oMatrix.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "", Optional ByVal Series As Integer = 0)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PDSP_H")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PDSP_L")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@@TNX_PDSP_BT")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@@TNX_PDSP_APR")
            oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@@TNX_PDSP_ATT")

            objMatrix1 = objForm.Items.Item("MTX_1").Specific
            objMatrix2 = objForm.Items.Item("MTX_2").Specific
            objMatrix3 = objForm.Items.Item("MTX_3").Specific
            objMatrix4 = objForm.Items.Item("MTX_4").Specific

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_TNX_PDSP", "Primary"))
            oDBs_Head.SetValue("U_EDate", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            objMatrix1.Clear()
            objMatrix2.Clear()
            objMatrix3.Clear()
            objMatrix4.Clear()

            oDBs_Details1.Clear()
            oDBs_Details2.Clear()
            oDBs_Details3.Clear()
            oDBs_Details4.Clear()
            objMatrix1.FlushToDataSource()
            objMatrix2.FlushToDataSource()
            objMatrix3.FlushToDataSource()
            objMatrix4.FlushToDataSource()

            Me.SetNewLine1(objForm.UniqueID)
            Me.SetNewLine2(objForm.UniqueID)
            Me.SetNewLine3(objForm.UniqueID)
            Me.SetNewLine4(objForm.UniqueID)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class


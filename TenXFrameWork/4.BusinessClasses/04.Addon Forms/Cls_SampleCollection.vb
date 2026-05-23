Imports System.Configuration
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports Newtonsoft.Json.Linq

Public Class Cls_SampleCollection

#Region "Declaration"

    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBS1, oDBS2 As SAPbouiCOM.DBDataSource
    Dim oDBs_Details1, oDBs_Details2, oDBs_Details3 As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1, objMatrix2 As SAPbouiCOM.Matrix
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

            objMain.objUtilities.LoadForm("SampleCollection.xml", "SCOLLN", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("SCOLLN",
                                                       objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            objutilities = New Utilities
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_QCSC_H")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_QCSC_L")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_QCSC_COC")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_QCSC_ATT")

            objForm.EnableMenu("1292", True)
            objForm.EnableMenu("774", True)

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

            objMatrix = objForm.Items.Item("Mtx").Specific
            objMatrix1 = objForm.Items.Item("Mtx1").Specific
            objMatrix2 = objForm.Items.Item("Mtx2").Specific

            objForm.Freeze(False)

            objForm.EnableMenu("1292", True)

            Me.SetNewLine(objForm.UniqueID)
            Me.SetNewLine2(objForm.UniqueID)
            Me.SetNewLine3(objForm.UniqueID)

            ' Me.AutoDocentryNumber()

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





    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Dim DocNum As String = ""
            If pVal.MenuUID = "10X_SAMPCOL" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'Me.GetNextDocNum()
                Me.SetDefault(objForm.UniqueID)
                Me.SetNewLine(objForm.UniqueID)
                Me.SetNewLine2(objForm.UniqueID)
                Me.SetNewLine3(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then
                objMatrix = objForm.Items.Item("Mtx").Specific
                objMatrix1 = objForm.Items.Item("Mtx1").Specific
                objMatrix2 = objForm.Items.Item("Mtx2").Specific
                Me.SetNewLine(objForm.UniqueID)
                Me.SetNewLine2(objForm.UniqueID)
                Me.SetNewLine3(objForm.UniqueID)

            ElseIf pVal.MenuUID = "774" And pVal.BeforeAction = False Then

                objForm.Freeze(True)

                Try

                    Dim objMatrix As SAPbouiCOM.Matrix
                    Dim objMatrix1 As SAPbouiCOM.Matrix

                    '=====================================================
                    ' MATRIX 1
                    '=====================================================
                    objMatrix = objForm.Items.Item("Mtx").Specific

                    Dim delRows As New List(Of Integer)

                    ' capture selection safely
                    For i As Integer = 1 To objMatrix.RowCount
                        If objMatrix.IsRowSelected(i) Then
                            delRows.Add(i)
                        End If
                    Next

                    ' delete backward (CRITICAL)
                    For i As Integer = delRows.Count - 1 To 0 Step -1
                        objMatrix.DeleteRow(delRows(i))
                    Next

                    ' refresh UI
                    objMatrix.LoadFromDataSource()

                    If objMatrix.RowCount = 0 Then objMatrix.AddRow()

                    For i As Integer = 1 To objMatrix.RowCount
                        objMatrix.Columns.Item("LineID").Cells.Item(i).Specific.Value = i
                    Next


                    '=====================================================
                    ' MATRIX 2
                    '=====================================================
                    objMatrix1 = objForm.Items.Item("Mtx1").Specific

                    Dim delRows1 As New List(Of Integer)

                    For i As Integer = 1 To objMatrix1.RowCount
                        If objMatrix1.IsRowSelected(i) Then
                            delRows1.Add(i)
                        End If
                    Next

                    For i As Integer = delRows1.Count - 1 To 0 Step -1
                        objMatrix1.DeleteRow(delRows1(i))
                    Next

                    objMatrix1.LoadFromDataSource()

                    If objMatrix1.RowCount = 0 Then objMatrix1.AddRow()

                    For i As Integer = 1 To objMatrix1.RowCount
                        objMatrix1.Columns.Item("LineID").Cells.Item(i).Specific.Value = i
                    Next


                    '=====================================================
                    ' FORM MODE UPDATE
                    '=====================================================
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If

                Catch ex As Exception



                Finally
                    objForm.Freeze(False)
                End Try

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK

                    If (pVal.ItemUID = "Mtx") And pVal.BeforeAction = False Then

                        MATRIXS = pVal.ItemUID
                    End If

                    If (pVal.ItemUID = "Mtx1") And pVal.BeforeAction = False Then ' Or pVal.ItemUID = "Matrx5" Or pVal.ItemUID = "Matrix" Then
                        MATRIXS = pVal.ItemUID
                    End If
                    If (pVal.ItemUID = "Mtx2") And pVal.BeforeAction = False Then ' Or pVal.ItemUID = "Matrx5" Or pVal.ItemUID = "Matrix" Then
                        MATRIXS = pVal.ItemUID
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "1" AndAlso pVal.BeforeAction = False Then
                        Me.SetDefault(objForm.UniqueID)
                        Me.SetNewLine(objForm.UniqueID)

                        Me.SetNewLine2(objForm.UniqueID)
                        Me.SetNewLine3(objForm.UniqueID)
                    End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub SetNewLine(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_QCSC_L")
            objMatrix = objForm.Items.Item("Mtx").Specific

            objMatrix.AddRow()

            oDBs_Details1.SetValue("LineID",
                                       oDBs_Details1.Offset,
                                       objMatrix.VisualRowCount.ToString())

            oDBs_Details1.SetValue("U_CNTNO", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_CNTYPE", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_PCKSZ", oDBs_Details1.Offset, 0)
            oDBs_Details1.SetValue("U_PUOM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_SAQTY", oDBs_Details1.Offset, 0)
            oDBs_Details1.SetValue("U_UOM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_SAMPT", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_SELNO", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_SELIN", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VISOB", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_LINST", oDBs_Details1.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)

            objMatrix.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

    Sub SetNewLine2(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_QCSC_COC")

            objMatrix1 = CType(objForm.Items.Item("Mtx1").Specific,
                                   SAPbouiCOM.Matrix)

            objMatrix1.AddRow()

            oDBs_Details2.SetValue("LineID",
                                       oDBs_Details2.Offset,
                                       objMatrix1.VisualRowCount.ToString())

            oDBs_Details2.SetValue("U_Action", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_FUSER", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_TOUSR", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_ADATE", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_ATIME", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_LCTN", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_REMK", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_Esign", oDBs_Details2.Offset, "")

            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
            objMatrix1.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

    Sub SetNewLine3(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_QCSC_ATT")

            objMatrix2 = CType(objForm.Items.Item("Mtx2").Specific, SAPbouiCOM.Matrix)

            objMatrix2.AddRow()

            oDBs_Details3.SetValue("LineId",
                               oDBs_Details3.Offset,
                               objMatrix2.VisualRowCount.ToString())

            oDBs_Details3.SetValue("U_FNAME", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_FILTY", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_ATPTH", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_SAPAT", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_UPBY", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_UPDATE", oDBs_Details3.Offset, "")

            objMatrix2.SetLineData(objMatrix2.VisualRowCount)
            objMatrix2.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub

    Sub SetDefault(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_QCSC_H")

            oDBs_Head.SetValue("DocNum",
                               0,
                               objMain.objUtilities.GetNextDocNum(objForm,
                                                                  "TNX_QC_SC",
                                                                  "Primary"))

            objForm.DataBrowser.BrowseBy = "DocNum"
            oDBs_Head.SetValue("U_DATE", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))

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

        Return True

    End Function

End Class
Imports SAPbouiCOM

Public Class ClerenceMaster

#Region "Declaration"

    Public objForm As SAPbouiCOM.Form
    Public objMatrix As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim oDBs_Attach As SAPbouiCOM.DBDataSource
    Dim oDS As SAPbouiCOM.DBDataSource

#End Region

#Region "Create Form"

    Public Sub CreateForm()

        Try
            objMain.objUtilities.LoadForm("CleranceMaster.xml", "TNX_PCLM", ResourceType.Embeded)

            ' Robustly find the loaded form: XML FormType is TNX_PEQP, ObjectType is TNX_PCLMUDO

            objForm = objMain.objApplication.Forms.GetForm("TNX_PCLM",
                  objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)



            '================================================================
            ' DATASOURCES - guarded
            '================================================================
            Try
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCLM_H")
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText("ClerenceMaster: missing DB datasource @TNX_PCLM_H. " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                oDBs_Head = Nothing
            End Try

            Try
                oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCLM_L")
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText("ClerenceMaster: missing DB datasource @TNX_PCLM_L. " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                oDBs_Details = Nothing
            End Try

            '================================================================
            ' Set DocNum safely
            '================================================================
            Dim rs As SAPbobsCOM.Recordset

            rs = objMain.objCompany.GetBusinessObject(
             SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(
        "SELECT TOP 1 ""Code"" FROM ""@TNX_PCLM_H""")
            If rs.RecordCount = 0 Then

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE

                CType(objForm.Items.Item("Code").Specific,
                   SAPbouiCOM.EditText).Value = "1"

                'CType(objForm.Items.Item("txtName").Specific,
                '   SAPbouiCOM.EditText).Value = "1"

            Else

                Me.LoadExistingRecord()

            End If

            '================================================================
            ' SAFE UI ITEM ACCESS: guard each item before using
            '================================================================
            Dim itemsToGuard As String() = {"0_U_E", "17_U_E", "18_U_E", "19_U_E", "20_U_E", "21_U_E", "22_U_E", "23_U_E", "24_U_E", "25_U_E", "26_U_E"}
            For Each id As String In itemsToGuard
                If ItemExists(objForm, id) Then
                    Try
                        objForm.Items.Item(id).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        objForm.Items.Item(id).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                    Catch
                    End Try
                End If
            Next

            If ItemExists(objForm, "Item_2") Then
                Try
                    objForm.Items.Item("Item_2").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                    objForm.Items.Item("Item_2").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                Catch
                End Try
            End If

            If ItemExists(objForm, "Item_1") Then
                Try
                    objForm.Items.Item("Item_1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                Catch
                End Try
            End If

            If ItemExists(objForm, "1") Then
                Try
                    objForm.Items.Item("1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                Catch
                End Try
            End If

            If ItemExists(objForm, "2") Then
                Try
                    objForm.Items.Item("2").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                Catch
                End Try
            End If

            Me.objForm.EnableMenu("1282", True)
            Me.objForm.EnableMenu("1292", True)
            Me.objForm.EnableMenu("1293", True)

            If ItemExists(objForm, "Item_5") Then
                Try
                    objForm.Items.Item("Item_5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                Catch
                End Try
            End If

            If ItemExists(objForm, "APPI") Then
                Try
                    objForm.Items.Item("APPI").Enabled = False
                Catch
                End Try
            End If

            If ItemExists(objForm, "Item_7") Then
                Try
                    objForm.Items.Item("Item_7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Catch
                End Try
            End If
            '  Me.SetNewLine(objForm.UniqueID)
            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
                "Clerance Master Form Loaded Successfully",
                BoMessageTime.bmt_Short,
                BoStatusBarMessageType.smt_Success)

        Catch ex As Exception

            Try
                If objForm IsNot Nothing Then objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(
                "ClerenceMaster.CreateForm: " & ex.Message,
                BoMessageTime.bmt_Short,
                BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Public Sub SetNewLine(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCLM_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCLM_L")

            objMatrix = objForm.Items.Item("0_U_G").Specific

            objMatrix.AddRow()

            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount.ToString())

            oDBs_Details.SetValue("U_StepNo", oDBs_Details.Offset, "")


            oDBs_Details.SetValue("U_CheckPoint", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Method", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Chemical", oDBs_Details.Offset, "")

            oDBs_Details.SetValue("U_Contacttime", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Acceptcriteria", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_ResultType", oDBs_Details.Offset, "")

            oDBs_Details.SetValue("U_Mandatory", oDBs_Details.Offset, "N")

            oDBs_Details.SetValue("U_ChkPoint", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_CntTime", oDBs_Details.Offset, "")

            oDBs_Details.SetValue("U_AccCrit", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_ResType", oDBs_Details.Offset, "")

            oDBs_Details.SetValue("U_Mandtry", oDBs_Details.Offset, "N")

            objMatrix.SetLineData(objMatrix.VisualRowCount)

            objMatrix.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub
    ' Helper to check for item existence safely
    Private Function ItemExists(frm As SAPbouiCOM.Form, id As String) As Boolean
        If frm Is Nothing Then Return False
        Try
            Dim tmp = frm.Items.Item(id)
            Return True
        Catch
            Return False
        End Try
    End Function

#Region "Menu Event"
#Region "Menu Event"
    Sub LoadExistingRecord()

        Try

            Dim rs As SAPbobsCOM.Recordset

            rs = objMain.objCompany.GetBusinessObject(
                 SAPbobsCOM.BoObjectTypes.BoRecordset)

            rs.DoQuery(
            "SELECT TOP 1 * FROM ""@TNX_PCLM_H""")

            If rs.RecordCount > 0 Then

                Dim strCode As String

                strCode = rs.Fields.Item("Code").Value.ToString()

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

                CType(objForm.Items.Item("Code").Specific,
       SAPbouiCOM.EditText).Value = strCode

                objForm.Items.Item("1").Click()

                'objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

            End If

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try

    End Sub
    Public Sub SetDefault(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            '========================================================================
            ' DATASOURCES
            '========================================================================

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PCLM_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCLM_L")

            objForm.Freeze(True)

            '========================================================================
            ' AUTO DOCUMENT NUMBER
            '========================================================================

            oDBs_Head.SetValue("Code",
                       0,
                       objMain.objUtilities.GetNextDocNum(objForm,
                                                          "TNX_PEQP",
                                                          "Primary"))



            '========================================================================
            ' DEFAULT VALUES


            '========================================================================
            ' DEFAULT PANE
            '========================================================================

            objForm.PaneLevel = 1

            '========================================================================
            ' MATRIX DEFAULT ROW
            '========================================================================

            SetNewLine(FormUID)

            '========================================================================
            ' BUTTON SETTINGS
            '========================================================================

            objForm.Items.Item("1").Enabled = True
            objForm.Items.Item("2").Enabled = True

            '========================================================================
            ' MENU SETTINGS
            '========================================================================

            Me.objForm.EnableMenu("1282", True)
            Me.objForm.EnableMenu("519", True)
            Me.objForm.EnableMenu("520", True)
            Me.objForm.EnableMenu("1292", True)
            Me.objForm.EnableMenu("1293", True)

            objForm.Freeze(False)

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

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "10X_PMS_CLEAN" And pVal.BeforeAction = False Then
                Me.CreateForm()

            ElseIf pVal.MenuUID = "1292" And pVal.BeforeAction = False Then

                SetNewLine(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

                Me.SetDefault(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1293" AndAlso pVal.BeforeAction = True Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "TNX_PCLM" Then Exit Sub

                    BubbleEvent = False

                    objForm.Freeze(True)

                    objMatrix = CType(objForm.Items.Item("0_U_G").Specific,
                                    SAPbouiCOM.Matrix)

                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PCLM_L")

                    Dim selectedRow As Integer =
                        objMatrix.GetNextSelectedRow(0,
                        SAPbouiCOM.BoOrderType.ot_RowOrder)

                    If selectedRow <= 0 Then

                        objMain.objApplication.StatusBar.SetText(
                        "Please select row to delete",
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        Exit Try

                    End If

                    'Delete selected row
                    objMatrix.DeleteRow(selectedRow)



                    'Remove extra datasource rows
                    While oDBs_Details.Size > objMatrix.VisualRowCount
                        oDBs_Details.RemoveRecord(oDBs_Details.Size - 1)
                    End While

                    'Keep minimum one row
                    If oDBs_Details.Size = 0 Then

                        oDBs_Details.InsertRecord(0)

                        oDBs_Details.SetValue("LineId", 0, "1")

                        oDBs_Details.SetValue("U_StepNo", 0, "")


                        oDBs_Details.SetValue("U_CheckPoint", 0, "")
                        oDBs_Details.SetValue("U_Method", 0, "")
                        oDBs_Details.SetValue("U_Chemical", 0, "")

                        oDBs_Details.SetValue("U_Contacttime", 0, "")
                        oDBs_Details.SetValue("U_Acceptcriteria", 0, "")
                        oDBs_Details.SetValue("U_ResultType", 0, "")

                        oDBs_Details.SetValue("U_Mandatory", 0, "N")

                        oDBs_Details.SetValue("U_ChkPoint", 0, "")
                        oDBs_Details.SetValue("U_CntTime", 0, "")

                        oDBs_Details.SetValue("U_AccCrit", 0, "")
                        oDBs_Details.SetValue("U_ResType", 0, "")

                        oDBs_Details.SetValue("U_Mandtry", 0, "N")

                    End If

                    'Re-sequence line numbers
                    For i As Integer = 0 To oDBs_Details.Size - 1

                        oDBs_Details.SetValue("LineId",
                                           i,
                                           (i + 1).ToString())

                        ' ensure step number keeps in sync with line number
                        oDBs_Details.SetValue("U_StepNo",
                                              i,
                                              (i + 1).ToString())

                    Next

                    ' reload matrix from datasource so UI reflects new sequence
                    Try
                        objMatrix.LoadFromDataSource()
                    Catch
                    End Try

                    objMatrix.AutoResizeColumns()

                    'UPDATE MODE
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
            End If
        Catch ex As Exception

            Try
                If objForm IsNot Nothing Then objForm.Freeze(False)
            Catch
            End Try

            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try


    End Sub

#End Region


#End Region
    Public Sub ItemEvent(ByVal FormUID As String,
                 ByRef pVal As SAPbouiCOM.ItemEvent,
                 ByRef BubbleEvent As Boolean)
        Try

            If pVal.EventType = BoEventTypes.et_ITEM_PRESSED _
                    AndAlso pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.Item(FormUID)
                If pVal.ItemUID = "1" Then
                    Me.SetDefault(objForm.UniqueID)
                End If
            End If
        Catch ex As Exception
            Try
                If objForm IsNot Nothing Then objForm.Freeze(False)
            Catch
            End Try
            objMain.objApplication.StatusBar.SetText(
            ex.Message,
            BoMessageTime.bmt_Short,
            BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

#End Region
End Class

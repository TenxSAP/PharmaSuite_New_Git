Imports SAPbouiCOM



Public Class Countryregulatoryconfig


#Region "Declaration"

        Public objForm As SAPbouiCOM.Form
        Public oDBs_Content As SAPbouiCOM.DBDataSource
        Public Prgbar As SAPbouiCOM.ProgressBar
        Dim objutilities As Utilities

        Public objMatrix As SAPbouiCOM.Matrix

        Dim oDBs_Head As SAPbouiCOM.DBDataSource
        Dim oDBs_Details As SAPbouiCOM.DBDataSource
        Dim oDS As SAPbouiCOM.DBDataSource

#End Region

        Public Sub CreateForm()

            Try

            objMain.objUtilities.LoadForm("CountryRegulatoryconfig.xml",
                                              "REG_CNFG",
                                              ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm(
                          "REG_CNFG",
                          objMain.objApplication.Forms.ActiveForm.TypeCount)

                '========================================================================
                ' DATASOURCES
                '========================================================================

                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_CNFG")
                oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_REG_CNFL")

                '========================================================================
                ' DEFAULT VALUES
                '========================================================================

                oDBs_Head.SetValue("DocEntry",
                                   oDBs_Head.Offset,
                                   objMain.objUtilities.GetNextDocNum(
                                   objForm,
                                   "REG_CNFG",
                                   "Primary"))

                oDBs_Head.SetValue("U_Status", 0, "Active")

                objForm.PaneLevel = 1

                '========================================================================
                ' DOCENTRY CONTROL
                '========================================================================

                objForm.Items.Item("0_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                -1,
                SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                objForm.Items.Item("0_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                SAPbouiCOM.BoAutoFormMode.afm_Find,
                SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                '========================================================================
                ' HEADER FIELD CONTROLS
                '========================================================================

                'Country Code
                objForm.Items.Item("23_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                SAPbouiCOM.BoAutoFormMode.afm_Add,
                SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                'Authority Code
                objForm.Items.Item("24_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                SAPbouiCOM.BoAutoFormMode.afm_Add,
                SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                'Regulatory Format
                objForm.Items.Item("25_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                SAPbouiCOM.BoAutoFormMode.afm_Add,
                SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                'Default Validity Years
                objForm.Items.Item("26_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                SAPbouiCOM.BoAutoFormMode.afm_Add,
                SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                'Renewal Before Days
                objForm.Items.Item("27_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                SAPbouiCOM.BoAutoFormMode.afm_Add,
                SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                'Artwork Required
                objForm.Items.Item("28_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                SAPbouiCOM.BoAutoFormMode.afm_Add,
                SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                'Local Agent Required
                objForm.Items.Item("29_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                SAPbouiCOM.BoAutoFormMode.afm_Add,
                SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                'Language Requirement
                objForm.Items.Item("30_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                SAPbouiCOM.BoAutoFormMode.afm_Add,
                SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                'Status
                objForm.Items.Item("31_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                -1,
                SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                'Remarks
                objForm.Items.Item("32_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                SAPbouiCOM.BoAutoFormMode.afm_Add,
                SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                '========================================================================
                ' FIND MODE CONTROLS
                '========================================================================

                objForm.Items.Item("24_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                SAPbouiCOM.BoAutoFormMode.afm_Find,
                SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                objForm.Items.Item("25_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                SAPbouiCOM.BoAutoFormMode.afm_Find,
                SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                objForm.Items.Item("31_U_E").SetAutoManagedAttribute(
                SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                SAPbouiCOM.BoAutoFormMode.afm_Find,
                SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                '========================================================================
                ' MATRIX SETTINGS
                '========================================================================

                objMatrix = CType(objForm.Items.Item("0_U_G").Specific,
                                  SAPbouiCOM.Matrix)

                objMatrix.SelectionMode =
                SAPbouiCOM.BoMatrixSelect.ms_Single

                '========================================================================
                ' MATRIX COLUMN SETTINGS
                '========================================================================

                objMatrix.Columns.Item("LineId").Editable = False
                objMatrix.Columns.Item("C_0_1").Editable = True
                objMatrix.Columns.Item("C_0_2").Editable = True
                objMatrix.Columns.Item("C_0_3").Editable = True
                objMatrix.Columns.Item("C_0_4").Editable = True
                objMatrix.Columns.Item("C_0_5").Editable = True
                objMatrix.Columns.Item("C_0_6").Editable = True
                objMatrix.Columns.Item("C_0_7").Editable = True

                '========================================================================
                ' BUTTON SETTINGS
                '========================================================================

                objForm.Items.Item("1").Enabled = True
                objForm.Items.Item("2").Enabled = True

                '========================================================================
                ' MENU SETTINGS
                '========================================================================

                Me.objForm.EnableMenu("1282", True)
                Me.objForm.EnableMenu("1281", True)
                Me.objForm.EnableMenu("1292", True)
                Me.objForm.EnableMenu("1293", True)
                Me.objForm.EnableMenu("519", True)
                Me.objForm.EnableMenu("520", True)

                '========================================================================
                ' DEFAULT MATRIX ROW
                '========================================================================

                Me.SetNewLine(objForm.UniqueID)

                objForm.Freeze(False)

                objMain.objApplication.StatusBar.SetText(
                "Regulatory Configuration Form Loaded Successfully",
                BoMessageTime.bmt_Short,
                BoStatusBarMessageType.smt_Success)

            Catch ex As Exception

                Try
                    objForm.Freeze(False)
                Catch
                End Try

                objMain.objApplication.StatusBar.SetText(
                ex.Message,
                BoMessageTime.bmt_Short,
                BoStatusBarMessageType.smt_Error)

            End Try

        End Sub

        Public Sub SetNewLine(ByVal FormUID As String)

            Try

                objForm = objMain.objApplication.Forms.Item(FormUID)

                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_CNFG")
                oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_REG_CNFL")

                objMatrix = CType(objForm.Items.Item("0_U_G").Specific,
                                  SAPbouiCOM.Matrix)

                objMatrix.AddRow()

                oDBs_Details.SetValue("LineId",
                                      oDBs_Details.Offset,
                                      objMatrix.VisualRowCount.ToString())

                oDBs_Details.SetValue("U_DocTypCod",
                                      oDBs_Details.Offset,
                                      "")

                oDBs_Details.SetValue("U_DocTypNam",
                                      oDBs_Details.Offset,
                                      "")

                oDBs_Details.SetValue("U_Mandatory",
                                      oDBs_Details.Offset,
                                      "Y")

                oDBs_Details.SetValue("U_ValidReq",
                                      oDBs_Details.Offset,
                                      "Y")

                oDBs_Details.SetValue("U_MinValid",
                                      oDBs_Details.Offset,
                                      "0")

                oDBs_Details.SetValue("U_AttachReq",
                                      oDBs_Details.Offset,
                                      "N")

                oDBs_Details.SetValue("U_Remarks",
                                      oDBs_Details.Offset,
                                      "")

                objMatrix.SetLineData(objMatrix.VisualRowCount)

                objMatrix.AutoResizeColumns()

            Catch ex As Exception

                objMain.objApplication.StatusBar.SetText(
                ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            End Try

        End Sub

#Region "Menu Event"

        Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent,
                      ByRef BubbleEvent As Boolean)

            Try

            '========================================================================
            ' OPEN FORM
            '========================================================================

            If pVal.MenuUID = "10X_RMS_CNFG" _
                AndAlso pVal.BeforeAction = False Then

                Me.CreateForm()

                '========================================================================
                ' ADD MODE
                '========================================================================

            ElseIf pVal.MenuUID = "1282" _
                AndAlso pVal.BeforeAction = False Then

                Me.SetDefault(objForm.UniqueID)

                '========================================================================
                ' ADD ROW
                '========================================================================

            ElseIf pVal.MenuUID = "1292" _
                AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "REG_CNFG" Then Exit Sub

                Me.SetNewLine(objForm.UniqueID)

                '========================================================================
                ' DELETE ROW
                '========================================================================

            ElseIf pVal.MenuUID = "1293" _
                AndAlso pVal.BeforeAction = True Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "REG_CNFG" Then Exit Sub

                    BubbleEvent = False

                    objForm.Freeze(True)

                    objMatrix = CType(objForm.Items.Item("0_U_G").Specific,
                                          SAPbouiCOM.Matrix)

                    oDBs_Details =
                        objForm.DataSources.DBDataSources.Item("@TNX_REG_CNFL")

                    Dim selectedRow As Integer =
                        objMatrix.GetNextSelectedRow(
                        0,
                        SAPbouiCOM.BoOrderType.ot_RowOrder)

                    If selectedRow <= 0 Then

                        objMain.objApplication.StatusBar.SetText(
                            "Please select row to delete",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        Exit Try

                    End If

                    objMatrix.DeleteRow(selectedRow)

                    While oDBs_Details.Size > objMatrix.VisualRowCount

                        oDBs_Details.RemoveRecord(
                            oDBs_Details.Size - 1)

                    End While

                    If oDBs_Details.Size = 0 Then

                        oDBs_Details.InsertRecord(0)

                        oDBs_Details.SetValue("LineId", 0, "1")
                        oDBs_Details.SetValue("U_DocTypCod", 0, "")
                        oDBs_Details.SetValue("U_DocTypNam", 0, "")
                        oDBs_Details.SetValue("U_Mandatory", 0, "Y")
                        oDBs_Details.SetValue("U_ValidReq", 0, "Y")
                        oDBs_Details.SetValue("U_MinValid", 0, "0")
                        oDBs_Details.SetValue("U_AttachReq", 0, "N")
                        oDBs_Details.SetValue("U_Remarks", 0, "")

                    End If

                    For i As Integer = 0 To oDBs_Details.Size - 1

                        oDBs_Details.SetValue(
                            "LineId",
                            i,
                            (i + 1).ToString())

                    Next

                    objMatrix.LoadFromDataSource()

                    objMatrix.AutoResizeColumns()

                    If objForm.Mode =
                        SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                        objForm.Mode =
                            SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

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

                '========================================================================
                ' FIND MODE
                '========================================================================

            ElseIf pVal.MenuUID = "1281" _
                AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "REG_CNFG" Then Exit Sub

                    objForm.PaneLevel = 1

                End If

            Catch ex As Exception

                Try
                    If objForm IsNot Nothing Then
                        objForm.Freeze(False)
                    End If
                Catch
                End Try

                objMain.objApplication.StatusBar.SetText(
                ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            End Try

        End Sub

#End Region

        Public Sub SetDefault(ByVal FormUID As String)

            Try

                objForm = objMain.objApplication.Forms.Item(FormUID)

                '========================================================================
                ' DATASOURCES
                '========================================================================

                oDBs_Head =
                objForm.DataSources.DBDataSources.Item("@TNX_REG_CNFG")

                oDBs_Details =
                objForm.DataSources.DBDataSources.Item("@TNX_REG_CNFL")

                objForm.Freeze(True)

                '========================================================================
                ' AUTO DOCUMENT NUMBER
                '========================================================================

                oDBs_Head.SetValue(
                "DocEntry",
                0,
                objMain.objUtilities.GetNextDocNum(
                objForm,
                "REG_CNFG",
                "Primary"))

                '========================================================================
                ' DEFAULT VALUES
                '========================================================================

                oDBs_Head.SetValue("U_Status", 0, "Active")

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
                    If objForm IsNot Nothing Then
                        objForm.Freeze(False)
                    End If
                Catch
                End Try

                objMain.objApplication.StatusBar.SetText(
                ex.Message,
                BoMessageTime.bmt_Short,
                BoStatusBarMessageType.smt_Error)

            End Try

        End Sub

    End Class


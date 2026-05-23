Public Class ProductRegistration


#Region "Declaration"
    Public objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details, oDBs_Details2, oDBs_Details3 As SAPbouiCOM.DBDataSource
    Dim objMatrix, ObjMatrix2, ObjMatrix3, oMatrix As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim str, str1 As String
    Public rs, RsNum As SAPbobsCOM.Recordset
    Dim LostFocusFlag As Boolean = False
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim objutilities As Utilities
    Dim matrixid As String
    Public objFormx As SAPbouiCOM.Form
    Dim SERIES As String = ""
    Dim DOCKEY As String = ""
#End Region

#Region "Form Initialization"
    Public Sub CreateForm(Optional ByVal DocEntryNum As String = "", Optional ByVal DocSeries As String = "")
        Try
            ' Load form layout from embedded resource using your target template definitions
            objMain.objUtilities.LoadForm("PRODUCTREGISTRATION.xml", "UDO_F_UDO_REG_PRDREG", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("UDO_F_UDO_REG_PRDREG", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            ' Find mode execution fallback mapping if explicit routing keys are passed
            If DocEntryNum <> "" Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("0_U_E").Specific.Value = DocEntryNum
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                objForm.Freeze(False)
                Exit Sub
            End If

            objutilities = New Utilities

            ' Initialize DataSources matching target UDO fields
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_PRH")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_REG_PRL")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_REG_PRD")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_REG_PRA")

            ' Bind Matrix variables to target matching Grid component UIDs 
            objMatrix = objForm.Items.Item("0_U_G").Specific
            ObjMatrix2 = objForm.Items.Item("1_U_G").Specific
            ObjMatrix3 = objForm.Items.Item("2_U_G").Specific

            ' Initialize AutoManaged entry states and primary field contexts
            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_REG_PRDREG"))
            oDBs_Head.SetValue("U_CreatedDate", 0, DateTime.Now.ToString("yyyyMMdd"))

            ' Configure standard editing access controls dynamically
            objForm.Items.Item("0_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("0_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            ' Enable right-click standard add row options
            objForm.EnableMenu("1292", True)

            ' Clean matrix layouts and inject the baseline data row arrays
            objMatrix.Clear()
            ObjMatrix2.Clear()
            ObjMatrix3.Clear()

            ' Force view level pane to load context tab index 1 natively
            objForm.PaneLevel = 1

            Me.SetNewLine(objForm.UniqueID)
            Me.SetNewLine1(objForm.UniqueID)
            Me.SetNewLine2(objForm.UniqueID)

            objMain.objApplication.StatusBar.SetText("Product Registration screens initialized successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Form Creation Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Events Handler"
    ''' <summary>
    ''' Intercepts standard application execution menus like add record, add row or right-click deletions.
    ''' </summary>
    Public Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case "10X_REG_PROD"
                        Me.CreateForm()

                    Case "1282" ' Add Record Option
                        objMatrix = objForm.Items.Item("0_U_G").Specific
                        ObjMatrix2 = objForm.Items.Item("1_U_G").Specific
                        ObjMatrix3 = objForm.Items.Item("2_U_G").Specific
                        Me.SetDefault(objForm.UniqueID)

                    Case "1292" ' Standard Add Row Parameter Trigger
                        If matrixid = "0_U_G" Then
                            Me.SetNewLine(objForm.UniqueID)
                        ElseIf matrixid = "1_U_G" Then
                            Me.SetNewLine1(objForm.UniqueID)
                        ElseIf matrixid = "2_U_G" Then
                            Me.SetNewLine2(objForm.UniqueID)
                        End If

                    Case "Delete Row" ' Contextual removal interceptor
                        Dim targetMatrix As SAPbouiCOM.Matrix = Nothing
                        Dim targetDS As SAPbouiCOM.DBDataSource = Nothing

                        ' Dynamic targeted lookup route based on screen tracking position focus variables
                        If matrixid = "0_U_G" Then
                            targetMatrix = CType(objForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix)
                            targetDS = objForm.DataSources.DBDataSources.Item("@TNX_REG_PRL")
                        ElseIf matrixid = "1_U_G" Then
                            targetMatrix = CType(objForm.Items.Item("1_U_G").Specific, SAPbouiCOM.Matrix)
                            targetDS = objForm.DataSources.DBDataSources.Item("@TNX_REG_PRD")
                        ElseIf matrixid = "2_U_G" Then
                            targetMatrix = CType(objForm.Items.Item("2_U_G").Specific, SAPbouiCOM.Matrix)
                            targetDS = objForm.DataSources.DBDataSources.Item("@TNX_REG_PRA")
                        End If

                        If targetMatrix Is Nothing OrElse targetMatrix.VisualRowCount = 0 Then
                            objMain.objApplication.MessageBox("No active lines available to target for deletion.")
                            Return
                        End If

                        Dim rowDeleted As Boolean = False

                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            For i As Integer = targetMatrix.VisualRowCount To 1 Step -1
                                If targetMatrix.IsRowSelected(i) Then
                                    targetMatrix.DeleteRow(i)
                                    rowDeleted = True
                                End If
                            Next
                            If rowDeleted Then targetMatrix.FlushToDataSource()

                        ElseIf objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE OrElse objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            targetMatrix.FlushToDataSource()

                            For i As Integer = targetMatrix.VisualRowCount To 1 Step -1
                                If targetMatrix.IsRowSelected(i) Then
                                    Dim dsIndex As Integer = i - 1
                                    If dsIndex < targetDS.Size Then
                                        targetDS.RemoveRecord(dsIndex)
                                        rowDeleted = True
                                    End If
                                End If
                            Next
                            If rowDeleted Then targetMatrix.LoadFromDataSource()
                        End If

                        If Not rowDeleted Then
                            objMain.objApplication.MessageBox("Please select a valid grid row pointer to remove.")
                            Return
                        End If

                        ' Re-number the unique index identifier mappings sequentially
                        For i As Integer = 1 To targetMatrix.VisualRowCount
                            targetMatrix.Columns.Item("#").Cells.Item(i).Specific.Value = i.ToString()
                        Next
                End Select
            End If
        Catch ex As Exception
            If objForm IsNot Nothing Then objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Menu Interface Fail: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Form Field Validation Routing"

    Public Function Validate() As Boolean
        Try
            ' Header contextual string checks
            If oDBs_Head.GetValue("U_RegCode", 0).Trim() = "" Then
                objMain.objApplication.StatusBar.SetText("Registration Code criteria definition required.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oDBs_Head.GetValue("U_ItemCode", 0).Trim() = "" Then
                objMain.objApplication.StatusBar.SetText("Item Code field validation failed: Missing entry.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oDBs_Head.GetValue("U_Status", 0).Trim() = "" Then
                objMain.objApplication.StatusBar.SetText("Global Form Status parameter mandatory context.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            ' Matrix line layer presence asserts
            If oDBs_Details.Size = 0 OrElse oDBs_Details.GetValue("U_Country", 0).Trim() = "" Then
                objMain.objApplication.StatusBar.SetText("At least one registration destination country line entry must be structured.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Validation Processing Failure: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region

#Region "Item Events Logic Block"
    
    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    ' Folder Switch management via target XML structures
                    If pVal.BeforeAction = False Then
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        Select Case pVal.ItemUID
                            Case "0_U_FD"
                                objForm.PaneLevel = 1
                            Case "1_U_FD"
                                objForm.PaneLevel = 2
                            Case "2_U_FD"
                                objForm.PaneLevel = 3
                        End Select
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    ' Retain tracking context focus point mapping values on matrix items
                    If pVal.BeforeAction = False Then
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If pVal.ItemUID = "0_U_G" OrElse pVal.ItemUID = "1_U_G" OrElse pVal.ItemUID = "2_U_G" Then
                            matrixid = pVal.ItemUID
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    ' Standard system ChooseFromList data handling pipelines
                    If pVal.BeforeAction = False Then
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                        Dim oDT As SAPbouiCOM.DataTable = CFLEvent.SelectedObjects

                        If Not oDT Is Nothing AndAlso pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                            ' Example placeholder hook for item code configuration linkage lookup operations
                            If CFLEvent.ChooseFromListUID = "CFL_ITEM" Then
                                oDBs_Head.SetValue("U_ItemCode", oDBs_Head.Offset, oDT.GetValue("ItemCode", 0))
                                oDBs_Head.SetValue("U_ItemName", oDBs_Head.Offset, oDT.GetValue("ItemName", 0))
                            End If
                        End If
                    End If
            End Select

        Catch ex As Exception
            If objForm IsNot Nothing Then objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("ItemEvent Exception Raised: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "New Grid Row Allocation Blueprints"
    ''' <summary>
    ''' Generates empty placeholder allocation space targeting Details Matrix 1 (Tab index 1)
    ''' </summary>
    Public Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_REG_PRL")
            objMatrix = objForm.Items.Item("0_U_G").Specific

            objForm.Freeze(True)
            objMatrix.AddRow()
            oDBs_Details.SetValue("LineID", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_Country", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_RegNo", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_Status", oDBs_Details.Offset, "Active")

            objMatrix.SetLineData(objMatrix.VisualRowCount)
            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Line Insertion 1 Failure: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Generates empty placeholder allocation space targeting Details Matrix 2 (Tab index 2)
    ''' </summary>
    Public Sub SetNewLine1(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@TNX_REG_PRD")
            ObjMatrix2 = objForm.Items.Item("1_U_G").Specific

            objForm.Freeze(True)
            ObjMatrix2.AddRow()
            oDBs_Details2.SetValue("LineID", oDBs_Details2.Offset, ObjMatrix2.VisualRowCount)
            oDBs_Details2.SetValue("U_DocType", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_DocName", oDBs_Details2.Offset, "")
            oDBs_Details2.SetValue("U_DocDate", oDBs_Details2.Offset, DateTime.Now.ToString("yyyyMMdd"))

            ObjMatrix2.SetLineData(ObjMatrix2.VisualRowCount)
            ObjMatrix2.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Line Insertion 2 Failure: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Generates empty placeholder allocation space targeting Details Matrix 3 (Tab index 3)
    ''' </summary>
    Public Sub SetNewLine2(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@TNX_REG_PRA")
            ObjMatrix3 = objForm.Items.Item("2_U_G").Specific

            objForm.Freeze(True)
            ObjMatrix3.AddRow()
            oDBs_Details3.SetValue("LineID", oDBs_Details3.Offset, ObjMatrix3.VisualRowCount)
            oDBs_Details3.SetValue("U_Level", oDBs_Details3.Offset, "")
            oDBs_Details3.SetValue("U_ApproverUser", oDBs_Details3.Offset, "")

            ObjMatrix3.SetLineData(ObjMatrix3.VisualRowCount)
            ObjMatrix3.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Line Insertion 3 Failure: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Form State Reset Configuration"
    ''' <summary>
    ''' Falls back control definitions to baseline operational system configurations.
    ''' </summary>
    Public Sub SetDefault(ByVal FormUID As String, Optional ByVal Flag As String = "")
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_PRH")
            oDBs_Head.SetValue("DocEntry", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "UDO_REG_PRDREG"))
            oDBs_Head.SetValue("U_CreatedDate", 0, DateTime.Now.ToString("yyyyMMdd"))

            objMatrix = objForm.Items.Item("0_U_G").Specific
            ObjMatrix2 = objForm.Items.Item("1_U_G").Specific
            ObjMatrix3 = objForm.Items.Item("2_U_G").Specific

            objMatrix.Clear()
            ObjMatrix2.Clear()
            ObjMatrix3.Clear()

            Me.SetNewLine(objForm.UniqueID)
            Me.SetNewLine1(objForm.UniqueID)
            Me.SetNewLine2(objForm.UniqueID)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText("Defaults Reset Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Context Click Modifiers"
    ''' <summary>
    ''' Controls context menu modification access parameters on active coordinate points.
    ''' </summary>
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            Dim oMenuItem As SAPbouiCOM.MenuItem
            Dim oMenus As SAPbouiCOM.Menus
            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams

            oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)

            If eventInfo.FormUID = objForm.UniqueID AndAlso eventInfo.BeforeAction = True Then
                If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE AndAlso objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then

                    ' Verify focus indicator highlights row indices inside our matrices
                    If (eventInfo.ItemUID = "0_U_G" OrElse eventInfo.ItemUID = "1_U_G" OrElse eventInfo.ItemUID = "2_U_G") AndAlso eventInfo.ColUID = "#" Then
                        oMenuItem = objMain.objApplication.Menus.Item("1280") ' Main 'Data' block 
                        oMenus = oMenuItem.SubMenus
                        If Not oMenus.Exists("Delete Row") Then
                            oCreationPackage.UniqueID = "Delete Row"
                            oCreationPackage.String = "Delete Row"
                            oCreationPackage.Enabled = True
                            oMenus.AddEx(oCreationPackage)
                        End If
                    Else
                        ' Remove element context dynamically if clicking blank regions
                        Try
                            objMain.objApplication.Menus.RemoveEx("Delete Row")
                        Catch
                        End Try
                    End If
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("RightClick Handler Error: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Form Data Execution Events Listener"
    ''' <summary>
    ''' Monitors background serialization events like data retrieval loading or operations commits.
    ''' </summary>
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("UDO_F_UDO_REG_PRDREG", objMain.objApplication.Forms.ActiveForm.TypeCount)

            If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD AndAlso BusinessObjectInfo.BeforeAction = False AndAlso BusinessObjectInfo.ActionSuccess = True Then
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_REG_PRH")
                Dim statusValue As String = oDBs_Head.GetValue("U_Status", 0).Trim().ToUpper()

                ' Block interaction paths automatically if visual record indicator represents closed status
                If statusValue = "CLOSED" Then
                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("FormDataEvent Trace Error: " & ex.Message)
        End Try
    End Sub
#End Region

End Class

Imports SAPbouiCOM
Imports SAPbobsCOM
Imports System.IO
Imports System.Text
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Public Class LineClerence
#Region "Declaration"

    Public objForm As SAPbouiCOM.Form
    Public objMatrix, objMatrix1, objMatrix2, objMatrix3 As SAPbouiCOM.Matrix
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim oDBs_Details As SAPbouiCOM.DBDataSource
    Dim oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim oDBs_Attach As SAPbouiCOM.DBDataSource
    Dim oDS As SAPbouiCOM.DBDataSource


#End Region

    Public Sub CreateForm()

        Try

            objMain.objUtilities.LoadForm("Lineclerance.xml", "TNX_PLCL", ResourceType.Embeded)

            objForm = objMain.objApplication.Forms.GetForm("TNX_PLCL",
                  objMain.objApplication.Forms.ActiveForm.TypeCount)
            'objForm.DataBrowser.BrowseBy = "DocNum"
            objForm.Freeze(True)

            Try
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_H")
                oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_L")
                oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_EQP")
                oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_ATT")
            Catch ex As Exception
                Throw New Exception("Missing DBDataSource. Verify UDT names: @TNX_PLCL_H, @TNX_PLCL_L, @TNX_PLCL_EQP, @TNX_PLCL_ATT. " & ex.Message)
            End Try

            Try
                objMatrix = CType(objForm.Items.Item("matChk").Specific, SAPbouiCOM.Matrix)
            Catch ex As Exception
                Throw New Exception("Matrix 'matChk' not found on the form. Confirm UID in LineClerence1.xml. " & ex.Message)
            End Try

            Try
                objMatrix1 = CType(objForm.Items.Item("Item_1").Specific, SAPbouiCOM.Matrix)
            Catch ex As Exception
                Throw New Exception("Matrix 'Item_1' not found on the form. Confirm UID in LineClerence1.xml. " & ex.Message)
            End Try

            Try
                objMatrix2 = CType(objForm.Items.Item("Item_0").Specific, SAPbouiCOM.Matrix)
            Catch ex As Exception
                Throw New Exception("Matrix 'Item_0' not found on the form. Confirm UID in LineClerence1.xml. " & ex.Message)
            End Try

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset,
                               objMain.objUtilities.GetNextDocNum(objForm, "TNX_PLCL", "Primary"))

            oDBs_Head.SetValue("U_ClearanceDate", 0, DateTime.Now.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_ClearanceTime", 0, DateTime.Now.ToString("HHmm"))
            oDBs_Head.SetValue("U_Status", 0, "Draft")
            oDBs_Head.SetValue("U_ApprovalStatus", 0, "Pending")

            objForm.PaneLevel = 1

            SetDefault(objForm.UniqueID)

            If ItemExists(objForm, "ItemName") Then objForm.Items.Item("ItemName").Enabled = False
            If ItemExists(objForm, "AreaName") Then objForm.Items.Item("AreaName").Enabled = False

            If ItemExists(objForm, "ClrDate") Then
                objForm.Items.Item("ClrDate").SetAutoManagedAttribute(
                    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                    SAPbouiCOM.BoAutoFormMode.afm_Add,
                    SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If ItemExists(objForm, "ClrTime") Then
                objForm.Items.Item("ClrTime").SetAutoManagedAttribute(
                    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                    SAPbouiCOM.BoAutoFormMode.afm_Add,
                    SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If ItemExists(objForm, "Status") Then
                objForm.Items.Item("Status").SetAutoManagedAttribute(
                    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                    -1,
                    SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If ItemExists(objForm, "AppStat") Then
                objForm.Items.Item("AppStat").SetAutoManagedAttribute(
                    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                    -1,
                    SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If ItemExists(objForm, "AprBy") Then
                objForm.Items.Item("AprBy").SetAutoManagedAttribute(
                    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                    -1,
                    SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If ItemExists(objForm, "AprDate") Then
                objForm.Items.Item("AprDate").SetAutoManagedAttribute(
                    SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
                    -1,
                    SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If


            If ItemExists(objForm, "btnSub") Then objForm.Items.Item("btnSub").Enabled = True
            If ItemExists(objForm, "btnCopy") Then objForm.Items.Item("btnCopy").Enabled = True
            If ItemExists(objForm, "btn_dlt") Then objForm.Items.Item("btn_dlt").Enabled = True

            Me.objForm.EnableMenu("1282", True) ' Add
            Me.objForm.EnableMenu("1281", True) ' Find
            Me.objForm.EnableMenu("1288", True) ' Next Record
            Me.objForm.EnableMenu("1289", True) ' Previous Record
            Me.objForm.EnableMenu("1290", True) ' First Record
            Me.objForm.EnableMenu("1291", True) ' Last Record
            Me.objForm.EnableMenu("1292", True) ' Add Row
            Me.objForm.EnableMenu("1293", True) ' Delete Row

            ' Normalize date fields (truncate/format) to avoid length errors
            objForm.Items.Item("DocNum").SetAutoManagedAttribute(
            SAPbouiCOM.BoAutoManagedAttr.ama_Editable,
            SAPbouiCOM.BoAutoFormMode.afm_Find,
            SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Freeze(False)

            objMain.objApplication.StatusBar.SetText(
                "10X Pharma Line Clearance Form Loaded Successfully",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Success)

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

    Private Function ItemExists(ByVal frm As SAPbouiCOM.Form, ByVal itemId As String) As Boolean
        Try
            If frm Is Nothing Then Return False
            Dim dummy = frm.Items.Item(itemId)
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Function GetDBValueAt(oDB As SAPbouiCOM.DBDataSource, rowIndex As Integer, ParamArray aliases() As String) As String
        If oDB Is Nothing Then Return ""
        For Each a As String In aliases
            Try
                Dim v As String = oDB.GetValue(a, rowIndex)
                If v IsNot Nothing Then Return v
            Catch
            End Try
        Next
        Return ""
    End Function

    ' Try to safely set value on a DBDataSource using the first alias that works.
    Private Sub SetDBValueAt(oDB As SAPbouiCOM.DBDataSource, rowIndex As Integer, value As String, ParamArray aliases() As String)
        If oDB Is Nothing Then Return
        For Each a As String In aliases
            Try
                oDB.SetValue(a, rowIndex, value)
                Return
            Catch
            End Try
        Next
    End Sub

    Private Sub NormalizeCheckedDates(frm As SAPbouiCOM.Form)
        If frm Is Nothing Then Return
        Dim ds As SAPbouiCOM.DBDataSource = Nothing
        Try
            ds = frm.DataSources.DBDataSources.Item("@TNX_PLCL_L")
        Catch
        End Try
        If ds Is Nothing Then Return

        Dim aliases As String() = New String() {"U_CheckedDate", "U_CheckedDate ", "U_CheckedDate  ", "CheckedDate", "CheckedDate "}

        For i As Integer = 0 To ds.Size - 1
            Dim raw As String = GetDBValueAt(ds, i, aliases)
            If String.IsNullOrEmpty(raw) Then Continue For

            raw = raw.Trim()
            If raw.Length >= 10 Then

                Dim parsed As DateTime
                If DateTime.TryParse(raw, parsed) Then
                    Dim formatted As String = parsed.ToString("yyyyMMdd")
                    SetDBValueAt(ds, i, formatted, aliases)
                Else

                    Dim truncated As String = raw.Substring(0, Math.Min(8, raw.Length))
                    SetDBValueAt(ds, i, truncated, aliases)
                End If
            ElseIf raw.Length > 0 And raw.Length <= 9 Then

                Dim parsed As DateTime
                If DateTime.TryParse(raw, parsed) Then
                    SetDBValueAt(ds, i, parsed.ToString("yyyyMMdd"), aliases)
                Else

                End If
            End If
        Next
    End Sub

#Region "Default Values"

    Public Sub SetDefault(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_L")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_EQP")
            oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_ATT")

            objForm.Freeze(True)

            oDBs_Head.SetValue("U_Status", 0, "Draft")
            oDBs_Head.SetValue("U_ApprovalStatus", 0, "Pending")
            oDBs_Head.SetValue("U_ClearanceDate", 0, DateTime.Now.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_ClearanceTime", 0, DateTime.Now.ToString("HHmm"))
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset,
                           objMain.objUtilities.GetNextDocNum(objForm, "TNX_PLCL", "Primary"))
            If String.IsNullOrEmpty(oDBs_Head.GetValue("U_RequestedBy", 0).Trim()) Then
                oDBs_Head.SetValue("U_RequestedBy", 0, objMain.objCompany.UserName)
            End If


            objForm.PaneLevel = 1

            SetNewLine(FormUID)
            SetNewLine1(FormUID)
            SetNewLine2(FormUID)


            If ItemExists(objForm, "btnSub") Then objForm.Items.Item("btnSub").Enabled = True
            If ItemExists(objForm, "btnCopy") Then objForm.Items.Item("btnCopy").Enabled = True
            If ItemExists(objForm, "btn_dlt") Then objForm.Items.Item("btn_dlt").Enabled = True

            If ItemExists(objForm, "AprBy") Then objForm.Items.Item("AprBy").Enabled = False
            If ItemExists(objForm, "AprDate") Then objForm.Items.Item("AprDate").Enabled = False
            If ItemExists(objForm, "Status") Then objForm.Items.Item("Status").Enabled = False
            If ItemExists(objForm, "AppStat") Then objForm.Items.Item("AppStat").Enabled = False


            If ItemExists(objForm, "fldChk") Then objForm.Items.Item("fldChk").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

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

#End Region

#Region "Add New Line"

    Sub SetNewLine(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_L")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_EQP")
            oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_ATT")

            objMatrix = objForm.Items.Item("matChk").Specific

            objMatrix.AddRow()


            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_CheckCode", 0, "")
            oDBs_Details.SetValue("U_CheckPoint", 0, "")
            oDBs_Details.SetValue("U_Category", 0, "")
            oDBs_Details.SetValue("U_Expected", 0, "")
            oDBs_Details.SetValue("U_Observed", 0, "")
            oDBs_Details.SetValue("U_CheckedBy", 0, "")
            oDBs_Details.SetValue("U_Result", 0, "")
            oDBs_Details.SetValue("U_Remarks", 0, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)

            objMatrix.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
                ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub


    Sub SetNewLine1(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_L")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_EQP")
            oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_ATT")
            objMatrix1 = objForm.Items.Item("Item_1").Specific


            objMatrix1.AddRow()


            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
            oDBs_Details1.SetValue("U_EquipCode", 0, "")
            oDBs_Details1.SetValue("U_EquipName", 0, "")
            oDBs_Details1.SetValue("U_CleaningLogNo", 0, "")
            oDBs_Details1.SetValue("U_CleaningStatus", 0, "")
            oDBs_Details1.SetValue("U_CalibDueDate", 0, "")
            oDBs_Details1.SetValue("U_CalibStatus", 0, "")
            oDBs_Details1.SetValue("U_ReadyStatus", 0, "")
            oDBs_Details1.SetValue("U_Rmarks", 0, "")

            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
            objMatrix1.AutoResizeColumns()
        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
                ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
    Sub SetNewLine2(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_H")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_L")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_EQP")
            oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_PLCL_ATT")

            objMatrix2 = objForm.Items.Item("Item_0").Specific


            objMatrix2.AddRow()

            oDBs_Attach.SetValue("LineId", oDBs_Attach.Offset, objMatrix2.VisualRowCount)
            oDBs_Attach.SetValue("U_TPA", 0, "")
                oDBs_Attach.SetValue("U_FN", 0, "")
                oDBs_Attach.SetValue("U_ATD", 0, "")
                oDBs_Attach.SetValue("U_FTT", 0, "")


            objMatrix2.SetLineData(objMatrix2.VisualRowCount)
            objMatrix2.AutoResizeColumns()

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
                ex.Message,
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

#End Region
#Region "Item Event"

    Public Sub ItemEvent(ByVal FormUID As String,
                     ByRef pVal As SAPbouiCOM.ItemEvent,
                     ByRef BubbleEvent As Boolean)

        Try
            ' Guard: try to obtain form safely. Some events may fire when form is not loaded.
            Try
                objForm = objMain.objApplication.Forms.Item(FormUID)
            Catch ex As Exception
                objMain.objApplication.StatusBar.SetText("LineClerence.ItemEvent: Form not found or already closed (UID=" & FormUID & "). Ignoring event.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End Try

            ' Guard: ensure this handler only processes the expected form type
            If objForm Is Nothing OrElse objForm.TypeEx <> "TNX_PLCL" Then
                ' Not the Line Clearance form - ignore
                Exit Sub
            End If

            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.BeforeAction = False Then

                        Select Case pVal.ItemUID

                            Case "1"
                                objForm = objMain.objApplication.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" Then
                                    SetDefault(objForm.UniqueID)
                                End If


                            Case "Item_1"     'General Tab

                                objForm.PaneLevel = 1

                            Case "Item_2"     'Details Tab

                                objForm.PaneLevel = 2

                            Case "Item_3"     'Attachment Tab

                                objForm.PaneLevel = 3

                            Case "btn_Add"

                                SetNewLine(FormUID)

                                SetNewLine1(FormUID)
                                SetNewLine2(FormUID)
                            Case "btn_Del"

                                Try

                                    objForm.Freeze(True)

                                    objMatrix2 = CType(objForm.Items.Item("Item_0").Specific,
                                                      SAPbouiCOM.Matrix)

                                    oDBs_Attach = objForm.DataSources.DBDataSources.Item("@TNX_ATTACH_C3")

                                    Dim selectedRow As Integer = 0

                                    For i As Integer = 1 To objMatrix2.VisualRowCount

                                        If objMatrix2.IsRowSelected(i) = True Then

                                            selectedRow = i
                                            Exit For

                                        End If

                                    Next

                                    If selectedRow = 0 Then

                                        objMain.objApplication.StatusBar.SetText(
                                        "Please select attachment row.",
                                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                        Exit Try

                                    End If

                                    objMatrix2.FlushToDataSource()

                                    oDBs_Attach.RemoveRecord(selectedRow - 1)

                                    If oDBs_Attach.Size = 0 Then

                                        oDBs_Attach.InsertRecord(0)

                                        oDBs_Attach.SetValue("LineId", 0, "1")
                                        oDBs_Attach.SetValue("U_TPA", 0, "")
                                        oDBs_Attach.SetValue("U_FN", 0, "")
                                        oDBs_Attach.SetValue("U_FTT", 0, "")
                                        oDBs_Attach.SetValue("U_ATD", 0, "")

                                    Else

                                        For i As Integer = 0 To oDBs_Attach.Size - 1

                                            oDBs_Attach.SetValue("LineId",
                                                             i,
                                                             (i + 1).ToString())

                                        Next

                                    End If

                                    objMatrix2.LoadFromDataSource()
                                    objMatrix2.AutoResizeColumns()

                                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                                        objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                                    End If

                                    If objForm.Items.Exists("btn_Del") Then objForm.Items.Item("btn_Del").Enabled = False

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

                        End Select

                    End If

                Case SAPbouiCOM.BoEventTypes.et_CLICK

                    If pVal.BeforeAction = False Then

                        If pVal.ItemUID = "Item_0" Then

                            If objForm.Items.Exists("btn_Del") Then objForm.Items.Item("btn_Del").Enabled = True

                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                    If pVal.BeforeAction = False Then

                        If pVal.ItemUID = "Item_0" _
                        AndAlso pVal.ColUID = "FPATH" Then

                            Try

                                Dim objMatrix As SAPbouiCOM.Matrix =
                                CType(objForm.Items.Item("Item_0").Specific,
                                      SAPbouiCOM.Matrix)

                                If pVal.Row > 0 AndAlso
                               pVal.Row <= objMatrix.VisualRowCount Then

                                    Dim fullPath As String =
                                    CType(objMatrix.Columns.Item("FPATH").
                                    Cells.Item(pVal.Row).Specific,
                                    SAPbouiCOM.EditText).Value

                                    If fullPath <> "" AndAlso
                                   fullPath.Contains("\") Then

                                        Dim lastIndex As Integer =
                                        fullPath.LastIndexOf("\")

                                        Dim fileName As String =
                                        fullPath.Substring(lastIndex + 1)

                                        CType(objMatrix.Columns.Item("FNAME").
                                    Cells.Item(pVal.Row).Specific,
                                    SAPbouiCOM.EditText).Value = fileName

                                        CType(objMatrix.Columns.Item("ATD").
                                    Cells.Item(pVal.Row).Specific,
                                    SAPbouiCOM.EditText).Value =
                                    DateTime.Now.ToString("yyyyMMdd")

                                        If objForm.Items.Exists("btn_Del") Then objForm.Items.Item("btn_Del").Enabled = True

                                    End If

                                End If

                            Catch ex As Exception

                                objMain.objApplication.StatusBar.SetText(
                                "Attachment Error : " & ex.Message,
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            End Try

                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                    If pVal.BeforeAction = False Then

                        Select Case pVal.ItemUID

                            Case "DocDate"

                                Try

                                    Dim docDate As String =
                                    CType(objForm.Items.Item("DocDate").Specific,
                                          SAPbouiCOM.EditText).Value

                                    If docDate = "" Then

                                        objMain.objApplication.StatusBar.SetText(
                                        "Document Date should not be empty.",
                                        SAPbouiCOM.BoMessageTime.bmt_Short,
                                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                    End If

                                Catch ex As Exception

                                End Try

                        End Select

                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                    If pVal.BeforeAction = False Then

                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent =
                        CType(pVal, SAPbouiCOM.IChooseFromListEvent)

                        Dim oDataTable As SAPbouiCOM.DataTable =
                        oCFLEvento.SelectedObjects

                        If oDataTable IsNot Nothing Then

                            Try

                                If pVal.ItemUID = "CardCode" Then

                                    Dim CardCode As String =
                                    oDataTable.GetValue("CardCode", 0).ToString()

                                    Dim CardName As String =
                                    oDataTable.GetValue("CardName", 0).ToString()

                                    CType(objForm.Items.Item("CardCode").Specific,
                                      SAPbouiCOM.EditText).Value = CardCode

                                    CType(objForm.Items.Item("CardName").Specific,
                                      SAPbouiCOM.EditText).Value = CardName

                                End If

                            Catch ex As Exception

                                objMain.objApplication.StatusBar.SetText(
                                ex.Message,
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                            End Try

                        End If

                    End If

            End Select

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(
            "Item Event Error : " & ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
#Region "Menu Event"

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent,
              ByRef BubbleEvent As Boolean)

        Try

            If pVal.MenuUID = "10X_LINE_CLR" _
            AndAlso pVal.BeforeAction = False Then

                Me.CreateForm()

            ElseIf pVal.MenuUID = "1282" _
            AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "TNX_PLCL" Then Exit Sub

                Me.SetDefault(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1281" _
            AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "TNX_PLCL" Then Exit Sub

            ElseIf pVal.MenuUID = "1292" _
            AndAlso pVal.BeforeAction = False Then

                objForm = objMain.objApplication.Forms.ActiveForm

                If objForm.TypeEx <> "TNX_PLCL" Then Exit Sub

                SetNewLine(objForm.UniqueID)
                SetNewLine1(objForm.UniqueID)
                SetNewLine2(objForm.UniqueID)

            ElseIf pVal.MenuUID = "1293" _
     AndAlso pVal.BeforeAction = True Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "TNX_PLCL" Then Exit Sub

                    BubbleEvent = False

                    objForm.Freeze(True)

                    If objForm.PaneLevel = 1 Then

                        objMatrix = CType(
                            objForm.Items.Item("matChk").Specific,
                            SAPbouiCOM.Matrix)

                        oDBs_Details =
                            objForm.DataSources.DBDataSources.Item("@TNX_PLCL_L")

                        Dim selectedRow As Integer =
                            objMatrix.GetNextSelectedRow(
                                0,
                                SAPbouiCOM.BoOrderType.ot_RowOrder)

                        If selectedRow <= 0 Then

                            objMain.objApplication.StatusBar.SetText(
                                "Please select Result row.",
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            Exit Try

                        End If

                        objMatrix.FlushToDataSource()

                        objMatrix.DeleteRow(selectedRow)

                        While oDBs_Details.Size > objMatrix.VisualRowCount

                            oDBs_Details.RemoveRecord(
                                oDBs_Details.Size - 1)

                        End While

                        If oDBs_Details.Size = 0 Then

                            oDBs_Details.InsertRecord(0)

                            oDBs_Details.SetValue("LineId", 0, "1")
                            oDBs_Details.SetValue("U_TestCode", 0, "")
                            oDBs_Details.SetValue("U_TestName", 0, "")
                            oDBs_Details.SetValue("U_Parameter", 0, "")
                            oDBs_Details.SetValue("U_TestMethod", 0, "")
                            oDBs_Details.SetValue("U_Unit", 0, "")
                            oDBs_Details.SetValue("U_MinValue", 0, "")
                            oDBs_Details.SetValue("U_MaxValue", 0, "")
                            oDBs_Details.SetValue("U_ActualValue", 0, "")
                            oDBs_Details.SetValue("U_Result", 0, "")
                            oDBs_Details.SetValue("U_Status", 0, "Pending")

                        End If

                        For i As Integer = 0 To oDBs_Details.Size - 1

                            oDBs_Details.SetValue(
                                "LineId",
                                i,
                                (i + 1).ToString())

                        Next

                        objMatrix.LoadFromDataSource()
                        objMatrix.AutoResizeColumns()

                    End If

                    If objForm.PaneLevel = 2 Then

                        objMatrix1 = CType(
                            objForm.Items.Item("Item_1").Specific,
                            SAPbouiCOM.Matrix)

                        oDBs_Details1 =
                            objForm.DataSources.DBDataSources.Item("@TNX_PLCL_EQP")

                        Dim selectedRow As Integer =
                            objMatrix1.GetNextSelectedRow(
                                0,
                                SAPbouiCOM.BoOrderType.ot_RowOrder)

                        If selectedRow <= 0 Then

                            objMain.objApplication.StatusBar.SetText(
                                "Please select Equipment row.",
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            Exit Try

                        End If

                        objMatrix1.FlushToDataSource()

                        objMatrix1.DeleteRow(selectedRow)

                        While oDBs_Details1.Size > objMatrix1.VisualRowCount

                            oDBs_Details1.RemoveRecord(
                                oDBs_Details1.Size - 1)

                        End While

                        If oDBs_Details1.Size = 0 Then

                            oDBs_Details1.InsertRecord(0)

                            oDBs_Details1.SetValue("LineId", 0, "1")
                            oDBs_Details1.SetValue("U_EquipCode", 0, "")
                            oDBs_Details1.SetValue("U_EquipName", 0, "")
                            oDBs_Details1.SetValue("U_CleaningLogNo", 0, "")
                            oDBs_Details1.SetValue("U_CleaningStatus", 0, "")
                            oDBs_Details1.SetValue("U_CalibDueDate", 0, "")
                            oDBs_Details1.SetValue("U_CalibStatus", 0, "")
                            oDBs_Details1.SetValue("U_ReadyStatus", 0, "")
                            oDBs_Details1.SetValue("U_Rmarks", 0, "")

                        End If

                        For i As Integer = 0 To oDBs_Details1.Size - 1

                            oDBs_Details1.SetValue(
                                "LineId",
                                i,
                                (i + 1).ToString())

                        Next

                        objMatrix1.LoadFromDataSource()
                        objMatrix1.AutoResizeColumns()

                    End If

                    If objForm.PaneLevel = 3 Then

                        objMatrix2 = CType(
                            objForm.Items.Item("Item_0").Specific,
                            SAPbouiCOM.Matrix)

                        oDBs_Attach =
                            objForm.DataSources.DBDataSources.Item("@TNX_PLCL_ATT")

                        Dim selectedRow As Integer =
                            objMatrix2.GetNextSelectedRow(
                                0,
                                SAPbouiCOM.BoOrderType.ot_RowOrder)

                        If selectedRow <= 0 Then

                            objMain.objApplication.StatusBar.SetText(
                                "Please select Attachment row.",
                                SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            Exit Try

                        End If

                        objMatrix2.FlushToDataSource()

                        objMatrix2.DeleteRow(selectedRow)

                        While oDBs_Attach.Size > objMatrix2.VisualRowCount

                            oDBs_Attach.RemoveRecord(
                                oDBs_Attach.Size - 1)

                        End While

                        If oDBs_Attach.Size = 0 Then

                            oDBs_Attach.InsertRecord(0)

                            oDBs_Attach.SetValue("LineId", 0, "1")
                            oDBs_Attach.SetValue("U_TPA", 0, "")
                            oDBs_Attach.SetValue("U_FN", 0, "")
                            oDBs_Attach.SetValue("U_FTT", 0, "")
                            oDBs_Attach.SetValue("U_ATD", 0, "")

                        End If

                        For i As Integer = 0 To oDBs_Attach.Size - 1

                            oDBs_Attach.SetValue(
                                "LineId",
                                i,
                                (i + 1).ToString())

                        Next

                        objMatrix2.LoadFromDataSource()
                        objMatrix2.AutoResizeColumns()

                    End If

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

            ElseIf pVal.MenuUID = "519" Then

                Try

                    objForm = objMain.objApplication.Forms.ActiveForm

                    If objForm.TypeEx <> "TNX_PLCL" Then Exit Sub

                    objMain.objApplication.StatusBar.SetText(
                    "Preview functionality triggered.",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                Catch ex As Exception

                    objMain.objApplication.StatusBar.SetText(
                    "Preview Error : " & ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                End Try

            End If

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

#End Region

#End Region
End Class

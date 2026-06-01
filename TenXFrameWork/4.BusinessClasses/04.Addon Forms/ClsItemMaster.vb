
Imports SAPbouiCOM
Imports SAPbobsCOM

Public Class ClsItemMaster

#Region "Declaration"

    Dim objForm As SAPbouiCOM.Form
    Dim oItem As SAPbouiCOM.Item

#End Region

    Public Sub ItemEvent(ByVal FormUID As String,
                         ByRef pVal As SAPbouiCOM.ItemEvent,
                         ByRef BubbleEvent As Boolean)

        Try

            Select Case pVal.EventType

                '====================================================
                ' FORM LOAD
                '====================================================

                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                    If pVal.BeforeAction = False Then

                        objForm = objMain.objApplication.Forms.Item(FormUID)

                        If objForm.TypeEx = "150" Then

                            AddItems(FormUID)

                        End If

                    End If

                '====================================================
                ' TAB CLICK
                '====================================================

                Case SAPbouiCOM.BoEventTypes.et_CLICK

                    If pVal.ItemUID = "PHARMA" And
                       pVal.BeforeAction = False Then

                        objForm = objMain.objApplication.Forms.Item(FormUID)

                        objForm.PaneLevel = 60

                    End If

            End Select

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Private Sub AddItems(ByVal FormUID As String)

        Try

            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oFolder As SAPbouiCOM.Folder

            '====================================================
            ' CHECK TAB ALREADY EXISTS
            '====================================================

            If ItemExists(objForm, "PHARMA") = True Then
                Exit Sub
            End If

            '====================================================
            ' CREATE PHARMA TAB
            '====================================================

            oItem = objForm.Items.Add("PHARMA",
            SAPbouiCOM.BoFormItemTypes.it_FOLDER)

            ' Reference Existing Folder

            oItem.Top = objForm.Items.Item("26").Top

            ' Change this value if tab position needs adjustment

            oItem.Left = objForm.Items.Item("26").Left + 300

            oItem.Width = 100

            oItem.Height = objForm.Items.Item("26").Height

            oItem.FromPane = 0
            oItem.ToPane = 0

            oItem.AffectsFormMode = False

            oFolder = objForm.Items.Item("PHARMA").Specific

            oFolder.Caption = "Pharma"

            oFolder.GroupWith("26")

            oFolder.Pane = 60

            '====================================================
            ' LABEL
            '====================================================

            oItem = objForm.Items.Add("LBLTYPE",
            SAPbouiCOM.BoFormItemTypes.it_STATIC)

            oItem.Top = 120
            oItem.Left = 20
            oItem.Width = 120
            oItem.Height = 15

            oItem.FromPane = 60
            oItem.ToPane = 60

            Dim lbl As SAPbouiCOM.StaticText

            lbl = objForm.Items.Item("LBLTYPE").Specific

            lbl.Caption = "Pharma Type"

            '====================================================
            ' EDIT TEXT
            '====================================================

            oItem = objForm.Items.Add("TXTPHTYP",
            SAPbouiCOM.BoFormItemTypes.it_EDIT)

            oItem.Top = 120
            oItem.Left = 140
            oItem.Width = 150
            oItem.Height = 15

            oItem.FromPane = 60
            oItem.ToPane = 60

            Dim txt As SAPbouiCOM.EditText

            txt = objForm.Items.Item("TXTPHTYP").Specific

        Catch ex As Exception

            objMain.objApplication.StatusBar.SetText(ex.Message,
            SAPbouiCOM.BoMessageTime.bmt_Short,
            SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Private Function ItemExists(ByVal oForm As SAPbouiCOM.Form,
                                ByVal ItemUID As String) As Boolean

        Try

            Dim oItem As SAPbouiCOM.Item

            oItem = oForm.Items.Item(ItemUID)

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

End Class


Option Strict Off
Option Explicit On
Public Class MainCls

#Region "Declaration"
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Public ExePath As String
    Public objUtilities As Utilities
    Public objDatabaseCreation As DatabaseCreation
    Public Shared ohtLookUpForm As Hashtable = New Hashtable
    'Addon Files
    'Public ObjFORCAST As clsFORCAST
    'Public ObjSUBTYPE As clssubmenutype
    'Public oClsParameterSelection As ClsParameterSelection
    'Public oApprovedForecating As ApprovedForecating
    'Public oSubParameterSelection As SubParameterSelection
    Public ObjCorporateTaxConfiguration As clsCorporateTexConfig
    Public ObjCorporateTaxCalculation As ClsCorporateTaxCalcu
    Public ObjclsPilotBatch As clsPilotBatch
    Public ObjclsSOPManagement As clsSOPManagement
    Public ObjclsTrainingPlan As clsTrainingPlan
    Public ObjclsTrainingExecution As clsTrainingExecution
    Public ObjclsTrainingCertificate As clsTrainingCertificate
    Public ObjclsAudit As ClsAuditChecklist
    Public ObjclsEmployeeTraining As clsEmployeeTraining
    Public objClerenceMaster As ClerenceMaster
    Public objproductionStage As ProductionStage
    Public objEquipmentMaster As EquipmentMaster
    Public objLineClerencechecklist As LineClerenceChecklist
    Public objYeildtolorance As Yeildtolorance
    Public objRegulatoryAuthority As RegulatoryAuthority
    Public objcountryregulatoryconfig As Countryregulatoryconfig
    Public objDowntimeReason As DowntimeReason
    Public ObjclsDevidation As ClsDevidationCategory
    Public ObjCAPAMaster As ClsCAPAMaster
    Public ObjSOP As ClsSOPMaster
    Public ObjRisk As ClsRiskMaster
    Public ObjSample As ClsSampleType
    'Manohar
    Public objProductRegistration As ProductRegistration

    'vsm
    Public objInprogresschecklist As Inprocesschecklist
    Public ObjclsMstrProductCategory As ClsMstrProductCategory
    Public ObjClsMstrDosageForm As ClsMstrDosageForm
    'vsm
    Public ObjClsCorpTax As ClsCorpTax
    Public ObjclsFtaVat As clsFtaVat
    Public ObjClsLkMstr As ClsLkMstr
    Public ObjclsCOMTemplate As ClsCOMTemplate

    Public objLicenceNew As Cfrm_LicenceAdministrationNew
    Public objDevice As DeviceMaster
    Public objLicenceAdministration As cfrm_LicenseAdministration
    Public Objclsmanage As ClsCOAManagement
    Public Objclscontrol As ClsChangeControl
    Public objSpecificationMaster As Specificationmaster
    Public ObjLabTesting As LabTesting
    Public YieldAnalysis As Cfrm_YieldAnalysis
    Public objPharmaDispensing As ClsPharmaDispensing
    Public objPhamaBpr As ClsPharmaBPR
    Public SampleCollection As Cls_SampleCollection

    Public SampleRegistration As Cls_SampleRegistration
    Public ObjclsMstrCorrectiveAction As clsMstrCorrectiveAction
    Public ObjclsMstrPreventiveAction As clsMstrPreventiveAction
    Public ObjclsMstrDocumentNumSetup As clsMstrDocumentNumSetup
    Public ObjclsMstrRetentionPolicy As clsMstrRetentionPolicy
    Public ObjclsMstrIncidentCategory As clsMstrIncidentCategory
    Public ObjclsMstrComplianceSett As clsMstrComplianceSett
    Public ObjclsElectronicSignaturePolicy As clsElectronicSignaturePolicy


    Public ObjclsBatch As ClsBatchRelease

    Public ObjclsCAPAManage As ClsCAPAManagement
    Public ObjBmr As Cfrm_BMR
    'vsm
    Public Stabilityprotocal As Cls_StabilityProtocol
    Public Stabilitystudy As ClS_StabilityStudy
    Public Shelflife As Cls_ShelfLifeAnalysis
    Public ObjVatReport As ClsVatReports
    Public objformulacosting As ClsFormulaCosting
    Public ObjPayloadD As ClsPayloadD
    ' Public ObjClsCorpTax As ClsCorpTax
    Public ObjClsCorpTaxMstr As ClsCropTaxMstr
    Public ObjclsFtaVatMstr As ClsFtaVatMstr


    Public objEstimation As clsEstimation
    Public objARInvoice As clsARInvoice
    Public objARCreditMemo As clsARCreditMemo
    Public objARDownPayment As clsARDownPayment
    Public objPayLoad As ClsPayLoad

    Public ObjApproval As clsApproval
    Public ObjGRID As clsGRID
    Public ObjADR As clsApprovalDecision
    Public ObjAStg As ClsAPPROVALSTAGES
    Public ObjAPTEMP As clsAPPROVALTEMP
    Public ObjGRIDES As ClsGRIDES
    Public ObjDraftProcedure As ClsDraftProcedureStages
    Public ObjVIEW As ClsVIEWFORAPP
    Public objSAPAlertWindow As clsSAPAlertWindow
    Public objFormulaMaster As ClsFormulaMaster

    Public objExperimentManagement As ClsExperimentManagement
    Public Objclsvalidation As ClsValidationManage
    Public ObjclsIncident As ClsIncidentManage

    Public objAREinvoice As CLSEinvoiceButton

    Public objInvPost As ClsInvPost
    Public objOnboarding As ClsOnboarding
    'Vamshi Sai
    Public objInvoicePosting As ClsInvoicePsoting

    Public oGeneralService As SAPbobsCOM.GeneralService
    Public oGeneralData As SAPbobsCOM.GeneralData
    Public oSons As SAPbobsCOM.GeneralDataCollection
    Public oSon As SAPbobsCOM.GeneralData
    Public oChildren As SAPbobsCOM.GeneralDataCollection
    Public oChild As SAPbobsCOM.GeneralData
    Public sCmp As SAPbobsCOM.CompanyService
    Public oGeneralParams As SAPbobsCOM.GeneralDataParams
    Public oGeneralService1 As SAPbobsCOM.GeneralService
    Public oGeneralData1 As SAPbobsCOM.GeneralData
    Public oChildren1 As SAPbobsCOM.GeneralDataCollection
    Public oChild1 As SAPbobsCOM.GeneralData
    Public sCmp1 As SAPbobsCOM.CompanyService
    Public oGeneralParams1 As SAPbobsCOM.GeneralDataParams
    Public oitem As SAPbouiCOM.Item
    Dim SOSeries As String = ""
    Dim SODocNum As String = ""
    Dim PaymentType As String = ""
#End Region
    Public Sub New()
        objUtilities = New Utilities
        objDatabaseCreation = New DatabaseCreation
    End Sub

#Region "Initialilse"
    Public Function Initialise() As Boolean
        objApplication = objUtilities.GetApplication()
        If objApplication Is Nothing Then Return False
        objCompany = objUtilities.GetCompany(objApplication)
        If objCompany Is Nothing Then : Return False : Exit Function : End If
        If Not objDatabaseCreation.CreateTables() Then Return False
        Me.LoadFromXML("Menu.xml")
        If objMain.objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            IsNull = "IFNULL"
            GetDate = "NOW"
            HanaLen = "Length"
            HanaInt = "Integer"
            Hana = "Hana"
            Concate = "||"
            SelectCase = ""
        Else
            IsNull = "IsNull"
            GetDate = "GetDate"
            HanaLen = "Len"
            HanaInt = "Int"
            Sql = "Sql"
            Concate = "+"
            SelectCase = "Select"
        End If
        CreateObjects()

        'Try
        '    If objMain.objApplication.Menus.Exists("ME_TenXFrameWork") = True Then
        '        objMain.objApplication.Menus.Item("ME_TenXFrameWork").Image = System.Windows.Forms.Application.StartupPath & "/TenXFrameWork.JPG"
        '    End If
        'Catch ex As Exception
        'End Try
        'objApplication.StatusBar.SetText(Me.GetUserName() + " ! :-)  TenXFrameWork  Addon has been Connected.......You can continue your work ..........", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        objApplication.StatusBar.SetText("Pharma Suite Addon is connected....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True
    End Function
#End Region

    'Public Function GetUserName1() As String
    '    Dim oRsGetUserNames As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    Try
    '        Dim mHour As Integer
    '        Dim Greeting As String
    '        mHour = Hour(Now)
    '        If mHour < 12 Then
    '            Greeting = "Good Morning"
    '        ElseIf mHour < 16 Then
    '            Greeting = "Good Afternoon"
    '        Else
    '            Greeting = "Good Evening"
    '        End If

    '        Dim GetUserNames As String = "Select ""U_NAME"" From OUSR Where ""USER_CODE"" = '" & objMain.objCompany.UserName & "'"
    '        oRsGetUserNames.DoQuery(GetUserNames)
    '        Return Greeting + "...." + oRsGetUserNames.Fields.Item(0).Value.ToString
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        oRsGetUserNames = Nothing
    '    End Try
    'End Function

#Region "Create Object"
    Private Sub CreateObjects()
        'ObjFORCAST = New clsFORCAST
        'oClsParameterSelection = New ClsParameterSelection
        'oSubParameterSelection = New SubParameterSelection
        'oApprovedForecating = New ApprovedForecating
        'ObjSUBTYPE = New clssubmenutype
        ObjClsCorpTax = New ClsCorpTax
        ObjclsFtaVat = New clsFtaVat
        Objclsvalidation = New ClsValidationManage
        ObjclsIncident = New ClsIncidentManage
        ObjApproval = New clsApproval
        ObjGRID = New clsGRID
        ObjADR = New clsApprovalDecision
        ObjAStg = New ClsAPPROVALSTAGES
        ObjAPTEMP = New clsAPPROVALTEMP
        ObjGRIDES = New ClsGRIDES
        ObjDraftProcedure = New ClsDraftProcedureStages
        ObjVIEW = New ClsVIEWFORAPP
        objSAPAlertWindow = New clsSAPAlertWindow
        ObjclsFtaVatMstr = New ClsFtaVatMstr
        ObjClsCorpTaxMstr = New ClsCropTaxMstr
        objEstimation = New clsEstimation
        objARInvoice = New clsARInvoice
        objARCreditMemo = New clsARCreditMemo
        objARDownPayment = New clsARDownPayment
        objAREinvoice = New CLSEinvoiceButton
        objPayLoad = New ClsPayLoad
        ObjCorporateTaxConfiguration = New clsCorporateTexConfig
        ObjCorporateTaxCalculation = New ClsCorporateTaxCalcu
        ObjVatReport = New ClsVatReports
        OBJFormulaCosting = New ClsFormulaCosting
        objInvPost = New ClsInvPost
        objOnboarding = New ClsOnboarding
        ObjPayloadD = New ClsPayloadD
        'Vamshi Sai
        objInvoicePosting = New ClsInvoicePsoting
        SampleCollection = New Cls_SampleCollection

        SampleRegistration = New Cls_SampleRegistration
        ObjclsCOMTemplate = New ClsCOMTemplate

        ObjclsDevidation = New ClsDevidationCategory

        ObjclsAudit = New ClsAuditChecklist
        objProductRegistration = New ProductRegistration

        ObjclsMstrCorrectiveAction = New clsMstrCorrectiveAction
        ObjclsMstrPreventiveAction = New clsMstrPreventiveAction
        ObjclsMstrDocumentNumSetup = New clsMstrDocumentNumSetup
        ObjclsMstrRetentionPolicy = New clsMstrRetentionPolicy
        ObjclsMstrIncidentCategory = New clsMstrIncidentCategory
        ObjclsMstrComplianceSett = New clsMstrComplianceSett
        ObjclsElectronicSignaturePolicy = New clsElectronicSignaturePolicy
        ObjclsMstrProductCategory = New ClsMstrProductCategory
        ObjClsMstrDosageForm = New ClsMstrDosageForm
        'vsm
        ObjClsCorpTax = New ClsCorpTax
        ObjclsFtaVat = New clsFtaVat
        ObjClsLkMstr = New ClsLkMstr

        objproductionStage = New ProductionStage
        objEquipmentMaster = New EquipmentMaster
        objLineClerencechecklist = New LineClerenceChecklist
        objRegulatoryAuthority = New RegulatoryAuthority
        objcountryregulatoryconfig = New Countryregulatoryconfig
        'objRegulatoryDoctype = New RegulatoryDocumentType
        objInprogresschecklist = New Inprocesschecklist
        objClerenceMaster = New ClerenceMaster
        objYeildtolorance = New Yeildtolorance

        objLicenceNew = New Cfrm_LicenceAdministrationNew
        objDevice = New DeviceMaster
        objLicenceAdministration = New cfrm_LicenseAdministration
        objFormulaMaster = New ClsFormulaMaster
        objExperimentManagement = New ClsExperimentManagement
        'vsm
        Objclsmanage = New ClsCOAManagement
        Objclscontrol = New ClsChangeControl
        ObjclsPilotBatch = New clsPilotBatch
        ObjclsSOPManagement = New clsSOPManagement
        ObjclsTrainingPlan = New clsTrainingPlan
        ObjclsTrainingExecution = New clsTrainingExecution
        ObjclsTrainingCertificate = New clsTrainingCertificate
        ObjclsEmployeeTraining = New clsEmployeeTraining
        objSpecificationMaster = New Specificationmaster
        ObjLabTesting = New LabTesting
        ObjclsBatch = New ClsBatchRelease
        ObjclsCAPAManage = New ClsCAPAManagement
        ObjBmr = New Cfrm_BMR
        YieldAnalysis = New Cfrm_YieldAnalysis
        objPhamaBpr = New ClsPharmaBPR
        objPharmaDispensing = New ClsPharmaDispensing

        Stabilitystudy = New ClS_StabilityStudy
        Stabilityprotocal = New Cls_StabilityProtocol
        Shelflife = New Cls_ShelfLifeAnalysis
    End Sub
#End Region

#Region "Create UDO"
    Public Sub CreateApprovalTemplatesUDO()
        If Not Me.UDOExists("SBOAPPUDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("SBOAPPUDO", "SBOAPPUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "SBO_APPHDR", "SBO_APPREQ", "SBO_APPDOC", "SBO_APPAUT")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub TNXQASampleRegistration()
        If Not Me.UDOExists("TNX_QA_SAMPLE") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_QA_SAMPLE", "TNX QA Sample Registration UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_QASMPH", "TNX_QASMPL", "TNX_QASMPA")
            findaliasdescription = Nothing
        End If
    End Sub



    Public Sub TNXQCSampleCollection()
        If Not Me.UDOExists("TNX_QC_SC") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_QC_SC", "TNX QC Sample Collection", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_QCSC_H", "TNX_QCSC_L", "TNX_QCSC_COC", "TNX_QCSC_ATT")
            findaliasdescription = Nothing
        End If
    End Sub


    Public Sub TNXPharmaYieldAnalysis()
        If Not Me.UDOExists("UDO_TNX_PYLD") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("UDO_TNX_PYLD", "TNX Pharma Yield Analysis", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_PYLD_H", "TNX_PYLD_MAT", "TNX_PYLD_OUT", "TNX_PYLD_VAR", "TNX_PYLD_APR", "TNX_PYLD_ATT")
            findaliasdescription = Nothing
        End If
    End Sub


    Public Sub CreateFormulaMasterUDO()
        If Not Me.UDOExists("TNX_PHUDO") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDONoLog("TNX_PHUDO", "TNX_PHUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_PH_FORMULA", "TNX_PH_FORMULA_D1", "TNX_PH_FORMULA_D2", "TNX_PH_FORMULA_D3", "TNX_PH_FORMULA_D4", "TNX_PH_FORMULA_D5", "TNX_PH_FORMULA_D6")
            findaliasdescription = Nothing
        End If
    End Sub

    Public Sub ProductionStageMasterUDO()

        Try

            If Not Me.UDOExists("TNX_PSTG") Then

                Dim findAliasNDescription = New String(,) {
              {"DocNum", "Document Number"},
              {"U_StageSeq", "Stage Sequence"},
              {"U_DosageFrm", "Dosage Form"}
          }

                Me.registerUDO(
              "TNX_PSTG",
              "Production Stage Master",
              SAPbobsCOM.BoUDOObjType.boud_Document,
              findAliasNDescription,
              "TNX_PSTG_H",
              "TNX_PSTG_L")

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub


    Public Sub CreateStabilityProtocolUDO()

        If Not Me.UDOExists("UDO_STAB_PROTOCOL") Then
            Dim findAliasDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDONoLog("UDO_STAB_PROTOCOL", "10X Stability Protocol", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasDescription, "TNX_STAB_PROTO", "TNX_STAB_PROTO_T", "TNX_STAB_PROTO_S")
            findAliasDescription = Nothing

        End If

    End Sub

    Public Sub CreateStabilityStudyUDO()
        If Not Me.UDOExists("UDO_STAB_STUDY") Then
            Dim findAliasDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("UDO_STAB_STUDY", "10X Stability Study", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasDescription, "TNX_STAB_STUDY", "TNX_STAB_STUDY_B", "TNX_STAB_STUDY_C", "TNX_STAB_STUDY_T")
            findAliasDescription = Nothing
        End If
    End Sub

    Public Sub CreateShelfLifeUDO()
        If Not Me.UDOExists("UDO_STAB_SHELF") Then
            Dim findAliasDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("UDO_STAB_SHELF", "10X Shelf Life Analysis", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasDescription, "TNX_STAB_SHELF", "TNX_STAB_SHELF_L")
            findAliasDescription = Nothing
        End If
    End Sub


    '============================================================
    ' Regulatory Document Type Master
    '============================================================
    Public Sub TNXRegulatoryDocumentUDO()
        If Not Me.UDOExists("UDO_REG_DOCTYP") Then

            Dim findAlias3 = New String(,) {{"Code", "Code"}}

            Me.registerUDO("UDO_REG_DOCTYP",
                        "Regulatory Document Type Master",
                        SAPbobsCOM.BoUDOObjType.boud_MasterData,
                        findAlias3,
                        "TNX_REG_DOCTYP")

            findAlias3 = Nothing

        End If
    End Sub

    '============================================================
    ' Dossier Section Master
    '============================================================
    Public Sub TNXDossierUDO()
        If Not Me.UDOExists("UDO_REG_DOSSEC") Then

            Dim findAlias4 = New String(,) {{"Code", "Code"}}

            Me.registerUDONoLog("UDO_REG_DOSSEC",
                        "Dossier Section Master",
                        SAPbobsCOM.BoUDOObjType.boud_MasterData,
                        findAlias4,
                        "TNX_REG_DOSSEC")

            findAlias4 = Nothing

        End If
    End Sub

    '============================================================
    ' CTD/eCTD Template Master
    '============================================================
    Public Sub TNXTemplateUDO()
        If Not Me.UDOExists("UDO_REG_CTDTMP") Then

            Dim findAlias5 = New String(,) {{"DocNum", "DocNum"}}

            Me.registerUDO("UDO_REG_CTDTMP",
                        "CTD eCTD Template Master",
                        SAPbobsCOM.BoUDOObjType.boud_Document,
                        findAlias5,
                        "TNX_REG_CTDTMP",
                        "TNX_REG_CTDL")

            findAlias5 = Nothing

        End If
    End Sub

    '============================================================
    ' Artwork Type Master
    '============================================================
    Public Sub TNXArtworkUDO()
        If Not Me.UDOExists("UDO_REG_ARTTYP") Then

            Dim findAlias6 = New String(,) {{"Code", "Code"}}

            Me.registerUDONoLog("UDO_REG_ARTTYP",
                        "Artwork Type Master",
                        SAPbobsCOM.BoUDOObjType.boud_MasterData,
                        findAlias6,
                        "TNX_REG_ARTTYP")

            findAlias6 = Nothing

        End If

    End Sub
    Public Sub TNXSubmissionUDO()
        '============================================================
        ' Submission Type Master
        '============================================================
        If Not Me.UDOExists("UDO_REG_SUBTYP") Then

            Dim findAlias7 = New String(,) {{"Code", "Code"}}

            Me.registerUDO("UDO_REG_SUBTYP",
                        "Submission Type Master",
                        SAPbobsCOM.BoUDOObjType.boud_MasterData,
                        findAlias7,
                        "TNX_REG_SUBTYP")

            findAlias7 = Nothing

        End If
    End Sub
    Public Sub TNXRegistrationUDO()
        '============================================================
        ' Registration Status Master
        '============================================================
        If Not Me.UDOExists("UDO_REG_STAT") Then

            Dim findAlias8 = New String(,) {{"Code", "Code"}}

            Me.registerUDONoLog("UDO_REG_STAT",
                        "Registration Status Master",
                        SAPbobsCOM.BoUDOObjType.boud_MasterData,
                        findAlias8,
                        "TNX_REG_STAT")

            findAlias8 = Nothing

        End If
    End Sub
    Public Sub LineClearanceUDO()

        If Not Me.UDOExists("TNX_PLCL") Then

            Dim findAliasNDescription = New String(,) {{"DocNum", "Document Number"}}

            Me.registerUDO("TNX_PLCL", "10X Pharma Line Clearance", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_PLCL_H", "TNX_PLCL_L", "TNX_PLCL_EQP", "TNX_PLCL_ATT")

            findAliasNDescription = Nothing

        End If

    End Sub
    Public Sub TNXApprovalUDO()

        '============================================================
        ' Approval Matrix Master
        '============================================================
        If Not Me.UDOExists("UDO_REG_APRMAT") Then

            Dim findAlias9 = New String(,) {{"DocNum", "DocNum"}}

            Me.registerUDO("UDO_REG_APRMAT",
                        "Approval Matrix Master",
                        SAPbobsCOM.BoUDOObjType.boud_Document,
                        findAlias9,
                        "TNX_REG_APRH",
                        "TNX_REG_APRL")

            findAlias9 = Nothing

        End If
    End Sub
    Public Sub TNXRegulatoryUDO()

        '============================================================
        ' Regulatory Authority Master
        '============================================================
        If Not Me.UDOExists("UDO_REG_AUTH") Then

            Dim findAlias1 = New String(,) {{"DocNum", "DocNum"}}

            Me.registerUDONoLog("UDO_REG_AUTH",
                      "Regulatory Authority Master",
                      SAPbobsCOM.BoUDOObjType.boud_Document,
                      findAlias1,
                      "TNX_REG_AUTH")

            findAlias1 = Nothing

        End If
    End Sub
    Public Sub EquipmentMasterUDO1()

        Try

            If Not Me.UDOExists("TNX_PEQP") Then

                Dim findAliasNDescription = New String(,) {
              {"DocNum", "Document Number"},
              {"U_EquipType", "Equipment Type"},
              {"U_AreaCode", "Production Area"}
          }

                Me.registerUDONoLog(
              "TNX_PEQP",
              "Equipment Master",
              SAPbobsCOM.BoUDOObjType.boud_Document,
              findAliasNDescription,
              "TNX_PEQP_H",
              "TNX_PEQP_L")

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    Public Sub YieldToleranceMasterUDO()

        Try

            If Not Me.UDOExists("TNX_PYTM") Then

                Dim findAliasNDescription = New String(,) {
              {"DocNum", "Document Number"},
              {"U_ItemCode", "Item Code"},
              {"U_ItemGroup", "Product Group"}
          }

                Me.registerUDO(
              "TNX_PYTM",
              "Yield Tolerance Master",
              SAPbobsCOM.BoUDOObjType.boud_Document,
              findAliasNDescription,
              "TNX_PYTM_H",
              "TNX_PYTM_L")

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    Public Sub InProcessQCChecklistUDO()

        Try

            If Not Me.UDOExists("TNX_PIQC") Then

                Dim findAliasNDescription = New String(,) {
              {"DocNum", "Document Number"},
              {"U_ItemCode", "Item Code"},
              {"U_StageCode", "Stage Code"}
          }

                Me.registerUDO(
              "TNX_PIQC",
              "Inprocess QC Checklist",
              SAPbobsCOM.BoUDOObjType.boud_Document,
              findAliasNDescription,
              "TNX_PIQC_H",
              "TNX_PIQC_L")

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    Public Sub LineClearanceChecklistUDO()

        Try

            If Not Me.UDOExists("TNX_PLCC") Then

                Dim findAliasNDescription = New String(,) {
              {"DocNum", "Document Number"},
              {"U_ClrType", "Clearance Type"},
              {"U_StageCode", "Stage Code"}
          }

                Me.registerUDO(
              "TNX_PLCC",
              "Line Clearance Checklist",
              SAPbobsCOM.BoUDOObjType.boud_Document,
              findAliasNDescription,
              "TNX_PLCC_H",
              "TNX_PLCC_L")

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    Public Sub DowntimeReasonMasterUDO()

        Try

            If Not Me.UDOExists("TNX_PDTR") Then

                Dim findAliasNDescription = New String(,) {
              {"Code", "Code"},
              {"Name", "Name"},
              {"U_Category", "Category"}
          }

                Me.registerUDONoLog(
              "TNX_PDTR",
              "Downtime Reason Master",
              SAPbobsCOM.BoUDOObjType.boud_MasterData,
              findAliasNDescription,
              "TNX_PDTR_H")

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    Public Sub CleaningMethodMasterUDO()

        Try

            If Not Me.UDOExists("TNX_PCLMUDO") Then

                Dim findAliasNDescription = New String(,) {
              {"Code", "Code"},
              {"Name", "Name"}
          }

                Me.registerUDONoLog(
              "TNX_PCLMUDO",
              "Cleaning Method Master",
              SAPbobsCOM.BoUDOObjType.boud_MasterData,
              findAliasNDescription,
              "TNX_PCLM_H",
              "TNX_PCLM_L")

            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub
    '============================================================


    '============================================================
    ' Country Regulatory Configuration
    '============================================================
    Public Sub TNXCountryRegulatoryUDO()
        If Not Me.UDOExists("UDO_REG_CNFG") Then

            Dim findAlias2 = New String(,) {{"DocNum", "DocNum"}}

            Me.registerUDO("UDO_REG_CNFG",
                      "Country Regulatory Configuration",
                      SAPbobsCOM.BoUDOObjType.boud_Document,
                      findAlias2,
                      "TNX_REG_CNFG",
                      "TNX_REG_CNFL")

            findAlias2 = Nothing

        End If

    End Sub

    Public Sub CreatePharmaBMRExecutionUDO()
        If Not Me.UDOExists("UDO_TNX_PBMR") Then
            Dim findAliasDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("UDO_TNX_PBMR", "10X Pharma BMR Execution", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasDescription, "TNX_PBMR_H", "TNX_PBMR_STAGE", "TNX_PBMR_MAT", "TNX_PBMR_EQP", "TNX_PBMR_IPQC", "TNX_PBMR_DEV", "TNX_PBMR_APP")
            findAliasDescription = Nothing
        End If
    End Sub
    Public Sub QCLabTestingUDO()

        If Not Me.UDOExists("TNXPH_QCLAB") Then

            Dim findAliasNDescription = New String(,) {{"DocNum", "Document Number"}}
            Me.registerUDO("TNXPH_QCLAB", "10X Pharma QC Lab Testing", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNXPH_QCLABH", "TNXPH_QCLABL", "TNXPH_QCLABATT", "TNXPH_QCLABAPP")

            findAliasNDescription = Nothing

        End If

    End Sub
    Public Sub LkMsterUDO()
        If Not Me.UDOExists("TNX_LKMTR") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNX_LKMTR", "TNX_LKMTR", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_LKMTR")
            findaliasdescription = Nothing
        End If
    End Sub

    Public Sub QABatchUDO()
        If Not Me.UDOExists(" TNXQABRUDO") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNXQABRUDO", "TNXQABRUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_QABR_H", "TNX_QABR_TST", "TNX_QABR_APR", "TNX_QABR_DOC")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CAPAManageUDO()
        If Not Me.UDOExists("TNX_CAPA_UDO") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_CAPA_UDO", "TNX_CAPA_UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_CAPAH", "TNX_CAPAL", "TNX_CAPAE", "TNX_CAPAA", "TNX_CAPAW")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub FtaVatUDO1()
        If Not Me.UDOExists("TNX_FTAVAT") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNX_FTAVAT", "TNX_FTAVAT", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_FTAVAT")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CorpTaxUDO1()
        If Not Me.UDOExists("TNX_CORPTAX") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNX_CORPTAX", "TNX_CORPTAX", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_CORPTAX")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CreateAPPROVALSTAGESUDO()
        If Not Me.UDOExists("SBO_ASTGUDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("SBO_ASTGUDO", "SBO_ASTGUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "SBO_AST", "SBO_AST_C0")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub FtaVatUDO()
        If Not Me.UDOExists("UDO_FTAVAT") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("UDO_FTAVAT", "UDO_FTAVAT", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "FTAVAT")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CorpTaxUDO()
        If Not Me.UDOExists("UDO_CORPTAX") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("UDO_CORPTAX", "UDO_CORPTAX", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "CORPTAX")
            findaliasdescription = Nothing
        End If
    End Sub

    Public Sub DDUDO()
        If Not Me.UDOExists("SBO_DDUDO") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("SBO_DDUDO", "SBO_DDUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "SBO_DD", "SBO_DD1")
            findAliasNDescription = Nothing
        End If

    End Sub
    Public Sub CreateONBPUDO()
        If Not Me.UDOExists("TNX_ONBP_UDO") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNX_ONBP_UDO", "Onboarding Process UDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_ONBP", "TNX_ONBP_C0")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CreateIPUDO()
        If Not Me.UDOExists("TNX_IPUDO") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_IPUDO", "Invoice PostingUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_IP", "TNX_IP_C0")
            findaliasdescription = Nothing
        End If
    End Sub


    Public Sub ManagementUDO()
        If Not Me.UDOExists("TNXCOAUDO") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNXCOAUDO", "TNXCOAUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_COA_H", "TNX_COA_T", "TNX_COA_A", "TNX_COA_APP")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub IncidentUDO()
        If Not Me.UDOExists("TNX_INC_UDO ") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_INC_UDO", "TNX_INC_UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_INCH", "TNX_INCL", "TNX_INCACT", "TNX_INCAP", "TNX_INCAT")
            findaliasdescription = Nothing
        End If
    End Sub

    Public Sub CreatePilotBatcUDO()
        If Not Me.UDOExists("UDO_TNX_PILOT") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDONoLog("UDO_TNX_PILOT", "UDO_TNX_PILOT", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_PB_HDR", "TNX_PB_YIELD", "TNX_PB_MAT", "TNX_PB_ISS", "TNX_PB_PROC", "TNX_PB_QC")
            findaliasdescription = Nothing
        End If
    End Sub

    Public Sub TNXTrainingPlanUDO()

        If Not Me.UDOExists("UDO_TNX_TRNPLAN") Then

            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}

            Me.registerUDO("UDO_TNX_TRNPLAN", "10X Training Plan", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_TRNPH", "TNX_TRNPL", "TNX_TRNP_Att")

        End If

    End Sub

    Public Sub CreateTrainingMatrUDO()
        If Not Me.UDOExists("UDO_TNX_TRNMAT") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDONoLog("UDO_TNX_TRNMAT", "TENX Employee Training", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_TRNMH", "TNX_TRNML")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CreateTrainingCertificateUDO()
        If Not Me.UDOExists("UDO_TNX_TRNCERT") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDONoLog("UDO_TNX_TRNCERT", "10X Training Certificate", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_TRNCH", "TNX_TRNCL")
            findaliasdescription = Nothing
        End If
    End Sub

    Public Sub SpecificMasterUDO()
        If Not Me.UDOExists("XPH_QSPEC") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("XPH_QSPEC", "XPH_QSPEC ", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "TNX_PH_QSPECH", "TNX_PH_QSPECL", "TNX_PH_QSPECM", "TNX_QSPECM_ATT")
            findAliasNDescription = Nothing
        End If
    End Sub
    Public Sub CreateTrainingExecutionUDO()
        If Not Me.UDOExists("UDO_TNX_TRNEXE") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDONoLog("UDO_TNX_TRNEXE", "10X Training Execution", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_TRNEH", "TNX_TRNEL", "TNX_TRNASM")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CreateSOPUDO()
        If Not Me.UDOExists("UDO_TENX_SOP") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDONoLog("UDO_TENX_SOP", "UDO_TENX_SOP", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_SOPH", "TNX_SOP_REV", "TNX_SOP_APR", "TNX_SOP_TRN", "TNX_SOP_DIST", "TNX_SOP_CAT")
            findaliasdescription = Nothing
        End If
    End Sub

    Public Sub ValidationUDO()
        If Not Me.UDOExists("TNX_VAL_UDO ") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_VAL_UDO", "TNX_VAL_UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_VALH", "TNX_VALP", "TNX_VALSTEP", "TNX_VALRES", "TNX_VALAP", "TNX_VALDOC", "TNX_VALAPP")
            findaliasdescription = Nothing
        End If
    End Sub

    Public Sub ControlUDO()
        If Not Me.UDOExists("TNX_CHG_UDO") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_CHG_UDO", "TNX_CHG_UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_CHGH", "TNX_CHGIMP", "TNX_CHGAP", "TNX_CHGTASK", "TNX_CHGVER", "TNX_CHGDOC", "TNX_CHGATT")
            findaliasdescription = Nothing
        End If
    End Sub


    '=========================================================
    ' 1. SOP Category Master
    '=========================================================
    Public Sub SOPCategoryMasterUDO()
        If Not Me.UDOExists("TNX_SOPCAT_UDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_SOPCAT_UDO", "SOPCategoryMaster", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_SOPCAT")
            findAliasNDescription = Nothing
        End If
    End Sub

    '=========================================================
    ' 2. Department Master
    '=========================================================
    Public Sub DepartmentMasterUDO()
        If Not Me.UDOExists("TNX_DEPT_UDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_DEPT_UDO", "DepartmentMaster", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_DEPT")
            findAliasNDescription = Nothing
        End If
    End Sub

    '=========================================================
    ' 3. Training Type Master
    '=========================================================
    Public Sub TrainingTypeMasterUDO()
        If Not Me.UDOExists("TNX_TRNTYP_UDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_TRNTYP_UDO", "TrainingTypeMaster", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_TRNTYP")
            findAliasNDescription = Nothing
        End If
    End Sub

    '=========================================================
    ' 4. Validation Type Master
    '=========================================================
    Public Sub ValidationTypeMasterUDO()
        If Not Me.UDOExists("TNX_VALTYP_UDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_VALTYP_UDO", "ValidationTypeMaster", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_VALTYP")
            findAliasNDescription = Nothing
        End If
    End Sub

    '=========================================================
    ' 5. Equipment Master
    '=========================================================
    Public Sub EquipmentMasterUDO()
        If Not Me.UDOExists("TNX_EQP_UDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_EQP_UDO", "EquipmentMaster", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_EQP")
            findAliasNDescription = Nothing
        End If
    End Sub

    '=========================================================
    ' 6. Risk Classification Master
    '=========================================================
    Public Sub RiskClassificationMasterUDO()
        If Not Me.UDOExists("TNX_RISK_UDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_RISK_UDO", "RiskClassificationMaster", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_RISK")
            findAliasNDescription = Nothing
        End If
    End Sub

    '=========================================================
    ' 7. CAPA Category Master
    '=========================================================
    Public Sub CAPACategoryMasterUDO()
        If Not Me.UDOExists("TNX_CAPACAT_UDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_CAPACAT_UDO", "CAPACategoryMaster", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_CAPACAT")
            findAliasNDescription = Nothing
        End If
    End Sub

    '=========================================================
    ' 8. Audit Type Master
    '=========================================================
    Public Sub AuditTypeMasterUDO()
        If Not Me.UDOExists("TNX_AUDTYP_UDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_AUDTYP_UDO", "AuditTypeMaster", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_AUDTYP")
            findAliasNDescription = Nothing
        End If
    End Sub

    '=========================================================
    ' 9. Root Cause Master
    '=========================================================
    Public Sub RootCauseMasterUDO()
        If Not Me.UDOExists("TNX_ROOT_UDO") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_ROOT_UDO", "RootCauseMaster", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "TNX_ROOT")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub FormulaCostingUDO()
        If Not Me.UDOExists("UDO_TNX_FRM_VER") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("UDO_TNX_FRM_VER", "UDO_TNX_FRM_VER", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_FRM_VER_H", "TNX_FRM_VER_D1", "TNX_FRM_VER_D2", "TNX_FRM_VER_D3", "TNX_FRM_VER_D4", "TNX_FRM_VER_D5", "TNX_FRM_VER_AUD")
            findaliasdescription = Nothing
        End If
    End Sub

    Public Sub ExperimentManagementUDO()
        If Not Me.UDOExists("UDO_TNX_EXP") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}

            Me.registerUDO(
            "UDO_TNX_EXP",
            "UDO_TNX_EXP",
            SAPbobsCOM.BoUDOObjType.boud_Document,
            findaliasdescription,
            "TNX_EXP_HDR",
            "TNX_EXP_ING",
            "TNX_EXP_PROC",
            "TNX_EXP_TEST",
            "TNX_EXP_OBS",
            "TNX_EXP_ATTACH1")

            findaliasdescription = Nothing
        End If
    End Sub

    'Submission Tracker UDO
    Public Sub SubmissionTrackerUDO()
        If Not Me.UDOExists("UDO_REG_SUB") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("UDO_REG_SUB", "10XRegulatorySubmissionTracker", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_REG_SUBH", "TNX_REG_SUBL", "TNX_REG_QRY", "TNX_REG_STAT1", "TNX_REG_APRV", "TNX_ATTACHMENT_C3")
            findaliasdescription = Nothing
        End If
    End Sub
    '=1
    Public Sub StabilityStudyUDO()
        If Not Me.UDOExists("TNX_PH_STAB") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_PH_STAB", "TNX_PH_STAB", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_PH_STAB", "TNX_PH_STAB_D1", "TNX_PH_STAB_D2")
            findaliasdescription = Nothing
        End If
    End Sub



    Public Sub ValidationyUDO()
        If Not Me.UDOExists("TNX_PH_VAL") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_PH_VAL", "TNX_PH_VAL", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_PH_VAL", "TNX_PH_VAL_D1", "TNX_PH_VAL_D2")
            findaliasdescription = Nothing
        End If
    End Sub

    Public Sub ABCCostingUDO()
        If Not Me.UDOExists("TNX_PH_ABC_COST") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_PH_ABC_COST", "TNX_PH_ABC_COST", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_PH_ABC_COST", "TNX_PH_ABC_D1", "TNX_PH_ABC_D2")
            findaliasdescription = Nothing
        End If
    End Sub

    Public Sub CAPAUDO()
        If Not Me.UDOExists("TNX_PH_CAPA") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_PH_CAPA", "TNX_PH_CAPA", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_PH_CAPA", "TNX_PH_CAPA_D1", "TNX_PH_CAPA_D2")
            findaliasdescription = Nothing
        End If
    End Sub



    Public Sub CorporateTaxCalculationUDO()
        If Not Me.UDOExists("TNX_CTCAUDO") Then
            Dim findaliasdescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_CTCAUDO", "TNX_CTCAUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasdescription, "TNX_CTAXCALCU", "TNX_CTAXCALCU_C2")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CTAXConifgUDO()
        If Not Me.UDOExists("TNX_CTXUDO") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDONoLog("TNX_CTXUDO", "TNX_CTXUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_CTAXCNF")
            findaliasdescription = Nothing
        End If
    End Sub


    Public Sub CreateDatMTUDO()
        If Not Me.UDOExists("TNX_DBM_UDO") Then
            Dim findaliasndescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNX_DBM_UDO", "TNX_DBM_UDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasndescription, "TNX_DM")
            findaliasndescription = Nothing
        End If
    End Sub

    'vamshi sai
    Public Sub PAYUDO()
        If Not Me.UDOExists("TNXPAYUDO") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNXPAYUDO", "TNXPAYUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_PAY")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub PAYLOADUDO()
        If Not Me.UDOExists("TNXPAYLDUDO") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNXPAYLDUDO", "TNXPAYLDUDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_PAYLD")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CreateINVFUDO()
        If Not Me.UDOExists("TNX_INVF_UDO") Then
            Dim findaliasdescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("TNX_INVF_UDO", "Invoicing Configuration UDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findaliasdescription, "TNX_INVF")
            findaliasdescription = Nothing
        End If
    End Sub
    Public Sub CreateLicenceNewUDO()
        If Not Me.UDOExists("TNX_LICENSE_UDO") Then
            Dim findaliasndescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("TNX_LICENSE_UDO", "Licence  Administration New1 UDO ", SAPbobsCOM.BoUDOObjType.boud_Document, findaliasndescription, "TNX_LICENCE", "TNX_LICENCE_C0")
            findaliasndescription = Nothing
        End If
    End Sub

#End Region

#Region "UDO Exists"
    Public Function UDOExists(ByVal code As String) As Boolean
        GC.Collect()
        Dim v_UDOMD As SAPbobsCOM.UserObjectsMD
        Dim v_ReturnCode As Boolean
        v_UDOMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        v_ReturnCode = v_UDOMD.GetByKey(code)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UDOMD)
        v_UDOMD = Nothing
        Return v_ReturnCode
    End Function
#End Region

#Region "Register UDO"

    Function registerUDO(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal findAliasNDescription As String(,), ByVal parentTableName As String, Optional ByVal childTable1 As String = "", Optional ByVal childTable2 As String = "", Optional ByVal childTable3 As String = "", Optional ByVal childTable4 As String = "", Optional ByVal childTable5 As String = "", Optional ByVal childTable6 As String = "", Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim actionSuccess As Boolean = False
        Try
            registerUDO = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = LogOption
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = parentTableName
            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.LogTableName = "L" & parentTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To findAliasNDescription.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = findAliasNDescription(i, 0)
                v_udoMD.FindColumns.ColumnDescription = findAliasNDescription(i, 1)
            Next
            If childTable1 <> "" Then
                v_udoMD.ChildTables.TableName = childTable1
                v_udoMD.ChildTables.Add()
            End If
            If childTable2 <> "" Then
                v_udoMD.ChildTables.TableName = childTable2
                v_udoMD.ChildTables.Add()
            End If
            If childTable3 <> "" Then
                v_udoMD.ChildTables.TableName = childTable3
                v_udoMD.ChildTables.Add()
            End If
            If childTable4 <> "" Then
                v_udoMD.ChildTables.TableName = childTable4
                v_udoMD.ChildTables.Add()
            End If
            If childTable5 <> "" Then
                v_udoMD.ChildTables.TableName = childTable5
                v_udoMD.ChildTables.Add()
            End If
            If childTable6 <> "" Then
                v_udoMD.ChildTables.TableName = childTable6
                v_udoMD.ChildTables.Add()
            End If

            If v_udoMD.Add() = 0 Then
                registerUDO = True
                objMain.objApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objMain.objApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                registerUDO = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Function

    Function registerUDONoLog(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal findAliasNDescription As String(,), ByVal parentTableName As String, Optional ByVal childTable1 As String = "", Optional ByVal childTable2 As String = "", Optional ByVal childTable3 As String = "", Optional ByVal childTable4 As String = "", Optional ByVal childTable5 As String = "", Optional ByVal childTable6 As String = "", Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim actionSuccess As Boolean = False
        Try
            registerUDONoLog = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = LogOption
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = parentTableName
            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.LogTableName = "A" & parentTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To findAliasNDescription.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = findAliasNDescription(i, 0)
                v_udoMD.FindColumns.ColumnDescription = findAliasNDescription(i, 1)
            Next
            If childTable1 <> "" Then
                v_udoMD.ChildTables.TableName = childTable1
                v_udoMD.ChildTables.Add()
            End If
            If childTable2 <> "" Then
                v_udoMD.ChildTables.TableName = childTable2
                v_udoMD.ChildTables.Add()
            End If
            If childTable3 <> "" Then
                v_udoMD.ChildTables.TableName = childTable3
                v_udoMD.ChildTables.Add()
            End If
            If childTable4 <> "" Then
                v_udoMD.ChildTables.TableName = childTable4
                v_udoMD.ChildTables.Add()
            End If
            If childTable5 <> "" Then
                v_udoMD.ChildTables.TableName = childTable5
                v_udoMD.ChildTables.Add()
            End If
            If childTable6 <> "" Then
                v_udoMD.ChildTables.TableName = childTable6
                v_udoMD.ChildTables.Add()
            End If

            If v_udoMD.Add() = 0 Then
                registerUDONoLog = True
                objMain.objApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objMain.objApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                registerUDONoLog = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Function

#End Region

#Region "Add Menu's With XML"

    Private Sub LoadFromXML(ByRef FileName As String)
        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        '// load the content of the XML File
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
        ExePath = sPath
        oXmlDoc.Load(sPath & "\" & FileName)
        '// load the form to the SBO application in one batch
        objApplication.LoadBatchActions(oXmlDoc.InnerXml)
        sPath = objApplication.GetLastBatchResults()
    End Sub
#End Region


#Region "Item Event"
    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try
            '------------------------------------------------------------------------
            Try
                If TenXFrameWork.MainCls.ohtLookUpForm.ContainsValue(FormUID) = True Then
                    Dim keys As ICollection = TenXFrameWork.MainCls.ohtLookUpForm.Keys
                    Dim keysArray(TenXFrameWork.MainCls.ohtLookUpForm.Count - 1) As String
                    keys.CopyTo(keysArray, 0)
                    For Each key As String In keysArray
                        If FormUID = TenXFrameWork.MainCls.ohtLookUpForm(key) Then
                            While TenXFrameWork.MainCls.ohtLookUpForm.ContainsValue(key) = True
                                For Each dKey As String In keysArray
                                    If key = TenXFrameWork.MainCls.ohtLookUpForm(dKey) Then
                                        key = dKey
                                        Exit For
                                    End If
                                Next
                            End While
                            objMain.objApplication.Forms.Item(key).Select()
                            BubbleEvent = False
                            Exit Sub
                        End If
                    Next
                End If
            Catch ex As Exception
            End Try
            Select Case pVal.FormTypeEx
                'Addon Files
                'Case "FORCAST"
                '    ObjFORCAST.ItemEvent(FormUID, pVal, BubbleEvent)
                'Case "TNX_FSLC"
                '    oClsParameterSelection.ItemEvent(FormUID, pVal, BubbleEvent)
                'Case "TNX_USR"
                '    oSubParameterSelection.ItemEvent(FormUID, pVal, BubbleEvent)
                'Case "TNX_OAFC"
                '    oApprovedForecating.ItemEvent(FormUID, pVal, BubbleEvent
                Case "133"
                    objARInvoice.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "179"
                    objARCreditMemo.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "65300"
                    objARDownPayment.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_Approve"
                    ObjVIEW.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "198"
                    If objSAPAlertWindow IsNot Nothing Then
                        objSAPAlertWindow.ItemEvent(FormUID, pVal, BubbleEvent)
                    End If

                Case "frm_CORACT"
                    ObjclsMstrCorrectiveAction.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_PREACT"
                    ObjclsMstrPreventiveAction.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_DOCNUM"
                    ObjclsMstrDocumentNumSetup.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_RETPOL"
                    ObjclsMstrRetentionPolicy.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_INCCAT"
                    ObjclsMstrIncidentCategory.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_COMPSET"
                    ObjclsMstrComplianceSett.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_CMS_ESIGN"
                    ObjclsElectronicSignaturePolicy.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_BPR"
                    objPhamaBpr.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_DISP"
                    objPharmaDispensing.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_AUDCHK"
                    ObjclsAudit.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_COAT"
                    ObjclsCOMTemplate.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_DEVCAT"
                    ObjclsDevidation.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_VAL"
                    Objclsvalidation.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_INC"
                    ObjclsIncident.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_PRDCAT"
                    ObjclsMstrProductCategory.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_DOSFORM"
                    ObjClsMstrDosageForm.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "TNX_PEQP"
                    objEquipmentMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "UDO_FT_TNX_PDTR"
                    objDowntimeReason.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "TNX_PSTG"
                    objproductionStage.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "TNX_PIQC"
                    objInprogresschecklist.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "TNX_PEQP"
                    objClerenceMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "REG_AUTH"
                    objRegulatoryAuthority.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "TNX_PLCC"
                    objLineClerencechecklist.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "TNX_PYTM"
                    objYeildtolorance.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "REG_CNFG"
                    objcountryregulatoryconfig.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_CAPACAT"
                    ObjCAPAMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_SOP"
                    ObjSOP.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_RISK"
                    ObjRisk.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_STYPE"
                    ObjSample.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "UDO_F_UDO_REG_PRDREG"
                    objProductRegistration.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "FORMULACOSTING"
                    objformulacosting.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "BMRR"
                    ObjBmr.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_ASR"
                    ObjApproval.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GRID"
                    ObjGRID.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ME_ADR"
                    ObjADR.ItemEvent(FormUID, pVal, BubbleEvent)
                'Case "SBO_AST"
                '    ObjAStg.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "GRIDES"
                    ObjGRIDES.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SBO_Draft"
                    ObjDraftProcedure.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "Temp"
                    ObjAPTEMP.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "IK_ESTMT"
                    objEstimation.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "INV_P"
                    objInvPost.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "PROTOCOL"
                    Stabilityprotocal.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "SSTUD"
                    Stabilitystudy.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "STAB_SHELF"
                    Shelflife.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "ONBP"
                    objOnboarding.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "PAYR"
                    objPayLoad.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "133"
                    objAREinvoice.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "CTAXC"
                    ObjCorporateTaxConfiguration.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "CTAXCAL"
                    ObjCorporateTaxCalculation.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VATR"
                    ObjVatReport.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "PAYLD"
                    ObjPayloadD.ItemEvent(FormUID, pVal, BubbleEvent)
                    'vsm
                Case "COTAX"
                    ObjClsCorpTax.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_FTAVM"
                    ObjclsFtaVat.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_LKMTR"
                    ObjClsLkMstr.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_QAB"
                    ObjclsBatch.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_CAPA"
                    ObjclsCAPAManage.ItemEvent(FormUID, pVal, BubbleEvent)

                'Case "License"
                '    objLicenceNew.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "DEVICE"
                    objDevice.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case "License"
                    '    objLicenceAdministration.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_EXP_MGT"
                    objExperimentManagement.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "PHFARM"
                    objFormulaMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_COA"
                    Objclsmanage.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "10X_CHG"
                    Objclscontrol.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "PHPILOT"
                    ObjclsPilotBatch.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_SOPMGT"
                    ObjclsSOPManagement.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_TRANIPLN"
                    ObjclsTrainingPlan.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_TRNEXE"
                    ObjclsTrainingExecution.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_TRNCERT"
                    ObjclsTrainingCertificate.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "frm_TRNMAT"
                    ObjclsEmployeeTraining.ItemEvent(FormUID, pVal, BubbleEvent)
            End Select
        Catch ex As Exception
            objApplication.MessageBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Menu Events"
    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        'Try
        Dim objform As SAPbouiCOM.Form
        ' Catch ex As Exception

        ' End Try

        Try
            '  objform = objMain.objApplication.Forms.ActiveForm
            Select Case pVal.MenuUID

                'Case "FORCAST"
                '    ObjFORCAST.MenuEvent(pVal, BubbleEvent)
                'Case "TNX_FSLC"
                '    oClsParameterSelection.MenuEvent(pVal, BubbleEvent)
                'Case "TNX_USR"
                '    oSubParameterSelection.MenuEvent(pVal, BubbleEvent)
                'Case "TNX_OAFC"
                '    oApprovedForecating.MenuEvent(pVal, BubbleEvent)
                'Case "SUBTYPE"
                '    ObjSUBTYPE.MenuEvent(pVal, BubbleEvent)
                '    'Find
                Case "10X_FRM_MST"
                    objFormulaMaster.MenuEvent(pVal, BubbleEvent)
                Case "10X_EXP_MGT"
                    objExperimentManagement.MenuEvent(pVal, BubbleEvent)
                Case "10X_SPEC"
                    objSpecificationMaster.MenuEvent(pVal, BubbleEvent)
                Case "10X_LABTEST"
                    ObjLabTesting.MenuEvent(pVal, BubbleEvent)
                Case "10X_COST"
                    objformulacosting.MenuEvent(pVal, BubbleEvent)
                Case "10X_BATCHREL"
                    ObjclsBatch.MenuEvent(pVal, BubbleEvent)
                Case "10X_CAPA"
                    ObjclsCAPAManage.MenuEvent(pVal, BubbleEvent)
                Case "IK_ESTMT"
                    objEstimation.MenuEvent(pVal, BubbleEvent)
                Case "10X_RD_PCAT"
                    ObjclsMstrProductCategory.MenuEvent(pVal, BubbleEvent)
                Case "10X_RD_DFORM"
                    ObjClsMstrDosageForm.MenuEvent(pVal, BubbleEvent)
                Case "Inv_Posting"
                    objInvPost.MenuEvent(pVal, BubbleEvent)
                Case "10X_QC_CAPACAT"
                    ObjCAPAMaster.MenuEvent(pVal, BubbleEvent)
                Case "10X_QC_SOP"
                    ObjSOP.MenuEvent(pVal, BubbleEvent)
                Case "10X_REG_PROD"
                    objProductRegistration.MenuEvent(pVal, BubbleEvent)
                Case "10X_QC_RISK"
                    ObjRisk.MenuEvent(pVal, BubbleEvent)
                Case "10X_QC_STYPE"
                    ObjSample.MenuEvent(pVal, BubbleEvent)
                Case "On_Process"
                    objOnboarding.MenuEvent(pVal, BubbleEvent)
                Case "PAYR"
                    objPayLoad.MenuEvent(pVal, BubbleEvent)
                Case "CTAXC"
                    ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)
                Case "CTAXCAL"
                    ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                Case "VATR"
                    ObjVatReport.MenuEvent(pVal, BubbleEvent)
                Case "PAYLD"
                    ObjPayloadD.MenuEvent(pVal, BubbleEvent)
                Case "COTX"
                    ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                Case "FTAV"
                    ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                       'Approval
                Case "ME_ASR"
                    ObjApproval.MenuEvent(pVal, BubbleEvent)
                Case "GRID"
                    'ObjGRID.MenuEvent(pVal, BubbleEvent)
                Case "ME_ADR"
                    ObjADR.MenuEvent(pVal, BubbleEvent)
                'Case "SBO_AST"
                '    ObjAStg.MenuEvent(pVal, BubbleEvent)
                Case "Temp"
                    ObjAPTEMP.MenuEvent(pVal, BubbleEvent)
                Case "GRIDES"
                    ObjGRIDES.MenuEvent(pVal, BubbleEvent)
                    'vsm
                Case "COTX"
                    ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                Case "FTAV"
                    ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                Case "LKMT"
                    ObjClsLkMstr.MenuEvent(pVal, BubbleEvent)
                'Case "License"
                '    objLicenceNew.MenuEvent(pVal, BubbleEvent)
                     'Case "License"
                '    objLicenceAdministration.MenuEvent(pVal, BubbleEvent)
                    'Vamshi Sai
                Case "Invoice_Posting"
                    objInvoicePosting.MenuEvent(pVal, BubbleEvent)
                Case "DEVICE"
                    objDevice.MenuEvent(pVal, BubbleEvent)
                Case "10X_COA"
                    Objclsmanage.MenuEvent(pVal, BubbleEvent)
                Case "10X_COMP_CC"
                    Objclscontrol.MenuEvent(pVal, BubbleEvent)
                Case "10X_YIELD"
                    YieldAnalysis.MenuEvent(pVal, BubbleEvent)
                Case "10X_PILOT"
                    ObjclsPilotBatch.MenuEvent(pVal, BubbleEvent)
                Case "10X_COMP_SOP"
                    ObjclsSOPManagement.MenuEvent(pVal, BubbleEvent)
                Case "TRAINP"
                    ObjclsTrainingPlan.MenuEvent(pVal, BubbleEvent)
                Case "TRAINE"
                    ObjclsTrainingExecution.MenuEvent(pVal, BubbleEvent)
                Case "TRAINC"
                    ObjclsTrainingCertificate.MenuEvent(pVal, BubbleEvent)
                Case "TRAINET"
                    ObjclsEmployeeTraining.MenuEvent(pVal, BubbleEvent)
                Case "10X_BMR"
                    ObjBmr.MenuEvent(pVal, BubbleEvent)
                Case "10X_SAMPLE"
                    SampleRegistration.MenuEvent(pVal, BubbleEvent)
                Case "10X_SAMPCOL"
                    SampleCollection.MenuEvent(pVal, BubbleEvent)

                Case "10X_BPR"
                    objPhamaBpr.MenuEvent(pVal, BubbleEvent)
                Case "10X_DISP"
                    objPharmaDispensing.MenuEvent(pVal, BubbleEvent)
                Case "10X_PMS_CLEAN"
                    objClerenceMaster.MenuEvent(pVal, BubbleEvent)

                Case "10X_RMS_CNFG"
                    objcountryregulatoryconfig.MenuEvent(pVal, BubbleEvent)
                Case "10X_PMS_LCC"
                    objLineClerencechecklist.MenuEvent(pVal, BubbleEvent)
                Case "10X_PMS_STAGE"
                    objproductionStage.MenuEvent(pVal, BubbleEvent)

                Case "10X_RMS_AUTH"
                    objRegulatoryAuthority.MenuEvent(pVal, BubbleEvent)
                Case "10X_PMS_EQP"
                    objEquipmentMaster.MenuEvent(pVal, BubbleEvent)
                Case "10X_PMS_DTR"
                    objDowntimeReason.MenuEvent(pVal, BubbleEvent)
                Case "10X_PMS_IPQC"
                    objInprogresschecklist.MenuEvent(pVal, BubbleEvent)
                Case "10X_PMS_YIELD"
                    objYeildtolorance.MenuEvent(pVal, BubbleEvent)
                Case "10X_CMS_ESIGN"
                    ObjclsElectronicSignaturePolicy.MenuEvent(pVal, BubbleEvent)
                Case "10X_CMS_CORACT"
                    ObjclsMstrCorrectiveAction.MenuEvent(pVal, BubbleEvent)
                Case "10X_CMS_PREACT"
                    ObjclsMstrPreventiveAction.MenuEvent(pVal, BubbleEvent)
                Case "10X_CMS_DOCNUM"
                    ObjclsMstrDocumentNumSetup.MenuEvent(pVal, BubbleEvent)
                Case "10X_CMS_RETPOL"
                    ObjclsMstrRetentionPolicy.MenuEvent(pVal, BubbleEvent)
                Case "10X_CMS_INCCAT"
                    ObjclsMstrIncidentCategory.MenuEvent(pVal, BubbleEvent)
                Case "10X_CMS_COMPSET"
                    ObjclsMstrComplianceSett.MenuEvent(pVal, BubbleEvent)
                Case "10X_COMP_VAL"
                    Objclsvalidation.MenuEvent(pVal, BubbleEvent)
                Case "10X_COMP_INC"
                    ObjclsIncident.MenuEvent(pVal, BubbleEvent)
                Case "10X_QC_AUDCHK"
                    ObjclsAudit.MenuEvent(pVal, BubbleEvent)
                Case "10X_QC_COAT"
                    ObjclsCOMTemplate.MenuEvent(pVal, BubbleEvent)
                Case "10X_QC_DEVCAT"
                    ObjclsDevidation.MenuEvent(pVal, BubbleEvent)

                Case "10X_STB_PROTO"
                    Stabilityprotocal.MenuEvent(pVal, BubbleEvent)
                Case "10X_STB_STUDY"
                    Stabilitystudy.MenuEvent(pVal, BubbleEvent)

                Case "10X_STB_SHELFL"
                    Shelflife.MenuEvent(pVal, BubbleEvent)

                Case "1282"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "133" Then
                        objARInvoice.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "COTX" Then
                        ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "FTAV" Then
                        ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "LKMT" Then
                        ObjClsLkMstr.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_COAT" Then
                        ObjclsCOMTemplate.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_DEVCAT" Then
                        ObjclsDevidation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_VAL" Then
                        Objclsvalidation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_INC" Then
                        ObjclsIncident.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_AUDCHK" Then
                        ObjclsAudit.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PEQP" Then
                        objEquipmentMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PSTG" Then
                        objproductionStage.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PLCC" Then
                        objLineClerencechecklist.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PCLM" Then
                        objClerenceMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PYTM" Then
                        objYeildtolorance.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "REG_CNFG" Then
                        objcountryregulatoryconfig.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PIQC" Then
                        objInprogresschecklist.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "UDO_FT_TNX_PDTR" Then
                        objDowntimeReason.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "REG_AUTH" Then
                        objRegulatoryAuthority.MenuEvent(pVal, BubbleEvent)
                        'ElseIf objform.TypeEx = "License" Then
                        '    objLicenceNew.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "FORMULACOSTING" Then
                        objformulacosting.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PHFARM" Then
                        objFormulaMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_PRDCAT" Then
                        ObjclsMstrProductCategory.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_DOSFORM" Then
                        ObjClsMstrDosageForm.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PROTOCOL" Then
                        Stabilityprotocal.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "SSTUD" Then
                        Stabilitystudy.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "STAB_SHELF" Then
                        Shelflife.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "EXPMNG" Then
                        objExperimentManagement.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_COMP_CC" Then
                        Objclscontrol.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_COA" Then
                        Objclsmanage.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PHPILOT" Then
                        ObjclsPilotBatch.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_SOPMGT" Then
                        ObjclsSOPManagement.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_TRANIPLN" Then
                        ObjclsTrainingPlan.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_TRNEXE" Then
                        ObjclsTrainingExecution.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_TRNCERT" Then
                        ObjclsTrainingCertificate.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_EXP_MGT" Then
                        objExperimentManagement.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "XPH_QSPEC" Then
                        objSpecificationMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNXPH_QCLAB" Then
                        ObjLabTesting.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_QAB" Then
                        ObjclsBatch.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_CAPA" Then
                        ObjclsCAPAManage.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "TNXPYLD" Then
                        YieldAnalysis.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "SCOLLN" Then
                        SampleCollection.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_CMS_ESIGN" Then
                        ObjclsElectronicSignaturePolicy.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "UDO_F_UDO_REG_PRDREG" Then
                        objProductRegistration.MenuEvent(pVal, BubbleEvent)
                    End If



                Case "1281"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "EXPMNG" Then
                        objExperimentManagement.MenuEvent(pVal, BubbleEvent)

                    End If
                    'Navigations
                    'Case "1288"
                    '    objform = objMain.objApplication.Forms.ActiveForm
                    '    If objform.TypeEx = "TNX_USR" Then
                    '        oSubParameterSelection.MenuEvent(pVal, BubbleEvent)
                    '    End If
                    '    If objform.TypeEx = "TNX_OAFC" Then
                    '        oApprovedForecating.MenuEvent(pVal, BubbleEvent)
                    '    End If
                    'Case "1289"
                    '    objform = objMain.objApplication.Forms.ActiveForm
                    '    If objform.TypeEx = "TNX_USR" Then
                    '        oSubParameterSelection.MenuEvent(pVal, BubbleEvent)
                    '    End If
                    '    If objform.TypeEx = "TNX_OAFC" Then
                    '        oApprovedForecating.MenuEvent(pVal, BubbleEvent)
                    '    End If
                    'Case "1290"
                    '    objform = objMain.objApplication.Forms.ActiveForm
                    '    If objform.TypeEx = "TNX_USR" Then
                    '        oSubParameterSelection.MenuEvent(pVal, BubbleEvent)
                    '    End If
                    '    If objform.TypeEx = "TNX_OAFC" Then
                    '        oApprovedForecating.MenuEvent(pVal, BubbleEvent)
                    '    End If
                Case "1291"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                        'ElseIf objform.TypeEx = "CTAXCAL" Then
                        '    ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_CMS_ESIGN" Then
                        ObjclsElectronicSignaturePolicy.MenuEvent(pVal, BubbleEvent)

                    End If
                    'If objform.TypeEx = "TNX_OAFC" Then
                    '    oApprovedForecating.MenuEvent(pVal, BubbleEvent)
                    'End If
                    '    'ADD ROW
                Case "1293"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_PRDCAT" Then
                        ObjclsMstrProductCategory.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_CAPACAT" Then
                        ObjCAPAMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_SOP" Then
                        ObjSOP.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_RISK" Then
                        ObjRisk.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_STYPE" Then
                        ObjSample.MenuEvent(pVal, BubbleEvent)
                        'ElseIf objform.TypeEx = "UDO_F_UDO_REG_PRDREG" Then
                        '    objProductRegistration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_DOSFORM" Then
                        ObjClsMstrDosageForm.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "COTX" Then
                        ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "FTAV" Then
                        ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PHPILOT" Then
                        ObjclsPilotBatch.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_SOPMGT" Then
                        ObjclsSOPManagement.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_TRANIPLN" Then
                        ObjclsTrainingPlan.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_TRNEXE" Then
                        ObjclsTrainingExecution.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_TRNCERT" Then
                        ObjclsTrainingCertificate.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_TRNMAT" Then
                        ObjclsEmployeeTraining.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_CMS_ESIGN" Then
                        ObjclsElectronicSignaturePolicy.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_COAT" Then
                        ObjclsCOMTemplate.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_DEVCAT" Then
                        ObjclsDevidation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_VAL" Then
                        Objclsvalidation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_INC" Then
                        ObjclsIncident.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_AUDCHK" Then
                        ObjclsAudit.MenuEvent(pVal, BubbleEvent)

                    ElseIf objform.TypeEx = "frm_CORACT" Then
                        ObjclsMstrCorrectiveAction.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_PREACT" Then
                        ObjclsMstrPreventiveAction.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_DOCNUM" Then
                        ObjclsMstrDocumentNumSetup.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_RETPOL" Then
                        ObjclsMstrRetentionPolicy.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_INCCAT" Then
                        ObjclsMstrIncidentCategory.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_COMPSET" Then
                        ObjclsMstrComplianceSett.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PEQP" Then
                        objEquipmentMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PSTG" Then
                        objproductionStage.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PLCC" Then
                        objLineClerencechecklist.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PCLM" Then
                        objClerenceMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PYTM" Then
                        objYeildtolorance.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "REG_CNFG" Then
                        objcountryregulatoryconfig.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PIQC" Then
                        objInprogresschecklist.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "UDO_FT_TNX_PDTR" Then
                        objDowntimeReason.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "REG_AUTH" Then
                        objRegulatoryAuthority.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "774"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "FORMULACOSTING" Then
                        objformulacosting.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "TNXPYLD" Then
                        YieldAnalysis.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "SAMPLE" Then
                        SampleRegistration.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "SCOLLN" Then
                        SampleCollection.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "PROTOCOL" Then
                        Stabilityprotocal.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "SSTUD" Then
                        Stabilitystudy.MenuEvent(pVal, BubbleEvent)
                    End If
                    If objform.TypeEx = "STAB_SHELF" Then
                        Shelflife.MenuEvent(pVal, BubbleEvent)
                    End If

                    '    'ADD ROW
                Case "1282"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_BPR" Then
                        objPhamaBpr.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_DISP" Then
                        objPharmaDispensing.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                        'ElseIf objform.TypeEx = "CTAXCAL" Then
                        '    ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_COMP_CC" Then
                        Objclscontrol.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_COA" Then
                        Objclsmanage.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "SAMPLE" Then
                        SampleRegistration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_CAPACAT" Then
                        ObjCAPAMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_SOP" Then
                        ObjSOP.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "UDO_F_UDO_REG_PRDREG" Then
                        objProductRegistration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_RISK" Then
                        ObjRisk.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_STYPE" Then
                        ObjSample.MenuEvent(pVal, BubbleEvent)

                    End If

                    'Case "Add Row"
                    '    objform = objMain.objApplication.Forms.ActiveForm
                    '    If objform.TypeEx = "FORCAST" Then
                    '        ObjFORCAST.MenuEvent(pVal, BubbleEvent)
                    '    End If
                Case "519"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "COTX" Then
                        ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "FTAV" Then
                        ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "520"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "COTX" Then
                        ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "FTAV" Then
                        ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "UDO_F_UDO_REG_PRDREG" Then
                        objProductRegistration.MenuEvent(pVal, BubbleEvent)
                    End If
                Case "1284"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "1292"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "CTAXC" Then
                        ObjCorporateTaxConfiguration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_PRDCAT" Then
                        ObjclsMstrProductCategory.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_DOSFORM" Then
                        ObjClsMstrDosageForm.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VATR" Then
                        ObjVatReport.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "CTAXCAL" Then
                        ObjCorporateTaxCalculation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "COTAX" Then
                        ObjClsCorpTax.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_FTAVM" Then
                        ObjclsFtaVat.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "Temp" Then
                        ObjAPTEMP.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "DEVICE" Then
                        objDevice.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "FORMULACOSTING" Then
                        objformulacosting.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_BPR" Then
                        objPhamaBpr.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_DISP" Then
                        objPharmaDispensing.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_CAPACAT" Then
                        ObjCAPAMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_SOP" Then
                        ObjSOP.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_RISK" Then
                        ObjRisk.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "UDO_F_UDO_REG_PRDREG" Then
                        objProductRegistration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_STYPE" Then
                        ObjSample.MenuEvent(pVal, BubbleEvent)

                        'ElseIf objform.TypeEx = "TNX_USR" Then
                        '    oSubParameterSelection.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PEQP" Then
                        objEquipmentMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PSTG" Then
                        objproductionStage.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PLCC" Then
                        objLineClerencechecklist.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PCLM" Then
                        objClerenceMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PYTM" Then
                        objYeildtolorance.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "REG_CNFG" Then
                        objcountryregulatoryconfig.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNX_PIQC" Then
                        objInprogresschecklist.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "UDO_FT_TNX_PDTR" Then
                        objDowntimeReason.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "REG_AUTH" Then
                        objRegulatoryAuthority.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_COA" Then
                        Objclsmanage.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_CHG" Then
                        Objclscontrol.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_QAB" Then
                        ObjclsBatch.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_CAPA" Then
                        ObjclsCAPAManage.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "SCOLLN" Then
                        SampleCollection.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNXPYLD" Then
                        YieldAnalysis.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "SAMPLE" Then
                        SampleRegistration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "EXPMNG" Then
                        objExperimentManagement.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PHPILOT" Then
                        ObjclsPilotBatch.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_SOPMGT" Then
                        ObjclsSOPManagement.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_TRANIPLN" Then
                        ObjclsTrainingPlan.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_TRNEXE" Then
                        ObjclsTrainingExecution.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_TRNCERT" Then
                        ObjclsTrainingCertificate.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_TRNMAT" Then
                        ObjclsEmployeeTraining.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "XPH_QSPEC" Then
                        objSpecificationMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "TNXPH_QCLAB" Then
                        ObjLabTesting.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_CMS_ESIGN" Then
                        ObjclsElectronicSignaturePolicy.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_COAT" Then
                        ObjclsCOMTemplate.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_DEVCAT" Then
                        ObjclsDevidation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_VAL" Then
                        Objclsvalidation.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_INC" Then
                        ObjclsIncident.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_AUDCHK" Then
                        ObjclsAudit.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "PROTOCOL" Then
                        Stabilityprotocal.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "SSTUD" Then
                        Stabilitystudy.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "STAB_SHELF" Then
                        Shelflife.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_CORACT" Then
                        ObjclsMstrCorrectiveAction.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_PREACT" Then
                        ObjclsMstrPreventiveAction.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_DOCNUM" Then
                        ObjclsMstrDocumentNumSetup.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_RETPOL" Then
                        ObjclsMstrRetentionPolicy.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_INCCAT" Then
                        ObjclsMstrIncidentCategory.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "frm_COMPSET" Then
                        ObjclsMstrComplianceSett.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_BPR" Then
                        objPhamaBpr.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "10X_DISP" Then
                        objPharmaDispensing.MenuEvent(pVal, BubbleEvent)

                    End If
                Case "Delete Row"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "UDO_F_UDO_REG_PRDREG" Then
                        objProductRegistration.MenuEvent(pVal, BubbleEvent)
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region "Form Data Event"
    Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objApplication.FormDataEvent
        If BusinessObjectInfo.BeforeAction = False Then
            Select Case BusinessObjectInfo.FormTypeEx
                Case "TNX_OAFC"
                    'oApprovedForecating.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "133"
                    objARInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                    objARInvoice.FormDataEvent1(BusinessObjectInfo, BubbleEvent)
                Case "179"
                    objARCreditMemo.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "CTAXCAL"
                    ObjCorporateTaxCalculation.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "65300"
                    objARDownPayment.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                Case "VATR"
                    ObjVatReport.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            End Select
        End If
    End Sub
#End Region

#Region "Application Event"
    Private Sub oApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                objCompany.Disconnect()
                End
        End Select
    End Sub
#End Region

#Region "Right Click Event"
    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        Dim objForm As SAPbouiCOM.Form
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        'If objForm.TypeEx = "FORCAST" Then
        '    'ObjFORCAST.RightClickEvent(eventInfo, BubbleEvent)
        'End If
        If objForm.TypeEx = "UDO_F_UDO_REG_PRDREG" Then
            objProductRegistration.RightClickEvent(eventInfo, BubbleEvent)
        End If
    End Sub
#End Region

End Class

Imports SAPbobsCOM

Public Class DatabaseCreation

#Region "Declaration"
    Private objUtilities As Utilities
    Dim DBCode As String = "v0.400"
    Dim DBName As String = "v0.400"
    Dim Version As String = "v0.706"
#End Region

#Region "DB Creation Main"
    Public Sub New()
        objUtilities = New Utilities
    End Sub
    Public Function CreateTables() As Boolean
        Try
            objUtilities.CreateTable("TNX_DB", "DBCONFIG(EInvoice)TABLE", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            objUtilities.AddAlphaField("@TNX_DB", "VERSION", "VERSION", 100)
            Dim oRs As SAPbobsCOM.Recordset
            oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("SELECT * FROM ""@TNX_DB"" where ""U_VERSION"" = '" & Version & "'")
            Dim iDBConfigRecordCount As Integer = oRs.RecordCount
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
            If iDBConfigRecordCount = 0 Then

                'Me.LicenceNew()
                'objMain.CreateLicenceNewUDO()
                'Me.CorpTax()
                'objMain.CorpTaxUDO()
                'Me.FtaVat()
                'objMain.FtaVatUDO()

                'Me.OnboardingProcessTable()
                'objMain.CreateONBPUDO()
                'Me.InvoicePosting()
                'objMain.CreateIPUDO()
                'Me.InvoicePostingTable()
                'objMain.CreateINVFUDO()
                'PayPosting()
                'objMain.PAYUDO()
                ''VatReport()
                ''objMain.VATREPORTUDO()
                Me.CreateAPPROVALSTAGES()
                objMain.CreateAPPROVALSTAGESUDO()
                objMain.CreateApprovalTemplatesUDO()

                Me.CreateFormulaVersionControl()
                objMain.FormulaCostingUDO()

                'Me.CreateSubmissionTracker()
                'objMain.SubmissionTrackerUDO()
                '=1
                Me.CreateStabilityStudy()
                objMain.StabilityStudyUDO()

                '=2
                CreateValidationFramework()
                objMain.ValidationyUDO()

                '=3
                Me.CreateABCCosting()
                objMain.ABCCostingUDO()

                '=4
                Me.CreateCAPA()
                objMain.CAPAUDO()

                ' 1. SOP Category Master
                CreateSOPCategoryMaster()

                ' 2. Department Master
                CreateDepartmentMaster()

                ' 3. Training Type Master
                CreateTrainingTypeMaster()

                ' 4. Validation Type Master
                CreateValidationTypeMaster()

                ' 5. Equipment Master
                CreateEquipmentMaster()

                ' 6. Risk Classification Master
                CreateRiskClassificationMaster()

                ' 7. CAPA Category Master
                CreateCAPACategoryMaster()

                ' 8. Audit Type Master
                CreateAuditTypeMaster()

                ' 9. Root Cause Master
                CreateRootCauseMaster()
                Me.SpecificationMaster1()
                objMain.SpecificMasterUDO()

                objMain.SOPCategoryMasterUDO()
                objMain.DepartmentMasterUDO()
                objMain.TrainingTypeMasterUDO()
                objMain.ValidationTypeMasterUDO()
                objMain.EquipmentMasterUDO()
                objMain.RiskClassificationMasterUDO()
                objMain.CAPACategoryMasterUDO()
                objMain.AuditTypeMasterUDO()
                objMain.RootCauseMasterUDO()
                CAPAManagementTables()
                objMain.CAPAManageUDO()
                QABatchReleaseTables()
                objMain.QABatchUDO()

                IncidentManagementTables()
                objMain.IncidentUDO()
                ValidationManagementTables()
                objMain.ValidationUDO()
                COAManagementTables()
                objMain.ManagementUDO()
                ChangeControlTables()
                objMain.ControlUDO()
                CreateExperimentManagement()
                objMain.ExperimentManagementUDO()
                CreateFormulaMaster()
                Me.CreatePharmaBMRExecution()
                objMain.CreatePharmaBMRExecutionUDO()
                objMain.CreateFormulaMasterUDO()
                Me.CreateAPPROVALTemplates()

                Me.CreateShelfLife()
                objMain.CreateShelfLifeUDO()

                Me.CreateStabilityStudy11()
                objMain.CreateStabilityStudyUDO()

                Me.CreateStabilityProtocol()
                objMain.CreateStabilityProtocolUDO()

                Me.YeildAnalysis()
                objMain.TNXPharmaYieldAnalysis()
                'objMain.DDUDO()
                'objMain.CreateApprovalTemplatesUDO()
                'Me.CreateAPPROVALSTAGES()
                'Me.CreateDDTable()
                '' Me.CreateDraftTable()
                ''objMain.DraftUDO()
                CreateTrainingTables()
                objMain.CreateTrainingExecutionUDO()
                objMain.CreateTrainingCertificateUDO()
                objMain.CreateTrainingMatrUDO()

                TNXTrainingPlan()
                objMain.TNXTrainingPlanUDO()

                SampleCollection()
                objMain.TNXQCSampleCollection()

                SampleRegistration()
                objMain.TNXQASampleRegistration()

                RemincgPilotBatchfields()
                objMain.CreatePilotBatcUDO()

                CreateSOPManagement()
                objMain.CreateSOPUDO()

                Me.QCLabTestingMaster()
                objMain.QCLabTestingUDO()

                Me.LineClearanceChecklistMaster()
                objMain.LineClearanceChecklistUDO()
                Me.DowntimeReasonMaster()
                objMain.DowntimeReasonMasterUDO()
                Me.CleaningMethodMaster()
                objMain.CleaningMethodMasterUDO()
                Me.EquipmentMaster()
                objMain.EquipmentMasterUDO1()
                Me.ProductionStageMaster()
                objMain.ProductionStageMasterUDO()
                Me.InProcessQCChecklistMaster()
                Me.YieldToleranceMaster()
                Me.TNXRegulatoryAuthority()
                Me.TNXCountryRegConfig()
                Me.TNXApprovalMatrix()
                Me.CreateMissingUserFieldsWithUtilities()
                Me.TNXRegistrationStatus()
                Me.TNXSubmissionType()
                Me.TNXDossierSection()
                Me.TNXCTDTemplate()
                Me.TNXRegulatoryDocType()

                Me.TNXArtworkType()
                Me.YieldToleranceMaster()
                objMain.InProcessQCChecklistUDO()
                objMain.YieldToleranceMasterUDO()
                objMain.TNXCountryRegulatoryUDO()
                objMain.TNXRegulatoryUDO()

                CreateMaterialQualificationMaster()
                MaterialTechnicalEvaluation()
                MaterialSpecifications()
                TrialSampleManagement()
                MaterialApproval()
                MaterialRequalification()
                MaterialRiskAssessment()
                CreateVendorQualificationMaster()
                VendorAudit()
                ApprovedVendorList()
                VendorRequalification()
                VendorRiskAssessment()
                VendorPerformance()

                ' Create Material Qualification UDOs
                objMain.CreateMaterialNewRequestUDO()
                objMain.CreateMaterialTechnicalEvaluationUDO()
                objMain.CreateMaterialSpecUDO()
                objMain.CreateTrialSampleUDO()
                objMain.CreateMaterialApprovalUDO()
                objMain.CreateMaterialRequalificationUDO()
                objMain.CreateMaterialRiskAssessmentUDO()
                objMain.CreateVendorQualificationUDO()
                objMain.CreateVendorAuditUDO()
                objMain.CreateApprovedVendorListUDO()
                objMain.CreateVendorRequalificationUDO()
                objMain.CreateVendorRiskAssessmentUDO()
                objMain.CreateVendorPerformanceReviewUDO()

                objMain.objUtilities.AddAlphaField("OITM", "TNX_MatQual", "Material Qualification Status", 30)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_MatType", "Material Type", 30)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_SpecNo", "Linked Spec No", 30)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_SpecVer", "Spec Version", 20)
                objMain.objUtilities.AddDateField("OITM", "TNX_QualDate", "Qualification Date", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddDateField("OITM", "TNX_ValidUpto", "Qualification Expiry", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_RiskCls", "Risk Class", 20)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_POBlock", "PO Block", 1)

                objMain.objUtilities.AddDateField("@TNX_QC_COAT_H", "DTN", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddDateField("@TNX_QC_AUDCHK_H", "DNL", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddDateField("@TNX_ROUTE", "DCM", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddDateField("@TNX_ARTWRK", "DER", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddDateField("@TNX_EMON", "DRT", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddDateField("@TNX_BNUM", "DUE", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)


                ' Business Partner (OCRD)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_VendQual", "Vendor Qualification Status", 30)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_GMPCert", "GMP Certified", 1)
                objMain.objUtilities.AddDateField("OCRD", "TNX_ValidUpto", "Vendor Approval Validity", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_RiskCls", "Vendor Risk Class", 20)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_AuditReq", "Audit Required", 1)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_VQRNo", "VQR No", 30)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_VendStatus", "Vendor Status", 30)

                ' Purchase Order Header (OPOR)
                objMain.objUtilities.AddAlphaField("OPOR", "TNX_PharmaChk", "Pharma Compliance Checked", 1)
                objMain.objUtilities.AddAlphaField("OPOR", "TNX_POStatus", "PO Status", 30)
                objMain.objUtilities.AddAlphaField("OPOR", "TNX_BlockReason", "Block Reason", 254)

                ' Purchase Order Lines (POR1)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_MatQual", "Material Qualification Status", 30)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_VendMap", "Vendor-Material Mapping", 30)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_SpecNo", "Specification No", 30)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_RiskCls", "Risk Class", 20)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_AVLCode", "AVL Reference", 30)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_AVLStatus", "AVL Status", 30)

                ' GRPO Header (OPDN) and Lines (PDN1)
                objMain.objUtilities.AddAlphaField("OPDN", "TNX_QCReq", "QC Required", 1)
                objMain.objUtilities.AddAlphaField("OPDN", "TNX_QCStatus", "QC Status", 30)
                objMain.objUtilities.AddAlphaField("PDN1", "TNX_SpecNo", "Specification No", 30)
                objMain.objUtilities.AddAlphaField("PDN1", "TNX_SampleID", "QC Sample ID", 30)
                objMain.objUtilities.AddAlphaField("PDN1", "TNX_Release", "Batch Release Status", 30)
                objMain.objUtilities.AddAlphaField("PDN1", "TNX_AVLCode", "AVL Code", 30)

                ' Batch Master (OBTN)
                objMain.objUtilities.AddAlphaField("OBTN", "TNX_QCStatus", "QC Status", 30)
                objMain.objUtilities.AddAlphaField("OBTN", "TNX_COARef", "COA Reference", 50)
                objMain.objUtilities.AddAlphaField("OBTN", "TNX_SpecNo", "Specification No", 30)
                objMain.objUtilities.AddDateField("OBTN", "TNX_RelDate", "Batch Release Date", SAPbobsCOM.BoFldSubTypes.st_None)


                CreateMaterialQualificationMaster()
                MaterialTechnicalEvaluation()
                MaterialSpecifications()
                TrialSampleManagement()
                MaterialApproval()
                MaterialRequalification()
                MaterialRiskAssessment()
                CreateVendorQualificationMaster()
                VendorAudit()
                ApprovedVendorList()
                VendorRequalification()
                VendorRiskAssessment()
                VendorPerformance()

                ' Create Material Qualification UDOs
                objMain.CreateMaterialNewRequestUDO()
                objMain.CreateMaterialTechnicalEvaluationUDO()
                objMain.CreateMaterialSpecUDO()
                objMain.CreateTrialSampleUDO()
                objMain.CreateMaterialApprovalUDO()
                objMain.CreateMaterialRequalificationUDO()
                objMain.CreateMaterialRiskAssessmentUDO()
                objMain.CreateVendorQualificationUDO()
                objMain.CreateVendorAuditUDO()
                objMain.CreateApprovedVendorListUDO()
                objMain.CreateVendorRequalificationUDO()
                objMain.CreateVendorRiskAssessmentUDO()
                objMain.CreateVendorPerformanceReviewUDO()

                objMain.objUtilities.AddAlphaField("OITM", "TNX_MatQual", "Material Qualification Status", 30)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_MatType", "Material Type", 30)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_SpecNo", "Linked Spec No", 30)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_SpecVer", "Spec Version", 20)
                objMain.objUtilities.AddDateField("OITM", "TNX_QualDate", "Qualification Date", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddDateField("OITM", "TNX_ValidUpto", "Qualification Expiry", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_RiskCls", "Risk Class", 20)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_POBlock", "PO Block", 1)

                ' Business Partner (OCRD)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_VendQual", "Vendor Qualification Status", 30)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_GMPCert", "GMP Certified", 1)
                objMain.objUtilities.AddDateField("OCRD", "TNX_ValidUpto", "Vendor Approval Validity", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_RiskCls", "Vendor Risk Class", 20)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_AuditReq", "Audit Required", 1)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_VQRNo", "VQR No", 30)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_VendStatus", "Vendor Status", 30)

                ' Purchase Order Header (OPOR)
                objMain.objUtilities.AddAlphaField("OPOR", "TNX_PharmaChk", "Pharma Compliance Checked", 1)
                objMain.objUtilities.AddAlphaField("OPOR", "TNX_POStatus", "PO Status", 30)
                objMain.objUtilities.AddAlphaField("OPOR", "TNX_BlockReason", "Block Reason", 254)

                ' Purchase Order Lines (POR1)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_MatQual", "Material Qualification Status", 30)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_VendMap", "Vendor-Material Mapping", 30)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_SpecNo", "Specification No", 30)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_RiskCls", "Risk Class", 20)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_AVLCode", "AVL Reference", 30)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_AVLStatus", "AVL Status", 30)

                ' GRPO Header (OPDN) and Lines (PDN1)
                objMain.objUtilities.AddAlphaField("OPDN", "TNX_QCReq", "QC Required", 1)
                objMain.objUtilities.AddAlphaField("OPDN", "TNX_QCStatus", "QC Status", 30)
                objMain.objUtilities.AddAlphaField("PDN1", "TNX_SpecNo", "Specification No", 30)
                objMain.objUtilities.AddAlphaField("PDN1", "TNX_SampleID", "QC Sample ID", 30)
                objMain.objUtilities.AddAlphaField("PDN1", "TNX_Release", "Batch Release Status", 30)
                objMain.objUtilities.AddAlphaField("PDN1", "TNX_AVLCode", "AVL Code", 30)

                ' Batch Master (OBTN)
                objMain.objUtilities.AddAlphaField("OBTN", "TNX_QCStatus", "QC Status", 30)
                objMain.objUtilities.AddAlphaField("OBTN", "TNX_COARef", "COA Reference", 50)
                objMain.objUtilities.AddAlphaField("OBTN", "TNX_SpecNo", "Specification No", 30)
                objMain.objUtilities.AddDateField("OBTN", "TNX_RelDate", "Batch Release Date", SAPbobsCOM.BoFldSubTypes.st_None)


                CreateMaterialQualificationMaster()
                MaterialTechnicalEvaluation()
                MaterialSpecifications()
                TrialSampleManagement()
                MaterialApproval()
                MaterialRequalification()
                MaterialRiskAssessment()
                CreateVendorQualificationMaster()
                VendorAudit()
                ApprovedVendorList()
                VendorRequalification()
                VendorRiskAssessment()
                VendorPerformance()

                ' Create Material Qualification UDOs
                objMain.CreateMaterialNewRequestUDO()
                objMain.CreateMaterialTechnicalEvaluationUDO()
                objMain.CreateMaterialSpecUDO()
                objMain.CreateTrialSampleUDO()
                objMain.CreateMaterialApprovalUDO()
                objMain.CreateMaterialRequalificationUDO()
                objMain.CreateMaterialRiskAssessmentUDO()
                objMain.CreateVendorQualificationUDO()
                objMain.CreateVendorAuditUDO()
                objMain.CreateApprovedVendorListUDO()
                objMain.CreateVendorRequalificationUDO()
                objMain.CreateVendorRiskAssessmentUDO()
                objMain.CreateVendorPerformanceReviewUDO()

                objMain.objUtilities.AddAlphaField("OITM", "TNX_MatQual", "Material Qualification Status", 30)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_MatType", "Material Type", 30)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_SpecNo", "Linked Spec No", 30)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_SpecVer", "Spec Version", 20)
                objMain.objUtilities.AddDateField("OITM", "TNX_QualDate", "Qualification Date", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddDateField("OITM", "TNX_ValidUpto", "Qualification Expiry", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_RiskCls", "Risk Class", 20)
                objMain.objUtilities.AddAlphaField("OITM", "TNX_POBlock", "PO Block", 1)

                ' Business Partner (OCRD)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_VendQual", "Vendor Qualification Status", 30)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_GMPCert", "GMP Certified", 1)
                objMain.objUtilities.AddDateField("OCRD", "TNX_ValidUpto", "Vendor Approval Validity", SAPbobsCOM.BoFldSubTypes.st_None)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_RiskCls", "Vendor Risk Class", 20)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_AuditReq", "Audit Required", 1)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_VQRNo", "VQR No", 30)
                objMain.objUtilities.AddAlphaField("OCRD", "TNX_VendStatus", "Vendor Status", 30)

                ' Purchase Order Header (OPOR)
                objMain.objUtilities.AddAlphaField("OPOR", "TNX_PharmaChk", "Pharma Compliance Checked", 1)
                objMain.objUtilities.AddAlphaField("OPOR", "TNX_POStatus", "PO Status", 30)
                objMain.objUtilities.AddAlphaField("OPOR", "TNX_BlockReason", "Block Reason", 254)

                ' Purchase Order Lines (POR1)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_MatQual", "Material Qualification Status", 30)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_VendMap", "Vendor-Material Mapping", 30)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_SpecNo", "Specification No", 30)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_RiskCls", "Risk Class", 20)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_AVLCode", "AVL Reference", 30)
                objMain.objUtilities.AddAlphaField("POR1", "TNX_AVLStatus", "AVL Status", 30)

                ' GRPO Header (OPDN) and Lines (PDN1)
                objMain.objUtilities.AddAlphaField("OPDN", "TNX_QCReq", "QC Required", 1)
                objMain.objUtilities.AddAlphaField("OPDN", "TNX_QCStatus", "QC Status", 30)
                objMain.objUtilities.AddAlphaField("PDN1", "TNX_SpecNo", "Specification No", 30)
                objMain.objUtilities.AddAlphaField("PDN1", "TNX_SampleID", "QC Sample ID", 30)
                objMain.objUtilities.AddAlphaField("PDN1", "TNX_Release", "Batch Release Status", 30)
                objMain.objUtilities.AddAlphaField("PDN1", "TNX_AVLCode", "AVL Code", 30)

                ' Batch Master (OBTN)
                objMain.objUtilities.AddAlphaField("OBTN", "TNX_QCStatus", "QC Status", 30)
                objMain.objUtilities.AddAlphaField("OBTN", "TNX_COARef", "COA Reference", 50)
                objMain.objUtilities.AddAlphaField("OBTN", "TNX_SpecNo", "Specification No", 30)
                objMain.objUtilities.AddDateField("OBTN", "TNX_RelDate", "Batch Release Date", SAPbobsCOM.BoFldSubTypes.st_None)


                objMain.TNXRegulatoryDocumentUDO()  ' Regulatory Document Type Master (Master Data)
                objMain.TNXDossierUDO()             ' Dossier Section Master (Master Data)
                objMain.TNXTemplateUDO()            ' CTD/eCTD Template Master (Document + Lines)
                objMain.TNXArtworkUDO()             ' Artwork Type Master (Master Data)
                objMain.TNXSubmissionUDO()          ' Submission Type Master (Master Data)
                objMain.TNXRegistrationUDO()        ' Registration Status Master (Master Data)
                objMain.TNXApprovalUDO()            ' Approval Matrix Master (Document + Lines)
                'Me.CorpTax1()
                'objMain.CorpTaxUDO1()
                'Me.FtaVat1()
                'objMain.FtaVatUDO1()
                'Me.LkMster()
                'objMain.LkMsterUDO()
                'Me.CorporateTaxConfiguration()
                'objMain.CTAXConifgUDO()

                Me.LineClearanceMaster()
                objMain.LineClearanceUDO()
                Me.Errorlogs()
                'objMain.objUtilities.AddDateField("OINV", "EIDate", "EInvoice Date", SAPbobsCOM.BoFldSubTypes.st_None)
                'objMain.objUtilities.AddDateField("OINV", "EITime", "EInvoice Time", SAPbobsCOM.BoFldSubTypes.st_Time)

                'objMain.objUtilities.AddAlphaField("OINV", "XMLPath", "XML Path", 254)
                'objMain.objUtilities.AddAlphaField("OINV", "INVHASH", "ZATCA Invoice Hash", 254)
                'objMain.objUtilities.AddAlphaField("OINV", "UUID", "ZATCA UUID", 254)
                'objMain.objUtilities.AddAlphaField("OINV", "UUID1", "ZATCA UUID1", 254)
                'objMain.objUtilities.AddAlphaField("OINV", "MSG", "ZATCA Message", 254)
                'objMain.objUtilities.AddAlphaField("OINV", "STATUS1", "ZATCA Status", 254)
                'objMain.objUtilities.AddAlphaMemoField("OINV", "XML", "ZATCA XML ", 254)
                'objMain.objUtilities.AddAlphaMemoField("OINV", "QRCODE", "QRCODE", 254Ap
                'objMain.objUtilities.AddAlphaMemoField("OINV", "RESPONSE", "API Response", 254)
                'objMain.objUtilities.AddAlphaMemoField("OINV", "ARQT", "API Request", 254)


                'objMain.objUtilities.addField("OINV", "PAYMEANS", "Payment Means", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "10,30,42,48,1", "In cash,Credit transfer,Payment to a bank account,Bank card, The Instrument is not defined", "")

                'objMain.objUtilities.AddAlphaField("OINV", "STATUS", "ZATCA Status", 254)

                'objMain.objUtilities.addField("OACT", "CTAX", "Corporate Tax Include", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "Yes")

                'objMain.objUtilities.addField("OALT", "WAAlert", "WhatsApp Alert", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")
                'objMain.objUtilities.addField("AOB1", "WASent", "Whatsapp Sent", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Yes,No", "Yes,No", "No")

                'objMain.objUtilities.AddAlphaField("ORIN", "REASON", "Reason", 254)
                'objMain.objUtilities.AddAlphaField("ORIN", "INVNO", "Invoice No", 30)
                'objMain.objUtilities.AddAlphaField("ORIN", "UUID", "ZATCA UUID", 254)
                'objMain.objUtilities.AddAlphaField("ORIN", "UUID1", "ZATCA UUID1", 254)
                'objMain.objUtilities.AddAlphaField("ORIN", "MSG", "ZATCA Message", 254)
                'objMain.objUtilities.AddAlphaField("ORIN", "STATUS1", "ZATCA Status", 254)
                'objMain.objUtilities.AddAlphaMemoField("ORIN", "XML", "ZATCA XML ", 254)
                'objMain.objUtilities.AddAlphaMemoField("ORIN", "QRCODE", "QRCODE", 254)
                'objMain.objUtilities.AddAlphaMemoField("ORIN", "RESPONSE", "API Response", 254)
                'objMain.objUtilities.addField("ORIN", "PAYMEANS", "Payment Means", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "10,30,42,48,1", "In cash,Credit transfer,Payment to a bank account,Bank card, The Instrument is not defined", "")

                'objMain.objUtilities.AddAlphaField("CRD1", "StreetAr", "Street Arabic", 254)
                'objMain.objUtilities.AddAlphaField("CRD1", "SubDivAr", "Sub Division Arabic", 254)
                'objMain.objUtilities.AddAlphaField("CRD1", "CityAr", "City Arabic", 254)

                '' objMain.objUtilities.addField("OALT", "WAAlert", "WhatsApp Alert", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "NO")

                '' objMain.objUtilities.addField("AOB1", "WASent", "Whatsapp Sent", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "NO")
                'objMain.objUtilities.addField("OCRD", "InvType", "Invoice Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, SAPbobsCOM.BoFldSubTypes.st_None, "B2C,B2B", "B2C,B2B", "B2C")

                ' RunSQLScripts()
                'UpdateTransactionNotification()
                'UpdateTransactionNotification()
                ' Me.AddUserFormAuthorizations()

                'Me.AddAlerts()
                'Me.QueryManager()
                Dim err As String

                err = objUtilities.AddDataToNoObjectTable("TNX_DB", Version, Version, "U_VERSION", Version)

                If err <> "" Then
                    MsgBox(err)
                End If
                objMain.objApplication.StatusBar.SetText("Your Database has now been upgraded to Version " + Version + ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
        Return True
    End Function

#End Region

#Region "Create Tables"


    Public Sub TNXRegistrationStatus()

        objMain.objUtilities.CreateTable("TNX_REG_STAT",
                                    "Registration Status Master",
                                    SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT", "StatGroup", "Status Group", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT", "AppForm", "Applicable Form", 100)

        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT", "Editable", "Is Editable", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT", "Final", "Is Final", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT", "AllowCopy", "Allow Copy To", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT", "ReqAppr", "Require Approval", 1)

        objMain.objUtilities.AddInteger("@TNX_REG_STAT", "SeqNo", "Sequence Number", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT", "Status", "Status", 20)

    End Sub
    '================================================================
    ' 7. SUBMISSION TYPE MASTER
    '================================================================
    Public Sub TNXSubmissionType()

        objMain.objUtilities.CreateTable("TNX_REG_SUBTYP",
                                    "Submission Type Master",
                                    SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBTYP", "Category", "Category", 100)

        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBTYP", "DossReq", "Dossier Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBTYP", "ArtReq", "Artwork Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBTYP", "ApprReq", "Approval Required", 1)

        objMain.objUtilities.AddInteger("@TNX_REG_SUBTYP", "ExpDays", "Expected Days", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBTYP", "QueryAll", "Query Allowed", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBTYP", "CCReq", "Change Control Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBTYP", "Status", "Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBTYP", "Remarks", "Remarks", 254)

    End Sub

    Public Sub CreateMissingUserFieldsWithUtilities()
        '======================================================================================
        ' Missing User-Defined Fields Generation Using objUtilities Factories
        '======================================================================================

        ' 1. @SBO_APPAUT
        objMain.objUtilities.AddAlphaField("@SBO_APPAUT", "NAME", "NAME", 50)

        ' 2. @SBO_APPHDR
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Conds", "Conds", 50)

        ' 3-9. @TNX_PCLM_H (Header)
        objMain.objUtilities.AddAlphaField("@TNX_PCLM_H", "ApprovalReq", "ApprovalReq", 10)
        ' Note: "Code" and "Name" are skipped as they are system primary keys for Master Data tables.
        objMain.objUtilities.AddAlphaField("@TNX_PCLM_H", "Equipmenttype", "Equipment type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PCLM_H", "QARequired", "QA Required", 10)
        objMain.objUtilities.AddAlphaField("@TNX_PCLM_H", "RinseRequirement", "RinseReq", 1)
        objMain.objUtilities.AddInteger("@TNX_PCLM_H", "ValidHours", "Valid Hours", SAPbobsCOM.BoFldSubTypes.st_None, 8)

        ' 10-15. @TNX_PCLM_L (Lines)
        objMain.objUtilities.AddAlphaField("@TNX_PCLM_L", "Acceptcriteria", "Acceptance criteria", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PCLM_L", "CheckPoint", "Check Point", 150)
        objMain.objUtilities.AddInteger("@TNX_PCLM_L", "Contacttime", "Contact time", SAPbobsCOM.BoFldSubTypes.st_None, 8)
        ' Note: "LineId" is skipped as it is managed natively by SAP for line child tables.
        objMain.objUtilities.AddAlphaField("@TNX_PCLM_L", "Mandatory", "Mandatory", 1)
        objMain.objUtilities.AddAlphaField("@TNX_PCLM_L", "ResultType", "Result Type", 30)

        ' 16. @TNX_QC_SPLAN_H
        objMain.objUtilities.AddDateField("@TNX_QC_SPLAN_H", "DSR", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)

        ' 17-18. @TNX_REG_ARTTYP
        objMain.objUtilities.AddAlphaField("@TNX_REG_ARTTYP", "CompType", "Component Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_ARTTYP", "LangReq", "Language Required", 1)

        ' 19. @TNX_REG_RENQ
        objMain.objUtilities.AddAlphaMemoField("@TNX_REG_RENQ", "QueryDescription", "Query Details", 254)

        ' 20-28. @TNX_REG_STAT
        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT", "ChangBy", "Changed By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT", "ChangedBy", "User", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT", "FromStat", "From Status", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT", "FromStatus", "Previous Status", 50)
        ' Note: "LineId" skipped (System property)
        objMain.objUtilities.AddAlphaMemoField("@TNX_REG_STAT", "Reason", "Reason", 254)
        objMain.objUtilities.AddDateField("@TNX_REG_STAT", "StatDate", "Status Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_REG_STAT", "StatusDate", "Status Change Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT", "ToStatus", "To Status", 50)

        ' 29. @TNX_STAB_CHMBR
        ' Note: "DocNum" skipped as it is managed natively by SAP for document headers.

    End Sub

    '================================================================
    ' 9. APPROVAL MATRIX MASTER
    '================================================================
    Public Sub TNXApprovalMatrix()

        objMain.objUtilities.CreateTable("TNX_REG_APRH",
                                    "Approval Matrix Header",
                                    SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRH", "FormType", "Form Type", 100)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRH", "Country", "Country", 100)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRH", "SubType", "Submission Type", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRH", "DocTypCod", "Document Type Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRH", "ArtTyp", "Artwork Type", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRH", "RiskClass", "Risk Class", 50)

        objMain.objUtilities.AddDateField("@TNX_REG_APRH", "ActFrom", "Active From", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddDateField("@TNX_REG_APRH", "ActTo", "Active To", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRH", "Status", "Status", 20)


        objMain.objUtilities.CreateTable("TNX_REG_APRL",
                                    "Approval Matrix Lines",
                                    SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddInteger("@TNX_REG_APRL", "LevelNo", "Level Number", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRL", "ApprRole", "Approver Role", 100)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRL", "ApprUser", "Approver User", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRL", "Mandatory", "Mandatory", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRL", "Parallel", "Parallel Approval", 1)

        objMain.objUtilities.AddInteger("@TNX_REG_APRL", "EscDays", "Escalation Days", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRL", "EscUser", "Escalation User", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRL", "Remarks", "Remarks", 254)

    End Sub
    '================================================================
    ' 3. REGULATORY DOCUMENT TYPE MASTER
    '================================================================
    Public Sub TNXRegulatoryDocType()

        objMain.objUtilities.CreateTable("TNX_REG_DOCTYP",
                                    "Regulatory Document Type",
                                    SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOCTYP", "DocCat", "Document Category", 100)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOCTYP", "ValidReq", "Default Validity Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOCTYP", "VerCtrl", "Version Control Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOCTYP", "ApprReq", "Approval Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOCTYP", "AttachMan", "Attachment Mandatory", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOCTYP", "Confid", "Confidential", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOCTYP", "Status", "Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOCTYP", "Remarks", "Remarks", 254)

    End Sub
    '================================================================
    ' 4. DOSSIER SECTION MASTER
    '================================================================
    Public Sub TNXDossierSection()

        objMain.objUtilities.CreateTable("TNX_REG_DOSSEC",
                                    "Dossier Section Master",
                                    SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOSSEC", "ModuleNo", "Module Number", 20)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOSSEC", "ParentSec", "Parent Section", 50)

        objMain.objUtilities.AddInteger("@TNX_REG_DOSSEC", "SecSeq", "Section Sequence", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOSSEC", "DocTypCod", "Document Type Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOSSEC", "Mandatory", "Mandatory", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOSSEC", "AllowMulti", "Allow Multiple Documents", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOSSEC", "ApprReq", "Approval Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOSSEC", "Status", "Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_REG_DOSSEC", "Remarks", "Remarks", 254)

    End Sub
    '================================================================
    ' 5. CTD/eCTD TEMPLATE MASTER
    '================================================================
    Public Sub TNXCTDTemplate()

        objMain.objUtilities.CreateTable("TNX_REG_CTDTMP",
                                    "CTD Template Master",
                                    SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDTMP", "Country", "Country", 100)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDTMP", "AuthCode", "Authority Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDTMP", "TempType", "Template Type", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDTMP", "ProdType", "Product Type", 100)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDTMP", "VersionNo", "Version Number", 20)

        objMain.objUtilities.AddDateField("@TNX_REG_CTDTMP", "EffDate", "Effective Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDTMP", "Status", "Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDTMP", "Remarks", "Remarks", 254)


        objMain.objUtilities.CreateTable("TNX_REG_CTDL",
                                    "CTD Template Lines",
                                    SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDL", "ModuleNo", "Module Number", 20)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDL", "SecCode", "Section Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDL", "SecName", "Section Name", 200)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDL", "DocTypCod", "Document Type Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDL", "Mandatory", "Mandatory", 1)

        objMain.objUtilities.AddInteger("@TNX_REG_CTDL", "SeqNo", "Sequence Number", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDL", "AllowMulti", "Allow Multiple Docs", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDL", "AttachReq", "Attachment Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CTDL", "ApprReq", "Approval Required", 1)

    End Sub
    Public Sub TNXArtworkType()

        objMain.objUtilities.CreateTable("TNX_REG_ARTTYP",
                                   "Artwork Type Master",
                                   SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_REG_ARTTYP", "VerCtrl", "Version Control Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_ARTTYP", "QAAppr", "QA Approval Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_ARTTYP", "RegAppr", "Regulatory Approval Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_ARTTYP", "MktAppr", "Marketing Approval Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_ARTTYP", "AttachMan", "Attachment Mandatory", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_ARTTYP", "Status", "Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_REG_ARTTYP", "Remarks", "Remarks", 254)

    End Sub
    Sub RemincgPilotBatchfields()
        ' =====================================================
        ' PILOT BATCH HEADER TABLE
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_PB_HDR", "Pilot Batch Header", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "PBNo", "Pilot Batch No", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "ForCode", "Formula Code", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "ForVer", "Formula Version", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "ProCode", "Product Item Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "ProName", "Product Name", 150)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "BatType", "Batch Type", 30)

        objMain.objUtilities.AddFloatField("@TNX_PB_HDR", "PlanQty", "Planned Batch Size", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "UOM", "Unit of Measure", 20)

        objMain.objUtilities.AddDateField("@TNX_PB_HDR", "PlanStDt", "Planned Start Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddDateField("@TNX_PB_HDR", "PlanEnDt", "Planned End Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddDateField("@TNX_PB_HDR", "ActStDt", "Actual Start Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddDateField("@TNX_PB_HDR", "ActEnDt", "Actual End Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "IssWhs", "Issue Warehouse", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "ProdWhs", "Production Warehouse", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "QAWhs", "QA Warehouse", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "SAPPrOrd", "SAP Production Order No", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "Status", "Batch Status", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "QAStatus", "QA Status", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "FinDec", "Final Decision", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "Remarks", "Remarks", 254)

        objMain.objUtilities.AddAlphaField("@TNX_PB_HDR", "CrtBy", "Created By", 50)

        objMain.objUtilities.AddDateField("@TNX_PB_HDR", "CrtDate", "Created Date", SAPbobsCOM.BoFldSubTypes.st_None)



        ' =====================================================
        ' YIELD ANALYSIS TABLE
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_PB_YIELD", "Yield Analysis", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddFloatField("@TNX_PB_YIELD", "PlanOut", "Planned Output Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PB_YIELD", "ActOut", "Actual Output Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PB_YIELD", "RejQty", "Rejected Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PB_YIELD", "SampQty", "Sample Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PB_YIELD", "LossQty", "Loss Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PB_YIELD", "YieldPer", "Yield %", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PB_YIELD", "LossPer", "Loss %", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PB_YIELD", "YieldStat", "Yield Status", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_YIELD", "Remarks", "Remarks", 254)



        ' =====================================================
        ' APPROVAL TABLE
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_PB_APPR", "Pilot Batch Approval", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PB_APPR", "Stage", "Approval Stage", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PB_APPR", "Approver", "Approver User", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PB_APPR", "Status", "Status", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_APPR", "ActDate", "Action Date", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_APPR", "Remarks", "Approval Remarks", 254)

        objMain.objUtilities.AddAlphaField("@TNX_PB_APPR", "SignID", "E-Signature Reference", 50)
        ' =====================================================
        ' MATERIAL REQUIREMENT TABLE
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_PB_MAT", "Material Requirement", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PB_MAT", "ItemCode", "Ingredient Item Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PB_MAT", "ItemName", "Ingredient Name", 150)

        objMain.objUtilities.AddAlphaField("@TNX_PB_MAT", "IngType", "Ingredient Type", 30)

        objMain.objUtilities.AddFloatField("@TNX_PB_MAT", "ForQty", "Formula Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PB_MAT", "ForUOM", "Formula UOM", 20)

        objMain.objUtilities.AddFloatField("@TNX_PB_MAT", "ScaleFac", "Scale Factor", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PB_MAT", "ReqQty", "Required Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PB_MAT", "IssWhs", "Issue Warehouse", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PB_MAT", "BatMng", "Batch Managed", 1)

        objMain.objUtilities.AddFloatField("@TNX_PB_MAT", "TolPer", "Tolerance %", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PB_MAT", "Remarks", "Remarks", 254)



        ' =====================================================
        ' MATERIAL ISSUE / DISPENSING TABLE
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_PB_ISS", "Material Issue", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PB_ISS", "ItemCode", "Item Code", 50)

        objMain.objUtilities.AddFloatField("@TNX_PB_ISS", "ReqQty", "Required Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PB_ISS", "IssQty", "Actual Issued Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PB_ISS", "BatchNo", "SAP Batch No", 50)

        objMain.objUtilities.AddDateField("@TNX_PB_ISS", "ExpDate", "Expiry Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PB_ISS", "WhsCode", "Warehouse", 20)

        objMain.objUtilities.AddFloatField("@TNX_PB_ISS", "DiffQty", "Difference Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PB_ISS", "TolStat", "Tolerance Status", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_ISS", "IssBy", "Issued By", 50)

        objMain.objUtilities.AddDateField("@TNX_PB_ISS", "IssDate", "Issue Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PB_ISS", "SAPIssNo", "SAP Goods Issue No", 30)



        ' =====================================================
        ' PROCESS STEPS TABLE
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_PB_PROC", "Process Steps", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddFloatField("@TNX_PB_PROC", "StepNo", "Step No", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PB_PROC", "ProcStage", "Process Stage", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PB_PROC", "Instruc", "Instruction", 254)

        objMain.objUtilities.AddAlphaField("@TNX_PB_PROC", "MacCode", "Machine Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PB_PROC", "MacName", "Machine Name", 100)

        objMain.objUtilities.AddFloatField("@TNX_PB_PROC", "PlanTime", "Planned Time", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PB_PROC", "ActStart", "Actual Start Time", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_PROC", "ActEnd", "Actual End Time", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_PROC", "Oper", "Operator", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PB_PROC", "Status", "Status", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_PROC", "Remarks", "Remarks", 254)



        ' =====================================================
        ' IN PROCESS QC TABLE
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_PB_QC", "In Process QC", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PB_QC", "TestCode", "Test Code", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PB_QC", "TestName", "Test Name", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PB_QC", "StdVal", "Standard Value", 100)

        objMain.objUtilities.AddFloatField("@TNX_PB_QC", "MinVal", "Minimum Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PB_QC", "MaxVal", "Maximum Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PB_QC", "ActVal", "Actual Value", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PB_QC", "Result", "Pass / Fail", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PB_QC", "ChkBy", "Checked By", 50)

        objMain.objUtilities.AddDateField("@TNX_PB_QC", "ChkDate", "Checked Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PB_QC", "Remarks", "Remarks", 254)
    End Sub
    Sub TNXTrainingPlan()

        objMain.objUtilities.CreateTable("TNX_TRNPH", "Training Plan Header", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPH", "PlanCode", "Plan Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPH", "PlanName", "Plan Name", 150)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPH", "TrainType", "Training Type", 50)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPH", "SourceType", "Source Type", 50)

        objMain.objUtilities.AddInteger("@TNX_TRNPH", "SourceEntry", "Source DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPH", "SourceNum", "Source Document Number", 50)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPH", "Department", "Department", 100)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPH", "Trainer", "Trainer Employee ID", 100)

        objMain.objUtilities.AddDateField("@TNX_TRNPH", "PlanDate", "Plan Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddDateField("@TNX_TRNPH", "TargetDate", "Target Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPH", "Mode", "Training Mode", 50)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPH", "Status", "Status", 50)

        objMain.objUtilities.AddInteger("@TNX_TRNPH", "AttachEntry", "Attachment Entry", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPH", "Remarks", "Remarks", 254)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPH", "CreatedBy", "Created By", 100)

        objMain.objUtilities.AddDateField("@TNX_TRNPH", "CreatedOn", "Created On", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPH", "UpdatedBy", "Updated By", 100)

        objMain.objUtilities.AddDateField("@TNX_TRNPH", "UpdatedOn", "Updated On", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("TNX_TRNPL", "Training Plan child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPL", "EmpID", "Employee ID", 50)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPL", "EmpName", "Employee Name", 150)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPL", "Position", "Position", 100)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPL", "Dept", "Department", 100)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPL", "Required", "Required", 10)

        objMain.objUtilities.AddDateField("@TNX_TRNPL", "DueDate", "Due Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPL", "Status", "Status", 50)

        objMain.objUtilities.AddAlphaField("@TNX_TRNPL", "Remarks", "Remarks", 254)


        objMain.objUtilities.CreateTable("TNX_TRNP_Att", "Training Plan Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        ' objMain.objUtilities.CreateTable("TNX_PLCL_ATT", "Line Clearance Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddLinkField("@TNX_TRNP_Att", "TPA", "Target Path", 254, SAPbobsCOM.BoFldSubTypes.st_Link)

        objMain.objUtilities.AddAlphaField("@TNX_TRNP_Att", "FN", "File Name", 254)

        objMain.objUtilities.AddDateField("@TNX_TRNP_Att", "ATD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_TRNP_Att", "FTT", "Free Text", 254)

    End Sub


    Private Sub CreateStabilityStudy11()
        objMain.objUtilities.CreateTable("TNX_STAB_STUDY", "Stability Study", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddInteger("@TNX_STAB_STUDY", "DocNum", "Document Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY", "StudyNo", "Study Number", 30)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY", "PCODE", "Protocol Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY", "StudyType", "Study Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY", "ItemCode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY", "ItemName", "Item Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY", "BatchNum", "Batch No", 50)
        objMain.objUtilities.AddDateField("@TNX_STAB_STUDY", "MfgDate", "Manufacturing Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_STAB_STUDY", "ExpDate", "Existing Expiry Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY", "PORDE", "Production Order", 30)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY", "QC", "QC Batch No", 30)
        objMain.objUtilities.AddDateField("@TNX_STAB_STUDY", "StartDate", "Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_STAB_STUDY", "EndDate", "Planned End Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY", "Status", "Status", 20)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY", "APPS", "Approval Status", 20)

        objMain.objUtilities.CreateTable("TNX_STAB_STUDY_B", "Stability Study Batches", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddInteger("@TNX_STAB_STUDY_B", "LineId", "LineId", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY_B", "BatchNum", "Batch No", 50)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY_B", "WhsCode", "Warehouse", 20)
        objMain.objUtilities.AddFloatField("@TNX_STAB_STUDY_B", "BatchQty", "Batch Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_STAB_STUDY_B", "SampleQty", "Sample Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY_B", "UOM", "UOM", 20)

        objMain.objUtilities.CreateTable("TNX_STAB_STUDY_C", "Stability Study Conditions", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddInteger("@TNX_STAB_STUDY_C", "LineId", "LineId", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY_C", "CNDC", "Condition Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY_C", "CHMBR", "Chamber Code", 30)
        objMain.objUtilities.AddDateField("@TNX_STAB_STUDY_C", "StartDate", "Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_STAB_STUDY_C", "EndDate", "End Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY_C", "Status", "Status", 20)

        objMain.objUtilities.CreateTable("TNX_STAB_STUDY_T", "Stability Study Tests", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddInteger("@TNX_STAB_STUDY_T", "LineId", "LineId", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY_T", "TestCode", "Test Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY_T", "TestName", "Test Name", 100)
        objMain.objUtilities.AddFloatField("@TNX_STAB_STUDY_T", "SpecMin", "Spec Min", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_STAB_STUDY_T", "SpecMax", "Spec Max", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_STUDY_T", "UOM", "UOM", 20)
    End Sub

    Private Sub CreateStabilityProtocol()

        objMain.objUtilities.CreateTable("TNX_STAB_PROTO", "Stability Protocol", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO", "Code", "Protocol Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO", "DocNum", "Protocol num", 30)
        objMain.objUtilities.AddDateField("@TNX_STAB_PROTO", "DATE", "Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO", "Name", "Protocol Name", 100)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO", "ItemGroup", "Product Group", 50)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO", "DFMM", "Dosage Form", 50)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO", "PackType", "Pack Type", 50)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO", "StudyType", "Study Type", 30)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO", "Version", "Version", 20)

        objMain.objUtilities.AddDateField("@TNX_STAB_PROTO", "EFD", "Effective Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO", "Status", "Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO", "APPB", "Approved By", 50)

        objMain.objUtilities.AddDateField("@TNX_STAB_PROTO", "APPD", "Approved Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaMemoField("@TNX_STAB_PROTO", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_STAB_PROTO_T", "Stability Protocol Tests", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO_T", "Code", "Protocol Code", 30)

        objMain.objUtilities.AddInteger("@TNX_STAB_PROTO_T", "LineId", "LineId", SAPbobsCOM.BoFldSubTypes.st_None, 10)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO_T", "TestCode", "Test Code", 30)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO_T", "TestName", "Test Name", 100)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO_T", "MethodCode", "Method Reference", 30)

        objMain.objUtilities.AddFloatField("@TNX_STAB_PROTO_T", "SpecMin", "Spec Min", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.AddFloatField("@TNX_STAB_PROTO_T", "SpecMax", "Spec Max", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO_T", "UOM", "UOM", 20)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO_T", "Critical", "Critical", 1)

        objMain.objUtilities.CreateTable("TNX_STAB_PROTO_S", "Stability Protocol Schedule", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO_S", "Code", "Protocol Code", 30)

        objMain.objUtilities.AddInteger("@TNX_STAB_PROTO_S", "LineId", "LineId", SAPbobsCOM.BoFldSubTypes.st_None, 10)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO_S", "TIME", "Time Point", 20)

        ' objMain.objUtilities.AddInteger("@TNX_STAB_PROTO_S", "DaysFromStart", "Days From Start", SAPbobsCOM.BoFldSubTypes.st_None, 10)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO_S", "DAYSS", "Days From Start", 20)

        objMain.objUtilities.AddFloatField("@TNX_STAB_PROTO_S", "SAMPQ", "Sample Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_STAB_PROTO_S", "TESTR", "Test Required", 1)

    End Sub

    Private Sub CreateShelfLife()
        objMain.objUtilities.CreateTable("TNX_STAB_SHELF", "Shelf Life Analysis", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddInteger("@TNX_STAB_SHELF", "DocNum", "Document Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_SHELF", "ANLY", "Analysis No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_SHELF", "STDY", "Study No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_SHELF", "ItemCode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_SHELF", "BATCH", "Batch No", 50)
        objMain.objUtilities.AddDateField("@TNX_STAB_SHELF", "CNDAT", "Current Expiry", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_STAB_SHELF", "PXDA", "Proposed Expiry", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddInteger("@TNX_STAB_SHELF", "SHLIF", "Shelf Life Months", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_SHELF", "TREND", "Trend", 30)
        objMain.objUtilities.AddAlphaMemoField("@TNX_STAB_SHELF", "RECM", "Recommendation", 254)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_SHELF", "APPST", "Approval Status", 20)

        objMain.objUtilities.CreateTable("TNX_STAB_SHELF_L", "Shelf Life Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddInteger("@TNX_STAB_SHELF_L", "LineId", "LineId", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_SHELF_L", "TestCode", "Test Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_SHELF_L", "TimePoint", "Time Point", 20)
        objMain.objUtilities.AddFloatField("@TNX_STAB_SHELF_L", "RESLT", "Result Value", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_STAB_SHELF_L", "TRENDS", "Trend", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNX_STAB_SHELF_L", "RISK", "Risk Level", 20)
    End Sub

    Public Sub CleaningMethodMaster()

        '====================================================
        ' HEADER TABLE
        '====================================================

        objMain.objUtilities.CreateTable("TNX_PCLM_H", "Cleaning Method Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_PCLM_H", "CleanType", "Clean Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PCLM_H", "EquipType", "Equipment Type", 100)

        objMain.objUtilities.AddInteger("@TNX_PCLM_H", "ValidHrs", "Validity Hours", SAPbobsCOM.BoFldSubTypes.st_None, 10)

        objMain.objUtilities.AddAlphaField("@TNX_PCLM_H", "QAReq", "QA Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_PCLM_H", "SwabReq", "Swab Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_PCLM_H", "RinseReq", "Rinse Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_PCLM_H", "ApprReq", "Approval Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PCLM_H", "Status", "Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PCLM_H", "Remarks", "Remarks", 254)

        '====================================================
        ' LINE TABLE
        '====================================================

        objMain.objUtilities.CreateTable("TNX_PCLM_L", "Cleaning Method Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

        objMain.objUtilities.AddInteger("@TNX_PCLM_L", "StepNo", "Step Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)

        objMain.objUtilities.AddAlphaField("@TNX_PCLM_L", "ChkPoint", "Cleaning Check Point", 200)

        objMain.objUtilities.AddAlphaField("@TNX_PCLM_L", "Method", "Cleaning Method", 254)

        objMain.objUtilities.AddAlphaField("@TNX_PCLM_L", "Chemical", "Cleaning Chemical", 100)

        objMain.objUtilities.AddInteger("@TNX_PCLM_L", "CntTime", "Contact Time", SAPbobsCOM.BoFldSubTypes.st_None, 10)

        objMain.objUtilities.AddAlphaField("@TNX_PCLM_L", "AccCrit", "Acceptance Criteria", 254)

        objMain.objUtilities.AddAlphaField("@TNX_PCLM_L", "ResType", "Result Type", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PCLM_L", "Mandtry", "Mandatory", 1)

    End Sub
    Public Sub EquipmentMaster()

        '====================================================
        ' HEADER TABLE
        '====================================================

        objMain.objUtilities.CreateTable("TNX_PEQP_H", "Equipment Master", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_H", "EquipType", "Equipment Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PEQP_H", "EquipCode", "Equipment Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PEQP_H", "EquipName", "Equipment Name", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_H", "AreaCode", "Production Area", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_H", "Location", "Physical Location", 100)

        objMain.objUtilities.AddFloatField("@TNX_PEQP_H", "Capacity", "Equipment Capacity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_H", "CapUOM", "Capacity UOM", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_H", "MfgSerNo", "Manufacturer Serial No", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_H", "ModelNo", "Model Number", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_H", "CalibReq", "Calibration Required", 1)

        objMain.objUtilities.AddDateField("@TNX_PEQP_H", "CalibDue", "Calibration Due Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddDateField("@TNX_PEQP_H", "LstClean", "Last Cleaning Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_H", "CleanStat", "Cleaning Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_H", "MaintStat", "Maintenance Status", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_H", "Status", "Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_H", "Remarks", "Remarks", 254)


        '====================================================
        ' LINE TABLE
        '====================================================

        objMain.objUtilities.CreateTable("TNX_PEQP_L", "Equipment Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_L", "ItemGroup", "Product Group", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_L", "DosageFrm", "Dosage Form", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_L", "StageCode", "Production Stage", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_L", "AllowFlg", "Allowed Flag", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PEQP_L", "Remarks", "Remarks", 254)

    End Sub
    Public Sub YieldToleranceMaster()

        '====================================================
        ' HEADER TABLE
        '====================================================

        objMain.objUtilities.CreateTable("TNX_PYTM_H",
                                   "Yield Tolerance Master",
                                   SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_PYTM_H",
                                     "ItemCode",
                                     "Item Code",
                                     50)

        objMain.objUtilities.AddAlphaField("@TNX_PYTM_H",
                                     "ItemGroup",
                                     "Product Group",
                                     100)

        objMain.objUtilities.AddAlphaField("@TNX_PYTM_H",
                                     "DosageFrm",
                                     "Dosage Form",
                                     30)

        objMain.objUtilities.AddFloatField("@TNX_PYTM_H",
                                     "BatSizeFr",
                                     "Batch Size From",
                                     SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PYTM_H",
                                     "BatSizeTo",
                                     "Batch Size To",
                                     SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PYTM_H",
                                     "MinYield",
                                     "Minimum Yield Percent",
                                     SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.AddFloatField("@TNX_PYTM_H",
                                     "MaxYield",
                                     "Maximum Yield Percent",
                                     SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.AddFloatField("@TNX_PYTM_H",
                                     "MaxLoss",
                                     "Maximum Loss Percent",
                                     SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.AddFloatField("@TNX_PYTM_H",
                                     "MaxReject",
                                     "Maximum Reject Percent",
                                     SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.AddAlphaField("@TNX_PYTM_H",
                                     "DevReq",
                                     "Deviation Required",
                                     1)

        objMain.objUtilities.AddAlphaField("@TNX_PYTM_H",
                                     "ApprReq",
                                     "Approval Required",
                                     1)

        objMain.objUtilities.AddAlphaField("@TNX_PYTM_H",
                                     "Status",
                                     "Status",
                                     20)



        '====================================================
        ' LINE TABLE
        '====================================================

        objMain.objUtilities.CreateTable("TNX_PYTM_L",
                                   "Yield Tolerance Lines",
                                   SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PYTM_L",
                                     "StageCode",
                                     "Stage Code",
                                     50)

        objMain.objUtilities.AddFloatField("@TNX_PYTM_L",
                                     "AllLoss",
                                     "Allowed Loss Percent",
                                     SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.AddFloatField("@TNX_PYTM_L",
                                     "AllReject",
                                     "Allowed Reject Percent",
                                     SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.AddAlphaField("@TNX_PYTM_L",
                                     "DevReq",
                                     "Deviation Required",
                                     1)

        objMain.objUtilities.AddAlphaField("@TNX_PYTM_L",
                                     "Remarks",
                                     "Remarks",
                                     254)

    End Sub
    Public Sub InProcessQCChecklistMaster()

        '====================================================
        ' HEADER TABLE
        '====================================================

        objMain.objUtilities.CreateTable("TNX_PIQC_H",
                                   "Inprocess QC Checklist",
                                   SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_H",
                                     "ItemCode",
                                     "Item Code",
                                     50)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_H",
                                     "ItemGroup",
                                     "Product Group",
                                     100)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_H",
                                     "DosageFrm",
                                     "Dosage Form",
                                     30)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_H",
                                     "StageCode",
                                     "Production Stage",
                                     50)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_H",
                                     "QCFreq",
                                     "QC Frequency",
                                     30)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_H",
                                     "ApprReq",
                                     "Approval Required",
                                     1)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_H",
                                     "Status",
                                     "Status",
                                     20)



        '====================================================
        ' LINE TABLE
        '====================================================

        objMain.objUtilities.CreateTable("TNX_PIQC_L",
                                   "Inprocess QC Checklist Lines",
                                   SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_L",
                                     "TestCode",
                                     "Test Code",
                                     50)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_L",
                                     "TestName",
                                     "Test Name",
                                     100)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_L",
                                     "Spec",
                                     "Specification",
                                     254)

        objMain.objUtilities.AddFloatField("@TNX_PIQC_L",
                                     "MinValue",
                                     "Minimum Value",
                                     SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PIQC_L",
                                     "MaxValue",
                                     "Maximum Value",
                                     SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_L",
                                     "UOM",
                                     "UOM",
                                     30)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_L",
                                     "ResType",
                                     "Result Type",
                                     30)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_L",
                                     "IsMand",
                                     "Mandatory",
                                     1)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_L",
                                     "IsCrit",
                                     "Critical Test",
                                     1)

        objMain.objUtilities.AddAlphaField("@TNX_PIQC_L",
                                     "DevReq",
                                     "Deviation Required",
                                     1)

    End Sub
    Public Sub LineClearanceChecklistMaster()

        '====================================================
        ' HEADER TABLE
        '====================================================

        objMain.objUtilities.CreateTable("TNX_PLCC_H", "Line Clearance Checklist", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_H", "ClrType", "Clearance Type", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_H", "DosageFrm", "Dosage Form", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_H", "StageCode", "Production Stage", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_H", "AreaCode", "Production Area", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_H", "QAReq", "QA Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_H", "ApprReq", "Approval Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_H", "Status", "Status", 20)


        '====================================================
        ' LINE TABLE
        '====================================================

        objMain.objUtilities.CreateTable("TNX_PLCC_L", "Line Clearance Checklist Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_L", "ChkPoint", "Checklist Check Point", 254)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_L", "ExpResult", "Expected Result", 254)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_L", "ResType", "Result Type", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_L", "IsMand", "Mandatory", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_L", "IsCrit", "Critical Checkpoint", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_L", "DevReq", "Deviation Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PLCC_L", "Remarks", "Remarks", 254)

    End Sub
    Public Sub DowntimeReasonMaster()

        '====================================================
        ' HEADER TABLE
        '====================================================

        objMain.objUtilities.CreateTable("TNX_PDTR_H", "Downtime Reason Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_PDTR_H", "Category", "Category", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PDTR_H", "IsPlan", "Planned Downtime", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PDTR_H", "DevReq", "Deviation Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PDTR_H", "CAPAReq", "CAPA Required", 1)

        objMain.objUtilities.AddInteger("@TNX_PDTR_H", "ThresMin", "Threshold Minutes", SAPbobsCOM.BoFldSubTypes.st_None, 10)

        objMain.objUtilities.AddAlphaField("@TNX_PDTR_H", "ApprReq", "Approval Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PDTR_H", "Status", "Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PDTR_H", "Remarks", "Remarks", 254)

    End Sub
    Public Sub ProductionStageMaster()

        '====================================================
        ' HEADER TABLE
        '====================================================

        objMain.objUtilities.CreateTable("TNX_PSTG_H", "Production Stage Master", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddInteger("@TNX_PSTG_H", "StageSeq", "Stage Sequence", SAPbobsCOM.BoFldSubTypes.st_None, 10)

        objMain.objUtilities.AddAlphaField("@TNX_PSTG_H", "DosageFrm", "Dosage Form", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PSTG_H", "IsMand", "Mandatory Stage", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PSTG_H", "QCReq", "QC Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PSTG_H", "LCReq", "Line Clearance Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PSTG_H", "CleanReq", "Cleaning Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PSTG_H", "ApprReq", "Approval Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PSTG_H", "Status", "Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PSTG_H", "Remarks", "Remarks", 254)


        '====================================================
        ' LINE TABLE
        '====================================================

        objMain.objUtilities.CreateTable("TNX_PSTG_L", "Production Stage Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PSTG_L", "ParamCode", "Parameter Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PSTG_L", "ParamName", "Parameter Name", 100)

        objMain.objUtilities.AddFloatField("@TNX_PSTG_L", "MinValue", "Minimum Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PSTG_L", "MaxValue", "Maximum Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PSTG_L", "UOM", "UOM", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PSTG_L", "IsCrit", "Critical Parameter", 1)

        objMain.objUtilities.AddAlphaField("@TNX_PSTG_L", "DevReq", "Deviation Required", 1)

    End Sub
    Private Sub LineClearanceMaster()

        '========================================================
        ' HEADER TABLE
        '========================================================

        objMain.objUtilities.CreateTable("TNX_PLCL_H", "Line Clearance Header", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "LCType", "LC Type", 30)

        objMain.objUtilities.AddInteger("@TNX_PLCL_H", "ProdOrdEntry", "Production Order Entry", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddInteger("@TNX_PLCL_H", "ProdOrdNo", "Production Order No", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "ItemCode", "Item Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "ItemName", "Item Name", 150)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "BatchNo", "Batch Number", 50)

        objMain.objUtilities.AddFloatField("@TNX_PLCL_H", "PlannedQty", "Planned Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "UOM", "UOM", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "WhsCode", "Warehouse Code", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "AreaCode", "Area Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "AreaName", "Area Name", 150)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "LineCode", "Line Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "EquipCode", "Main Equipment", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "PreviousBatch", "Previous Batch", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "PreviousItem", "Previous Product", 150)

        objMain.objUtilities.AddDateField("@TNX_PLCL_H", "ClearanceDate", "Clearance Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "ClearanceTime", "Clearance Time", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "RequestedBy", "Requested By", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "CheckedBy", "Checked By", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "VerifiedBy", "Verified By", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "Status", "Status", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "ApprovalStatus", "Approval Status", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "ApprovedBy", "Approved By", 100)

        objMain.objUtilities.AddDateField("@TNX_PLCL_H", "ApprovedDate", "Approved Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_H", "Remarks", "Remarks", 254)



        '========================================================
        ' CHECKLIST LINES
        '========================================================

        objMain.objUtilities.CreateTable("TNX_PLCL_L", "Line Clearance Checklist", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_L", "CheckCode", "Check Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_L", "CheckPoint", "Checkpoint", 200)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_L", "Category", "Category", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_L", "Expected", "Expected Condition", 254)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_L", "Observed", "Observed Condition", 254)

        ' objMain.objUtilities.AddDateField("@TNX_PLCL_L", "CheckedDate", "Checked Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_L", "CheckedBy", "Checked By", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_L", "Result", "Result", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_L", "Remarks", "Remarks", 254)



        '========================================================
        ' EQUIPMENT LINES
        '========================================================

        objMain.objUtilities.CreateTable("TNX_PLCL_EQP", "Line Clearance Equipment", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_EQP", "EquipCode", "Equipment Code", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_EQP", "EquipName", "Equipment Name", 150)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_EQP", "CleaningLogNo", "Cleaning Log Number", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_EQP", "CleaningStatus", "Cleaning Status", 50)

        objMain.objUtilities.AddDateField("@TNX_PLCL_EQP", "CalibDueDate", "Calibration Due Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_EQP", "CalibStatus", "Calibration Status", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_EQP", "ReadyStatus", "Ready Status", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_EQP", "Rmarks", "Remarks", 254)



        '========================================================
        ' ATTACHMENTS
        '========================================================

        objMain.objUtilities.CreateTable("TNX_PLCL_ATT", "Line Clearance Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddLinkField("@TNX_PLCL_ATT", "TPA", "Target Path", 254, SAPbobsCOM.BoFldSubTypes.st_Link)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_ATT", "FN", "File Name", 254)

        objMain.objUtilities.AddDateField("@TNX_PLCL_ATT", "ATD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PLCL_ATT", "FTT", "Free Text", 254)

    End Sub
    '================================================================
    ' 1. REGULATORY AUTHORITY MASTER
    '================================================================
    Public Sub TNXRegulatoryAuthority()

        objMain.objUtilities.CreateTable("TNX_REG_AUTH",
                                   "Regulatory Authority Master",
                                   SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_REG_AUTH", "Country", "Country", 100)

        objMain.objUtilities.AddAlphaField("@TNX_REG_AUTH", "AuthType", "Authority Type", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_AUTH", "Website", "Website", 200)

        objMain.objUtilities.AddAlphaField("@TNX_REG_AUTH", "Email", "Email", 100)

        objMain.objUtilities.AddAlphaField("@TNX_REG_AUTH", "Phone", "Phone", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_AUTH", "SubMode", "Submission Mode", 50)

        objMain.objUtilities.AddInteger("@TNX_REG_AUTH", "DefTime", "Default Timeline", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_REG_AUTH", "AgentReq", "Local Agent Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_AUTH", "Status", "Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_REG_AUTH", "Remarks", "Remarks", 254)

    End Sub
    '================================================================
    ' 2. COUNTRY REGULATORY CONFIGURATION
    '================================================================
    Public Sub TNXCountryRegConfig()

        objMain.objUtilities.CreateTable("TNX_REG_CNFG",
                                   "Country Regulatory Config",
                                   SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFG", "CntryCode", "Country Code", 10)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFG", "AuthCode", "Authority Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFG", "RegFormat", "Regulatory Format", 50)

        objMain.objUtilities.AddInteger("@TNX_REG_CNFG", "ValidYrs", "Default Validity Years", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddInteger("@TNX_REG_CNFG", "RenewDays", "Renewal Before Days", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFG", "ArtReq", "Artwork Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFG", "AgentReq", "Local Agent Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFG", "LangReq", "Language Requirement", 100)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFG", "Status", "Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFG", "Remarks", "Remarks", 254)


        objMain.objUtilities.CreateTable("TNX_REG_CNFL",
                                   "Country Config Lines",
                                   SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFL", "DocTypCod", "Document Type Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFL", "DocTypNam", "Document Type Name", 200)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFL", "Mandatory", "Mandatory", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFL", "ValidReq", "Validity Required", 1)

        objMain.objUtilities.AddInteger("@TNX_REG_CNFL", "MinValid", "Minimum Validity Days", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFL", "AttachReq", "Attachment Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_REG_CNFL", "Remarks", "Remarks", 254)

    End Sub
    Sub SampleCollection()
        objMain.objUtilities.CreateTable("TNX_QCSC_H", "Sample Collection", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "DocNum", "Sample Collection No", 11)
        objMain.objUtilities.AddDateField("@TNX_QCSC_H", "DATE", "Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "SREGN", "Sample Registration No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "SREGDE", "Sample Registration DocEntry", 11)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "SPTYE", "Sample Type", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "STYPE", "Source Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "SODCEN", "Source Document Entry", 11)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "SORCDC", "Source Document No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "ITMCD", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "ITMNM", "Item Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "BATNO", "Batch No", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "WhsCode", "Warehouse", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "BINCD", "Bin Location", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "COLBY", "Collected By", 50)
        objMain.objUtilities.AddDateField("@TNX_QCSC_H", "COLLDT", "Collection Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_QCSC_H", "COLLTM", "Collection Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddFloatField("@TNX_QCSC_H", "ReqQty", "Required Sample Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_QCSC_H", "CollQty", "Collected Sample Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "UOM", "UOM", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "SAMTD", "Sampling Method", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "SAMPLN", "Sampling Plan", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "CNCNT", "No. of Containers Sampled", 11)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "TCNT", "Total Containers / Packs", 11)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "SSTS", "Seal Status", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "LBLSTS", "Label Status", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "STCDN", "Storage Condition", 100)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "TEMPT", "Temperature at Collection", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "HUMDT", "Humidity %", 100)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "VIOBS", "Visual Observation", 254)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "Status", "Document Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "LABHN", "Lab Handover Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "RMRKS", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_H", "ATTEN", "SAP Attachment Entry", 11)

        objMain.objUtilities.CreateTable("TNX_QCSC_L", "Sample Container Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_QCSC_L", "CNTNO", "Container No", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_L", "CNTYPE", "Container Type", 50)
        objMain.objUtilities.AddFloatField("@TNX_QCSC_L", "PCKSZ", "Pack Size", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_L", "PUOM", "Pack UOM", 20)
        objMain.objUtilities.AddFloatField("@TNX_QCSC_L", "SAQTY", "Sample Qty from Container", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_L", "UOM", "UOM", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_L", "SAMPT", "Sample Point", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_L", "SELNO", "Seal No", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_L", "SELIN", "Seal Inact", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_L", "VISOB", "Visual Observation", 254)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_L", "LINST", "Line Status", 20)

        objMain.objUtilities.CreateTable("TNX_QCSC_COC", "Chain of Custody", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_QCSC_COC", "Action", "Action", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_COC", "FUSER", "From User", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_COC", "TOUSR", "To User", 50)

        objMain.objUtilities.AddDateField("@TNX_QCSC_COC", "ADATE", "Action Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_QCSC_COC", "ATIME", "Action Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_COC", "LCTN", "Location", 10)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_COC", "REMK", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_COC", "Esign", "Electronic Signature", 50)

        objMain.objUtilities.CreateTable("TNX_QCSC_ATT", "Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_QCSC_ATT", "FNAME", "File Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_ATT", "FILTY", "File Type", 30)
        objMain.objUtilities.AddLinkField("@TNX_QCSC_ATT", "ATPTH", "Attachment Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)

        objMain.objUtilities.AddAlphaField("@TNX_QCSC_ATT", "SAPAT", "SAP Attachment", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QCSC_ATT", "UPBY", "Updated By", 50)
        objMain.objUtilities.AddDateField("@TNX_QCSC_ATT", "UPDATE", "Uploaded Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_QCSC_ATT", "REMK", "Remarks", 254)

    End Sub



    Sub SampleRegistration()
        objMain.objUtilities.CreateTable("TNX_QASMPH", "Sample Registration", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "DocNum", "Document Number", 100)
        objMain.objUtilities.AddDateField("@TNX_QASMPH", "DATE", "Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "SMPNO", "Sample number", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "SAMCT", "Sample category", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "SOUDC", "Source document", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "ITMCDE", "SAP item code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "ITMNM", "Item name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "BATNO", "SAP batch number", 100)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "PRODR", "Production order number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "UOM", "Unit of measure", 20)
        objMain.objUtilities.AddDateField("@TNX_QASMPH", "SAMDT", "Sample collection date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "COLLBY", "QA person", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "SPECD", "Linked specification", 30)


        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "WhsCode", "Warehouse", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "BPCode", "Vendor/customer code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "BPName", "Vendor/customer name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "GRPON", "GRPO number", 100)

        objMain.objUtilities.AddFloatField("@TNX_QASMPH", "RQty", "Received quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_QASMPH", "SAMQTY", "Sample quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "PRITY", "Priority", 20)

        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "QCSTS", "QC status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "RESTS", "Batch release status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPH", "SPVRN", "Specification version", 20)


        objMain.objUtilities.CreateTable("TNX_QASMPL", "ItemDetails Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPL", "TESTCD", "QC test code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPL", "TENM", "Test name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPL", "TESTCA", "Chemical/Micro/Physical", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPL", "MEDCD", "Testing method code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPL", "MTDNM", "Method name", 100)

        objMain.objUtilities.AddFloatField("@TNX_QASMPL", "MINVAL", "Minimum Limit", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_QASMPL", "MAXVAL", "Maximum Limit", SAPbobsCOM.BoFldSubTypes.st_Sum)


        objMain.objUtilities.AddAlphaField("@TNX_QASMPL", "STDVAL", "Standard expected value", 100)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPL", "UOM", "Test UOM", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPL", "ISMNDT", "Mandatory test", 1)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPL", "TESTS", "Test status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPL", "ASSIGN", "Lab analyst", 50)
        objMain.objUtilities.AddDateField("@TNX_QASMPL", "TADATE", "Expected test date", SAPbobsCOM.BoFldSubTypes.st_None)


        objMain.objUtilities.CreateTable("TNX_QASMPA", "Attachment Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddLinkField("@TNX_QASMPA", "FIPTH", "File Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPA", "FINM", "File name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPA", "UPBY", "Uploaded user", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QASMPA", "AType", "Attachment type", 50)
        objMain.objUtilities.AddDateField("@TNX_QASMPA", "UPDATE", "Upload date", SAPbobsCOM.BoFldSubTypes.st_None)
        'objMain.objUtilities.AddAlphaField("@TNX_QASMPA", "REMK", "Remarks", 300)


    End Sub

    Private Sub CreateTrainingTables()

        ' =====================================================
        ' TRAINING EXECUTION HEADER
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_TRNEH", "Training Execution", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEH", "TrainPlan", "Training Plan Entry", 20)
        objMain.objUtilities.AddAlphaField("@TNX_TRNEH", "TrainPlNo", "Training Plan Number", 30)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEH", "SessCode", "Session Code", 30)

        objMain.objUtilities.AddDateField("@TNX_TRNEH", "SessDate", "Session Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEH", "Trainer", "Trainer", 100)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEH", "Location", "Training Location", 100)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEH", "Mode", "Training Mode", 30)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEH", "StartTime", "Start Time", 20)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEH", "EndTime", "End Time", 20)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEH", "Status", "Status", 30)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEH", "AttEntry", "Attachment Entry", 20)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEH", "Remarks", "Remarks", 254)



        ' =====================================================
        ' TRAINING ATTENDANCE LINES
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_TRNEL", "Training Attendance", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEL", "EmpID", "Employee ID", 30)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEL", "EmpName", "Employee Name", 100)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEL", "Attended", "Attended", 1)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEL", "AttTime", "Attendance Time", 20)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEL", "Sign", "E-Signature", 100)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEL", "Status", "Status", 30)

        objMain.objUtilities.AddAlphaField("@TNX_TRNEL", "Remarks", "Remarks", 254)



        ' =====================================================
        ' TRAINING ASSESSMENT LINES
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_TRNASM", "Training Assessment", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_TRNASM", "EmpID", "Employee ID", 30)

        objMain.objUtilities.AddAlphaField("@TNX_TRNASM", "AssType", "Assessment Type", 30)

        objMain.objUtilities.AddFloatField("@TNX_TRNASM", "TotMarks", "Total Marks", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_TRNASM", "MarksOb", "Marks Obtained", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_TRNASM", "PassPer", "Pass Percentage", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_TRNASM", "Result", "Result", 20)

        objMain.objUtilities.AddAlphaField("@TNX_TRNASM", "AssBy", "Assessed By", 100)

        objMain.objUtilities.AddDateField("@TNX_TRNASM", "AssDate", "Assessment Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_TRNASM", "Remarks", "Remarks", 254)



        ' =====================================================
        ' TRAINING CERTIFICATE HEADER
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_TRNCH", "Training Certificate", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCH", "CertCode", "Certificate Code", 30)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCH", "TrainExec", "Training Execution Entry", 20)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCH", "TrainPlan", "Training Plan Entry", 20)

        objMain.objUtilities.AddDateField("@TNX_TRNCH", "IssueDate", "Issue Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddDateField("@TNX_TRNCH", "ValidFrm", "Valid From", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddDateField("@TNX_TRNCH", "ValidTo", "Valid To", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCH", "CertStat", "Certificate Status", 30)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCH", "ApprBy", "Approved By", 100)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCH", "AttEntry", "Attachment Entry", 20)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCH", "Remarks", "Remarks", 254)



        ' =====================================================
        ' TRAINING CERTIFICATE LINES
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_TRNCL", "Training Certificate Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCL", "EmpID", "Employee ID", 30)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCL", "EmpName", "Employee Name", 100)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCL", "Result", "Result", 20)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCL", "CertNo", "Certificate Number", 50)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCL", "Qualify", "Qualified", 1)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCL", "QualArea", "Qualification Area", 50)

        objMain.objUtilities.AddAlphaField("@TNX_TRNCL", "Remarks", "Remarks", 254)



        ' =====================================================
        ' EMPLOYEE TRAINING MATRIX HEADER
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_TRNMH", "Employee Training Matrix", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_TRNMH", "EmpID", "Employee ID", 30)

        objMain.objUtilities.AddAlphaField("@TNX_TRNMH", "EmpName", "Employee Name", 100)

        objMain.objUtilities.AddAlphaField("@TNX_TRNMH", "Dept", "Department", 50)

        objMain.objUtilities.AddAlphaField("@TNX_TRNMH", "Positn", "Position", 50)

        objMain.objUtilities.AddDateField("@TNX_TRNMH", "JoinDate", "Joining Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_TRNMH", "Status", "Status", 20)



        ' =====================================================
        ' EMPLOYEE TRAINING MATRIX LINES
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_TRNML", "Employee Training Matrix Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_TRNML", "DocType", "Document Type", 50)

        objMain.objUtilities.AddAlphaField("@TNX_TRNML", "DocCode", "Document Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_TRNML", "TrainReq", "Training Required", 1)

        objMain.objUtilities.AddDateField("@TNX_TRNML", "LastTrDt", "Last Training Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddDateField("@TNX_TRNML", "NextDue", "Next Due Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_TRNML", "Result", "Result", 20)

        objMain.objUtilities.AddAlphaField("@TNX_TRNML", "Qualify", "Qualified", 1)

        objMain.objUtilities.AddAlphaField("@TNX_TRNML", "Status", "Status", 30)

    End Sub

    Sub QABatchReleaseTables()

        '================ HEADER TABLE =================
        objMain.objUtilities.CreateTable("TNX_QABR_H", "QA Batch Release Header", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "BRNo", "Batch Release No", 30)
        objMain.objUtilities.AddDateField("@TNX_QABR_H", "BRDate", "Release Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "Status", "Status", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "SourceType", "Source Type", 20)
        objMain.objUtilities.AddInteger("@TNX_QABR_H", "SrcDocEnt", "Source DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddInteger("@TNX_QABR_H", "SrcDocNum", "Source DocNum", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddDateField("@TNX_QABR_H", "DCD", "Documnet Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "ItemCode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "ItemName", "Item Name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "BatchNo", "Batch No", 100)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "WhsCode", "Warehouse Code", 20)
        objMain.objUtilities.AddFloatField("@TNX_QABR_H", "BatchQty", "Batch Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "SampleNo", "Sample No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "LabTestNo", "Lab Test No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "QCResult", "QC Result", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "RelDecision", "Release Decision", 20)

        objMain.objUtilities.AddFloatField("@TNX_QABR_H", "ReleaseQty", "Release Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_QABR_H", "RejectedQty", "Rejected Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_QABR_H", "HoldQty", "Hold Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "RetestReq", "Retest Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "DevReq", "Deviation Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "DeviationNo", "Deviation No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "CAPANo", "CAPA No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "COAReq", "COA Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "COANo", "COA No", 30)

        objMain.objUtilities.AddDateField("@TNX_QABR_H", "MfgDate", "Manufacturing Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_QABR_H", "ExpDate", "Expiry Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "ReleasedBy", "Released By", 50)
        objMain.objUtilities.AddDateField("@TNX_QABR_H", "RelDate", "Release Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "Remarks", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "SAPUpdated", "SAP Updated", 1)
        objMain.objUtilities.AddDateField("@TNX_QABR_H", "SAPUpdDate", "SAP Update Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "CreatedBy", "Created By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_H", "ApprovedBy", "Approved By", 50)


        '================ CHILD TABLE 1 : TEST SUMMARY =================
        objMain.objUtilities.CreateTable("TNX_QABR_TST", "QA Batch Release Test Summary", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_QABR_TST", "TestCode", "Test Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_TST", "TestName", "Test Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_TST", "Method", "Test Method", 100)
        objMain.objUtilities.AddFloatField("@TNX_QABR_TST", "MinValue", "Minimum Value", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_QABR_TST", "MaxValue", "Maximum Value", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_QABR_TST", "ActualVal", "Actual Value", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_TST", "UOM", "UOM", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_TST", "Result", "Result", 20)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_TST", "Remarks", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_TST", "Analyst", "Analyst", 50)
        objMain.objUtilities.AddDateField("@TNX_QABR_TST", "TestDate", "Test Date", SAPbobsCOM.BoFldSubTypes.st_None)


        '================ CHILD TABLE 2 : APPROVAL =================
        objMain.objUtilities.CreateTable("TNX_QABR_APR", "QA Batch Release Approval", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddInteger("@TNX_QABR_APR", "Level", "Approval Level", SAPbobsCOM.BoFldSubTypes.st_None, 3)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_APR", "AppCode", "Approver Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_APR", "AppName", "Approver Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_APR", "Role", "Role", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_APR", "Status", "Status", 20)
        objMain.objUtilities.AddDateField("@TNX_QABR_APR", "ActionDate", "Action Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_APR", "ActionTime", "Action Time", 10)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_APR", "Remarks", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_APR", "ESign", "E-Sign", 100)


        '================ CHILD TABLE 3 : DOCUMENTS =================
        objMain.objUtilities.CreateTable("TNX_QABR_DOC", "QA Batch Release Documents", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_QABR_DOC", "DocType", "Document Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_DOC", "FileName", "File Name", 200)
        objMain.objUtilities.AddLinkField("@TNX_QABR_DOC", "FilePath", "File Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddDateField("@TNX_QABR_DOC", "AttachDate", "Attach Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_DOC", "AttachedBy", "Attached By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_QABR_DOC", "Remarks", "Remarks", 254)



    End Sub
    Sub CAPAManagementTables()

        '================ HEADER TABLE =================
        objMain.objUtilities.CreateTable("TNX_CAPAH", "TNX Pharma CAPA Header", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "CAPA_NO", "CAPA Number", 30)
        objMain.objUtilities.AddDateField("@TNX_CAPAH", "CAPA_DATE", "CAPA Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_CAPAH", "DDK", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "SOURCE_TYPE", "Source Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "SOURCE_NO", "Source Document No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "DEV_NO", "Linked Deviation No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "BATCH_NO", "Batch No", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "ITEM_CODE", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "ITEM_NAME", "Item Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "DEPT", "Department", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "SEVERITY", "Severity", 20)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "PRIORITY", "Priority", 20)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "ROOT_CAUSE", "Root Cause Summary", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "IMPACT", "Impact Assessment", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "RISK_LEVEL", "Risk Level", 20)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "CAPA_OWNER", "CAPA Owner", 50)
        objMain.objUtilities.AddDateField("@TNX_CAPAH", "TARGET_DATE", "Target Closure Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "STATUS", "CAPA Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "QA_REVIEW_BY", "QA Reviewer", 50)
        objMain.objUtilities.AddDateField("@TNX_CAPAH", "QA_REVIEW_DT", "QA Review Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "APPROVED_BY", "Approved By", 50)
        objMain.objUtilities.AddDateField("@TNX_CAPAH", "APPROVED_DT", "Approved Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_CAPAH", "CLOSURE_DATE", "Closure Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAH", "REMARKS", "Remarks", 254)


        '================ CHILD TABLE 1 : CAPA ACTION LINES =================
        objMain.objUtilities.CreateTable("TNX_CAPAL", "CAPA Action Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_CAPAL", "ACTION_TYPE", "Corrective / Preventive", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAL", "ACTION_DESC", "Action Description", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAL", "RESP_PERSON", "Responsible Person", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAL", "DEPT", "Department", 50)
        objMain.objUtilities.AddDateField("@TNX_CAPAL", "DUE_DATE", "Due Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_CAPAL", "COMPLETION_DT", "Completion Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAL", "ACTION_STATUS", "Action Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAL", "EVIDENCE_REQ", "Evidence Required", 10)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAL", "REMARKS", "Remarks", 254)


        '================ CHILD TABLE 2 : EFFECTIVENESS CHECK =================
        objMain.objUtilities.CreateTable("TNX_CAPAE", "Effectiveness Check Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddDateField("@TNX_CAPAE", "CHECK_DATE", "Check Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAE", "CHECK_BY", "Checked By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAE", "METHOD", "Verification Method", 100)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAE", "RESULT", "Effective / Not Effective", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAE", "OBSERVATION", "Observation", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAE", "REOPEN_REQ", "Reopen Required", 10)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAE", "NEXT_ACTION", "Next Action", 254)


        '================ CHILD TABLE 3 : ATTACHMENT / EVIDENCE =================
        objMain.objUtilities.CreateTable("TNX_CAPAA", "Attachment Evidence Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_CAPAA", "DOC_TYPE", "Document Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAA", "FILE_NAME", "File Name", 150)
        objMain.objUtilities.AddLinkField("@TNX_CAPAA", "FILE_PATH", "File Path / Attachment Entry", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        ' objMain.objUtilities.AddAlphaField("@TNX_CAPAA", "FILE_PATH", "File Path / Attachment Entry", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAA", "UPLOADED_BY", "Uploaded By", 50)
        objMain.objUtilities.AddDateField("@TNX_CAPAA", "UPLOAD_DATE", "Upload Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAA", "REMARKS", "Remarks", 254)


        '================ CHILD TABLE 4 : WORKFLOW / APPROVAL =================
        objMain.objUtilities.CreateTable("TNX_CAPAW", "Workflow Approval Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_CAPAW", "STAGE", "Approval Stage", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAW", "APPROVER", "Approver", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAW", "APPROVAL_STATUS", "Status", 30)
        objMain.objUtilities.AddDateField("@TNX_CAPAW", "APPROVAL_DATE", "Approval Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CAPAW", "COMMENTS", "Comments", 254)

    End Sub

    Sub QCLabTestingMaster()


        objMain.objUtilities.CreateTable("TNXPH_QCLABH", "QC Lab Testing Header", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "TestNo", "Test Number", 30)
        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "SampleNo", "Sample Number", 30)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "ItemCode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "ItemName", "Item Name", 150)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "BatchNo", "Batch Number", 50)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "SpecCode", "Specification Code", 50)
        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "SpecVersion", "Specification Version", 20)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "SampleType", "Sample Type", 30)
        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "SourceType", "Source Type", 50)

        objMain.objUtilities.AddDateField("@TNXPH_QCLABH", "TestDate", "Test Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "Analyst", "Analyst", 100)
        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "Reviewer", "Reviewer", 100)
        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "warehouse", "Ware house", 100)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "OverallResult", "Overall Result", 30)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "Status", "Status", 50)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "Remarks", "Remarks", 254)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "CreatedBy", "Created By", 100)
        objMain.objUtilities.AddDateField("@TNXPH_QCLABH", "CreatedOn", "Created On", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "UpdatedBy", "Updated By", 100)
        objMain.objUtilities.AddDateField("@TNXPH_QCLABH", "UpdatedOn", "Updated On", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddInteger("@TNXPH_QCLABH", "SampleEntry", "Sample Entry", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddInteger("@TNXPH_QCLABH", "SourceDocEntry", "Source Doc Entry", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "SourceDocNum", "Source Doc Num", 30)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "TestMethod", "Test Method", 100)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "TestStartTime", "Test Start Time", 10)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "TestEndTime", "Test End Time", 10)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "DeviationReq", "Deviation Required", 1)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "DeviationNo", "Deviation No", 30)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "CAPAReq", "CAPA Required", 1)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "CAPANo", "CAPA No", 30)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "QAApprovedBy", "QA Approved By", 100)

        objMain.objUtilities.AddDateField("@TNXPH_QCLABH", "QAApprovedDate", "QA Approved Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "QAApprovalStatus", "QA Approval Status", 30)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABH", "ElectronicSign", "Electronic Sign", 254)


        objMain.objUtilities.CreateTable("TNXPH_QCLABL", "QC Lab Test Results", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABL", "TestCode", "Test Code", 30)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABL", "TestName", "Test Name", 100)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABL", "Parameter", "Parameter", 100)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABL", "TestMethod", "Test Method", 100)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABL", "Unit", "Unit", 30)

        objMain.objUtilities.AddFloatField("@TNXPH_QCLABL", "MinValue", "Minimum Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNXPH_QCLABL", "MaxValue", "Maximum Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNXPH_QCLABL", "ActualValue", "Actual Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABL", "TextResult", "Text Result", 100)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABL", "ResultStatus", "Result Status", 30)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABL", "InstrumentName", "Instrument Name", 100)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABL", "TestedBy", "Tested By", 100)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABL", "Remarks", "Remarks", 254)



        objMain.objUtilities.CreateTable("TNXPH_QCLABATT", "QC Lab Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABATT", "AttachType", "Attachment Type", 50)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABATT", "FileName", "File Name", 254)

        objMain.objUtilities.AddLinkField("@TNXPH_QCLABATT", "FilePath", "File Path", 254, SAPbobsCOM.BoFldSubTypes.st_Link)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABATT", "FileExt", "File Extension", 20)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABATT", "UploadedBy", "Uploaded By", 100)

        objMain.objUtilities.AddDateField("@TNXPH_QCLABATT", "UploadedDate", "Uploaded Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABATT", "Remarks", "Remarks", 254)



        objMain.objUtilities.CreateTable("TNXPH_QCLABAPP", "QC Lab Approvals", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddInteger("@TNXPH_QCLABAPP", "ApprovalLevel", "Approval Level", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABAPP", "ApproverRole", "Approver Role", 100)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABAPP", "ApproverUser", "Approver User", 100)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABAPP", "ApprovalStatus", "Approval Status", 50)

        objMain.objUtilities.AddDateField("@TNXPH_QCLABAPP", "ApprovalDate", "Approval Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABAPP", "ApprovalTime", "Approval Time", 20)

        objMain.objUtilities.AddAlphaField("@TNXPH_QCLABAPP", "Comments", "Comments", 254)

    End Sub

    Sub SpecificationMaster1()



        objMain.objUtilities.CreateTable("TNX_PH_QSPECH", "Specification Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECH", "SpcNum", "Specification no", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECH", "SpecCode", "Specification Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECH", "SpecName", "Specification Name", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECH", "ItemCode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECH", "ItemName", "Item Name", 150)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECH", "ItemType", "Item Type", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECH", "Category", "Product Category", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECH", "Version", "Specification Version", 10)

        objMain.objUtilities.AddDateField("@TNX_PH_QSPECH", "EffDate", "Effective Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PH_QSPECH", "ValidFrom", "Valid From", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PH_QSPECH", "ValidTo", "Valid To", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECH", "Status", "Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECH", "AppBy", "Approved By", 50)
        objMain.objUtilities.AddDateField("@TNX_PH_QSPECH", "AppDate", "Approved Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECH", "Remarks", "Remarks", 254)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECH", "CreatedBy", "Created By", 50)
        objMain.objUtilities.AddDateField("@TNX_PH_QSPECH", "CreatedOn", "Created On", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECH", "UpdatedBy", "Updated By", 50)
        objMain.objUtilities.AddDateField("@TNX_PH_QSPECH", "UpdatedOn", "Updated On", SAPbobsCOM.BoFldSubTypes.st_None)




        objMain.objUtilities.CreateTable("TNX_PH_QSPECL", "QC Specification Parameters", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECL", "TestCode", "Test Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECL", "TestName", "Test Name", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECL", "TestCat", "Test Category", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECL", "TestMethod", "Test Method", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECL", "Unit", "Unit", 20)

        objMain.objUtilities.AddFloatField("@TNX_PH_QSPECL", "MinValue", "Minimum Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PH_QSPECL", "MaxValue", "Maximum Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PH_QSPECL", "TargetVal", "Target Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECL", "TextLimit", "Text Limit", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECL", "ResultType", "Result Type", 20)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECL", "Mandatory", "Mandatory", 1, "Y")

        objMain.objUtilities.AddInteger("@TNX_PH_QSPECL", "SeqNo", "Sequence No", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECL", "Remarks", "Remarks", 254)



        objMain.objUtilities.CreateTable("TNX_PH_QSPECM", "QC Test Method Mapping", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECM", "TestCode", "Test Code", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECM", "MethodCode", "Method Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECM", "MethodName", "Method Name", 100)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECM", "SOPNo", "SOP Number", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECM", "Instrument", "Instrument", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECM", "CalibReq", "Calibration Required", 1, "N")

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECM", "Frequency", "Frequency", 30)

        objMain.objUtilities.AddAlphaField("@TNX_PH_QSPECM", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_QSPECM_ATT", "QC Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        '   '.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "TPA", "Target Path", 254)
        objMain.objUtilities.AddLinkField("@TNX_QSPECM_ATT", "TPH", "Target Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_QSPECM_ATT", "FNM", "File Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_QSPECM_ATT", "FTR", "Free Text", 254)
        objMain.objUtilities.AddDateField("@TNX_QSPECM_ATT ", "ATCD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)


    End Sub

    Private Sub CreateFormulaMaster()
        objMain.objUtilities.CreateTable("TNX_PH_FORMULA", "Formula Master", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA", "ForCode", "Formula Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA", "ForName", "Formula Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA", "ProCode", "Product Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA", "ProName", "Product Item Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA", "DosAge", "Dosage Form", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA", "Strength", "Strength", 50)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA", "BatchSize", "Batch Size", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA", "BatchUOM", "Batch UOM", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA", "FType", "Formula Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA", "FStatus", "Formula Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA", "Version", "Formula Version", 20)
        objMain.objUtilities.AddDateField("@TNX_PH_FORMULA", "EffDate", "Formula Effective Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA", "LUser", "Logged-in User", 100)
        objMain.objUtilities.AddDateField("@TNX_PH_FORMULA", "SysDate", "System Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("TNX_PH_FORMULA_D1", "Formula Composition", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D1", "ICode", "Ingredient Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D1", "IName", "Ingredient Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D1", "IType", "Ingredient Type", 30)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D1", "ReqQty", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D1", "UOM", "UOM", 20)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D1", "CompPrc", "% Composition", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D1", "OPrc", "Overages %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D1", "LossPrc", "Loss %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D1", "FQty", "Final Required Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D1", "Function", "Function", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D1", "Mandatory", "Mandatory", 10)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D1", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_PH_FORMULA_D2", "Process Instructions", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D2", "ProcessStage", "Process Stage", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D2", "Instruction", "Instruction", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D2", "MachineType", "Machine Type", 100)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D2", "Temperature", "Temperature", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D2", "Duration", "Duration", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D2", "DurationUOM", "Duration UOM", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D2", "CriticalParam", "Critical Parameter", 10)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D2", "QCCheckReq", "QC Check Required", 10)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D2", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_PH_FORMULA_D3", "Quality Parameters", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D3", "ParameterCode", "Parameter Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D3", "ParameterName", "Parameter Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D3", "StandardValue", "Standard Value", 100)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D3", "MinValue", "Min Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D3", "MaxValue", "Max Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D3", "TestMethod", "Test Method", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D3", "Mandatory", "Mandatory", 10)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D3", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_PH_FORMULA_D4", "Costing Summary", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D4", "MaterialCost", "Material Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D4", "ProcessCost", "Process Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D4", "TestingCost", "Testing Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D4", "PackagingCost", "Packaging Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D4", "TotalFormulaCost", "Total Formula Cost", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_PH_FORMULA_D4", "CostPerUnit", "Cost Per Unit", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D4", "Currency", "Currency", 10)

        objMain.objUtilities.CreateTable("TNX_PH_FORMULA_D5", "Approval Version", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D5", "CurrentVersion", "Current Version", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D5", "RevisionReason", "Revision Reason", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D5", "ApprovalStatus", "Approval Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D5", "ApprovedBy", "Approved By", 100)
        objMain.objUtilities.AddDateField("@TNX_PH_FORMULA_D5", "ApprovedDate", "Approved Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D5", "Locked", "Locked", 10)

        objMain.objUtilities.CreateTable("TNX_PH_FORMULA_D6", "Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D6", "DocumentType", "Document Type", 50)
        objMain.objUtilities.AddLinkField("@TNX_PH_FORMULA_D6", "FPath", "File Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D6", "FileName", "File Name", 80)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D6", "UploadedBy", "Uploaded By", 100)
        objMain.objUtilities.AddDateField("@TNX_PH_FORMULA_D6", "UploadedDate", "Uploaded Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PH_FORMULA_D6", "Remarks", "Remarks", 254)

    End Sub
    Private Sub CreateSOPManagement()

        ' =====================================================
        ' HEADER TABLE
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_SOPH", "SOP Management", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "SOPCode", "SOP Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "SOPTitle", "SOP Title", 200)
        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "SOPType", "SOP Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "Dept", "Department", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "Category", "SOP Category", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "Version", "SOP Version", 10)

        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "PrevDoc", "Previous DocEntry", 20)
        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "PrevVer", "Previous Version", 10)

        objMain.objUtilities.AddDateField("@TNX_SOPH", "EffDate", "Effective Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_SOPH", "RevDate", "Review Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddFloatField("@TNX_SOPH", "RevCycle", "Review Cycle", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "RiskLvl", "Risk Level", 20)
        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "TrainReq", "Training Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "CCReq", "Change Control Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "CCDoc", "Change Control DocEntry", 20)

        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "ApprStat", "Approval Status", 30)

        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "PrepBy", "Prepared By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "RevBy", "Reviewed By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "ApprBy", "Approved By", 50)

        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "AttEntry", "Attachment Entry", 20)

        objMain.objUtilities.AddAlphaField("@TNX_SOPH", "Remarks", "Remarks", 254)



        ' =====================================================
        ' SOP REVISION TABLE
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_SOP_REV", "SOP Revision", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_REV", "SecNo", "Section No", 20)
        objMain.objUtilities.AddAlphaField("@TNX_SOP_REV", "SecTitle", "Section Title", 150)
        objMain.objUtilities.AddAlphaField("@TNX_SOP_REV", "ChgType", "Change Type", 30)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_REV", "ChgDesc", "Change Description", 254)
        objMain.objUtilities.AddAlphaField("@TNX_SOP_REV", "Reason", "Reason", 254)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_REV", "ImpArea", "Impact Area", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SOP_REV", "ImpLvl", "Impact Level", 20)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_REV", "ChgBy", "Changed By", 50)

        objMain.objUtilities.AddDateField("@TNX_SOP_REV", "ChgDate", "Change Date", SAPbobsCOM.BoFldSubTypes.st_None)



        ' =====================================================
        ' SOP APPROVAL TABLE
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_SOP_APR", "SOP Approval", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddFloatField("@TNX_SOP_APR", "Level", "Approval Level", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_APR", "ApprRole", "Approver Role", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SOP_APR", "ApprUser", "Approver User", 50)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_APR", "Status", "Status", 20)

        objMain.objUtilities.AddDateField("@TNX_SOP_APR", "ActDate", "Action Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_APR", "ActTime", "Action Time", 20)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_APR", "Comments", "Approval Comments", 254)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_APR", "ESign", "E-Signature", 100)



        ' =====================================================
        ' SOP TRAINING TABLE
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_SOP_TRN", "SOP Training", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_TRN", "Dept", "Department", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SOP_TRN", "Positn", "Employee Position", 50)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_TRN", "EmpID", "Employee ID", 30)
        objMain.objUtilities.AddAlphaField("@TNX_SOP_TRN", "EmpName", "Employee Name", 100)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_TRN", "TrainReq", "Training Required", 1)

        objMain.objUtilities.AddDateField("@TNX_SOP_TRN", "DueDate", "Training Due Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_TRN", "TrainDoc", "Training DocEntry", 20)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_TRN", "Status", "Status", 30)

        objMain.objUtilities.AddDateField("@TNX_SOP_TRN", "CompDate", "Completed Date", SAPbobsCOM.BoFldSubTypes.st_None)



        ' =====================================================
        ' SOP DISTRIBUTION TABLE
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_SOP_DIST", "SOP Distribution", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_DIST", "Dept", "Department", 50)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_DIST", "UserCode", "User Code", 50)

        objMain.objUtilities.AddDateField("@TNX_SOP_DIST", "DistDate", "Distribution Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_DIST", "Ack", "Acknowledged", 1)

        objMain.objUtilities.AddDateField("@TNX_SOP_DIST", "AckDate", "Acknowledgement Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_DIST", "Remarks", "Remarks", 254)



        ' =====================================================
        ' SOP CATEGORY MASTER
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_SOP_CAT", "SOP Category Master", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_CAT", "Code", "Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SOP_CAT", "Name", "Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_SOP_CAT", "Dept", "Department", 50)

        objMain.objUtilities.AddFloatField("@TNX_SOP_CAT", "RevCycle", "Review Cycle", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_CAT", "TrainReq", "Training Required", 1)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_CAT", "ApprRoute", "Approval Route", 50)



        ' =====================================================
        ' SOP TEMPLATE MASTER
        ' =====================================================

        objMain.objUtilities.CreateTable("TNX_SOP_TMP", "SOP Template Master", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_TMP", "Code", "Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_SOP_TMP", "Name", "Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_SOP_TMP", "SOPType", "SOP Type", 30)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_TMP", "DefSect", "Default Sections", 254)

        objMain.objUtilities.AddAlphaField("@TNX_SOP_TMP", "AttEntry", "Attachment Entry", 20)

    End Sub


    Sub LicenceNew()
        objMain.objUtilities.CreateTable("TNX_LICENCE", "Licence Administration New one", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "Company", "Company Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "DocNum", "Document Number", 100)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "License", "License", 254)

        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "EIVC", "EInvoice", 1)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "CTTX", "Corporate Tax", 1)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "VRPT", "Vat Report", 1)


        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "DB", "Database Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "HKey", "Hardware Key", 100)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "Total", "Total Licenses", 10)
        objMain.objUtilities.AddDateField("@TNX_LICENCE", "SDate", "Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_LICENCE", "EDate", "Expiry Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "Addon", "Add-on Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE", "NOUSR", "No of Users", 254)

        objMain.objUtilities.CreateTable("TNX_LICENCE_C0", "License Administration Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE_C0", "Code", "Device Code", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE_C0", "USERC", "User Code", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE_C0", "Name", "User Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_LICENCE_C0", "Sts", "Status", 1)

        'objMain.objUtilities.CreateTable("SBO_ADDON", "License Addon", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        'objMain.objUtilities.AddAlphaField("@SBO_ADDON", "ADCODE", "Addon Code", 254)
        'objMain.objUtilities.AddAlphaField("@SBO_ADDON", "ADNAM", "Addon Name", 254)
        'objMain.objUtilities.AddAlphaField("@SBO_ADDON", "USER", "Used", 1)
        'objMain.objUtilities.AddAlphaField("@SBO_ADDON", "AVBLE", "Available", 254)

    End Sub

    Sub CreateDDTable()
        objMain.objUtilities.CreateTable("SBO_DD", "DDTable", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.CreateTable("SBO_DD1", "DD1Table", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@SBO_DD", "AppID", "Appln ID Field", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "AppStat", "Appln Status Field", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Table", "Table", 40)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Doc", "Document", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "statusN", "Status New", 20)

        objMain.objUtilities.AddAlphaField("@SBO_DD", "docnum", "Document Number", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "tempid", "Template ID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "status", "Status", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "userid", "UserID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "objtype", "objtype", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Isdraft", "Isdraft", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Remarks", "Remarks", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Appcnt", "Application", 100)
        objMain.objUtilities.AddDateField("@SBO_DD", "docdate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_DD", "Stage", "Stage", 100)

        objMain.objUtilities.AddAlphaField("@SBO_DD1", "EmpId", "Employee Id", 50)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "Stage", "Stage", 50)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "statusN", "Status New", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "TempId", "Template ID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "Userid", "UserID", 100)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "status", "Status", 100)
        objMain.objUtilities.AddDateField("@SBO_DD1", "cdate", "Create Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "ctime", "Create Time", 20)
        objMain.objUtilities.AddDateField("@SBO_DD1", "Udate", "Update Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "Utime", "Update Time", 20)
        objMain.objUtilities.AddAlphaField("@SBO_DD1", "Remarks", "Remarks", 200)

    End Sub
    Sub LkMster()
        objMain.objUtilities.CreateTable("TNX_LKMTR", "Link Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_LKMTR", "CTAX", "Corpurate Tax", 150)
        objMain.objUtilities.AddAlphaField("@TNX_LKMTR", "FVRT", "FTA VAT Report Tax", 150)
    End Sub
    Sub FtaVat1()
        objMain.objUtilities.CreateTable("TNX_FTAVAT", " FTA VAT Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_FTAVAT", "CMR", "From Period", 2)
        objMain.objUtilities.AddAlphaField("@TNX_FTAVAT", "SO", "To Period", 2)
        objMain.objUtilities.AddInteger("@TNX_FTAVAT", "GRPO", "Alert Before", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddDateField("@TNX_FTAVAT", "DCNF", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_FTAVAT", "CMN", "Status", 1)


    End Sub
    Sub CorpTax1()
        objMain.objUtilities.CreateTable("TNX_CORPTAX", "Corporate Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_CORPTAX", "CMR", "From Period", 2)
        objMain.objUtilities.AddAlphaField("@TNX_CORPTAX", "SO", "To Period", 2)
        objMain.objUtilities.AddInteger("@TNX_CORPTAX", "GRPO", "Alert Before", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddDateField("@TNX_CORPTAX", "DCNF", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CORPTAX", "CMN", "Status", 1)

    End Sub
    Sub FtaVat()
        objMain.objUtilities.CreateTable("FTAVAT", "User Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddDateField("@FTAVAT", "CMR", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@FTAVAT", "SO", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddInteger("@FTAVAT", "GRPO", "Alert Before", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddDateField("@FTAVAT", "DCNF", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@FTAVAT", "CMN", "Status", 1)
        objMain.objUtilities.AddAlphaField("@FTAVAT", "CODE", "Code", 20)
        objMain.objUtilities.AddAlphaField("@FTAVAT", "NAME", "Name", 100)

    End Sub
    Sub CorpTax()
        objMain.objUtilities.CreateTable("CORPTAX", "User Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddDateField("@CORPTAX", "CMR", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@CORPTAX", "SO", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddInteger("@CORPTAX", "GRPO", "Alert Before", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddDateField("@CORPTAX", "DCNF", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@CORPTAX", "CMN", "Status", 1)
        objMain.objUtilities.AddAlphaField("@CORPTAX", "CODE", "Code", 20)
        objMain.objUtilities.AddAlphaField("@CORPTAX", "NAME", "Name", 100)

    End Sub

    Sub InvoicePosting()
        objMain.objUtilities.CreateTable("TNX_IP", "Invoice Posting Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@TNX_IP", "Fdate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_IP", "Tdate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_IP", "Status", "Status", 50)
        objMain.objUtilities.AddAlphaField("@TNX_IP", "DocType", "DocumentType", 50)


        objMain.objUtilities.CreateTable("TNX_IP_C0", "Invoice Posting Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "Select", "Select", 5, "Y")
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "DNum", "Document Number", 100)
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "DocEntry", "Document Entry", 100)
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "CCode", "Customer Code", 100)
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "CName", "Customer Name", 100)
        objMain.objUtilities.AddFloatField("@TNX_IP_C0", "Tbdisc", "Total Before Discount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_IP_C0", "Disc", "Discount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_IP_C0", "Tax", "Tax", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@TNX_IP_C0", "Total", "Total", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "MSG", "ZATCA Message", 254)
        objMain.objUtilities.AddAlphaField("@TNX_IP_C0", "STATUS", "ZATCA Status", 254)


    End Sub
    Sub CreateAPPROVALSTAGES()
        objMain.objUtilities.CreateTable("SBO_AST", "Approval Stages", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "SName", "Stage Name", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "SDesc", "Stage Description", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "NAP", "No. of Approvals Required", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "NRJ", "No. of Rejections Required", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST", "NAME", "NAME", 40)


        objMain.objUtilities.CreateTable("SBO_AST_C0", "Approval Stages", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@SBO_AST_C0", "AUTH", "AUTH", 40)
        objMain.objUtilities.AddAlphaField("@SBO_AST_C0", "EmpId", "EmpId", 50)
        objMain.objUtilities.AddAlphaField("@SBO_AST_C0", "Dept", "Dept", 30)
        objMain.objUtilities.AddAlphaField("@SBO_AST_C0", "UKey", "UKey", 30)
    End Sub

    Sub YeildAnalysis()
        objMain.objUtilities.CreateTable("TNX_PYLD_H", "Yeild Analysis", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "DocNum", "DocNum", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "Series", "Series", 20)

        objMain.objUtilities.AddDateField("@TNX_PYLD_H", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "PORDR", "Production Order Entry", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "PRNO", "Production Order No", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "PRSTS", "Production Status", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "BMRD", "BMR DocEntry", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "BMRNO", "BMR Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "ItemCode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "ItemName", "Item Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "BatchNo", "Batch Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "UOM", "UOM", 20)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_H", "PLAQ", "Planned Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_H", "ACQTY", "Actual Receipt Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_H", "REJQ", "Rejected Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_H", "PROQY", "Process Loss Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_H", "SAMPQTY", "Sample Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_H", "NETYQ", "Net Yield Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_H", "YIEL", "Yield %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_H", "STANY", "Standard Yield %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_H", "MINIY", "Minimum Yield %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_H", "MAXY", "Maximum Yield %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_H", "VARIN", "Variance %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "VARS", "Variance Status", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "DEVR", "Deviation Required", 10)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "DEVN", "Deviation Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "QASTS", "QA Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "APPSTS", "Approval Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_H", "DOCSTS", "Document Status", 30)
        ' objMain.objUtilities.AddMemoFie("@TNX_PYLD_H", "Remarks", "Remarks")

        objMain.objUtilities.CreateTable("TNX_PYLD_MAT", "Material Yield / Consumption", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PYLD_MAT", "CMPCD", "Raw material / component code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_MAT", "COMNM", "Component name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_MAT", "ITMTY", "RM / PM / Intermediate", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_MAT", "BATNO", "Component batch number", 50)

        objMain.objUtilities.AddFloatField("@TNX_PYLD_MAT", "BASEQTY", "BOM base quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_MAT", "PLANQT", "Planned component quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_MAT", "ISSQT", "Actual issued quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_MAT", "RETQT", "Returned quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_MAT", "CONQT", "Net consumed quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_MAT", "ACLQT", "Actual loss quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_MAT", "VAQT", "Issued vs consumed variance", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddFloatField("@TNX_PYLD_MAT", "STLOSP", "Standard loss percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_MAT", "VARP", "Variance percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_MAT", "ALLTO", "Allowed tolerance percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)

        objMain.objUtilities.AddAlphaField("@TNX_PYLD_MAT", "Status", "Within Limit / Outside Limit", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_MAT", "RECOD", "Reason code", 50)

        'Output / Receipt Details
        objMain.objUtilities.CreateTable("TNX_PYLD_OUT", "Output / Receipt Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)


        objMain.objUtilities.AddAlphaField("@TNX_PYLD_OUT", "REFRM", "Receipt from Production DocEntry", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_OUT", "REPN", "Receipt from Production number", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_OUT", "RELI", "Receipt line number", 30)

        objMain.objUtilities.AddDateField("@TNX_PYLD_OUT", "REPDT", "Receipt posting date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_PYLD_OUT", "ItemCode", "Finished good item", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_OUT", "BatchNo", "Finished good batch", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_OUT", "WhsCode", "Warehouse", 20)

        objMain.objUtilities.AddFloatField("@TNX_PYLD_OUT", "REQTY", "Received quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_OUT", "QCQty", "QC sample quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_OUT", "REQTY", "Rejected quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_OUT", "ACCPT", "Accepted quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)

        objMain.objUtilities.AddAlphaField("@TNX_PYLD_OUT", "UOM", "UOM", 20)


        ' Child Table 3 - Variance / Loss Details
        objMain.objUtilities.CreateTable("TNX_PYLD_VAR", "Variance / Loss Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_VAR", "VAR_TYPE", "Variance Type", 50) ' Material / Output / Rejection / Process Loss
        objMain.objUtilities.AddFloatField("@TNX_PYLD_VAR", "EXPQTY", "Expected quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_VAR", "ACTQTY", "Actual quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_VAR", "VARQTY", "Difference", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_VAR", "VAPCT", "Difference percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_PYLD_VAR", "TOPCT", "Allowed tolerance", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_VAR", "IMPACT", "Impact", 20) ' Low / Medium / High
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_VAR", "REAE", "Reason code", 50)
        objMain.objUtilities.AddAlphaMemoField("@TNX_PYLD_VAR", "REESC", "Reason description", 5000)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_VAR", "ACED", "Action Required", 1) ' Yes / No
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_VAR", "CARED", "CAPA Required", 1) ' Yes / No
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_VAR", "DIRED", "Deviation Required", 1) ' Yes / No

        ' Child Table 4 - Approval Details
        objMain.objUtilities.CreateTable("TNX_PYLD_APR", "Approval Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_APR", "STAGE", "Stage", 50) ' Production / QA / Plant Head
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_APR", "APER", "Approver", 50) ' User code
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_APR", "APPS", "Approval Status", 20) ' Pending / Approved / Rejected
        objMain.objUtilities.AddDateField("@TNX_PYLD_APR", "ATPE", "Approval date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PYLD_APR", "AIME", "Approval time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddAlphaMemoField("@TNX_PYLD_APR", "RMKS", "Approval remarks", 5000)

        ' Child Table 5 - Attachments
        objMain.objUtilities.CreateTable("TNX_PYLD_ATT", "Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_ATT", "FILEE", "File name", 150)
        objMain.objUtilities.AddLinkField("@TNX_PYLD_ATT", "FITH", "Attachment path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_ATT", "FIYPE", "File type", 30) ' PDF / Image / Excel
        objMain.objUtilities.AddAlphaField("@TNX_PYLD_ATT", "ATTBY", "Attached by", 50)
        objMain.objUtilities.AddDateField("@TNX_PYLD_ATT", "AE", "Attached date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_PYLD_ATT", "REMS", "Remarks", 5000)




    End Sub
    Sub CreateAPPROVALTemplates()
        objMain.objUtilities.CreateTable("SBO_APPHDR", "Temp", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "T1", "Stage Name", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "T2", "Stage Description", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "NAP", "No. of Approvals Required", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "C1", "Active", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "PCR", "Price Change Request", 254)



        objMain.objUtilities.CreateTable("SBO_APPREQ", "Originator", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "M1_1", "Name", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "EmpId", "EmpId", 50)
        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "M1_2", "Department", 30)

        objMain.objUtilities.CreateTable("SBO_APPAUT", "Stages", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@SBO_APPAUT", "M3_1", "M3_1", 40)
        objMain.objUtilities.AddAlphaField("@SBO_APPAUT", "M3_2", "M3_2", 50)


        objMain.objUtilities.AddAlphaField("@SBO_APPAUT", "NAMES", "Names", 30)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Name", "Name", 50)

        objMain.objUtilities.CreateTable("SBO_APPHDR", "Conditions", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Conds", "Always", 50)


        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "Name", "EmpId", 50)

        objMain.objUtilities.AddAlphaField("@SBO_APPREQ", "Dept", "Department", 50)

        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Descrip", "Desciption", 100)

        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Acive", "Active", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPHDR", "Active", "Active", 1)

        'objMain.objUtilities.CreateTable("SBO_APPHDR", "Conditions", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        'objMain.objUtilities.AddButton("@SBO_APPHDR", "M3_1", "M3_1", 40)
        'objMain.objUtilities.AddButton(objform.UniqueID, "Btn_Reset", objform.Items.Item("71").Top + 230,_  objform.Items.Item("73").Left, 80, objform.Items.Item("71").Height + 10, "71", "&Reset", objform.Items.Item("71").FromPane, objform.Items.Item("71").ToPane)

        objMain.objUtilities.CreateTable("SBO_APPDOC", "Documents", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "RcReq", "Recruitment Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "RmSm", "Resume Submission", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "OL", "Offer Letter", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "VRQ", "Advertisement Request", 1) 'C'
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Price", "Price Change Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "SuppReb", "SupplierRebate", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "IMP", "ItemMasterUpload", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "ITEML", "ItemListing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "ITEML", "ItemListing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "BP", "BusinessPartner", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "CPA", "Corporate Tax", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "VPA", "VAT Report", 1)


        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "EmpMstr", "Employee Master Data", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "EmpSal", "Employee Salary Setup", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "ConNew", "Employee Contract Renewal", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "TS", "TimeSheet", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LvRj", "Leave Rejoin Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LvApp", "Leave Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "CompOff", "Compensatory off", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "AttSum", "Attendance Summary", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Shift", "Shift Schedule", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "AReg", "Attendance Regularisation", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LER", "Leave Encashment Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LPC", "Leave Period End Closing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LnApp", "Loans and Advances Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LnPr", "Loan Pre Closure Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PD", "Payments and Deductions", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PrPre", "Payroll Pre-Process", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Claim", "Claim Application", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Review", "Employee Review / Appraisal", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PerEval", "Performance Evaluation", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Encash", "Leave Encashment", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "EOS", "End of Service Process", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Resign", "Employee Resignation Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Clear", "Exit Clearance", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Air", "Airticket Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Doc", "Document Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PTC", "Petty Cash Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PTV", "Petty Cash Voucher", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "MTR", "Material Request", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "Trans", "Branch Transfer", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "GRPMTNG", "Group Meeting", 1)


        'Akhila
        'Pharna new fields creation 


        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "FMS", "Formula Master", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "EMG", "Experiment Management", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PBT", "Pilot Batch", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "FCTG", "Formula Costing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "SAM", "Sample Registration", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "SLCN", "Sample Collection", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LBC", "Lab Testing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "STBT", "Stability Testing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "COA", "COA Management", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "APPL", "Approval Matrix Matter", 1)


        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PROR", "Production Order", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "STTS", "Stability Testing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LICL", "Line Clearance", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "DISP", "Dispensing", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "YIELD", "Yield Analysis", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "BPI", "BPI", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "SOP", "SOP Management", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "TRMG", "Training Management", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "VMG", "Validation Management", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "CHCO", "Change Control", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "ICMG", "Incident Management", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "CAPA", "CAPA Linkage", 1)



        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "EQMS", "Equipment Master", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PDOR", "Production Stage Master", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "CLMD", "Cleaning Method Master", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "LCLR", "Line Clearance Checklist Master", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "INPR", "In-process QC Checklist Master", 1)

        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "YIEL", "Yield Tolerance Master", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "CONT", "Country Regulatory Configuration", 1)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "BPR", "BPR", 1)


        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "RD", "TNX Pharma R&D", 30)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PDMS", "TNX Pharma Production Masters", 30)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PHC", "TNX Pharma Compliance", 30)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PQA", "TNX Pharma QA/QC", 30)
        objMain.objUtilities.AddAlphaField("@SBO_APPDOC", "PPD", "TNX Pharma Production", 30)


    End Sub
    Sub CorporateTaxConfiguration()
        objMain.objUtilities.CreateTable("TNX_CTAXCNF", "Corporate Tax Configuration", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddFloatField("@TNX_CTAXCNF", "MnProfit", "Minimum Profit", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_CTAXCNF", "MxProfit", "Maximum Profit", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_CTAXCNF", "TaxPrc", "Tax Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCNF", "LAccount", "Liability Account", 100)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCNF", "EAccount", "Expenditure Account", 100)
        objMain.objUtilities.AddAlphaField("@TNX_CTAXCNF", "FINA", "Finanical Year", 10)
    End Sub

    'Sub CorporateTaxCalculation()
    '    objMain.objUtilities.CreateTable("TNX_PH_STAB", "Corporate Tax Calculation Header", SAPbobsCOM.BoUTBTableType.bott_Document)
    '    objMain.objUtilities.AddDateField("@TNX_CTAXCALCU", "FDate", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
    '    objMain.objUtilities.AddDateField("@TNX_CTAXCALCU", "TDate", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
    '    objMain.objUtilities.AddAlphaField("@TNX_CTAXCALCU", "Branch", "Branch", 80)
    '    objMain.objUtilities.AddFloatField("@TNX_CTAXCALCU", "PPeriod", "Profit Period", SAPbobsCOM.BoFldSubTypes.st_Sum)
    '    objMain.objUtilities.AddFloatField("@TNX_CTAXCALCU", "CTax", "Corporate Tax %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
    '    objMain.objUtilities.AddFloatField("@TNX_CTAXCALCU", "CTaxVal", "Corporate Tax Value", SAPbobsCOM.BoFldSubTypes.st_Sum)
    '    objMain.objUtilities.AddDateField("@TNX_CTAXCALCU", "JEPOSTD", "JE Posting Date", SAPbobsCOM.BoFldSubTypes.st_None)
    '    objMain.objUtilities.AddAlphaField("@TNX_CTAXCALCU", "JENo", "JE No", 30)
    '    objMain.objUtilities.AddAlphaField("@TNX_CTAXCALCU", "Status", "Status", 15)
    '    objMain.objUtilities.AddAlphaField("@TNX_CTAXCALCU", "RMC", "Remarks", 254)
    '    objMain.objUtilities.AddAlphaField("@TNX_CTAXCALCU", "DST", "Document Status", 1)
    '    objMain.objUtilities.AddAlphaField("@TNX_CTAXCALCU", "CURR", "currency", 30)
    '    objMain.objUtilities.AddAlphaField("@TNX_CTAXCALCU", "YEAR", "YEAR", 30)
    '    objMain.objUtilities.AddAlphaField("@TNX_CTAXCALCU", "AIPD", "Application ID", 50)
    '    objMain.objUtilities.AddDateField("@TNX_CTAXCALCU", "DAT", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)

    '    objMain.objUtilities.CreateTable("TNX_CTAXCALCU_C2", "Corporate Tax Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
    '    '.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "TPA", "Target Path", 254)
    '    objMain.objUtilities.AddLinkField("@TNX_CTAXCALCU_C2", "TPH", "Target Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
    '    objMain.objUtilities.AddAlphaField("@TNX_CTAXCALCU_C2", "FNM", "File Name", 254)
    '    objMain.objUtilities.AddAlphaField("@TNX_CTAXCALCU_C2", "FTR", "Free Text", 254)
    '    objMain.objUtilities.AddDateField("@TNX_CTAXCALCU_C2 ", "ATCD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)


    'End Sub

    '====================================================================
    ' 1. Specifications Master
    ' UDO Code     : TNX_QC_SPEC_UDO
    ' Header Table : @TNX_QC_SPEC_H
    ' Line Table   : @TNX_QC_SPEC_L
    '====================================================================


    'Formula Costing
    Sub CreateFormulaVersionControl()

        objMain.objUtilities.CreateTable("TNX_FRM_VER_H", "Formula Version Control", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "FRM_CODE", "Formula Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "FRM_NAME", "Formula Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "ITEM_CODE", "Product Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "ITEM_NAME", "Product Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "OLD_VER", "Old Version", 20)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "NEW_VER", "New Version", 20)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "REV_TYPE", "Revision Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "CHG_CAT", "Change Category", 30)
        objMain.objUtilities.AddDateField("@TNX_FRM_VER_H", "REV_DATE", "Revision Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "REQ_BY", "Requested By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "DEPT", "Department", 50)
        objMain.objUtilities.AddAlphaMemoField("@TNX_FRM_VER_H", "CHG_RSN", "Change Reason", 5000)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "IMPCT_REQ", "Impact Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "STATUS", "Status", 20)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "APPR_TEMP", "Approval Template", 50)
        objMain.objUtilities.AddDateField("@TNX_FRM_VER_H", "EFF_FROM", "Effective From", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "PREV_LOCK", "Previous Locked", 1)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_H", "RISK_LEVEL", "Risk Level", 20)
        objMain.objUtilities.AddAlphaMemoField("@TNX_FRM_VER_H", "REMARKS", "Remarks", 5000)

        objMain.objUtilities.CreateTable("TNX_FRM_VER_D1", "Ingredient Change Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D1", "ITEM_CODE", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D1", "ITEM_NAME", "Item Name", 100)
        objMain.objUtilities.AddFloatField("@TNX_FRM_VER_D1", "OLD_QTY", "Old Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_FRM_VER_D1", "NEW_QTY", "New Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_FRM_VER_D1", "DIFF_QTY", "Difference Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D1", "UOM", "UOM", 20)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D1", "CHG_TYPE", "Change Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D1", "ROLE", "Role", 50)
        objMain.objUtilities.AddAlphaMemoField("@TNX_FRM_VER_D1", "REMARKS", "Remarks", 5000)

        objMain.objUtilities.CreateTable("TNX_FRM_VER_D2", "Process Change Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddFloatField("@TNX_FRM_VER_D2", "STEP_NO", "Step No", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D2", "STAGE", "Process Stage", 50)
        objMain.objUtilities.AddAlphaMemoField("@TNX_FRM_VER_D2", "OLD_INST", "Old Instruction", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_FRM_VER_D2", "NEW_INST", "New Instruction", 5000)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D2", "CHG_TYPE", "Change Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D2", "MACHINE", "Machine", 50)
        objMain.objUtilities.AddFloatField("@TNX_FRM_VER_D2", "DURATION", "Duration", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaMemoField("@TNX_FRM_VER_D2", "REMARKS", "Remarks", 5000)

        objMain.objUtilities.CreateTable("TNX_FRM_VER_D3", "Quality Regulatory Impact", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D3", "IMP_AREA", "Impact Area", 50)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D3", "REQUIRED", "Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D3", "RISK_LVL", "Risk Level", 20)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D3", "RV_BY", "Review By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D3", "RV_STS", "Review Status", 20)
        objMain.objUtilities.AddDateField("@TNX_FRM_VER_D3", "RV_DATE", "Review Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_FRM_VER_D3", "COMMENTS", "Comments", 5000)

        objMain.objUtilities.CreateTable("TNX_FRM_VER_D4", "Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D4", "DOC_TYPE", "Document Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D4", "FILE_NAME", "File Name", 150)
        objMain.objUtilities.AddLinkField("@TNX_FRM_VER_D4", "FILE_PATH", "File Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D4", "UPLD_BY", "Uploaded By", 50)
        objMain.objUtilities.AddDateField("@TNX_FRM_VER_D4", "UPLD_DT", "Upload Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_FRM_VER_D4", "REMARKS", "Remarks", 5000)

        objMain.objUtilities.CreateTable("TNX_FRM_VER_D5", "Approval History", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D5", "STAGE", "Stage", 50)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D5", "APPROVER", "Approver", 50)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D5", "ROLE", "Role", 50)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_D5", "STATUS", "Status", 20)
        objMain.objUtilities.AddDateField("@TNX_FRM_VER_D5", "DSC_DT", "Decision Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_FRM_VER_D5", "REMARKS", "Remarks", 5000)

        objMain.objUtilities.CreateTable("TNX_FRM_VER_AUD", "Formula Version Audit", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        ' objMain.objUtilities.AddFloatField("@TNX_FRM_VER_AUD", "DOCENTRY", "DocEntry", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_FRM_VER_AUD", "ACT_DT", "Action Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_AUD", "USER", "User", 50)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_AUD", "ACTION", "Action", 50)
        objMain.objUtilities.AddAlphaField("@TNX_FRM_VER_AUD", "FIELD", "Field", 100)
        objMain.objUtilities.AddAlphaMemoField("@TNX_FRM_VER_AUD", "OLD_VAL", "Old Value", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_FRM_VER_AUD", "NEW_VAL", "New Value", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_FRM_VER_AUD", "REMARKS", "Remarks", 5000)

        objMain.objUtilities.CreateTable("TNX_ATTACH_C3", "VAT Report Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        '.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "TPA", "Target Path", 254)
        objMain.objUtilities.AddLinkField("@TNX_ATTACH_C3", "TPA", "Target Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "FN", "File Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "FTT", "Free Text", 254)
        objMain.objUtilities.AddDateField("@TNX_ATTACH_C3", "ATD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)
        'objMain.objUtilities.AddFloatField("@TNX_ATTACH_C3", "VATA", "VAT Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Quantity)
    End Sub


    Sub CreateExperimentManagement()

        '=================================================================
        ' Header Table : @TNX_EXP_HDR
        '=================================================================
        objMain.objUtilities.CreateTable("TNX_EXP_HDR", "Experiment Management Header", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "ExpNo", "Experiment No", 50)
        objMain.objUtilities.AddDateField("@TNX_EXP_HDR", "ExpDate", "Experiment Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "FormulaCode", "Formula Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "FormulaVer", "Formula Version", 20)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "ProductCode", "Product Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "ProductName", "Product Name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "DosageForm", "Dosage Form", 50)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "ExpType", "Experiment Type", 50)
        objMain.objUtilities.AddAlphaMemoField("@TNX_EXP_HDR", "Objective", "Objective", 5000)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "Scientist", "Scientist", 100)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "Department", "Department", 50)
        objMain.objUtilities.AddFloatField("@TNX_EXP_HDR", "BatchSize", "Batch Size", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "UOM", "UOM", 20)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "Priority", "Priority", 20)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "Result", "Result", 30)
        objMain.objUtilities.AddFloatField("@TNX_EXP_HDR", "YieldPer", "Yield %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_EXP_HDR", "LossPer", "Loss %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddFloatField("@TNX_EXP_HDR", "CostImpact", "Cost Impact", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "ChangeReq", "Change Req", 1)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "NewVersionReq", "New Version Req", 1)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "PilotReq", "Pilot Req", 1)
        objMain.objUtilities.AddAlphaMemoField("@TNX_EXP_HDR", "FinalConclusion", "Final Conclusion", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_EXP_HDR", "Remarks", "Remarks", 5000)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "CreatedBy", "Created By", 100)
        objMain.objUtilities.AddDateField("@TNX_EXP_HDR", "CreatedDate", "Created Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_HDR", "ApprovedBy", "Approved By", 100)
        objMain.objUtilities.AddDateField("@TNX_EXP_HDR", "ApprovedDate", "Approved Date", SAPbobsCOM.BoFldSubTypes.st_None)

        '=================================================================
        ' Child Table 1 : @TNX_EXP_ING
        '=================================================================
        objMain.objUtilities.CreateTable("TNX_EXP_ING", "Experiment Ingredients", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_EXP_ING", "ItemCode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_ING", "ItemName", "Item Name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_ING", "IngredientType", "Ingredient Type", 50)
        objMain.objUtilities.AddFloatField("@TNX_EXP_ING", "StdQty", "Standard Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_EXP_ING", "TrialQty", "Trial Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_ING", "UOM", "UOM", 20)
        objMain.objUtilities.AddFloatField("@TNX_EXP_ING", "VarQty", "Variance Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_EXP_ING", "VarPer", "Variance %", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_ING", "Function", "Function", 100)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_ING", "BatchStage", "Batch Stage", 100)
        objMain.objUtilities.AddAlphaMemoField("@TNX_EXP_ING", "Remarks", "Remarks", 5000)

        '=================================================================
        ' Child Table 2 : @TNX_EXP_PROC
        '=================================================================
        objMain.objUtilities.CreateTable("TNX_EXP_PROC", "Experiment Process Steps", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddFloatField("@TNX_EXP_PROC", "StepNo", "Step No", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_PROC", "ProcessStage", "Process Stage", 100)
        objMain.objUtilities.AddAlphaMemoField("@TNX_EXP_PROC", "Instruction", "Instruction", 5000)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_PROC", "Equipment", "Equipment", 100)
        objMain.objUtilities.AddFloatField("@TNX_EXP_PROC", "Temperature", "Temperature", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_EXP_PROC", "Humidity", "Humidity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_EXP_PROC", "Duration", "Duration", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_PROC", "Operator", "Operator", 100)
        objMain.objUtilities.AddAlphaMemoField("@TNX_EXP_PROC", "Observation", "Observation", 5000)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_PROC", "Status", "Status", 30)

        '=================================================================
        ' Child Table 3 : @TNX_EXP_TEST
        '=================================================================
        objMain.objUtilities.CreateTable("TNX_EXP_TEST", "Experiment Test Parameters", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_EXP_TEST", "ParamCode", "Parameter Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_TEST", "ParamName", "Parameter Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_TEST", "StdValue", "Standard Value", 100)
        objMain.objUtilities.AddFloatField("@TNX_EXP_TEST", "MinValue", "Min Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_EXP_TEST", "MaxValue", "Max Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_TEST", "ActualValue", "Actual Value", 100)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_TEST", "UOM", "UOM", 20)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_TEST", "Result", "Result", 20)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_TEST", "TestedBy", "Tested By", 100)
        objMain.objUtilities.AddDateField("@TNX_EXP_TEST", "TestDate", "Test Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_EXP_TEST", "Remarks", "Remarks", 5000)

        '=================================================================
        ' Child Table 4 : @TNX_EXP_OBS
        '=================================================================
        objMain.objUtilities.CreateTable("TNX_EXP_OBS", "Experiment Observations", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_EXP_OBS", "ObsType", "Observation Type", 50)
        objMain.objUtilities.AddAlphaMemoField("@TNX_EXP_OBS", "Observation", "Observation", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_EXP_OBS", "Deviation", "Deviation", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_EXP_OBS", "CorrectiveAction", "Corrective Action", 5000)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_OBS", "ObservedBy", "Observed By", 100)
        objMain.objUtilities.AddDateField("@TNX_EXP_OBS", "ObsDate", "Observation Date", SAPbobsCOM.BoFldSubTypes.st_None)

        '=================================================================
        ' Child Table 5 : @TNX_EXP_ATTACH
        '=================================================================
        objMain.objUtilities.CreateTable("TNX_EXP_ATTACH1", "Experiment Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_EXP_ATTACH1", "DocType", "Document Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_ATTACH1", "FileName", "File Name", 200)
        objMain.objUtilities.AddLinkField("@TNX_EXP_ATTACH1", "FilePath", "File Path", 500, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_EXP_ATTACH1", "UploadedBy", "Uploaded By", 100)
        objMain.objUtilities.AddDateField("@TNX_EXP_ATTACH1", "UploadDate", "Upload Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_EXP_ATTACH1", "Remarks", "Remarks", 5000)

    End Sub
    'Submission Tracker
    Sub CreateSubmissionTracker()

        '==========================================================
        ' Header Table : @TNX_REG_SUBH
        '==========================================================
        objMain.objUtilities.CreateTable("TNX_REG_SUBH", "Reg Submission Header", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "SubNo", "Submission No", 50)
        objMain.objUtilities.AddFloatField("@TNX_REG_SUBH", "PrRegDoc", "Product Registration Doc", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "PrRegNo", "Product Registration No", 50)
        objMain.objUtilities.AddFloatField("@TNX_REG_SUBH", "DsrDoc", "Dossier Doc", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "DsrNo", "Dossier No", 50)
        objMain.objUtilities.AddFloatField("@TNX_REG_SUBH", "ArtwDoc", "Artwork Doc", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "ArtwNo", "Artwork No", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "ItemCode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "ItemName", "Item Name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "GenName", "Generic Name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "DosaForm", "Dosage Form", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "Strength", "Strength", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "Country", "Country", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "AuthCode", "Authority Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "AuthName", "Authority Name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "SubType", "Submission Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "SubMode", "Submission Mode", 50)
        objMain.objUtilities.AddDateField("@TNX_REG_SUBH", "SubDate", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_REG_SUBH", "TgtAprDt", "Target Approval Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "RefNo", "Reference No", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "AckNo", "Acknowledgement No", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "ApprNo", "Approval No", 100)
        objMain.objUtilities.AddDateField("@TNX_REG_SUBH", "ApprDate", "Approval Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_REG_SUBH", "ExpDate", "Expiry Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "SubStatus", "Submission Status", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "RiskLvl", "Risk Level", 30)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "Priority", "Priority", 30)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "AssignTo", "Assigned To", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "LocAgent", "Local Agent", 100)
        objMain.objUtilities.AddFloatField("@TNX_REG_SUBH", "FeeAmt", "Fee Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "Currency", "Currency", 10)
        objMain.objUtilities.AddFloatField("@TNX_REG_SUBH", "AtcEntry", "Attachment Entry", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_REG_SUBH", "Remarks", "Remarks", 5000)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "CreateBy", "Created By", 50)
        objMain.objUtilities.AddDateField("@TNX_REG_SUBH", "CreateDt", "Created Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBH", "ApprBy", "Approved By", 50)
        objMain.objUtilities.AddDateField("@TNX_REG_SUBH", "ApprvDt", "Approved Date", SAPbobsCOM.BoFldSubTypes.st_None)

        '==========================================================
        ' Child Table 1 : @TNX_REG_SUBL
        '==========================================================
        objMain.objUtilities.CreateTable("TNX_REG_SUBL", "Submitted Document Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBL", "DocCtgy", "Document Category", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBL", "DocType", "Document Type", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBL", "DocName", "Document Name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBL", "DocNo", "Document No", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBL", "DocVer", "Document Version", 20)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBL", "SrcDocTy", "Source Doc Type", 50)
        objMain.objUtilities.AddFloatField("@TNX_REG_SUBL", "SrcDocEn", "Source Doc Entry", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@TNX_REG_SUBL", "SrcLineId", "Source Line Id", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBL", "Mandatry", "Mandatory", 1)
        objMain.objUtilities.AddFloatField("@TNX_REG_SUBL", "AtchEntr", "Attachment Entry", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBL", "Submitd", "Submitted", 1)
        objMain.objUtilities.AddDateField("@TNX_REG_SUBL", "SubmDate", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBL", "AcptAuth", "Accepted By Authority", 1)
        objMain.objUtilities.AddAlphaMemoField("@TNX_REG_SUBL", "AuthRmk", "Authority Remarks", 5000)
        objMain.objUtilities.AddAlphaField("@TNX_REG_SUBL", "Status", "Status", 50)
        objMain.objUtilities.AddAlphaMemoField("@TNX_REG_SUBL", "Remarks", "Remarks", 5000)

        '==========================================================
        ' Child Table 2 : @TNX_REG_QRY
        '==========================================================
        objMain.objUtilities.CreateTable("TNX_REG_QRY", "Authority Query", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_REG_QRY", "QueryNo", "Query No", 50)
        objMain.objUtilities.AddDateField("@TNX_REG_QRY", "QueryDt", "Query Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_QRY", "QueryTyp", "Query Type", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REG_QRY", "QuerySev", "Query Severity", 30)
        objMain.objUtilities.AddAlphaMemoField("@TNX_REG_QRY", "QueryDsc", "Query Description", 5000)
        objMain.objUtilities.AddDateField("@TNX_REG_QRY", "RespDueDt", "Response Due Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_QRY", "RespOwnr", "Response Owner", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_QRY", "RespStat", "Response Status", 50)
        objMain.objUtilities.AddDateField("@TNX_REG_QRY", "RespDate", "Response Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@TNX_REG_QRY", "RespDoc", "Response Document", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_QRY", "IntRevRq", "Internal Review Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_REG_QRY", "QARevBy", "QA Review By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_QRY", "RegRevBy", "Reg Review By", 50)
        objMain.objUtilities.AddDateField("@TNX_REG_QRY", "CloseDt", "Closed Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_REG_QRY", "Remarks", "Remarks", 5000)

        '==========================================================
        ' Child Table 3 : @TNX_REG_STAT
        '==========================================================
        objMain.objUtilities.CreateTable("TNX_REG_STAT1", "Submission Status History", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddDateField("@TNX_REG_STAT1", "StatDate", "Status Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT1", "FromStat", "From Status", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT1", "ToStatus", "To Status", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_STAT1", "ChangBy", "Changed By", 50)
        objMain.objUtilities.AddAlphaMemoField("@TNX_REG_STAT1", "Reason", "Reason", 5000)





        '==========================================================
        ' Child Table 4 : @TNX_REG_APRV
        '==========================================================
        objMain.objUtilities.CreateTable("TNX_REG_APRV", "Submission Approval History", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_REG_APRV", "ApprStg", "Approval Stage", 100)
        objMain.objUtilities.AddAlphaField("@TNX_REG_APRV", "ApprUser", "Approver User", 50)
        objMain.objUtilities.AddAlphaField("@TNX_REG_APRV", "ApprStat", "Approval Status", 50)
        objMain.objUtilities.AddDateField("@TNX_REG_APRV", "ApprDate", "Approval Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_REG_APRV", "ApprRmk", "Approval Remarks", 5000)
        objMain.objUtilities.AddAlphaField("@TNX_REG_APRV", "EscalTo", "Escalated To", 50)
        objMain.objUtilities.AddDateField("@TNX_REG_APRV", "EscalDt", "Escalation Date", SAPbobsCOM.BoFldSubTypes.st_None)

        '==========================================================
        ' Attachment Table : @TNX_ATTACHMENT_C0
        '==========================================================
        objMain.objUtilities.CreateTable("TNX_ATTACHMENT_C3", "Submission Attachment", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddLinkField("@TNX_ATTACHMENT_C3", "TPA", "Target Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_ATTACHMENT_C3", "FN", "File Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_ATTACHMENT_C3", "FTT", "Free Text", 254)
        objMain.objUtilities.AddDateField("@TNX_ATTACHMENT_C3", "ATD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)

    End Sub


    '=========================================================
    ' 1. SOP Category Master
    ' UDO Code   : UDO_TNX_SOPCAT
    ' Table Name : @TNX_SOPCAT
    ' Type       : Master Data
    '=========================================================
    '=========================================================
    ' 1. SOP Category Master
    ' UDO Code   : UDO_TNX_SOPCAT
    ' Table Name : @TNX_SOPCAT
    '=========================================================
    Public Sub CreateSOPCategoryMaster()

        objMain.objUtilities.CreateTable("TNX_SOPCAT", "SOP Category Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_SOPCAT", "DeptCode", "Department Code", 20)
        objMain.objUtilities.AddAlphaField("@TNX_SOPCAT", "GMPReq", "GMP Required", 1)
        objMain.objUtilities.AddFloatField("@TNX_SOPCAT", "ReviewFr", "Review Frequency", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_SOPCAT", "TrainReq", "Training Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_SOPCAT", "ApprRoute", "Approval Route", 30)
        objMain.objUtilities.AddAlphaField("@TNX_SOPCAT", "Active", "Active", 1)
        objMain.objUtilities.AddAlphaField("@TNX_SOPCAT", "Remarks", "Remarks", 254)

    End Sub

    '=========================================================
    ' 2. Department Master
    ' UDO Code   : UDO_TNX_DEPT
    ' Table Name : @TNX_DEPT
    '=========================================================
    Public Sub CreateDepartmentMaster()

        objMain.objUtilities.CreateTable("TNX_DEPT", "Department Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_DEPT", "DeptHead", "Department Head", 50)
        objMain.objUtilities.AddAlphaField("@TNX_DEPT", "QAReview", "QA Reviewer", 50)
        objMain.objUtilities.AddAlphaField("@TNX_DEPT", "CompOwnr", "Compliance Owner", 50)
        objMain.objUtilities.AddAlphaField("@TNX_DEPT", "CostCent", "Cost Center", 20)
        objMain.objUtilities.AddAlphaField("@TNX_DEPT", "Active", "Active", 1)
        objMain.objUtilities.AddAlphaField("@TNX_DEPT", "Remarks", "Remarks", 254)

    End Sub

    '=========================================================
    ' 3. Training Type Master
    ' UDO Code   : UDO_TNX_TRNTYP
    ' Table Name : @TNX_TRNTYP
    '=========================================================
    Public Sub CreateTrainingTypeMaster()

        objMain.objUtilities.CreateTable("TNX_TRNTYP", "Training Type Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_TRNTYP", "AssessReq", "Assessment Required", 1)
        objMain.objUtilities.AddFloatField("@TNX_TRNTYP", "PassScore", "Pass Score", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_TRNTYP", "CertReq", "Certificate Required", 1)
        objMain.objUtilities.AddFloatField("@TNX_TRNTYP", "RetrainF", "Retraining Frequency", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_TRNTYP", "TrainerR", "Trainer Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_TRNTYP", "AttachReq", "Attachment Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_TRNTYP", "Active", "Active", 1)

    End Sub

    '=========================================================
    ' 4. Validation Type Master
    ' UDO Code   : UDO_TNX_VALTYP
    ' Table Name : @TNX_VALTYP
    '=========================================================
    Public Sub CreateValidationTypeMaster()

        objMain.objUtilities.CreateTable("TNX_VALTYP", "Validation Type Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_VALTYP", "IQReq", "IQ Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_VALTYP", "OQReq", "OQ Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_VALTYP", "PQReq", "PQ Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_VALTYP", "ProtoReq", "Protocol Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_VALTYP", "ReportReq", "Report Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_VALTYP", "ApprRoute", "Approval Route", 30)
        objMain.objUtilities.AddFloatField("@TNX_VALTYP", "RevalFr", "Revalidation Frequency", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_VALTYP", "Active", "Active", 1)

    End Sub

    '=========================================================
    ' 5. Equipment Master
    ' UDO Code   : UDO_TNX_EQP
    ' Table Name : @TNX_EQP
    '=========================================================
    Public Sub CreateEquipmentMaster()

        objMain.objUtilities.CreateTable("TNX_EQP", "Equipment Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_EQP", "EquipType", "Equipment Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_EQP", "Departmnt", "Department", 20)
        objMain.objUtilities.AddAlphaField("@TNX_EQP", "Location", "Location", 100)
        objMain.objUtilities.AddAlphaField("@TNX_EQP", "SerialNo", "Serial Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_EQP", "Manufactr", "Manufacturer", 100)
        objMain.objUtilities.AddAlphaField("@TNX_EQP", "ModelNo", "Model Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_EQP", "CalibReq", "Calibration Required", 1)
        objMain.objUtilities.AddFloatField("@TNX_EQP", "CalibFrq", "Calibration Frequency", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_EQP", "ValReq", "Validation Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_EQP", "ValType", "Validation Type", 20)
        objMain.objUtilities.AddAlphaField("@TNX_EQP", "CleanReq", "Cleaning Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_EQP", "Status", "Status", 20)
        objMain.objUtilities.AddAlphaField("@TNX_EQP", "Active", "Active", 1)

    End Sub

    '=========================================================
    ' 6. Risk Classification Master
    ' UDO Code   : UDO_TNX_RISK
    ' Table Name : @TNX_RISK
    '=========================================================
    Public Sub CreateRiskClassificationMaster()

        objMain.objUtilities.CreateTable("TNX_RISK", "Risk Classification Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_RISK", "RiskLevel", "Risk Level", 20)
        objMain.objUtilities.AddFloatField("@TNX_RISK", "MinScore", "Minimum Score", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_RISK", "MaxScore", "Maximum Score", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_RISK", "QAApprvl", "QA Approval", 1)
        objMain.objUtilities.AddAlphaField("@TNX_RISK", "MgmtAppr", "Management Approval", 1)
        objMain.objUtilities.AddAlphaField("@TNX_RISK", "CAPAReq", "CAPA Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_RISK", "ValidReq", "Validation Required", 1)
        objMain.objUtilities.AddFloatField("@TNX_RISK", "EscDays", "Escalation Days", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_RISK", "Active", "Active", 1)

    End Sub

    '=========================================================
    ' 7. CAPA Category Master
    ' UDO Code   : UDO_TNX_CAPACAT
    ' Table Name : @TNX_CAPACAT
    '=========================================================
    Public Sub CreateCAPACategoryMaster()

        objMain.objUtilities.CreateTable("TNX_CAPACAT", "CAPA Category Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_CAPACAT", "DefRisk", "Default Risk", 20)
        objMain.objUtilities.AddAlphaField("@TNX_CAPACAT", "EffectReq", "Effectiveness Required", 1)
        objMain.objUtilities.AddFloatField("@TNX_CAPACAT", "TargetDay", "Target Days", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_CAPACAT", "ApprRoute", "Approval Route", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CAPACAT", "Active", "Active", 1)

    End Sub

    '=========================================================
    ' 8. Audit Type Master
    ' UDO Code   : UDO_TNX_AUDTYP
    ' Table Name : @TNX_AUDTYP
    '=========================================================
    Public Sub CreateAuditTypeMaster()

        objMain.objUtilities.CreateTable("TNX_AUDTYP", "Audit Type Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddFloatField("@TNX_AUDTYP", "AuditFrq", "Audit Frequency", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_AUDTYP", "CheckReq", "Checklist Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_AUDTYP", "CAPAReq", "CAPA Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_AUDTYP", "ReportReq", "Report Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_AUDTYP", "ApprRoute", "Approval Route", 30)
        objMain.objUtilities.AddAlphaField("@TNX_AUDTYP", "Active", "Active", 1)

    End Sub

    '=========================================================
    ' 9. Root Cause Master
    ' UDO Code   : UDO_TNX_ROOT
    ' Table Name : @TNX_ROOT
    '=========================================================
    Private Sub CreateRootCauseMaster()

        objMain.objUtilities.CreateTable("TNX_ROOT", "Root Cause Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@TNX_ROOT", "Category", "Category", 30)
        objMain.objUtilities.AddAlphaField("@TNX_ROOT", "Descript", "Description", 254)
        objMain.objUtilities.AddAlphaField("@TNX_ROOT", "DefCAPA", "Default CAPA", 20)
        objMain.objUtilities.AddAlphaField("@TNX_ROOT", "Active", "Active", 1)

    End Sub


    '==1
    Private Sub CreateStabilityStudy()
        objMain.objUtilities.CreateTable("TNX_PH_STAB", "Stability Study", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB", "StdCode", "Study Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB", "PCode", "Product Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB", "BNo", "Batch Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB", "SType", "Study Type", 30)
        objMain.objUtilities.AddDateField("@TNX_PH_STAB", "SDate", "Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PH_STAB", "EndDate", "End Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@TNX_PH_STAB", "Temp", "Temperature", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PH_STAB", "Hmdt", "Humidity", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_PH_STAB_D1", "Stability Study Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB_D1", "TPoint", "Time Point", 20)
        objMain.objUtilities.AddDateField("@TNX_PH_STAB_D1", "TDate", "Test Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB_D1", "TPrmtr", "Test Parameter", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB_D1", "Spctn", "Specification", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB_D1", "RValue", "Result Value", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB_D1", "Result", "Result", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB_D1", "Remarks", "Remarks", 254)


        objMain.objUtilities.CreateTable("TNX_PH_STAB_D2", "Stabiity Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        '.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "TPA", "Target Path", 254)
        objMain.objUtilities.AddLinkField("@TNX_PH_STAB_D2", "TPA", "Target Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB_D2", "FN", "File Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PH_STAB_D2", "FTT", "Free Text", 254)
        objMain.objUtilities.AddDateField("@TNX_PH_STAB_D2", "ATD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)


    End Sub

    '==2
    Private Sub CreateValidationFramework()
        objMain.objUtilities.CreateTable("TNX_PH_VAL", "Validation Framework", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL", "VType", "Validation Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL", "ECode", "Equipment Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL", "PName", "Process Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL", "DNo", "Document No", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL", "Vrsn", "Version", 20)
        objMain.objUtilities.AddDateField("@TNX_PH_VAL", "StDate", "Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PH_VAL", "EndDate", "End Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL", "ApvdBy", "Approved By", 100)
        objMain.objUtilities.AddDateField("@TNX_PH_VAL", "ApDate", "Approved Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_PH_VAL_D1", "Validation Framework Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL_D1", "CPoint", "Check Point", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL_D1", "Acria", "Acceptance Criteria", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL_D1", "AResult", "Actual Result", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL_D1", "Result", "Result", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL_D1", "ExBy", "Executed By", 100)
        objMain.objUtilities.AddDateField("@TNX_PH_VAL_D1", "ExDate", "Execution Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddLinkField("@TNX_PH_VAL_D1", "EvPath", "Evidence Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL_D1", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_PH_VAL_D2", "validation Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        '.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "TPA", "Target Path", 254)
        objMain.objUtilities.AddLinkField("@TNX_PH_VAL_D2", "TPA", "Target Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL_D2", "FN", "File Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PH_VAL_D2", "FTT", "Free Text", 254)
        objMain.objUtilities.AddDateField("@TNX_PH_VAL_D2", "ATD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)
    End Sub
    '==3
    Private Sub CreateABCCosting()
        objMain.objUtilities.CreateTable("TNX_PH_ABC_COST", "ABC Costing", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_PH_ABC_COST", "PCode", "Product Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_ABC_COST", "BatchNo", "Batch Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_ABC_COST", "PROrder", "Production Order", 50)
        objMain.objUtilities.AddFloatField("@TNX_PH_ABC_COST", "BSize", "Batch Size", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddDateField("@TNX_PH_ABC_COST", "CostDate", "Cost Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@TNX_PH_ABC_COST", "TMTRCost", "Total Material Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PH_ABC_COST", "TLCost", "Total Labor Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PH_ABC_COST", "TOHCost", "Total Overhead Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PH_ABC_COST", "QCCost", "QC Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PH_ABC_COST", "RDCost", "RD Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PH_ABC_COST", "TCost", "Total Cost", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_PH_ABC_COST", "CSPUnit", "Cost Per Unit", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@TNX_PH_ABC_COST", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_ABC_COST", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_PH_ABC_D1", "ABC Costing Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PH_ABC_D1", "CostType", "Cost Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_ABC_D1", "CSTDVR", "Cost Driver", 100)
        objMain.objUtilities.AddFloatField("@TNX_PH_ABC_D1", "CstAmt", "Cost Amount", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@TNX_PH_ABC_D1", "AlBis", "Allocation Basis", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PH_ABC_D1", "Remarks", "Remarks", 254)


        objMain.objUtilities.CreateTable("TNX_PH_ABC_D2", "ABCCost Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        '.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "TPA", "Target Path", 254)
        objMain.objUtilities.AddLinkField("@TNX_PH_ABC_D2", "TPA", "Target Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_PH_ABC_D2", "FN", "File Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PH_ABC_D2", "FTT", "Free Text", 254)
        objMain.objUtilities.AddDateField("@TNX_PH_ABC_D2", "ATD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)
    End Sub

    '==4
    Private Sub CreateCAPA()
        objMain.objUtilities.CreateTable("TNX_PH_CAPA", "Deviation CAPA", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA", "CAPAType", "CAPA Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA", "SDType", "Source Doc Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA", "SDocNum", "Source Doc No", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA", "SVRT", "Severity", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA", "Description", "Description", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA", "RTCS", "Root Cause", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA", "Owner", "Owner", 100)
        objMain.objUtilities.AddDateField("@TNX_PH_CAPA", "TDate", "Target Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PH_CAPA", "ClsDate", "Closed Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("TNX_PH_CAPA_D1", "CAPA Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA_D1", "ACTYPE", "Action Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA_D1", "ACTDTLS", "Action Details", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA_D1", "RSPBL", "Responsible", 100)
        objMain.objUtilities.AddDateField("@TNX_PH_CAPA_D1", "TDate", "Target Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA_D1", "Status", "Status", 20)
        objMain.objUtilities.AddLinkField("@TNX_PH_CAPA_D1", "EPath", "Evidence Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)

        objMain.objUtilities.CreateTable("TNX_PH_CAPA_D2", "CAPA Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        '.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "TPA", "Target Path", 254)
        objMain.objUtilities.AddLinkField("@TNX_PH_CAPA_D2", "TPA", "Target Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA_D2", "FN", "File Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_PH_CAPA_D2", "FTT", "Free Text", 254)
        objMain.objUtilities.AddDateField("@TNX_PH_CAPA_D2", "ATD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)
    End Sub
    'Sub VatReport()
    '    objMain.objUtilities.CreateTable("TNX_VATRP", "VAT Report Header", SAPbobsCOM.BoUTBTableType.bott_Document)
    '    objMain.objUtilities.AddDateField("@TNX_VATRP", "SD", "Submission Date", SAPbobsCOM.BoFldSubTypes.st_None)
    '    objMain.objUtilities.AddAlphaField("@TNX_VATRP", "TRNM", "TRNL", 254)
    '    objMain.objUtilities.AddAlphaField("@TNX_VATRP", "APPI", "Application Id", 254)
    '    objMain.objUtilities.AddAlphaField("@TNX_VATRP", "TPNE", "Taxable Person Name(English)", 254)
    '    objMain.objUtilities.AddAlphaField("@TNX_VATRP", "TPNA", "Taxable Person Name(Arabic)", 254)
    '    objMain.objUtilities.AddAlphaField("@TNX_VATRP", "TPA", "Taxable Person Address", 254)
    '    objMain.objUtilities.AddAlphaField("@TNX_VATRP", "FTY", "Document Status", 1)
    '    objMain.objUtilities.AddDateField("@TNX_VATRP", "VRPY", "VAT Return Period", SAPbobsCOM.BoFldSubTypes.st_None)
    '    objMain.objUtilities.AddDateField("@TNX_VATRP", "VATTO", "VAT Return Period TO", SAPbobsCOM.BoFldSubTypes.st_None)

    '    objMain.objUtilities.CreateTable("TNX_VATCTM_C1", "VAT Report Child1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
    '    objMain.objUtilities.AddAlphaField("@TNX_VATCTM_C1", "TRN", "Description", 254)
    '    objMain.objUtilities.AddAlphaField("@TNX_VATCTM_C1", "ACN", "Account Number", 254)
    '    ' objMain.objUtilities.AddAlphaField("@TNX_VATCTM_C1", "TLP", "Tool Tip", 254)
    '    objMain.objUtilities.AddFloatField("@TNX_VATCTM_C1", "AMT", "Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Sum)
    '    objMain.objUtilities.AddFloatField("@TNX_VATCTM_C1", "VATA", "VAT Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Sum)
    '    objMain.objUtilities.AddFloatField("@TNX_VATCTM_C1", "AST", "Adjustment(AED)", SAPbobsCOM.BoFldSubTypes.st_Sum)

    '    objMain.objUtilities.CreateTable("TNX_ATTACH_C3", "VAT Report Child2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
    '    '.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "TPA", "Target Path", 254)
    '    objMain.objUtilities.AddLinkField("@TNX_ATTACH_C3", "TPA", "Target Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
    '    objMain.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "FN", "File Name", 254)
    '    objMain.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "FTT", "Free Text", 254)
    '    objMain.objUtilities.AddDateField("@TNX_ATTACH_C3", "ATD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)
    '    'objMain.objUtilities.AddFloatField("@TNX_ATTACH_C3", "VATA", "VAT Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Quantity)

    '    objMain.objUtilities.CreateTable("TNX_VATRP_C0", "VAT Report Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
    '    objMain.objUtilities.AddAlphaField("@TNX_VATRP_C0", "VATE", "Description", 254)
    '    objMain.objUtilities.AddAlphaField("@TNX_VATRP_C0", "ARN", "Account Number", 254)
    '    ' objMain.objUtilities.AddAlphaField("@TNX_VATRP_C0", "TVAT", "Tool Tip", 254)
    '    objMain.objUtilities.AddFloatField("@TNX_VATRP_C0", "AUT", "Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Sum)
    '    objMain.objUtilities.AddFloatField("@TNX_VATRP_C0", "RVAT", "Recoverable VAT Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Sum)
    '    objMain.objUtilities.AddFloatField("@TNX_VATRP_C0", "AVAT", "Adjustment(AED)", SAPbobsCOM.BoFldSubTypes.st_Sum)
    '    'objMain.objUtilities.AddAlphaField("@TNX_VATRP_C0", "ASP", "Are you using the profit margin scheme ", 1)
    '    'objMain.objUtilities.AddAlphaField("@TNX_VATRP_C0", "VATGCC", "GCC VAT", 254)
    '    'objMain.objUtilities.AddAlphaField("@TNX_VATRP_C0", "TTVAT", "Tool Tip", 254)
    '    'objMain.objUtilities.AddFloatField("@TNX_VATRP_C0", "DVAT", "Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Quantity)
    '    'objMain.objUtilities.AddFloatField("@TNX_VATRP_C0", "EVATA", "VAT Amount(AED)", SAPbobsCOM.BoFldSubTypes.st_Quantity)
    '    'objMain.objUtilities.AddFloatField("@TNX_VATRP_C0", "AAVAT", "Adjustment(AED)", SAPbobsCOM.BoFldSubTypes.st_Quantity)

    '    objMain.objUtilities.AddFloatField("@TNX_VATRP", "TVD", "Total Vat", SAPbobsCOM.BoFldSubTypes.st_Sum)
    '    objMain.objUtilities.AddFloatField("@TNX_VATRP", "TVRP", "Total Value of recoverable tax for the period ", SAPbobsCOM.BoFldSubTypes.st_Sum)
    '    objMain.objUtilities.AddFloatField("@TNX_VATRP", "NVAT", "Total Value of recoverable", SAPbobsCOM.BoFldSubTypes.st_Sum)
    '    'objMain.objUtilities.AddAlphaField("@TNX_VATRP", "TAVMN", "Total vat refund", 1)




    'End Sub

    Sub OnboardingProcessTable()
        objMain.objUtilities.CreateTable("TNX_ONBP", "Onboarding Process Table", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "CName", "Company Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Email", "Email", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "City", "City", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "PCode", "Postal Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Sheet", "Additional Sheet", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Address", "Short Address", 100)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "ANum", "Additional Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "OTP", "One Time Password", 30)
        objMain.objUtilities.AddDateField("@TNX_ONBP", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Country", "Country", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Type", "Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "IdNum", "ID Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "District", "District", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "SName", "Street Name", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "BNum", "Building Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "BCat", "Business Category", 50)
        objMain.objUtilities.AddAlphaMemoField("@TNX_ONBP", "CSR", "CSR", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_ONBP", "PCSID", "PCSID", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_ONBP", "CSID", "CSID", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_ONBP", "PBKey", "Public Key", 5000)
        objMain.objUtilities.AddAlphaMemoField("@TNX_ONBP", "PRKey", "Private Key", 5000)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "SName", "Street Name", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP", "PType", "Posting Type", 50)


        objMain.objUtilities.CreateTable("TNX_ONBP_C0", "Onboarding Process Child Table", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "CCode", "Customer Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "PMode", "Posting Mode", 30)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "IActive", "In Active", 5)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "ILog", "Is Logged in", 10)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "Options", "Options", 10)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "Remarks", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "URL", "URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "UUID", "UUID", 254)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "Key", "Key", 254)
        objMain.objUtilities.AddAlphaField("@TNX_ONBP_C0", "PType", "PType", 254)

    End Sub
    Sub PayPosting()
        objMain.objUtilities.CreateTable("TNX_PAY", "Invoice Posting Table", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        '  objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPURL", "Invoice Posting URL", 254)
        'objMain.objUtilities.AddAlphaField("@TNX_INVF", "DBNE", "DB Name", 254)
        'objMain.objUtilities.AddAlphaField("@TNX_INVF", "CTY", "Country", 60)
        'objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPPRL", "Invoice Possting URL", 254)
    End Sub
    Sub PayLoadData()
        objMain.objUtilities.CreateTable("TNX_PAYLD", "Invoice Posting Table", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        '  objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPURL", "Invoice Posting URL", 254)
        'objMain.objUtilities.AddAlphaField("@TNX_INVF", "DBNE", "DB Name", 254)
        'objMain.objUtilities.AddAlphaField("@TNX_INVF", "CTY", "Country", 60)
        'objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPPRL", "Invoice Possting URL", 254)
    End Sub
    Sub InvoicePostingTable()
        objMain.objUtilities.CreateTable("TNX_INVF", "Invoice Posting Table", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        '  objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPURL", "Invoice Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "DBNE", "DB Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CTY", "Country", 60)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPPRL", "Invoice Possting URL", 254)


        objMain.objUtilities.AddAlphaField("@TNX_INVF", "IQRCG", "Invoice QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "IUUID", "Invoice UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMPRL", "Credit Memo Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMQRC", "Credit Memo QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMURL", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "BPID", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "APIK", "API Key", 60)
        'objMain.objUtilities.AddDateField("@TNX_ONBP", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        'objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Country", "Country", 50)

        'B2B
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPURLB", "Invoice Possting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "IQRCGB", "Invoice QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "UUIDB", "Invoice UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMPRLB", "Credit Memo Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMQRB", "Credit Memo QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMURLB", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "BPIDB", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "APIKB", "API Key", 60)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TYPE", "Posting Type", 60)

        'LIVE

        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TIPPRL", "Invoice Possting URL", 254)


        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TIQRCG", "Invoice QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TIUUID", "Invoice UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TCMPRL", "Credit Memo Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TCMQRC", "Credit Memo QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TCMURL", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TBPID", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TAPIK", "API Key", 60)
        'objMain.objUtilities.AddDateField("@TNX_ONBP", "DocDate", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        'objMain.objUtilities.AddAlphaField("@TNX_ONBP", "Country", "Country", 50)

        'B2B
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TIPURLB", "Invoice Possting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TIQRCGB", "Invoice QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TUUIDB", "Invoice UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TCMPRLB", "Credit Memo Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TCMQRB", "Credit Memo QR Code Generation URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TCMURLB", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TBPIDB", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "TAPIKB", "API Key", 60)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LTYPE", "Posting Type", 60)

        'UAE test
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "AIPPRL", "Access Token URL", 254)


        objMain.objUtilities.AddAlphaField("@TNX_INVF", "AIUUID", "Get UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "ACMQRC", "Password", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "ABPID", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "ALIQRCG", "Invoice Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "ALCMPRL", "User Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "ALCMURL", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "ALAPIK", "API Key", 60)

        'live
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LIPPRL", "Access Token URL", 254)


        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LIUUID", "Get UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LCMQRC", "Password", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LBPID", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LIQRCG", "Invoice Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LCMPRL", "User Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LCMURL", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LAPIK", "API Key", 60)

        objMain.objUtilities.AddAlphaField("@TNX_INVF", "UTA", "Access Token URL", 254)


        objMain.objUtilities.AddAlphaField("@TNX_INVF", "GURL", "Get UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "BPI", "Password", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "IPU", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CPL", "Invoice Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "USM", "User Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "CMRL", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "APID", "API Key", 60)

        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LUTA", "Access Token URL", 254)


        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LGURL", "Get UUIDs URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LBPI", "Password", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LIPU", "Business Profile ID", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LCPL", "Invoice Posting URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LUSM", "User Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LCMRL", "Credit Memo UUID URL", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INVF", "LAPID", "API Key", 60)
    End Sub
    Sub Errorlogs()

        objMain.objUtilities.CreateTable("TNX_ERRORLOGS", "Error Logs", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objMain.objUtilities.AddDateField("@TNX_ERRORLOGS", "DATE", "Error Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_ERRORLOGS", "TIME", "Error Time", 10)
        objMain.objUtilities.AddAlphaField("@TNX_ERRORLOGS", "ERROR_MSG", "Error Message", 254)

    End Sub





#End Region

    '#Region "SQL Scripts, Queries, Authorizations"
    Sub RunSQLScripts()
        Try

            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_ZATCA_OINV.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_ZATCA_ORIN.sql"))

            'objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("SQLQuery.PayLoad.txt"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNX_EINVOICE_TN.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("Corporate.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("VatReport.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("Tnx_Branch.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("EInvoice.sql"))
            objUtilities.ExecuteSQLScript(objMain.objCompany, objUtilities.LoadEmbeddedSQL("TNXPLByBranch.sql"))
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub QueryManager()
        Try

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Sales Orders Posted", 10, objUtilities.LoadEmbeddedSQL("SalesOrdersPosted.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Corporate Tax - Open Status", 12, objUtilities.LoadEmbeddedSQL("CorporateTaxOpenStatus.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Corporate Tax - Approved", 14, objUtilities.LoadEmbeddedSQL("CorporateTaxApproved.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "Corporate Tax JE post", 16, objUtilities.LoadEmbeddedSQL("CorporateTaxJEPost.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "FTA Vat Report", 18, objUtilities.LoadEmbeddedSQL("FTAVatReport.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "FTA Vat Report- Approved", 20, objUtilities.LoadEmbeddedSQL("FTAVatReportApproved.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "E Invoice - status", 22, objUtilities.LoadEmbeddedSQL("EInvoiceStatus.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "E Invoice AR credit", 24, objUtilities.LoadEmbeddedSQL("EInvoiceARCredit.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "E Invoice - pending", 26, objUtilities.LoadEmbeddedSQL("EInvoicePending.sql"), "Alert")

            objUtilities.AddOrUpdateQuery(objMain.objCompany, "AP Credit Memo Alert", 28, objUtilities.LoadEmbeddedSQL("APCreditMemoAlert.sql"), "Alert")
        Catch ex As Exception

        End Try
    End Sub
    Sub AddAlerts()
        Try


            objUtilities.CreateCustomAlert("AP Credit Memo Alert", 8, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("Sales Orders Posted", 10, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("Corporate Tax - Open Status", 12, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("Corporate Tax - Approved", 14, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("Corporate Tax JE post", 16, 1, 1, AlertManagementFrequencyType.atfi_Days)

            objUtilities.CreateCustomAlert("FTA Vat Report", 18, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("FTA Vat Report- Approved", 20, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("E Invoice - status", 22, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("E Invoice AR credit", 24, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("E Invoice - pending", 26, 1, 1, AlertManagementFrequencyType.atfi_Days)


            objUtilities.CreateCustomAlert("AP Credit Memo Alert", 28, 1, 1, AlertManagementFrequencyType.atfi_Days)

            objMain.objApplication.StatusBar.SetText("All Alerts created successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub UpdateTransactionNotification()
        Try

            Dim oRecordset As Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim procedureName As String = "SBO_SP_TRANSACTIONNOTIFICATION"

            ' 1. Get existing procedure source
            Dim DBName As String = objMain.objCompany.CompanyDB
            Dim Qry As String = "SELECT ""DEFINITION"" FROM ""SYS"".""PROCEDURES"" WHERE ""SCHEMA_NAME"" = '" & DBName & "' AND ""PROCEDURE_NAME"" = 'SBO_SP_TRANSACTIONNOTIFICATION'"
            oRecordset.DoQuery(Qry)

            If oRecordset.EoF Then
                Throw New Exception("Procedure not found.")
            End If

            Dim originalCode As String = oRecordset.Fields.Item(0).Value.ToString()
            If originalCode.Contains("TNX_EINVOICE_TN") Then
                Exit Sub
            End If
            ' 2. Your code to append
            Dim codeToInsert As String = "
    
    if :error = 0 then
        CALL TNX_EINVOICE_TN(:object_type, :transaction_type, :num_of_cols_in_key, :list_of_key_cols_tab_del, :list_of_cols_val_tab_del, :error, :error_message);
    end if;
select :error, :error_message FROM dummy;
    "

            ' 3. Insert before the final "select :error, :error_message FROM dummy;"
            Dim finalCode As String = ""
            If originalCode.Contains("select :error, :error_message FROM dummy;") Then
                finalCode = originalCode.Replace("select :error, :error_message FROM dummy;", codeToInsert)
            Else
                Throw New Exception("""select :error, :error_message FROM dummy;"" not found in procedure definition.")
            End If
            finalCode = finalCode.Replace("CREATE", "ALTER")

            oRecordset.DoQuery(finalCode)
            objMain.objApplication.StatusBar.SetText("Transaction Notification Procedure updated successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Sub AddUserFormAuthorizations()
        Try

            objUtilities.AddAuthorization("VATR", "FTA VAT REPORT", "VATR", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("CTAXCAL", "Corporate Tax Calculation", "CTAXCAL", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("CTAXC", "Corporate Tax Configuration", "CTAXC", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("CONFT", "E Invoicing Configuration", "CONFT", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("COTAX", "Corporate Tax", "COTAX", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("frm_FTAVM", "FTA VAT", "frm_FTAVM", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
            objUtilities.AddAuthorization("frm_LKMTR", "LINK MASTER", "frm_LKMTR", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub



    Sub COAManagementTables()



        objMain.objUtilities.CreateTable("TNX_COA_H", "COA Header", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "COANO", "COA Number", 30)
        objMain.objUtilities.AddDateField("@TNX_COA_H", "COADATE", "COA Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "COATYPE", "COA Type", 20)
        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "DS", "Status", 20)
        objMain.objUtilities.AddDateField("@TNX_COA_H", "DDS", "DocumenttDate", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "SourceT", "Source Type", 30)
        objMain.objUtilities.AddInteger("@TNX_COA_H", "SourceDE", "Source DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddInteger("@TNX_COA_H", "SourceNO", "Source DocNum", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "ItemC", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "ItemN", "Item Name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "BatchNO", "Batch No", 50)

        objMain.objUtilities.AddDateField("@TNX_COA_H", "MfgDate", "Manufacturing Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_COA_H", "ExpDate", "Expiry Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "CardC", "Customer/Vendor Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "VENNAME", "Customer/Vendor Name", 200)

        objMain.objUtilities.AddInteger("@TNX_COA_H", "DDNO", "Delivery No", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddInteger("@TNX_COA_H", "GRN", "GRPO / Receipt No", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "SampleNo", "Sample No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "TestDocNo", "Test Document No", 30)

        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "SpecCode", "Specification Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "SpecV", "Specification Version", 20)

        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "QCR", "QC Result", 20)
        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "RST", "Release Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "Remarks", "Remarks", 254)

        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "PBY", "Prepared By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "CHB", "Checked By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "APPB", "Approved By", 50)
        objMain.objUtilities.AddDateField("@TNX_COA_H", "APB", "Approved Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "EsignS", "E-Sign Status", 20)
        objMain.objUtilities.AddInteger("@TNX_COA_H", "PCT", "Print Count", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddInteger("@TNX_COA_H", "ATY", "Attachment Entry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_COA_H", "CUSER", "Created User", 50)
        objMain.objUtilities.AddDateField("@TNX_COA_H", "CDATE", "Created Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("TNX_COA_T", "COA Test Result Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_COA_T", "TestCode", "Test Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_COA_T", "TestName", "Test Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_COA_T", "TestM", "Test Method", 100)
        objMain.objUtilities.AddAlphaField("@TNX_COA_T", "Unit", "Unit", 20)

        objMain.objUtilities.AddFloatField("@TNX_COA_T", "SpecMin", "Spec Min", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_COA_T", "SpecMax", "Spec Max", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objMain.objUtilities.AddAlphaField("@TNX_COA_T", "SpecText", "Spec Text", 254)

        objMain.objUtilities.AddFloatField("@TNX_COA_T", "ResultV", "Result Value", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddAlphaField("@TNX_COA_T", "ResultT", "Result Text", 254)
        objMain.objUtilities.AddAlphaField("@TNX_COA_T", "ResultS", "Result Status", 20)

        objMain.objUtilities.AddAlphaField("@TNX_COA_T", "Analyst", "Analyst", 50)
        objMain.objUtilities.AddDateField("@TNX_COA_T", "TestD", "Test Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_COA_T", "Instrument", "Instrument", 100)
        objMain.objUtilities.AddAlphaField("@TNX_COA_T", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_COA_A", "COA Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_COA_A", "FileN", "File Name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_COA_A", "FileT", "File Type", 30)
        objMain.objUtilities.AddLinkField("@TNX_COA_A", "FileP", "File Path", 254, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddInteger("@TNX_COA_A", "AttachE", "Attachment Entry", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_COA_A", "UPU", "Uploaded By", 50)
        objMain.objUtilities.AddDateField("@TNX_COA_A", "UPD", "Uploaded Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_COA_A", "RM", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_COA_APP", "COA Approval Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddInteger("@TNX_COA_APP", "ALevel", "Approval Level", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_COA_APP", "AppR", "Approver Role", 50)
        objMain.objUtilities.AddAlphaField("@TNX_COA_APP", "AppU", "Approver User", 50)
        objMain.objUtilities.AddAlphaField("@TNX_COA_APP", "Status", "Status", 20)
        objMain.objUtilities.AddDateField("@TNX_COA_APP", "ActionA", "Action Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_COA_APP", "Esign", "E-Sign", 20)
        objMain.objUtilities.AddAlphaField("@TNX_COA_APP", "Com", "Comments", 254)



    End Sub

    Sub CreatePharmaBMRExecution()
        ' Header Table: @TNX_PBMR_H
        objMain.objUtilities.CreateTable("TNX_PBMR_H", "Pharma BMR Execution Header", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "DocNum", "Document Number", 50)

        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "BMRNo", "BMR Reference Number", 50)
        objMain.objUtilities.AddInteger("@TNX_PBMR_H", "ProdOrdEntry", "Production Order DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "ProdOrdNo", "Production Order Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "ProdOrdStatus", "Production Order Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "ItemCode", "Finished Goods Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "ItemName", "Finished Goods Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "BatchNo", "Manufacturing Batch Number", 50)
        objMain.objUtilities.AddFloatField("@TNX_PBMR_H", "PlannedQty", "Planned Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PBMR_H", "CompletedQty", "Completed Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "UOM", "UOM", 20)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "BOMType", "BOM Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "BOMNo", "BOM Reference Number", 50)
        objMain.objUtilities.AddDateField("@TNX_PBMR_H", "MfgDate", "Manufacturing Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PBMR_H", "ExpDate", "Expiry Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PBMR_H", "StartDate", "BMR Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PBMR_H", "EndDate", "BMR Completion Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "Shift", "Shift", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "ProdArea", "Production Area", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "Status", "Document Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "QCStatus", "QC Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_H", "ApprovalStatus", "Approval Status", 30)
        objMain.objUtilities.AddAlphaMemoField("@TNX_PBMR_H", "Remarks", "Remarks", 2000)

        ' Child 1: Manufacturing Stages
        objMain.objUtilities.CreateTable("TNX_PBMR_STAGE", "Manufacturing Stages", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddInteger("@TNX_PBMR_STAGE", "LineId", "Line Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_STAGE", "StageCode", "Stage Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_STAGE", "StageName", "Stage Name", 100)
        objMain.objUtilities.AddInteger("@TNX_PBMR_STAGE", "SeqNo", "Stage Sequence", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddDateField("@TNX_PBMR_STAGE", "StartTime", "Stage Start Time", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PBMR_STAGE", "EndTime", "Stage End Time", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddInteger("@TNX_PBMR_STAGE", "DurationMin", "Duration in Minutes", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_STAGE", "Operator", "Operator User", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_STAGE", "Supervisor", "Supervisor", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_STAGE", "EquipCode", "Equipment Code", 50)
        objMain.objUtilities.AddFloatField("@TNX_PBMR_STAGE", "Temp", "Temperature", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PBMR_STAGE", "Humidity", "Humidity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PBMR_STAGE", "RPM", "Machine Speed", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_STAGE", "Status", "Stage Status", 30)
        objMain.objUtilities.AddAlphaMemoField("@TNX_PBMR_STAGE", "Remarks", "Stage Remarks", 1000)

        ' Child 2: Material Consumption
        objMain.objUtilities.CreateTable("TNX_PBMR_MAT", "Material Consumption", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddInteger("@TNX_PBMR_MAT", "LineId", "Line Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddInteger("@TNX_PBMR_MAT", "BaseLine", "SAP Production Order Component Line", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_MAT", "ItemCode", "Raw Material Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_MAT", "ItemName", "Raw Material Name", 100)
        objMain.objUtilities.AddFloatField("@TNX_PBMR_MAT", "PlannedQty", "BOM Planned Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PBMR_MAT", "IssuedQty", "SAP Issued Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PBMR_MAT", "ConsumedQty", "Actual Consumed Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PBMR_MAT", "ReturnQty", "Return Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PBMR_MAT", "WasteQty", "Wastage Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_MAT", "BatchNo", "Raw Material Batch", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_MAT", "WhsCode", "Warehouse", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_MAT", "UOM", "UOM", 20)
        objMain.objUtilities.AddFloatField("@TNX_PBMR_MAT", "VarianceQty", "Variance Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_PBMR_MAT", "VariancePct", "Variance Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_MAT", "Status", "Material Status", 30)
        objMain.objUtilities.AddAlphaMemoField("@TNX_PBMR_MAT", "Remarks", "Material Remarks", 1000)

        ' Child 3: Equipment Usage
        objMain.objUtilities.CreateTable("TNX_PBMR_EQP", "Equipment Usage", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddInteger("@TNX_PBMR_EQP", "LineId", "Line Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_EQP", "EquipCode", "Equipment Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_EQP", "EquipName", "Equipment Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_EQP", "StageCode", "Linked Stage", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_EQP", "CleaningStatus", "Cleaning Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_EQP", "CalibrationStatus", "Calibration Status", 30)
        objMain.objUtilities.AddDateField("@TNX_PBMR_EQP", "UsedFrom", "Usage Start", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_PBMR_EQP", "UsedTo", "Usage End", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_EQP", "Operator", "Operator", 100)
        objMain.objUtilities.AddAlphaMemoField("@TNX_PBMR_EQP", "Remarks", "Equipment Remarks", 1000)

        ' Child 4: In-Process QC
        objMain.objUtilities.CreateTable("TNX_PBMR_IPQC", "In-Process QC", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddInteger("@TNX_PBMR_IPQC", "LineId", "Line Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_IPQC", "StageCode", "Manufacturing Stage", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_IPQC", "TestCode", "QC Test Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_IPQC", "TestName", "QC Test Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_IPQC", "Specification", "Required Specification", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_IPQC", "ResultValue", "Actual Result", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_IPQC", "ResultStatus", "Result Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_IPQC", "CheckedBy", "QC User", 100)
        objMain.objUtilities.AddDateField("@TNX_PBMR_IPQC", "CheckedDate", "Checked Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_PBMR_IPQC", "Remarks", "QC Remarks", 1000)

        ' Child 5: Deviation
        objMain.objUtilities.CreateTable("TNX_PBMR_DEV", "BMR Deviations", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddInteger("@TNX_PBMR_DEV", "LineId", "Line Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_DEV", "DeviationNo", "Deviation Reference", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_DEV", "StageCode", "Related Stage", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_DEV", "DeviationType", "Deviation Type", 30)
        objMain.objUtilities.AddAlphaMemoField("@TNX_PBMR_DEV", "Description", "Deviation Details", 1000)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_DEV", "Severity", "Severity Level", 30)
        objMain.objUtilities.AddAlphaMemoField("@TNX_PBMR_DEV", "ActionTaken", "Immediate Action", 1000)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_DEV", "CAPARequired", "CAPA Required", 1)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_DEV", "Status", "Deviation Status", 30)

        ' Child 6: Approval
        objMain.objUtilities.CreateTable("TNX_PBMR_APP", "BMR Approvals", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddInteger("@TNX_PBMR_APP", "LineId", "Line Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddInteger("@TNX_PBMR_APP", "ApprovalLevel", "Approval Level", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_APP", "ApproverRole", "Approver Role", 50)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_APP", "ApproverUser", "Approver User", 100)
        objMain.objUtilities.AddAlphaField("@TNX_PBMR_APP", "Status", "Approval Status", 30)
        objMain.objUtilities.AddDateField("@TNX_PBMR_APP", "ApprovedDate", "Approval Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_PBMR_APP", "Remarks", "Approval Remarks", 1000)
    End Sub


    Sub IncidentManagementTables()

        '================ HEADER TABLE =================
        objMain.objUtilities.CreateTable("TNX_INCH", "Incident Management Header", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_INCH", "IncCode", "Incident Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "IncTitle", "Incident Title", 200)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "IncType", "Incident Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "IncCat", "Incident Category", 50)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "SrcModule", "Source Module", 50)
        objMain.objUtilities.AddDateField("@TNX_INCH", "DDT", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)


        objMain.objUtilities.AddAlphaField("@TNX_INCH", "RefObjT", "Ref Object Type", 50)
        objMain.objUtilities.AddInteger("@TNX_INCH", "RefDE", "Ref DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "RefDNo", "Ref DocNum", 50)

        objMain.objUtilities.AddAlphaField("@TNX_INCH", "ItemCode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "ItemName", "Item Name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "BatchNum", "Batch Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "WhsCode", "Warehouse", 30)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "EquipCode", "Equipment Code", 50)

        objMain.objUtilities.AddAlphaField("@TNX_INCH", "Dept", "Department", 50)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "Location", "Incident Location", 100)
        objMain.objUtilities.AddDateField("@TNX_INCH", "IncDate", "Incident Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "IncTime", "Incident Time", 10)

        objMain.objUtilities.AddAlphaField("@TNX_INCH", "RepBy", "Reported By", 50)
        objMain.objUtilities.AddDateField("@TNX_INCH", "RepDate", "Reported Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_INCH", "Severity", "Severity", 20)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "Priority", "Priority", 20)

        objMain.objUtilities.AddAlphaField("@TNX_INCH", "BatchImp", "Batch Impact", 20)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "RegImp", "Regulatory Impact", 20)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "CustImp", "Customer Impact", 20)

        objMain.objUtilities.AddAlphaField("@TNX_INCH", "ImmAct", "Immediate Action", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "RootCause", "Root Cause", 254)

        objMain.objUtilities.AddAlphaField("@TNX_INCH", "CAPAReq", "CAPA Required", 10)
        objMain.objUtilities.AddInteger("@TNX_INCH", "CAPADE", "CAPA DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_INCH", "ChangeReq", "Change Required", 10)
        objMain.objUtilities.AddInteger("@TNX_INCH", "ChangeDE", "Change DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_INCH", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "QAOwner", "QA Owner", 50)

        objMain.objUtilities.AddDateField("@TNX_INCH", "TarClsDt", "Target Close Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "ClosedBy", "Closed By", 50)
        objMain.objUtilities.AddDateField("@TNX_INCH", "ClosedDt", "Closed Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddInteger("@TNX_INCH", "AttachE", "Attachment Entry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_INCH", "Remarks", "Remarks", 254)


        '================ CHILD TABLE 1 : INVESTIGATION LINES =================
        objMain.objUtilities.CreateTable("TNX_INCL", "Incident Investigation Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_INCL", "InvStep", "Investigation Step", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INCL", "Response", "Response", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INCL", "EvdType", "Evidence Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_INCL", "EvdRef", "Evidence Reference", 100)
        objMain.objUtilities.AddAlphaField("@TNX_INCL", "RCFlag", "Root Cause Flag", 10)
        objMain.objUtilities.AddAlphaField("@TNX_INCL", "RCType", "Root Cause Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_INCL", "Resp", "Responsible", 50)
        objMain.objUtilities.AddDateField("@TNX_INCL", "InvDate", "Investigation Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_INCL", "Status", "Status", 20)
        objMain.objUtilities.AddAlphaField("@TNX_INCL", "Remarks", "Remarks", 254)


        '================ CHILD TABLE 2 : ACTION LINES =================
        objMain.objUtilities.CreateTable("TNX_INCACT", "Incident Action Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_INCACT", "ActType", "Action Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_INCACT", "ActDesc", "Action Description", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INCACT", "ActOwner", "Action Owner", 50)
        objMain.objUtilities.AddAlphaField("@TNX_INCACT", "Dept", "Department", 50)
        objMain.objUtilities.AddDateField("@TNX_INCACT", "DueDate", "Due Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_INCACT", "CompDate", "Completed Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_INCACT", "Status", "Status", 30)
        objMain.objUtilities.AddInteger("@TNX_INCACT", "EvdAttach", "Evidence Attachment", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_INCACT", "VerBy", "Verified By", 50)
        objMain.objUtilities.AddDateField("@TNX_INCACT", "VerDate", "Verified Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_INCACT", "VerRem", "Verification Remarks", 254)


        '================ CHILD TABLE 3 : CAPA LINK =================
        objMain.objUtilities.CreateTable("TNX_INCAP", "Incident CAPA Link", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddInteger("@TNX_INCAP", "CAPADE", "CAPA DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_INCAP", "CAPANo", "CAPA Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_INCAP", "CAPAType", "CAPA Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_INCAP", "CAPAStat", "CAPA Status", 30)
        objMain.objUtilities.AddDateField("@TNX_INCAP", "LinkDate", "Link Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_INCAP", "Remarks", "Remarks", 254)


        '================ CHILD TABLE 4 : ATTACHMENT / EVIDENCE =================
        objMain.objUtilities.CreateTable("TNX_INCAT", "Incident Attachment Evidence", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_INCAT", "AttachType", "Attachment Type", 50)
        ' objMain.objUtilities.AddInteger("@TNX_INCAT", "AttachE", "Attachment Entry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddLinkField("@TNX_INCAT", "AttachE", "Attachment Entry", 254, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_INCAT", "FileName", "File Name", 200)
        objMain.objUtilities.AddAlphaField("@TNX_INCAT", "Descr", "Description", 254)
        objMain.objUtilities.AddAlphaField("@TNX_INCAT", "UploadBy", "Uploaded By", 50)
        objMain.objUtilities.AddDateField("@TNX_INCAT", "UploadDt", "Upload Date", SAPbobsCOM.BoFldSubTypes.st_None)

    End Sub

    Sub ValidationManagementTables()

        '================ HEADER TABLE =================
        objMain.objUtilities.CreateTable("TNX_VALH", "Validation Management Header", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_VALH", "ValCode", "Validation Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_VALH", "ValTitle", "Validation Title", 200)
        objMain.objUtilities.AddAlphaField("@TNX_VALH", "ValType", "Validation Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VALH", "ValCat", "Validation Category", 50)
        objMain.objUtilities.AddDateField("@TNX_VALH", "DDM", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_VALH", "RefObjT", "Ref Object Type", 50)
        objMain.objUtilities.AddInteger("@TNX_VALH", "RefDE", "Ref DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_VALH", "ItemCode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VALH", "BatchNum", "Batch Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VALH", "EquipCode", "Equipment Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VALH", "Dept", "Department", 50)

        objMain.objUtilities.AddAlphaField("@TNX_VALH", "RiskLvl", "Risk Level", 20)
        objMain.objUtilities.AddAlphaField("@TNX_VALH", "ProtNo", "Protocol Number", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VALH", "ReportNo", "Report Number", 50)

        objMain.objUtilities.AddDateField("@TNX_VALH", "StartDate", "Planned Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_VALH", "EndDate", "Planned End Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_VALH", "ActStart", "Actual Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_VALH", "ActEnd", "Actual End Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_VALH", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_VALH", "Result", "Result", 30)

        objMain.objUtilities.AddAlphaField("@TNX_VALH", "PrepBy", "Prepared By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VALH", "RevBy", "Reviewed By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VALH", "AppBy", "Approved By", 50)

        objMain.objUtilities.AddInteger("@TNX_VALH", "AttachE", "Attachment Entry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_VALH", "Remarks", "Remarks", 254)


        '================ CHILD TABLE 1 : VALIDATION PLAN =================
        objMain.objUtilities.CreateTable("TNX_VALP", "Validation Plan", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_VALP", "Objective", "Validation Objective", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VALP", "Scope", "Scope", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VALP", "AccCrit", "Acceptance Criteria", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VALP", "SampPlan", "Sampling Plan", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VALP", "RespDept", "Responsible Department", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VALP", "RespUser", "Responsible User", 50)
        objMain.objUtilities.AddDateField("@TNX_VALP", "TargetDate", "Target Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VALP", "Status", "Status", 20)


        '================ CHILD TABLE 2 : VALIDATION STEPS =================
        objMain.objUtilities.CreateTable("TNX_VALSTEP", "Validation Steps", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_VALSTEP", "StepNo", "Step Number", 20)
        objMain.objUtilities.AddAlphaField("@TNX_VALSTEP", "Phase", "Phase", 30)
        objMain.objUtilities.AddAlphaField("@TNX_VALSTEP", "TestName", "Test Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_VALSTEP", "TestMeth", "Test Method", 100)
        objMain.objUtilities.AddAlphaField("@TNX_VALSTEP", "Parameter", "Parameter", 100)
        objMain.objUtilities.AddAlphaField("@TNX_VALSTEP", "ExpValue", "Expected Value", 100)

        objMain.objUtilities.AddFloatField("@TNX_VALSTEP", "MinLimit", "Minimum Limit", SAPbobsCOM.BoFldSubTypes.st_Sum)
        objMain.objUtilities.AddFloatField("@TNX_VALSTEP", "MaxLimit", "Maximum Limit", SAPbobsCOM.BoFldSubTypes.st_Sum)

        objMain.objUtilities.AddAlphaField("@TNX_VALSTEP", "UOM", "UOM", 20)
        objMain.objUtilities.AddAlphaField("@TNX_VALSTEP", "Freq", "Frequency", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VALSTEP", "Resp", "Responsible", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VALSTEP", "Mandatory", "Mandatory", 10)
        objMain.objUtilities.AddAlphaField("@TNX_VALSTEP", "Status", "Status", 30)


        '================ CHILD TABLE 3 : VALIDATION RESULTS =================
        objMain.objUtilities.CreateTable("TNX_VALRES", "Validation Results", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_VALRES", "StepNo", "Step Number", 20)
        objMain.objUtilities.AddAlphaField("@TNX_VALRES", "TestName", "Test Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_VALRES", "ActValue", "Actual Value", 100)
        objMain.objUtilities.AddAlphaField("@TNX_VALRES", "ResultTxt", "Result Text", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VALRES", "Result", "Result", 20)
        objMain.objUtilities.AddAlphaField("@TNX_VALRES", "Deviation", "Deviation", 10)

        objMain.objUtilities.AddInteger("@TNX_VALRES", "DevDE", "Deviation DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddInteger("@TNX_VALRES", "CAPADE", "CAPA DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddInteger("@TNX_VALRES", "AttachE", "Attachment Entry", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.AddAlphaField("@TNX_VALRES", "TestedBy", "Tested By", 50)
        objMain.objUtilities.AddDateField("@TNX_VALRES", "TestDate", "Test Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VALRES", "VerBy", "Verified By", 50)
        objMain.objUtilities.AddDateField("@TNX_VALRES", "VerDate", "Verified Date", SAPbobsCOM.BoFldSubTypes.st_None)


        '================ CHILD TABLE 4 : APPROVAL LINES =================
        objMain.objUtilities.CreateTable("TNX_VALAP", "Validation Approval Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddInteger("@TNX_VALAP", "Level", "Approval Level", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_VALAP", "Role", "Role", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VALAP", "Approver", "Approver", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VALAP", "Status", "Status", 30)
        objMain.objUtilities.AddDateField("@TNX_VALAP", "ActDate", "Action Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VALAP", "Comments", "Comments", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VALAP", "ESignRef", "E-Sign Ref", 50)


        '================ CHILD TABLE 5 : VALIDATION DOCUMENTS =================
        objMain.objUtilities.CreateTable("TNX_VALDOC", "Validation Documents", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_VALDOC", "DocType", "Document Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VALDOC", "DocName", "Document Name", 200)
        objMain.objUtilities.AddInteger("@TNX_VALDOC", "AttachE", "Attachment Entry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_VALDOC", "UploadBy", "Uploaded By", 50)
        objMain.objUtilities.AddDateField("@TNX_VALDOC", "UploadDt", "Uploaded Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VALDOC", "Required", "Required", 10)
        objMain.objUtilities.AddAlphaField("@TNX_VALDOC", "Status", "Status", 20)

        objMain.objUtilities.CreateTable("TNX_VALAPP", "Validation Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        '.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "TPA", "Target Path", 254)
        objMain.objUtilities.AddLinkField("@TNX_VALAPP", "TPH", "Target Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_VALAPP", "FNM", "File Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VALAPP", "FTR", "Free Text", 254)
        objMain.objUtilities.AddDateField("@TNX_VALAPP ", "ATCD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)


    End Sub


    Sub ChangeControlTables()


        '================ HEADER TABLE =================
        objMain.objUtilities.CreateTable("TNX_CHGH", "Change Control Header", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "CHGCode", "Change Control Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "CHGTitle", "Change Title", 200)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "CHGType", "Change Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "CHGCat", "Change Category", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "Dept", "Department", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "ReqBy", "Requestor", 50)

        objMain.objUtilities.AddDateField("@TNX_CHGH", "ReqDate", "Request Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_CHGH", "DDN", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_CHGH", "ReqDDate", "Required Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "RefObjT", "Ref Object Type", 50)
        objMain.objUtilities.AddInteger("@TNX_CHGH", "RefDE", "Ref DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "RefDNo", "Ref DocNum", 50)

        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "ChgReason", "Change Reason", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "CurState", "Current State", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "PropChg", "Proposed Change", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "BusJust", "Business Justification", 254)

        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "GMPImp", "GMP Impact", 10)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "QImp", "Quality Impact", 10)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "RegImp", "Regulatory Impact", 10)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "ValImp", "Validation Impact", 10)

        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "RiskLvl", "Risk Level", 20)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "Priority", "Priority", 20)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "AppStatus", "Approval Status", 30)

        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "CAPAReq", "CAPA Required", 10)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "ValReq", "Validation Required", 10)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "TrainReq", "Training Required", 10)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "SOPRevReq", "SOP Revision Required", 10)

        objMain.objUtilities.AddInteger("@TNX_CHGH", "AttachE", "Attachment Entry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "FinalDec", "Final Decision", 30)
        objMain.objUtilities.AddDateField("@TNX_CHGH", "CloseDate", "Closure Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "ClosedBy", "Closed By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CHGH", "Remarks", "Remarks", 254)


        '================ CHILD TABLE 1 : IMPACT =================
        objMain.objUtilities.CreateTable("TNX_CHGIMP", "Change Impact Assessment", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_CHGIMP", "Area", "Area", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CHGIMP", "ImpType", "Impact Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CHGIMP", "ImpDesc", "Impact Description", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CHGIMP", "Severity", "Severity", 20)
        objMain.objUtilities.AddAlphaField("@TNX_CHGIMP", "Prob", "Probability", 20)
        objMain.objUtilities.AddAlphaField("@TNX_CHGIMP", "Detect", "Detectability", 20)
        objMain.objUtilities.AddInteger("@TNX_CHGIMP", "RiskScore", "Risk Score", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_CHGIMP", "ActReq", "Action Required", 10)
        objMain.objUtilities.AddAlphaField("@TNX_CHGIMP", "RecAction", "Recommended Action", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CHGIMP", "Resp", "Responsible", 50)
        objMain.objUtilities.AddDateField("@TNX_CHGIMP", "TargetDate", "Target Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CHGIMP", "Status", "Status", 20)


        '================ CHILD TABLE 2 : APPROVAL =================
        objMain.objUtilities.CreateTable("TNX_CHGAP", "Change Approval Matrix", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddInteger("@TNX_CHGAP", "Level", "Approval Level", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_CHGAP", "AppRole", "Approver Role", 100)
        objMain.objUtilities.AddAlphaField("@TNX_CHGAP", "AppUser", "Approver User", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CHGAP", "AppStatus", "Approval Status", 30)
        objMain.objUtilities.AddDateField("@TNX_CHGAP", "ActDate", "Action Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CHGAP", "ActTime", "Action Time", 10)
        objMain.objUtilities.AddAlphaField("@TNX_CHGAP", "Comments", "Comments", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CHGAP", "ESign", "E-Signature", 50)


        '================ CHILD TABLE 3 : TASKS =================
        objMain.objUtilities.CreateTable("TNX_CHGTASK", "Change Implementation Tasks", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_CHGTASK", "TaskCode", "Task Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CHGTASK", "TaskName", "Task Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_CHGTASK", "TaskType", "Task Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CHGTASK", "TaskDesc", "Task Description", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CHGTASK", "Owner", "Owner", 50)
        objMain.objUtilities.AddDateField("@TNX_CHGTASK", "StartDate", "Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_CHGTASK", "DueDate", "Due Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_CHGTASK", "CompDate", "Completion Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CHGTASK", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CHGTASK", "EvdReq", "Evidence Required", 10)
        objMain.objUtilities.AddInteger("@TNX_CHGTASK", "AttachE", "Attachment Entry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_CHGTASK", "VerBy", "Verified By", 50)
        objMain.objUtilities.AddDateField("@TNX_CHGTASK", "VerDate", "Verified Date", SAPbobsCOM.BoFldSubTypes.st_None)


        '================ CHILD TABLE 4 : QA VERIFICATION =================
        objMain.objUtilities.CreateTable("TNX_CHGVER", "Change QA Verification", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_CHGVER", "VerPoint", "Verification Point", 100)
        objMain.objUtilities.AddAlphaField("@TNX_CHGVER", "AccCrit", "Acceptance Criteria", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CHGVER", "Result", "Result", 20)
        objMain.objUtilities.AddAlphaField("@TNX_CHGVER", "Obs", "Observation", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CHGVER", "VerBy", "Verified By", 50)
        objMain.objUtilities.AddDateField("@TNX_CHGVER", "VerDate", "Verified Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_CHGVER", "CAPAReq", "CAPA Required", 10)
        objMain.objUtilities.AddInteger("@TNX_CHGVER", "CAPADE", "CAPA DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 11)


        '================ CHILD TABLE 5 : LINKED DOCUMENTS =================
        objMain.objUtilities.CreateTable("TNX_CHGDOC", "Change Linked Documents", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        objMain.objUtilities.AddAlphaField("@TNX_CHGDOC", "DocType", "Document Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CHGDOC", "ObjType", "Object Type", 50)
        objMain.objUtilities.AddInteger("@TNX_CHGDOC", "DocEntry", "DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaField("@TNX_CHGDOC", "DocNum", "DocNum", 50)
        objMain.objUtilities.AddAlphaField("@TNX_CHGDOC", "LinkType", "Link Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_CHGDOC", "Remarks", "Remarks", 254)
        objMain.objUtilities.CreateTable("TNX_CHGATT", "Corporate Tax Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        '.objUtilities.AddAlphaField("@TNX_ATTACH_C3", "TPA", "Target Path", 254)
        objMain.objUtilities.AddLinkField("@TNX_CHGATT", "TPH", "Target Path", 250, SAPbobsCOM.BoFldSubTypes.st_Link)
        objMain.objUtilities.AddAlphaField("@TNX_CHGATT", "FNM", "File Name", 254)
        objMain.objUtilities.AddAlphaField("@TNX_CHGATT", "FTR", "Free Text", 254)
        objMain.objUtilities.AddDateField("@TNX_CHGATT ", "ATCD", "Attachment Date", SAPbobsCOM.BoFldSubTypes.st_None)




    End Sub


    Private Sub CreateMaterialQualificationMaster()
        ' 1. New Material Request (Header + Attachments + Approval)
        objMain.objUtilities.CreateTable("TNX_MQ_NMR_H", "New Material Request", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_H", "NMRNO", "NMR Number", 30)
        objMain.objUtilities.AddDateField("@TNX_MQ_NMR_H", "REQDATE", "Request Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_H", "REQDEPT", "Requesting Department", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_H", "REQUSER", "Requested By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_H", "MATTYPE", "Material Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_H", "MATNAME", "Proposed Material Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_H", "MATDESC", "Material Description", 250)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_H", "PHARMA", "Pharmacopoeia", 20)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_H", "DOSFORM", "Dosage Form", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_H", "USAGE", "Intended Usage", 250)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_H", "REASON", "Business Justification", 250)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_H", "PRIORITY", "Priority", 20)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_H", "STATUS", "Status", 30)
        objMain.objUtilities.AddAlphaMemoField("@TNX_MQ_NMR_H", "REMARKS", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_MQ_NMR_ATT", "NMR Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_ATT", "DOCTYPE", "Document Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_ATT", "FILENAME", "File Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_ATT", "ATCPATH", "Attachment Path", 250)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_ATT", "UPLOADBY", "Uploaded By", 50)
        objMain.objUtilities.AddDateField("@TNX_MQ_NMR_ATT", "UPLOADDT", "Upload Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("TNX_MQ_NMR_APP", "NMR Approvals", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_APP", "STAGE", "Stage", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_APP", "APPROVER", "Approver", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_NMR_APP", "STATUS", "Status", 20)
        objMain.objUtilities.AddDateField("@TNX_MQ_NMR_APP", "APPDATE", "Approval Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_MQ_NMR_APP", "REMARKS", "Remarks", 254)
    End Sub

    Private Sub MaterialTechnicalEvaluation()

        ' 2. Material Technical Evaluation (Header + Test/Criteria + Approvals)
        objMain.objUtilities.CreateTable("TNX_MQ_MTE_H", "Material Technical Evaluation", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_H", "EVNO", "Evaluation No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_H", "NMRNO", "Base NMR No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_H", "MATNAME", "Material Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_H", "MATTYPE", "Material Type", 30)
        objMain.objUtilities.AddDateField("@TNX_MQ_MTE_H", "EVALDATE", "Evaluation Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_H", "EVALBY", "Evaluated By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_H", "PHARMA", "Pharmacopoeia", 20)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_H", "COAREV", "COA Reviewed", 10)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_H", "MSDSREV", "MSDS Reviewed", 10)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_H", "TDSREV", "TDS Reviewed", 10)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_H", "COMPAT", "Compatibility", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_H", "STABILITY", "Stability Required", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_H", "RESULT", "Result", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_H", "STATUS", "Status", 30)

        objMain.objUtilities.CreateTable("TNX_MQ_MTE_T", "MTE Test Criteria", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_T", "PARAM", "Parameter", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_T", "METHOD", "Method", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_T", "LIMIT", "Acceptance Limit", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_T", "OBSERVED", "Observed Value", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_T", "RESULT", "Result", 20)
        objMain.objUtilities.AddAlphaMemoField("@TNX_MQ_MTE_T", "REMARKS", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_MQ_MTE_APP", "MTE Approvals", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_APP", "STAGE", "Stage", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_APP", "APPROVER", "Approver", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MTE_APP", "STATUS", "Status", 20)
        objMain.objUtilities.AddDateField("@TNX_MQ_MTE_APP", "APPDATE", "Approval Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_MQ_MTE_APP", "REMARKS", "Remarks", 254)
    End Sub

    Private Sub MaterialSpecifications()
        ' 3. Material Specifications (Master + Lines + Revision)
        objMain.objUtilities.CreateTable("TNX_MQ_SPEC_H", "Material Specification", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_H", "Code", "Spec Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_H", "Name", "Spec Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_H", "ITEMCODE", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_H", "ITEMNAME", "Item Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_H", "MATTYPE", "Material Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_H", "SPECNO", "Specification No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_H", "VERSION", "Version", 20)
        objMain.objUtilities.AddDateField("@TNX_MQ_SPEC_H", "EFFDATE", "Effective Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_MQ_SPEC_H", "REVDATE", "Review Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_H", "PHARMA", "Pharmacopoeia", 20)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_H", "STATUS", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_H", "APPROVEDBY", "Approved By", 50)
        objMain.objUtilities.AddDateField("@TNX_MQ_SPEC_H", "APPROVEDDT", "Approved Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("TNX_MQ_SPEC_L", "Specification Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_L", "TESTCODE", "Test Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_L", "TESTNAME", "Test Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_L", "METHOD", "Method", 100)
        objMain.objUtilities.AddFloatField("@TNX_MQ_SPEC_L", "MINVAL", "Min Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_MQ_SPEC_L", "MAXVAL", "Max Value", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_L", "UOM", "UOM", 20)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_L", "LIMITTXT", "Limit Text", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_L", "MANDATORY", "Mandatory", 1)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_L", "RESULTTYPE", "Result Type", 30)

        objMain.objUtilities.CreateTable("TNX_MQ_SPEC_REV", "Specification Revision History", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_REV", "VERSION", "Version", 20)
        objMain.objUtilities.AddDateField("@TNX_MQ_SPEC_REV", "REVDATE", "Revision Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_MQ_SPEC_REV", "REASON", "Reason", 254)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_REV", "CHANGEDBY", "Changed By", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_SPEC_REV", "STATUS", "Status", 30)
    End Sub

    Private Sub TrialSampleManagement()
        ' 4. Trial Sample Management (Header + Lines + QC)
        objMain.objUtilities.CreateTable("TNX_MQ_TRL_H", "Trial Sample Management", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddInteger("@TNX_MQ_TRL_H", "DocNum", "Document Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_H", "TRIALNO", "Trial No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_H", "NMRNO", "Base NMR No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_H", "VENDOR", "Vendor Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_H", "VENDORNM", "Vendor Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_H", "ITEMCODE", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_H", "ITEMNAME", "Item Name", 150)
        objMain.objUtilities.AddFloatField("@TNX_MQ_TRL_H", "QTYREQ", "Sample Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_H", "UOM", "UOM", 20)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_H", "PURPOSE", "Purpose", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_H", "STATUS", "Status", 30)

        objMain.objUtilities.CreateTable("TNX_MQ_TRL_L", "Trial Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_L", "BASEDOC", "Base Doc Type", 30)
        objMain.objUtilities.AddInteger("@TNX_MQ_TRL_L", "BASEENTRY", "Base DocEntry", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddInteger("@TNX_MQ_TRL_L", "BASELINE", "Base Line", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddFloatField("@TNX_MQ_TRL_L", "QTY", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_L", "BATCHNO", "Batch No", 50)
        objMain.objUtilities.AddDateField("@TNX_MQ_TRL_L", "MFGDATE", "Mfg Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_MQ_TRL_L", "EXPDATE", "Expiry Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_L", "COA", "COA Received", 1)

        objMain.objUtilities.CreateTable("TNX_MQ_TRL_QC", "Trial QC Results", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_QC", "SAMPLEID", "Sample ID", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_QC", "TESTNAME", "Test Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_QC", "RESULT", "Result", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_QC", "STATUS", "Status", 20)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_TRL_QC", "TESTBY", "Tested By", 50)
        objMain.objUtilities.AddDateField("@TNX_MQ_TRL_QC", "TESTDATE", "Test Date", SAPbobsCOM.BoFldSubTypes.st_None)
    End Sub

    Private Sub MaterialApproval()

        ' 5. Material Approval (Header + Approval + Approved Vendor mapping)
        objMain.objUtilities.CreateTable("TNX_MQ_MAPR_H", "Material Approval", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddInteger("@TNX_MQ_MAPR_H", "DocNum", "Document Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_H", "APRNO", "Approval No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_H", "NMRNO", "Base NMR No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_H", "EVALNO", "Evaluation No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_H", "SPECNO", "Specification No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_H", "TRIALNO", "Trial No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_H", "ITEMCODE", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_H", "ITEMNAME", "Item Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_H", "MATTYPE", "Material Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_H", "APPROVAL", "Approval Decision", 30)
        objMain.objUtilities.AddDateField("@TNX_MQ_MAPR_H", "EFFDATE", "Effective Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_MQ_MAPR_H", "VALIDUPTO", "Valid Up To", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_H", "STATUS", "Status", 30)

        objMain.objUtilities.CreateTable("TNX_MQ_MAPR_APP", "Material Approval Approvals", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_APP", "STAGE", "Stage", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_APP", "APPROVER", "Approver", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_APP", "DECISION", "Decision", 20)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_APP", "SIGNID", "E-Sign ID", 50)
        objMain.objUtilities.AddDateField("@TNX_MQ_MAPR_APP", "APPDATE", "Approval Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_MQ_MAPR_APP", "REMARKS", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_MQ_MAPR_AVL", "Material Approved Vendors", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_AVL", "VENDOR", "Vendor Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_AVL", "VENDORNM", "Vendor Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_AVL", "APPROVED", "Approved", 1)
        objMain.objUtilities.AddDateField("@TNX_MQ_MAPR_AVL", "VALIDUPTO", "Valid Up To", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_MAPR_AVL", "RISK", "Risk", 20)
    End Sub

    Private Sub MaterialRequalification()
        ' 6. Material Requalification (Header + QC Trend + Approvals)
        objMain.objUtilities.CreateTable("TNX_MQ_REQ_H", "Material Requalification", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddInteger("@TNX_MQ_REQ_H", "DocNum", "Document Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_H", "REQNO", "Requalification No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_H", "ITEMCODE", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_H", "ITEMNAME", "Item Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_H", "VENDOR", "Vendor Code", 50)
        objMain.objUtilities.AddDateField("@TNX_MQ_REQ_H", "LASTAPR", "Last Approval Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_MQ_REQ_H", "VALIDUPTO", "Valid Up To", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_H", "TRIGGER", "Trigger", 50)
        objMain.objUtilities.AddDateField("@TNX_MQ_REQ_H", "REVIEWDT", "Review Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_H", "DECISION", "Decision", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_H", "STATUS", "Status", 30)

        objMain.objUtilities.CreateTable("TNX_MQ_REQ_QC", "Requalification QC Trend", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_QC", "GRPO", "GRPO No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_QC", "BATCHNO", "Batch No", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_QC", "SAMPLEID", "Sample ID", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_QC", "RESULT", "Result", 20)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_QC", "OOS", "OOS", 1)
        objMain.objUtilities.AddAlphaMemoField("@TNX_MQ_REQ_QC", "REMARKS", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_MQ_REQ_APP", "Requalification Approvals", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_APP", "STAGE", "Stage", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_APP", "APPROVER", "Approver", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_REQ_APP", "STATUS", "Status", 20)
        objMain.objUtilities.AddDateField("@TNX_MQ_REQ_APP", "APPDATE", "Approval Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@TNX_MQ_REQ_APP", "REMARKS", "Remarks", 254)
    End Sub

    Private Sub MaterialRiskAssessment()
        ' 7. Material Risk Assessment (Header + Lines + Controls)
        objMain.objUtilities.CreateTable("TNX_MQ_RISK_H", "Material Risk Assessment", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddInteger("@TNX_MQ_RISK_H", "DocNum", "Document Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_RISK_H", "RISKNO", "Risk No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_RISK_H", "ITEMCODE", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_RISK_H", "ITEMNAME", "Item Name", 150)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_RISK_H", "MATTYPE", "Material Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_RISK_H", "VENDOR", "Vendor Code", 50)
        objMain.objUtilities.AddDateField("@TNX_MQ_RISK_H", "ASSESSDT", "Assessment Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_RISK_H", "ASSESSBY", "Assessed By", 50)
        objMain.objUtilities.AddFloatField("@TNX_MQ_RISK_H", "TOTALSCR", "Total Risk Score", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_RISK_H", "RISKCLS", "Risk Class", 20)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_RISK_H", "STATUS", "Status", 30)

        objMain.objUtilities.CreateTable("TNX_MQ_RISK_L", "Risk Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_RISK_L", "FACTOR", "Risk Factor", 100)
        objMain.objUtilities.AddInteger("@TNX_MQ_RISK_L", "IMPACT", "Impact Score", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddInteger("@TNX_MQ_RISK_L", "PROB", "Probability Score", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddInteger("@TNX_MQ_RISK_L", "DETECT", "Detection Score", SAPbobsCOM.BoFldSubTypes.st_None, 5)
        objMain.objUtilities.AddInteger("@TNX_MQ_RISK_L", "RPN", "RPN", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddAlphaMemoField("@TNX_MQ_RISK_L", "REMARKS", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_MQ_RISK_CTRL", "Risk Controls", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_RISK_CTRL", "CONTROL", "Control Measure", 150)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_RISK_CTRL", "OWNER", "Owner", 50)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_RISK_CTRL", "FREQ", "Frequency", 30)
        objMain.objUtilities.AddAlphaField("@TNX_MQ_RISK_CTRL", "STATUS", "Status", 30)
        objMain.objUtilities.AddDateField("@TNX_MQ_RISK_CTRL", "DUEDATE", "Due Date", SAPbobsCOM.BoFldSubTypes.st_None)
    End Sub

    Private Sub CreateVendorQualificationMaster()
        ' 1. Vendor Qualification Request (VQR)
        objMain.objUtilities.CreateTable("TNX_VQR_H", "Vendor Qualification Request", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddInteger("@TNX_VQR_H", "DocNum", "Document Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "VendName", "Proposed Vendor Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "VendType", "Vendor Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "Country", "Country", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "CPerson", "Contact Person", 100)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "Email", "Email", 100)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "Phone", "Phone", 30)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "MfgSite", "Manufacturing Site", 200)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "GMPCert", "GMP Certified", 1)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "ISOCert", "ISO Certified", 1)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "FApproved", "FDA Approved", 1)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "RiskLevel", "Risk Level", 20)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaMemoField("@TNX_VQR_H", "Remarks", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "BPCode", "Vendor BP Code", 20)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_H", "ReqBy", "Requested By", 50)
        objMain.objUtilities.AddDateField("@TNX_VQR_H", "ReqDate", "Request Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("TNX_VQR_DOC", "VQR Documents", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_DOC", "DocType", "Document Type", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_DOC", "DocNo", "Document No", 50)
        objMain.objUtilities.AddDateField("@TNX_VQR_DOC", "IssueDate", "Issue Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_VQR_DOC", "ExpiryDate", "Expiry Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_DOC", "AttachPath", "Attachment Link", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_DOC", "Verified", "Verified", 1)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_DOC", "VerifiedBy", "Verified By", 50)

        objMain.objUtilities.CreateTable("TNX_VQR_REV", "VQR Reviews", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_REV", "Dept", "Dept", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_REV", "Reviewer", "Reviewer", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VQR_REV", "Decision", "Decision", 20)
        objMain.objUtilities.AddAlphaMemoField("@TNX_VQR_REV", "Remarks", "Remarks", 254)
        objMain.objUtilities.AddDateField("@TNX_VQR_REV", "RDate", "Review Date", SAPbobsCOM.BoFldSubTypes.st_None)
    End Sub

    Private Sub VendorAudit()
        ' 2. Vendor Audit (Audit header + checks + observations + CAPA)
        objMain.objUtilities.CreateTable("TNX_VAUD_H", "Vendor Audit", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddInteger("@TNX_VAUD_H", "DocNum", "Document Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_H", "VQRNo", "VQR No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_H", "BPCode", "Vendor Code", 20)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_H", "BPName", "Vendor Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_H", "AuType", "Audit Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_H", "AuMode", "Audit Mode", 30)
        objMain.objUtilities.AddDateField("@TNX_VAUD_H", "AuDate", "Audit Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_H", "Auditor", "Auditor", 50)
        objMain.objUtilities.AddFloatField("@TNX_VAUD_H", "Score", "Audit Score", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_H", "Result", "Result", 30)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_H", "Status", "Status", 30)
        objMain.objUtilities.AddDateField("@TNX_VAUD_H", "NDate", "Next Audit Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("TNX_VAUD_CHK", "Audit Checkpoints", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_CHK", "Area", "Area", 50)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_CHK", "CheckPoint", "Checkpoint", 200)
        objMain.objUtilities.AddFloatField("@TNX_VAUD_CHK", "MaxScore", "Max Score", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_VAUD_CHK", "ActualScore", "Actual Score", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_CHK", "Compliance", "Compliance", 20)
        objMain.objUtilities.AddAlphaMemoField("@TNX_VAUD_CHK", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_VAUD_OBS", "Audit Observations", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaMemoField("@TNX_VAUD_OBS", "Observation", "Observation", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_OBS", "Severity", "Severity", 20)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_OBS", "ActionReq", "CAPA Required", 1)
        objMain.objUtilities.AddDateField("@TNX_VAUD_OBS", "TargetDate", "Target Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_OBS", "Status", "Status", 30)

        objMain.objUtilities.CreateTable("TNX_VAUD_CAPA", "Audit CAPA", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_CAPA", "CAPANo", "CAPA No", 30)
        objMain.objUtilities.AddAlphaMemoField("@TNX_VAUD_CAPA", "Action", "Action", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_CAPA", "Owner", "Owner", 50)
        objMain.objUtilities.AddDateField("@TNX_VAUD_CAPA", "DueDate", "Due Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_VAUD_CAPA", "CloDate", "Closure Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VAUD_CAPA", "Status", "Status", 30)
    End Sub

    Private Sub ApprovedVendorList()
        ' 3. Approved Vendor List (AVL)
        objMain.objUtilities.CreateTable("TNX_AVL_H", "Approved Vendor List", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_H", "Code", "AVL Code", 30)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_H", "Name", "AVL Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_H", "BPCode", "Vendor Code", 20)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_H", "BPName", "Vendor Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_H", "VQRNo", "VQR No", 30)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_H", "AuditNo", "Audit No", 30)
        objMain.objUtilities.AddDateField("@TNX_AVL_H", "AplDate", "Approval Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_AVL_H", "ValidFrom", "Valid From", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_AVL_H", "ValidTo", "Valid To", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_H", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_H", "RiskLevel", "Risk Level", 20)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_H", "AppdBy", "Approved By", 50)

        objMain.objUtilities.CreateTable("TNX_AVL_MAT", "AVL Materials", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_MAT", "ItemCode", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_MAT", "ItemName", "Item Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_MAT", "ItemType", "Item Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_MAT", "Pharmaco", "Pharmacopoeia", 20)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_MAT", "MfgSite", "Manufacturing Site", 200)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_MAT", "Status", "Status", 30)
        objMain.objUtilities.AddDateField("@TNX_AVL_MAT", "ValidTo", "Valid To", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@TNX_AVL_MAT", "MQty", "MOQ", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddInteger("@TNX_AVL_MAT", "LeadTime", "Lead Time Days", SAPbobsCOM.BoFldSubTypes.st_None, 10)

        objMain.objUtilities.CreateTable("TNX_AVL_HIS", "AVL History", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddDateField("@TNX_AVL_HIS", "ChangeDate", "Change Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_HIS", "OldStatus", "Old Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_HIS", "NewStatus", "New Status", 30)
        objMain.objUtilities.AddAlphaMemoField("@TNX_AVL_HIS", "Reason", "Reason", 254)
        objMain.objUtilities.AddAlphaField("@TNX_AVL_HIS", "ChangedBy", "Changed By", 50)
    End Sub

    Private Sub VendorRequalification()
        ' 4. Vendor Requalification (Header + Documents + Actions)
        objMain.objUtilities.CreateTable("TNX_VREQ_H", "Vendor Requalification", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddInteger("@TNX_VREQ_H", "DocNum", "Document Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_VREQ_H", "BPCode", "Vendor Code", 20)
        objMain.objUtilities.AddAlphaField("@TNX_VREQ_H", "BPName", "Vendor Name", 100)
        objMain.objUtilities.AddAlphaField("@TNX_VREQ_H", "AVLCode", "AVL Code", 30)
        objMain.objUtilities.AddDateField("@TNX_VREQ_H", "LastQualDate", "Last Qual Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_VREQ_H", "ValidTo", "Valid To", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_VREQ_H", "RequalDueDate", "Requal Due", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VREQ_H", "RequalType", "Requal Type", 30)
        objMain.objUtilities.AddAlphaField("@TNX_VREQ_H", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_VREQ_H", "FinalDecision", "Final Decision", 30)
        objMain.objUtilities.AddAlphaField("@TNX_VREQ_H", "ApprovedBy", "Approved By", 50)

        objMain.objUtilities.CreateTable("TNX_VREQ_DOC", "VREQ Documents", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_VREQ_DOC", "DocType", "Doc Type", 50)
        objMain.objUtilities.AddDateField("@TNX_VREQ_DOC", "OldExpiry", "Old Expiry", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_VREQ_DOC", "NewExpiry", "New Expiry", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VREQ_DOC", "Verified", "Verified", 1)
        objMain.objUtilities.AddAlphaMemoField("@TNX_VREQ_DOC", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_VREQ_ACT", "VREQ Actions", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaMemoField("@TNX_VREQ_ACT", "Action", "Action", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VREQ_ACT", "Owner", "Owner", 50)
        objMain.objUtilities.AddDateField("@TNX_VREQ_ACT", "TargetDate", "Target Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VREQ_ACT", "Status", "Status", 30)
    End Sub

    Private Sub VendorRiskAssessment()
        ' 5. Vendor Risk Assessment & Performance Review (core tables for scores and KPIs)
        objMain.objUtilities.CreateTable("TNX_VRA_H", "Vendor Risk Assessment", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddInteger("@TNX_VRA_H", "DocNum", "Document Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_VRA_H", "BPCode", "Vendor Code", 20)
        objMain.objUtilities.AddAlphaField("@TNX_VRA_H", "BPName", "Vendor Name", 100)
        objMain.objUtilities.AddDateField("@TNX_VRA_H", "AssessDate", "Assessment Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VRA_H", "AssessType", "Assessment Type", 30)
        objMain.objUtilities.AddFloatField("@TNX_VRA_H", "TotalScore", "Total Score", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_VRA_H", "RiskLevel", "Risk Level", 20)
        objMain.objUtilities.AddAlphaField("@TNX_VRA_H", "Status", "Status", 30)
        objMain.objUtilities.AddAlphaField("@TNX_VRA_H", "ApprovedBy", "Approved By", 50)

        objMain.objUtilities.CreateTable("TNX_VRA_SCORE", "Vendor Risk Scores", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_VRA_SCORE", "Factor", "Factor", 50)
        objMain.objUtilities.AddFloatField("@TNX_VRA_SCORE", "Weight", "Weight", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_VRA_SCORE", "Score", "Score", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_VRA_SCORE", "WScore", "Weighted Score", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaMemoField("@TNX_VRA_SCORE", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_VRA_REC", "Vendor Recommendations", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaMemoField("@TNX_VRA_REC", "Recommendation", "Recommendation", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VRA_REC", "ActionType", "Action Type", 50)
        objMain.objUtilities.AddDateField("@TNX_VRA_REC", "TargetDate", "Target Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VRA_REC", "Status", "Status", 30)
    End Sub

    Private Sub VendorPerformance()
        ' 6. Vendor Performance Review (header + KPI + Actions)
        objMain.objUtilities.CreateTable("TNX_VPR_H", "Vendor Performance Review", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddInteger("@TNX_VPR_H", "DocNum", "Document Number", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddAlphaField("@TNX_VPR_H", "BPCode", "Vendor Code", 20)
        objMain.objUtilities.AddAlphaField("@TNX_VPR_H", "BPName", "Vendor Name", 100)
        objMain.objUtilities.AddDateField("@TNX_VPR_H", "PeriodFrom", "Period From", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@TNX_VPR_H", "PeriodTo", "Period To", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddInteger("@TNX_VPR_H", "TotalPO", "Total PO", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddInteger("@TNX_VPR_H", "TotalGRPO", "Total GRPO", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddInteger("@TNX_VPR_H", "RCount", "Rejection Count", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddInteger("@TNX_VPR_H", "DelayCount", "Delay Count", SAPbobsCOM.BoFldSubTypes.st_None, 10)
        objMain.objUtilities.AddFloatField("@TNX_VPR_H", "FinalScore", "Final Score", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@TNX_VPR_H", "Rating", "Rating", 30)
        objMain.objUtilities.AddAlphaField("@TNX_VPR_H", "Status", "Status", 30)

        objMain.objUtilities.CreateTable("TNX_VPR_KPI", "VPR KPIs", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@TNX_VPR_KPI", "KPI", "KPI", 50)
        objMain.objUtilities.AddFloatField("@TNX_VPR_KPI", "Target", "Target", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_VPR_KPI", "Actual", "Actual", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@TNX_VPR_KPI", "Score", "Score", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaMemoField("@TNX_VPR_KPI", "Remarks", "Remarks", 254)

        objMain.objUtilities.CreateTable("TNX_VPR_ACT", "VPR Actions", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaMemoField("@TNX_VPR_ACT", "Action", "Action", 254)
        objMain.objUtilities.AddAlphaField("@TNX_VPR_ACT", "Owner", "Owner", 50)
        objMain.objUtilities.AddDateField("@TNX_VPR_ACT", "DueDate", "Due Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@TNX_VPR_ACT", "Status", "Status", 30)
    End Sub


End Class







Namespace Finance_Payment

    Public Class clsTable

        Public Sub FieldCreation()
            'MultiBranchAPInvoice()
            'AddFields("OPCH", "MBAPNo", "Multi-Branch AP No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            'AddFields("NNM1", "APSeries", "AP Service Series", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            'AddFields("ORCT", "PayInNo", "In Payment No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            'AddFields("ORCT", "PayOutNo", "Out Payment No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            'AddFields("OJDT", "PayOutNo", "Out Payment No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            'AddFields("OJDT", "PayInNo", "In Payment No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("OJDT", "IntRecNo", "Internal Reconciliation No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("OJDT", "IntRecEntry", "Internal Reconciliation Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("OUSR", "PayWithReco", "Payment with Reconiciliation", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("ORCT", "RecRem", "Rec Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 100)
            'In_Payment()
            'Out_Payment()
            Internal_Reconciliation()
            AddFields("ORCT", "Select", "Multi Branch Selection", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("ORCT", "JENo", "Journal Entry No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("RCT4", "Branch", "Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, , , , , {"-, "})
            AddFields("RCT4", "AcctCode", "GL Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 15,,,,,, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone) 'SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulChartOfAccounts
            AddFields("RCT4", "AcctName", "GL Account Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("OJDT", "TransId", "Auto Post JE No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 15,,,,,, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulJournalEntries)
            AddFields("OJDT", "OEntry", "Outgoing Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 15,,,,,, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulVendorPayments)
            AddFields("OJDT", "IEntry", "Incoming Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 15,,,,,, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulIncomingPayments)
            AddFields("JDT1", "TBranch", "Transaction Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("OJDT", "ReconNum", "Reconciliation Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("RCT4", "Source", "Source", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "2", , {"1,Business Partner", "2,G/L Account"})
        End Sub

#Region "Document Data Creation"

        Private Sub MultiBranchAPInvoice()
            AddTables("MIPL_OAPI", "AP Service Inv Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("MIPL_API1", "AP Service Inv Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            AddFields("@MIPL_OAPI", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_OAPI", "DueDate", "DocDue Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_OAPI", "PosDate", "Posting Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("@MIPL_API1", "SACCode", "SAC Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "VendorCode", "Vendor Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "VendorName", "Vendor Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_API1", "Desc", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_API1", "GLCode", "GL Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_API1", "GLName", "GL Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_API1", "TaxCode", "Tax Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 40)
            AddFields("@MIPL_API1", "OTaxCode", "O Tax Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 40)
            AddFields("@MIPL_API1", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MIPL_API1", "GTotal", "Gross Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MIPL_API1", "Branch", "Branch Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_API1", "TranEntry", "Transaction Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MIPL_API1", "RefNo", "Header Reference No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "OcrCode", "Cost Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "OcrCode2", "Cost Center 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "OcrCode3", "Cost Center 3", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "OcrCode4", "Cost Center 4", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "OcrCode5", "Cost Center 5", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "LRefNo", "Line Reference No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_API1", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("OPCH", "ymhbpref", "YMH Reference No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("PCH1", "ymhbpref", "YMH Reference No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)

            AddUDO("MIAPSI", "AP Service Invoice", SAPbobsCOM.BoUDOObjType.boud_Document, "MIPL_OAPI", {"MIPL_API1"}, {"DocEntry", "DocNum", "U_PosDate", "U_DocDate", "U_DueDate"}, True, True)
        End Sub

        Private Sub In_Payment()
            AddTables("MI_ORCT", "MIPL In Payment Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("MI_RCT1", "MIPL In Payment Lines 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("MI_RCT2", "MIPL In Payment Lines 2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("MI_RCT3", "MIPL In Payment Lines 3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            AddFields("@MI_ORCT", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_ORCT", "PayDate", "Payment Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_ORCT", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ORCT", "TotalFC", "Total FC", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ORCT", "TrsfrAcct", "Transfer Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MI_ORCT", "TrsfrDate", "Transfer Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_ORCT", "TrsfrRef", "Transfer Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, 27)
            AddFields("@MI_ORCT", "TrsfrSum", "Transfer Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ORCT", "CashAcct", "Cash Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MI_ORCT", "CashSum", "Cash Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ORCT", "CheckAcct", "Cheque Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MI_ORCT", "CheckSum", "Cheque Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ORCT", "DiffCurr", "Diff Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MI_ORCT", "CreditSum", "Credit Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ORCT", "IncomNo", "Incoming Payment No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ORCT", "JENo", "Journal Entry No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ORCT", "RecoNo", "Reconciliation No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ORCT", "BcgAcct", "Bcg AcctNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ORCT", "BcgSum", "Bcg Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ORCT", "DocCurr", "Document Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MI_ORCT", "CurTotal", "Currency Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ORCT", "BOESum", "BOE Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
            AddFields("@MI_ORCT", "TotFC", "Tot FC", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_ORCT", "FxJENo", "Forex Journal Entry No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ORCT", "CFxJENo", "Curr Forex JE No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ORCT", "FxRecoNo", "Forex Reconciliation No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ORCT", "ActTotal", "Actual Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ORCT", "BalExRate", "Bal Ex Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ORCT", "InTotal", "In Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ORCT", "InSeries", "Incoming Series", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)

            AddFields("@MI_RCT1", "Select", "Selected", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MI_RCT1", "DocNum", "Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_RCT1", "DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT1", "ObjType", "Object Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_RCT1", "Object", "Object", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_RCT1", "DocEntry", "Document Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_RCT1", "DocCur", "Document Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MI_RCT1", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_RCT1", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_RCT1", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_RCT1", "DueDays", "Due Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MI_RCT1", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            'AddFields("@MI_RCT1", "TotalLC", "Total LC", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("@MI_RCT1", "BalDueLC", "Balance LC", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MI_RCT1", "Round", "Rounding Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_RCT1", "BalDue", "Balance Due", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_RCT1", "CashDisc", "Cash Discount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("@MI_RCT1", "PayTotal", "Payment Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_RCT1", "Pay", "Pay", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_RCT1", "BranchId", "Branch Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_RCT1", "BranchNam", "Branch Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_RCT1", "OcrCode1", "Cost Center 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT1", "OcrCode2", "Cost Center 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT1", "OcrCode3", "Cost Center 3", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT1", "OcrCode4", "Cost Center 4", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT1", "OcrCode5", "Cost Center 5", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT1", "RefNo", "Reference No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_RCT1", "TransId", "Transaction ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT1", "TLine", "Tran Line ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT1", "JENo", "Journal Entry No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT1", "RecoNo", "Reconciliation No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT1", "DiscJE", "Discount JE No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT1", "DebCred", "DebCred", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            AddFields("@MI_RCT1", "DiscRecoNo", "Discount ReconNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT1", "distrule", "dist rule", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)


            AddFields("@MI_RCT2", "DueDate", "Due Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_RCT2", "CheckSum", "Rounding Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_RCT2", "CountryCod", "Country Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT2", "BankCode", "Bank Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_RCT2", "Branch", "Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_RCT2", "AcctNum", "Account Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_RCT2", "CheckNum", "Check Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 11)
            AddFields("@MI_RCT2", "Trnsfr", "Transferrable", SAPbobsCOM.BoFieldTypes.db_Alpha, 5)
            AddFields("@MI_RCT2", "IssuedBy", "Originally Issuedby", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT2", "FiscalID", "FiscalID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_RCT2", "GLAcc", "GL Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            AddFields("@MI_RCT3", "CreditSum", "Credit Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_RCT3", "CreditAcct", "Credit Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT3", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT3", "CardNo", "Card Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT3", "Valid", "Validity", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_RCT3", "IDNo", "ID Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT3", "TelNo", "Telephone Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT3", "PayMet", "Payment Method", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_RCT3", "NOP", "Number of Payments", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT3", "FPP", "First Partial Pay", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_RCT3", "AppCode", "Approval Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_RCT3", "TranType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)


            AddUDO("MIORCT", "MIPL In Payment", SAPbobsCOM.BoUDOObjType.boud_Document, "MI_ORCT", {"MI_RCT1", "MI_RCT2", "MI_RCT3"}, {"DocEntry", "DocNum", "U_DocDate"}, True, True)
        End Sub

        Private Sub Out_Payment()
            AddTables("MI_OVPM", "MIPL Out Payment Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("MI_VPM1", "MIPL Out Payment Lines 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("MI_VPM2", "MIPL Out Payment Lines 2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("MI_VPM3", "MIPL Out Payment Lines 3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            AddFields("@MI_OVPM", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_OVPM", "PayDate", "Payment Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_OVPM", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_OVPM", "TotalFC", "Total FC", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_OVPM", "TrsfrAcct", "Transfer Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MI_OVPM", "TrsfrDate", "Transfer Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_OVPM", "TrsfrRef", "Transfer Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, 27)
            AddFields("@MI_OVPM", "TrsfrSum", "Transfer Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_OVPM", "CashAcct", "Cash Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MI_OVPM", "CashSum", "Cash Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_OVPM", "CheckAcct", "Cheque Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@MI_OVPM", "CheckSum", "Cheque Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_OVPM", "DiffCurr", "Diff Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MI_OVPM", "CreditSum", "Credit Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_OVPM", "OutcomNo", "Outgoing Payment No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_OVPM", "JENo", "Journal Entry No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_OVPM", "RecoNo", "Reconciliation No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_OVPM", "BcgAcct", "Bcg AcctNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_OVPM", "BcgSum", "Bcg Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_OVPM", "DocCurr", "Document Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MI_OVPM", "CurTotal", "Currency Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_OVPM", "BOESum", "BOE Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Rate)
            AddFields("@MI_OVPM", "TotFC", "Tot FC", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_OVPM", "FxJENo", "Forex Journal Entry No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_OVPM", "CFxJENo", "Curr Forex JE No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_OVPM", "FxRecoNo", "Forex Reconciliation No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_OVPM", "ActTotal", "Actual Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_OVPM", "BalExRate", "Bal Ex Rate", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            'AddFields("@MI_OVPM", "OutTotal", "Out Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MI_OVPM", "OutSeries", "Outgoing Series", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)

            AddFields("@MI_VPM1", "Select", "Selected", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MI_VPM1", "DocNum", "Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_VPM1", "DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM1", "ObjType", "Object Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_VPM1", "Object", "Object", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_VPM1", "DocEntry", "Document Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_VPM1", "DocCur", "Document Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MI_VPM1", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_VPM1", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_VPM1", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_VPM1", "DueDays", "Due Days", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MI_VPM1", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            'AddFields("@MI_VPM1", "TotalLC", "Total LC", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            'AddFields("@MI_VPM1", "BalDueLC", "Balance LC", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MI_VPM1", "Round", "Rounding Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_VPM1", "BalDue", "Balance Due", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_VPM1", "CashDisc", "Cash Discount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("@MI_VPM1", "PayTotal", "Payment Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_VPM1", "Pay", "Pay", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_VPM1", "BranchId", "Branch Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_VPM1", "BranchNam", "Branch Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_VPM1", "OcrCode1", "Cost Center 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM1", "OcrCode2", "Cost Center 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM1", "OcrCode3", "Cost Center 3", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM1", "OcrCode4", "Cost Center 4", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM1", "OcrCode5", "Cost Center 5", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM1", "RefNo", "Reference No", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_VPM1", "TransId", "Transaction ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM1", "TLine", "Tran Line ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM1", "JENo", "Journal Entry No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM1", "RecoNo", "Reconciliation No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM1", "DiscJE", "Discount JE No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM1", "DebCred", "DebCred", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            AddFields("@MI_VPM1", "DiscRecoNo", "Discount ReconNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM1", "distrule", "dist rule", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)


            AddFields("@MI_VPM2", "DueDate", "Due Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_VPM2", "CheckSum", "Rounding Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_VPM2", "CountryCod", "Country Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM2", "BankCode", "Bank Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_VPM2", "Branch", "Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_VPM2", "AcctNum", "Account Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_VPM2", "CheckNum", "Check Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 11)
            AddFields("@MI_VPM2", "Trnsfr", "Transferrable", SAPbobsCOM.BoFieldTypes.db_Alpha, 5)
            AddFields("@MI_VPM2", "IssuedBy", "Originally Issuedby", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM2", "FiscalID", "FiscalID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_VPM2", "GLAcc", "GL Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)


            AddFields("@MI_VPM3", "CreditSum", "Credit Sum", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_VPM3", "CreditAcct", "Credit Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM3", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM3", "CardNo", "Card Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM3", "Valid", "Validity", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_VPM3", "IDNo", "ID Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM3", "TelNo", "Telephone Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM3", "PayMet", "Payment Method", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_VPM3", "NOP", "Number of Payments", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM3", "FPP", "First Partial Pay", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_VPM3", "AppCode", "Approval Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_VPM3", "TranType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)


            AddUDO("MIOVPM", "MIPL Out Payment", SAPbobsCOM.BoUDOObjType.boud_Document, "MI_OVPM", {"MI_VPM1", "MI_VPM2", "MI_VPM3"}, {"DocEntry", "DocNum", "U_DocDate"}, True, True)
        End Sub

        Private Sub Internal_Reconciliation()
            AddTables("MI_OITR", "Int Reconciliation Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("MI_ITR1", "Int Reconciliation Lines 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("MI_ITR2", "Int Reconciliation Lines 2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)


            AddFields("OADM", "BranchSplit", "Branch Split in Reco Add-on", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "Y", , {"Y,Yes", "N,No"})
            AddFields("@MI_OITR", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_OITR", "PayDate", "Payment Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_OITR", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            'AddFields("@MI_OITR", "RecoNo", "Reconciliation No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            'AddFields("@MI_OITR", "BranchSplit", "Branch Split", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "Y", , {"Y,Yes", "N,No"})
            AddFields("@MI_OITR", "TransId", "Transaction ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_OITR", "RevTransId", "Reverse Transaction ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_OITR", "DocStatus", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "O", , {"O,Open", "C,Canceled"})
            AddFields("OITR", "TransId", "JE No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("OITR", "RevTransId", "Reverse JE No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            AddFields("@MI_ITR1", "Select", "Selected", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MI_ITR1", "TransId", "Transaction ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_ITR1", "TLine", "Tran Line ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ITR1", "Origin", "Origin", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ITR1", "OriginNo", "Origin No", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_ITR1", "DocEntry", "Document Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_ITR1", "Object", "Object", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_ITR1", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_ITR1", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_ITR1", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_ITR1", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ITR1", "BalDue", "Balance Due", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ITR1", "PayTotal", "Payment Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ITR1", "BranchId", "Branch Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_ITR1", "BranchNam", "Branch Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_ITR1", "Memo", "Journal Memo", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_ITR1", "RecoNo", "Reconciliation No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ITR1", "DebCred", "DebCred", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            AddFields("@MI_ITR1", "JENo", "Journal Entry No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ITR1", "CardType", "Card Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            AddFields("@MI_ITR1", "Pay", "Pay", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ITR1", "Ref1", "Reference 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_ITR1", "Ref2", "Reference 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_ITR1", "Ref3", "Reference 3", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_ITR1", "JEAdjust", "JE Adjustment", SAPbobsCOM.BoFieldTypes.db_Alpha, 5)
            AddFields("@MI_ITR1", "TranAdj", "Transaction Adjustment", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", , {"Y,Yes", "N,No"})
            AddFields("@MI_ITR1", "RevJENo", "Revers JE No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            AddFields("@MI_ITR2", "TransId", "Transaction ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_ITR2", "TLine", "Tran Line ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ITR2", "Origin", "Origin", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ITR2", "OriginNo", "Origin No", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_ITR2", "Object", "Object", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MI_ITR2", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_ITR2", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_ITR2", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MI_ITR2", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ITR2", "BalDue", "Balance Due", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ITR2", "PayTotal", "Payment Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@MI_ITR2", "BranchId", "Branch Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_ITR2", "BranchNam", "Branch Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_ITR2", "DebCred", "DebCred", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            AddFields("@MI_ITR2", "TBranchId", "Tran Branch Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MI_ITR2", "TBranchNam", "Tran Branch Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@MI_ITR2", "Row", "Row", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MI_ITR2", "CardType", "Card Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 2)
            AddFields("@MI_ITR2", "RecoNo", "Reconciliation No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            AddUDO("MIOITR", "MIPL Int Reconciliation", SAPbobsCOM.BoUDOObjType.boud_Document, "MI_OITR", {"MI_ITR1", "MI_ITR2"}, {"DocEntry", "DocNum", "U_DocDate"}, True, True)
        End Sub


#End Region

#Region "Master Data Creation"


#End Region

#Region "Table Creation Common Functions"

        Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
            Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
            Try
                oUserTablesMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                'Adding Table
                If Not oUserTablesMD.GetByKey(strTab) Then
                    oUserTablesMD.TableName = strTab
                    oUserTablesMD.TableDescription = strDesc
                    oUserTablesMD.TableType = nType

                    If oUserTablesMD.Add <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription & strTab)
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
                oUserTablesMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub

        Private Sub AddFields(ByVal strTab As String, ByVal strCol As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoFieldTypes,
                             Optional ByVal nEditSize As Integer = 10, Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO,
                              Optional ByVal defaultvalue As String = "", Optional ByVal Yesno As Boolean = False, Optional ByVal Validvalues() As String = Nothing, Optional ByVal LinkedSystemObject As SAPbobsCOM.UDFLinkedSystemObjectTypesEnum = Nothing)
            Dim oUserFieldMD1 As SAPbobsCOM.UserFieldsMD
            oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            Try
                'oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                'If Not (strTab = "OPDN" Or strTab = "OQUT" Or strTab = "OADM" Or strTab = "OPOR" Or strTab = "OWST" Or strTab = "OUSR" Or strTab = "OSRN" Or strTab = "OSPP" Or strTab = "WTR1" Or strTab = "OEDG" Or strTab = "OHEM" Or strTab = "OLCT" Or strTab = "ITM1" Or strTab = "OCRD" Or strTab = "SPP1" Or strTab = "SPP2" Or strTab = "RDR1" Or strTab = "ORDR" Or strTab = "OWHS" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OWDD" Or strTab = "OWOR" Or strTab = "OWTQ" Or strTab = "OMRV" Or strTab = "JDT1" Or strTab = "OIGN" Or strTab = "OCQG") Then
                '    strTab = "@" + strTab
                'End If
                If Not IsColumnExists(strTab, strCol) Then
                    'If Not oUserFieldMD1 Is Nothing Then
                    '    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                    'End If
                    'oUserFieldMD1 = Nothing
                    'oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    oUserFieldMD1.Description = strDesc
                    oUserFieldMD1.Name = strCol
                    oUserFieldMD1.Type = nType
                    oUserFieldMD1.SubType = nSubType
                    oUserFieldMD1.TableName = strTab
                    oUserFieldMD1.EditSize = nEditSize
                    oUserFieldMD1.Mandatory = Mandatory
                    oUserFieldMD1.DefaultValue = defaultvalue

                    If Yesno = True Then
                        oUserFieldMD1.ValidValues.Value = "Y"
                        oUserFieldMD1.ValidValues.Description = "Yes"
                        oUserFieldMD1.ValidValues.Add()
                        oUserFieldMD1.ValidValues.Value = "N"
                        oUserFieldMD1.ValidValues.Description = "No"
                        oUserFieldMD1.ValidValues.Add()
                    End If
                    If LinkedSystemObject <> 0 Then oUserFieldMD1.LinkedSystemObject = LinkedSystemObject

                    Dim split_char() As String
                    If Not Validvalues Is Nothing Then
                        If Validvalues.Length > 0 Then
                            For i = 0 To Validvalues.Length - 1
                                If Trim(Validvalues(i)) = "" Then Continue For
                                split_char = Validvalues(i).Split(",")
                                If split_char.Length <> 2 Then Continue For
                                oUserFieldMD1.ValidValues.Value = split_char(0)
                                oUserFieldMD1.ValidValues.Description = split_char(1)
                                oUserFieldMD1.ValidValues.Add()
                            Next
                        End If
                    End If
                    Dim val As Integer
                    val = oUserFieldMD1.Add
                    If val <> 0 Then
                        objaddon.objapplication.SetStatusBarMessage(objaddon.objcompany.GetLastErrorDescription & " " & strTab & " " & strCol, True)
                    End If
                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                End If
            Catch ex As Exception
                Throw ex
            Finally

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                oUserFieldMD1 = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub

        Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim strSQL As String
            Try
                If objaddon.HANA Then
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE ""TableID"" = '" & Table & "' AND ""AliasID"" = '" & Column & "'"
                Else
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & Table & "' AND AliasID = '" & Column & "'"
                End If

                oRecordSet = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(strSQL)

                If oRecordSet.Fields.Item(0).Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                oRecordSet = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Function

        Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
            Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

            Try
                '// The meta-data object must be initialized with a
                '// regular UserKeys object
                oUserKeysMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

                If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                    '// Set the table name and the key name
                    oUserKeysMD.TableName = strTab
                    oUserKeysMD.KeyName = strKey

                    '// Set the column's alias
                    oUserKeysMD.Elements.ColumnAlias = strColumn
                    oUserKeysMD.Elements.Add()
                    oUserKeysMD.Elements.ColumnAlias = "RentFac"

                    '// Determine whether the key is unique or not
                    oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                    '// Add the key
                    If oUserKeysMD.Add <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription)
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
                oUserKeysMD = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub AddUDO(ByVal strUDO As String, ByVal strUDODesc As String, ByVal nObjectType As SAPbobsCOM.BoUDOObjType, ByVal strTable As String, ByVal childTable() As String, ByVal sFind() As String,
                           Optional ByVal canlog As Boolean = False, Optional ByVal Manageseries As Boolean = False)

            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
            Dim tablecount As Integer = 0
            Try
                oUserObjectMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
                If oUserObjectMD.GetByKey(strUDO) = 0 Then

                    oUserObjectMD.Code = strUDO
                    oUserObjectMD.Name = strUDODesc
                    oUserObjectMD.ObjectType = nObjectType
                    oUserObjectMD.TableName = strTable

                    'oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES : oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES

                    If Manageseries Then oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES Else oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO

                    If canlog Then
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                        oUserObjectMD.LogTableName = "A" + strTable.ToString
                    Else
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                        oUserObjectMD.LogTableName = ""
                    End If

                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.ExtensionName = ""

                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    tablecount = 1
                    If sFind.Length > 0 Then
                        For i = 0 To sFind.Length - 1
                            If Trim(sFind(i)) = "" Then Continue For
                            oUserObjectMD.FindColumns.ColumnAlias = sFind(i)
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.SetCurrentLine(tablecount)
                            tablecount = tablecount + 1
                        Next
                    End If

                    tablecount = 0
                    If Not childTable Is Nothing Then
                        If childTable.Length > 0 Then
                            For i = 0 To childTable.Length - 1
                                If Trim(childTable(i)) = "" Then Continue For
                                oUserObjectMD.ChildTables.SetCurrentLine(tablecount)
                                oUserObjectMD.ChildTables.TableName = childTable(i)
                                oUserObjectMD.ChildTables.Add()
                                tablecount = tablecount + 1
                            Next
                        End If
                    End If

                    If oUserObjectMD.Add() <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription)
                    End If
                End If

            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
                oUserObjectMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try

        End Sub

#End Region

    End Class
End Namespace


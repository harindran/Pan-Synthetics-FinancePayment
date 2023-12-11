Namespace Finance_Payment
    Module ModGlobalVariables
        Public NewLink As String = "-1"
        Public Link_Value As String = "-1"
        Public Link_objtype As String = "-1"
        Public Localization As String = "-1"
        Public bModal As Boolean = False 'Cost Center
        Public CostCenter As String = "-1"
        Public Query As String = "-1"
        Public MainCurr As String = ""
        Public BCGAcct As String = ""
        Public CashAcct As String = ""
        Public RoundAcct As String = ""
        Public pModal As Boolean = False 'Payment Means
        Public OEForm As SAPbouiCOM.Form
        Public objPayDT As New DataTable
        Public Forexgain As String = ""
        Public Forexloss As String = ""
        Public ForexDiff As String = ""
        Public DocumentDate As Date
        Public SumRound As Integer
        Public RateRound As Integer
        Public PayInitDate As Date
        Public PaymentWithReco As String = ""
        Public objPayIntRecoDT As New DataTable
        Public BranchSplitReconciliation As String = ""
        Public TempForm As Boolean = False
    End Module
End Namespace

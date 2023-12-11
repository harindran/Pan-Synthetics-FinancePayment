Namespace Finance_Payment
    Public Class SysAPInvoice
        Public Const Formtype = "141"
        Dim objform As SAPbouiCOM.Form
        Dim FormCount As Integer = 0

        Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            objform = objaddon.objapplication.Forms.Item(FormUID)
            If pVal.BeforeAction Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                End Select
            Else
                Try
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                        Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Try
                                objform = objaddon.objapplication.Forms.GetForm("141", Me.FormCount)
                                Dim oUDFForm As SAPbouiCOM.Form
                                oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                                oUDFForm.Items.Item("U_MBAPNo").Enabled = False
                            Catch ex As Exception
                            End Try
                    End Select
                Catch ex As Exception

                End Try
            End If
        End Sub

        Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID)
                If BusinessObjectInfo.BeforeAction = True Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            Try
                                objform = objaddon.objapplication.Forms.GetForm("141", Me.FormCount)
                                Dim oUDFForm As SAPbouiCOM.Form
                                oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                                oUDFForm.Items.Item("U_MBAPNo").Enabled = False
                            Catch ex As Exception

                            End Try
                    End Select
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            End Try
        End Sub

    End Class
End Namespace
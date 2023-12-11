Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework
Namespace Finance_Payment
    Public Class ClsUDF_JE
        Public Const Formtype = "-392"
        Dim objform As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset
        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                'objmatrix = objform.Items.Item("76").Specific
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            If pVal.ItemUID = "lnkirnum" Then
                                Link_Value = objform.Items.Item("U_IntRecEntry").Specific.String
                                Link_objtype = "MIOITR"
                                Dim activeform As New FrmInternalReconciliation
                                activeform.Show()
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            If (pVal.ItemUID = "U_TransId" Or pVal.ItemUID = "U_IEntry" Or pVal.ItemUID = "U_OEntry" Or pVal.ItemUID = "U_IntRecNo" Or pVal.ItemUID = "U_ReconNum") Then
                                BubbleEvent = False
                            End If

                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Try
                                Dim objlink As SAPbouiCOM.LinkedButton
                                Dim objItem As SAPbouiCOM.Item
                                objItem = objform.Items.Add("lnkreco", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
                                objItem.Left = objform.Items.Item("U_ReconNum").Left - 15
                                objItem.Width = 12
                                objItem.Top = objform.Items.Item("U_ReconNum").Top + 2
                                objItem.Height = 10
                                objlink = objItem.Specific
                                objlink.LinkedObjectType = "321"
                                objlink.Item.LinkTo = "U_ReconNum"
                                objform.Items.Item("U_ReconNum").Enabled = False

                                objItem = objform.Items.Add("lnkirnum", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
                                objItem.Left = objform.Items.Item("U_IntRecEntry").Left - 15
                                objItem.Width = 12
                                objItem.Top = objform.Items.Item("U_IntRecEntry").Top + 2
                                objItem.Height = 10
                                objlink = objItem.Specific
                                objlink.LinkedObjectType = "MIOITR"
                                objlink.Item.LinkTo = "U_IntRecEntry"
                                objform.Items.Item("U_IntRecEntry").Enabled = False
                            Catch ex As Exception

                            End Try

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID)
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If BusinessObjectInfo.BeforeAction = True Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD


                    End Select
                End If

            Catch ex As Exception
                'objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

    End Class
End Namespace


---------------Incoming Payment Series Fetching----------------
Select T1."Series",T1."SeriesName",T1."Remark",T1."BPLId" from ONNM T0 join NNM1 T1 on T0."ObjectCode"=T1."ObjectCode" where T0."ObjectCode"='24' 
and "Indicator"=(Select "Indicator" from OFPR where TO_DATE($["@MI_ORCT"."U_DocDate"],'dd/MM/yyyy') Between "F_RefDate" and "T_RefDate")
and Ifnull("Locked",'')='N';
---------------Outgoing Payment Series Fetching----------------
Select T1."Series",T1."SeriesName",T1."Remark",T1."BPLId" from ONNM T0 join NNM1 T1 on T0."ObjectCode"=T1."ObjectCode" where T0."ObjectCode"='46' 
and "Indicator"=(Select "Indicator" from OFPR where TO_DATE($["@MI_OVPM"."U_DocDate"],'dd/MM/yyyy') Between "F_RefDate" and "T_RefDate")
and Ifnull("Locked",'')='N';
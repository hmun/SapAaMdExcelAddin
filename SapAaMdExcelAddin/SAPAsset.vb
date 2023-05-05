' Copyright 2022 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPAsset

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            log.Debug("New - " & "checking connection")
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            log.Error("New - Exception=" & ex.ToString)
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAsset")
        End Try
    End Sub

    Private Sub addToStrucDic(pArrayName As String, pRfcStructureMetadata As RfcStructureMetadata, ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        If pStrucDic.ContainsKey(pArrayName) Then
            pStrucDic.Remove(pArrayName)
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Private Sub addToFieldDic(pArrayName As String, pRfcStructureMetadata As RfcParameterMetadata, ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata))
        If pFieldDic.ContainsKey(pArrayName) Then
            pFieldDic.Remove(pArrayName)
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Public Sub getMeta_Change(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"GENERALDATA", "GENERALDATAX", "INVENTORY", "INVENTORYX", "POSTINGINFORMATION", "POSTINGINFORMATIONX", "TIMEDEPENDENTDATA", "TIMEDEPENDENTDATAX", "ALLOCATIONS", "ALLOCATIONSX", "ORIGIN", "ORIGINX", "INVESTACCTASSIGNMNT", "INVESTACCTASSIGNMNTX", "NETWORTHVALUATION", "NETWORTHVALUATIONX", "REALESTATE", "REALESTATEX", "INSURANCE", "INSURANCEX", "LEASING", "LEASINGX"}
        Dim aImports As String() = {"COMPANYCODE", "ASSET", "SUBNUMBER", "GROUPASSET"}
        Dim aTables As String() = {"DEPRECIATIONAREAS", "DEPRECIATIONAREASX", "INVESTMENT_SUPPORT", "EXTENSIONIN"}
        Try
            log.Debug("getMeta_Change - " & "creating Function BAPI_FIXEDASSET_CHANGE")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_FIXEDASSET_CHANGE")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_Change - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAsset")
        Finally
            log.Debug("getMeta_GetDetail - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Sub getMeta_Create(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"KEY", "REFERENCE", "GENERALDATA", "GENERALDATAX", "INVENTORY", "INVENTORYX", "POSTINGINFORMATION", "POSTINGINFORMATIONX", "TIMEDEPENDENTDATA", "TIMEDEPENDENTDATAX", "ALLOCATIONS", "ALLOCATIONSX", "ORIGIN", "ORIGINX", "INVESTACCTASSIGNMNT", "INVESTACCTASSIGNMNTX", "NETWORTHVALUATION", "NETWORTHVALUATIONX", "REALESTATE", "REALESTATEX", "INSURANCE", "INSURANCEX", "LEASING", "LEASINGX"}
        Dim aImports As String() = {"CREATESUBNUMBER", "POSTCAP", "CREATEGROUPASSET", "TESTRUN"}
        Dim aTables As String() = {"DEPRECIATIONAREAS", "DEPRECIATIONAREASX", "INVESTMENT_SUPPORT", "EXTENSIONIN"}
        Try
            log.Debug("getMeta_Create - " & "creating Function BAPI_FIXEDASSET_CREATE1")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_FIXEDASSET_CREATE1")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_Create - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAsset")
        Finally
            log.Debug("getMeta_Create - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Sub getMeta_LegacyValuesPost(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"KEY"}
        Dim aImports As String() = {"TESTRUN"}
        Dim aTables As String() = {"CUMULATEDVALUES", "POSTEDVALUES", "POSTINGHEADERS", "TRANSACTIONS", "PROPORTIONALVALUES"}
        Try
            log.Debug("getMeta_LegacyValuesPost - " & "creating Function BAPI_FIXEDASSET_OVRTAKE_POST")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_FIXEDASSET_OVRTAKE_POST")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_LegacyValuesPost - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAsset")
        Finally
            log.Debug("getMeta_LegacyValuesPost - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function Change(pData As TSAP_AssetData, Optional pOKMsg As String = "OK") As String
        Change = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_FIXEDASSET_CHANGE")
            RfcSessionManager.BeginContext(destination)
            Dim oDEPRECIATIONAREAS As IRfcTable = oRfcFunction.GetTable("DEPRECIATIONAREAS")
            Dim oDEPRECIATIONAREASX As IRfcTable = oRfcFunction.GetTable("DEPRECIATIONAREASX")
            Dim oINVESTMENT_SUPPORT As IRfcTable = oRfcFunction.GetTable("INVESTMENT_SUPPORT")
            Dim oEXTENSIONIN As IRfcTable = oRfcFunction.GetTable("EXTENSIONIN")
            oDEPRECIATIONAREAS.Clear()
            oDEPRECIATIONAREASX.Clear()
            oINVESTMENT_SUPPORT.Clear()
            oEXTENSIONIN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the table fields
            pData.aDataDic.to_IRfcTable(pKey:="DEPRECIATIONAREAS", pIRfcTable:=oDEPRECIATIONAREAS)
            pData.aDataDic.to_IRfcTable(pKey:="DEPRECIATIONAREASX", pIRfcTable:=oDEPRECIATIONAREASX)
            pData.aDataDic.to_IRfcTable(pKey:="INVESTMENT_SUPPORT", pIRfcTable:=oINVESTMENT_SUPPORT)
            pData.aDataDic.to_IRfcTable(pKey:="EXTENSIONIN", pIRfcTable:=oEXTENSIONIN)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim sRETURN As IRfcStructure = oRfcFunction.GetStructure("RETURN")
            Dim aErr As Boolean = False
            If sRETURN.GetValue("TYPE") = "E" Or sRETURN.GetValue("TYPE") = "A" Then
                aErr = True
            End If
            Change = ":" & sRETURN.GetValue("MESSAGE")
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            Change = If(Change = "", pOKMsg, If(aErr = False, pOKMsg & Change, "Error" & Change))

        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAsset")
            Change = "Error: Exception in Change"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function Create(pData As TSAP_AssetData, Optional pOKMsg As String = "OK") As String
        Create = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_FIXEDASSET_CREATE1")
            RfcSessionManager.BeginContext(destination)
            Dim oDEPRECIATIONAREAS As IRfcTable = oRfcFunction.GetTable("DEPRECIATIONAREAS")
            Dim oDEPRECIATIONAREASX As IRfcTable = oRfcFunction.GetTable("DEPRECIATIONAREASX")
            Dim oINVESTMENT_SUPPORT As IRfcTable = oRfcFunction.GetTable("INVESTMENT_SUPPORT")
            Dim oEXTENSIONIN As IRfcTable = oRfcFunction.GetTable("EXTENSIONIN")
            oDEPRECIATIONAREAS.Clear()
            oDEPRECIATIONAREASX.Clear()
            oINVESTMENT_SUPPORT.Clear()
            oEXTENSIONIN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the table fields
            pData.aDataDic.to_IRfcTable(pKey:="DEPRECIATIONAREAS", pIRfcTable:=oDEPRECIATIONAREAS)
            pData.aDataDic.to_IRfcTable(pKey:="DEPRECIATIONAREASX", pIRfcTable:=oDEPRECIATIONAREASX)
            pData.aDataDic.to_IRfcTable(pKey:="INVESTMENT_SUPPORT", pIRfcTable:=oINVESTMENT_SUPPORT)
            pData.aDataDic.to_IRfcTable(pKey:="EXTENSIONIN", pIRfcTable:=oEXTENSIONIN)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim sRETURN As IRfcStructure = oRfcFunction.GetStructure("RETURN")
            Dim aErr As Boolean = False
            If sRETURN.GetValue("TYPE") = "E" Or sRETURN.GetValue("TYPE") = "A" Then
                aErr = True
            End If
            Create = ":" & sRETURN.GetValue("MESSAGE") & ";" & CStr(oRfcFunction.GetValue("COMPANYCODE")) & "/" & CStr(oRfcFunction.GetValue("ASSET")) & "/" & CStr(oRfcFunction.GetValue("SUBNUMBER"))
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            Create = If(Create = "", pOKMsg, If(aErr = False, pOKMsg & Create, "Error" & Create))

        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAsset")
            Create = "Error: Exception in Create"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function LegacyValuesPost(pData As TSAP_AssetData, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        LegacyValuesPost = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_FIXEDASSET_OVRTAKE_POST")
            RfcSessionManager.BeginContext(destination)
            Dim oCUMULATEDVALUES As IRfcTable = oRfcFunction.GetTable("CUMULATEDVALUES")
            Dim oPOSTEDVALUES As IRfcTable = oRfcFunction.GetTable("POSTEDVALUES")
            Dim oPOSTINGHEADERS As IRfcTable = oRfcFunction.GetTable("POSTINGHEADERS")
            Dim oTRANSACTIONS As IRfcTable = oRfcFunction.GetTable("TRANSACTIONS")
            Dim oPROPORTIONALVALUES As IRfcTable = oRfcFunction.GetTable("PROPORTIONALVALUES")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oCUMULATEDVALUES.Clear()
            oPOSTEDVALUES.Clear()
            oPOSTINGHEADERS.Clear()
            oTRANSACTIONS.Clear()
            oPROPORTIONALVALUES.Clear()
            oRETURN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' test run
            If pCheck Then
                oRfcFunction.SetValue("TESTRUN", "X")
            End If
            ' set the table fields
            pData.aDataDic.to_IRfcTable(pKey:="CUMULATEDVALUES", pIRfcTable:=oCUMULATEDVALUES)
            pData.aDataDic.to_IRfcTable(pKey:="POSTEDVALUES", pIRfcTable:=oPOSTEDVALUES)
            pData.aDataDic.to_IRfcTable(pKey:="POSTINGHEADERS", pIRfcTable:=oPOSTINGHEADERS)
            pData.aDataDic.to_IRfcTable(pKey:="TRANSACTIONS", pIRfcTable:=oTRANSACTIONS)
            pData.aDataDic.to_IRfcTable(pKey:="PROPORTIONALVALUES", pIRfcTable:=oPROPORTIONALVALUES)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    LegacyValuesPost = LegacyValuesPost & ";" & oRETURN(i).GetValue("MESSAGE")
                    If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "W" Then
                        aErr = True
                    End If
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            LegacyValuesPost = If(LegacyValuesPost = "", pOKMsg, If(aErr = False, pOKMsg & LegacyValuesPost, "Error" & LegacyValuesPost))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAsset")
            LegacyValuesPost = "Error: Exception in LegacyValuesPost"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class

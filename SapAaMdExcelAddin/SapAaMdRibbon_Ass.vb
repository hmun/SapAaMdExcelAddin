' Copyright 2022 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SapAaMdRibbon_Ass
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapAaMdRibbon_Ass getGenParametrs - " & "reading Parameter")
        aWB = Globals.SapAaMdAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid Sap AA Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AA Md")
            getGenParameters = False
            Exit Function
        End Try
        aName = "SAPAaMdAsset"
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid Sap AA Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AA Md")
            getGenParameters = False
            Exit Function
        End If
        i = 2
        pPar = New SAPCommon.TStr
        Do While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
            pPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 4).value), pFORMAT:=CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop
        getGenParameters = True
    End Function

    Private Function getIntParameters(ByRef pIntPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = Globals.SapAaMdAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid Sap AA Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AA Md")
            getIntParameters = False
            Exit Function
        End Try
        i = 2
        pIntPar = New SAPCommon.TStr
        Do
            pIntPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters check - we should know what we are doing
        getIntParameters = True
    End Function

    Public Sub Change(ByRef pSapCon As SapCon)
        Dim aSAPAsset As New SAPAsset(pSapCon)

        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim jMax As UInt64 = 0
        Dim aAssLOff As Integer = If(aIntPar.value("LOFF", "ASS_DATA") <> "", CInt(aIntPar.value("LOFF", "ASS_DATA")), 4)
        Dim aAssWsName As String = If(aIntPar.value("WS", "ASS_DATA") <> "", aIntPar.value("WS", "ASS_DATA"), "Data")
        Dim aAssWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aAssClmnNr As Integer = If(aIntPar.value("COLNR", "DATAASS") <> "", CInt(aIntPar.value("COLNR", "DATAASS")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("RET", "OKMSG") <> "", aIntPar.value("RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapAaMdAddIn.Application.ActiveWorkbook
        Try
            aAssWs = aWB.Worksheets(aAssWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aAssWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Asset Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Asset")
            Exit Sub
        End Try
        parseHeaderLine(aAssWs, jMax, aMsgClmn, aMsgClmnNr, pHdrLine:=aAssLOff - 3)
        Try
            log.Debug("SapAaMdRibbon_Ass.Change - " & "processing data - disabling events, screen update, cursor")
            Globals.SapAaMdAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapAaMdAddIn.Application.EnableEvents = False
            Globals.SapAaMdAddIn.Application.ScreenUpdating = False
            Dim i As UInt64 = aAssLOff + 1
            Dim aKey As String
            Dim aAssItems As New TData(aIntPar)
            Dim aTSAP_AssetData As New TSAP_AssetData(aPar, aIntPar, aSAPAsset, "Change")
            Do
                If Left(CStr(aAssWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    aRetStr = ""
                    ' read DATA
                    aAssItems.ws_parse_line_simple(aAssWs, aAssLOff, i, jMax)
                    If CStr(aAssWs.Cells(i, aAssClmnNr).Value) <> CStr(aAssWs.Cells(i + 1, aAssClmnNr).Value) Then
                        If aTSAP_AssetData.fillHeader(aAssItems) And aTSAP_AssetData.fillData(aAssItems) Then
                            log.Debug("SapAaMdRibbon_Ass.Change - " & "calling aSAPAsset.Change")
                            aRetStr = aSAPAsset.Change(aTSAP_AssetData, pOKMsg:=aOKMsg)
                            log.Debug("SapAaMdRibbon_Ass.Change - " & "aSAPAsset.Change returned, aRetStr=" & aRetStr)
                            For Each aKey In aAssItems.aTDataDic.Keys
                                aAssWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                            Next
                            aAssItems = New TData(aIntPar)
                            aTSAP_AssetData = New TSAP_AssetData(aPar, aIntPar, aSAPAsset, "Change")
                        End If
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aAssWs.Cells(i, 1).value))
            log.Debug("SapAaMdRibbon_Ass.Change - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapAaMdAddIn.Application.EnableEvents = True
            Globals.SapAaMdAddIn.Application.ScreenUpdating = True
            Globals.SapAaMdAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapAaMdAddIn.Application.EnableEvents = True
            Globals.SapAaMdAddIn.Application.ScreenUpdating = True
            Globals.SapAaMdAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapAaMdRibbon_Ass.Change failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Asset AddIn")
            log.Error("SapAaMdRibbon_Ass.Change - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

    End Sub

    Public Sub Create(ByRef pSapCon As SapCon)
        Dim aSAPAsset As New SAPAsset(pSapCon)

        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim jMax As UInt64 = 0
        Dim aAssLOff As Integer = If(aIntPar.value("LOFF", "ASS_DATA") <> "", CInt(aIntPar.value("LOFF", "ASS_DATA")), 4)
        Dim aAssWsName As String = If(aIntPar.value("WS", "ASS_DATA") <> "", aIntPar.value("WS", "ASS_DATA"), "Data")
        Dim aAssWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aAssClmnNr As Integer = If(aIntPar.value("COLNR", "DATAASS") <> "", CInt(aIntPar.value("COLNR", "DATAASS")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("RET", "OKMSG") <> "", aIntPar.value("RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapAaMdAddIn.Application.ActiveWorkbook
        Try
            aAssWs = aWB.Worksheets(aAssWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aAssWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Asset Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Asset")
            Exit Sub
        End Try
        parseHeaderLine(aAssWs, jMax, aMsgClmn, aMsgClmnNr, pHdrLine:=aAssLOff - 3)
        Try
            log.Debug("SapAaMdRibbon_Ass.Create - " & "processing data - disabling events, screen update, cursor")
            Globals.SapAaMdAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapAaMdAddIn.Application.EnableEvents = False
            Globals.SapAaMdAddIn.Application.ScreenUpdating = False
            Dim i As UInt64 = aAssLOff + 1
            Dim aKey As String
            Dim aAssItems As New TData(aIntPar)
            Dim aTSAP_AssetData As New TSAP_AssetData(aPar, aIntPar, aSAPAsset, "Create")
            Do
                If Left(CStr(aAssWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    aRetStr = ""
                    ' read DATA
                    aAssItems.ws_parse_line_simple(aAssWs, aAssLOff, i, jMax)
                    If CStr(aAssWs.Cells(i, aAssClmnNr).Value) <> CStr(aAssWs.Cells(i + 1, aAssClmnNr).Value) Then
                        If aTSAP_AssetData.fillHeader(aAssItems) And aTSAP_AssetData.fillData(aAssItems) Then
                            log.Debug("SapAaMdRibbon_Ass.Create - " & "calling aSAPAsset.Create")
                            aRetStr = aSAPAsset.Create(aTSAP_AssetData, pOKMsg:=aOKMsg)
                            log.Debug("SapAaMdRibbon_Ass.Create - " & "aSAPAsset.Create returned, aRetStr=" & aRetStr)
                            For Each aKey In aAssItems.aTDataDic.Keys
                                aAssWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                            Next
                            aAssItems = New TData(aIntPar)
                            aTSAP_AssetData = New TSAP_AssetData(aPar, aIntPar, aSAPAsset, "Create")
                        End If
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aAssWs.Cells(i, aAssClmnNr).value))
            log.Debug("SapAaMdRibbon_Ass.Create - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapAaMdAddIn.Application.EnableEvents = True
            Globals.SapAaMdAddIn.Application.ScreenUpdating = True
            Globals.SapAaMdAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapAaMdAddIn.Application.EnableEvents = True
            Globals.SapAaMdAddIn.Application.ScreenUpdating = True
            Globals.SapAaMdAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapAaMdRibbon_Ass.Create failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Asset AddIn")
            log.Error("SapAaMdRibbon_Ass.Create - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

    End Sub

    Public Sub LegacyValuesPost(ByRef pSapCon As SapCon, pCheck As Boolean)
        Dim aSAPAsset As New SAPAsset(pSapCon)

        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim jMax As UInt64 = 0
        Dim aLegLOff As Integer = If(aIntPar.value("LOFF", "LEG_DATA") <> "", CInt(aIntPar.value("LOFF", "LEG_DATA")), 4)
        Dim aLegWsName As String = If(aIntPar.value("WS", "LEG_DATA") <> "", aIntPar.value("WS", "LEG_DATA"), "Data")
        Dim aLegWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aAssClmnNr As Integer = If(aIntPar.value("COLNR", "DATAASS") <> "", CInt(aIntPar.value("COLNR", "DATAASS")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("RET", "OKMSG") <> "", aIntPar.value("RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapAaMdAddIn.Application.ActiveWorkbook
        Try
            aLegWs = aWB.Worksheets(aLegWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aLegWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Leget Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Leget")
            Exit Sub
        End Try
        parseHeaderLine(aLegWs, jMax, aMsgClmn, aMsgClmnNr, pHdrLine:=aLegLOff - 3)
        Try
            log.Debug("SapAaMdRibbon_Ass.LegacyValuesPost - " & "processing data - disabling events, screen update, cursor")
            Globals.SapAaMdAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapAaMdAddIn.Application.EnableEvents = False
            Globals.SapAaMdAddIn.Application.ScreenUpdating = False
            Dim i As UInt64 = aLegLOff + 1
            Dim aKey As String
            Dim aLegItems As New TData(aIntPar)
            Dim aTSAP_AssetData As New TSAP_AssetData(aPar, aIntPar, aSAPAsset, "LegacyValuesPost")
            Do
                If Left(CStr(aLegWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    aRetStr = ""
                    ' read DATA
                    aLegItems.ws_parse_line_simple(aLegWs, aLegLOff, i, jMax)
                    If CStr(aLegWs.Cells(i, aAssClmnNr).Value) <> CStr(aLegWs.Cells(i + 1, aAssClmnNr).Value) Then
                        If aTSAP_AssetData.fillHeader(aLegItems) And aTSAP_AssetData.fillData(aLegItems) Then
                            log.Debug("SapAaMdRibbon_Ass.LegacyValuesPost - " & "calling aSAPAsset.LegacyValuesPost")
                            aRetStr = aSAPAsset.LegacyValuesPost(aTSAP_AssetData, pOKMsg:=aOKMsg, pCheck:=pCheck)
                            log.Debug("SapAaMdRibbon_Ass.LegacyValuesPost - " & "aSAPAsset.LegacyValuesPost returned, aRetStr=" & aRetStr)
                            For Each aKey In aLegItems.aTDataDic.Keys
                                aLegWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                            Next
                            aLegItems = New TData(aIntPar)
                            aTSAP_AssetData = New TSAP_AssetData(aPar, aIntPar, aSAPAsset, "LegacyValuesPost")
                        End If
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aLegWs.Cells(i, aAssClmnNr).value))
            log.Debug("SapAaMdRibbon_Ass.LegacyValuesPost - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapAaMdAddIn.Application.EnableEvents = True
            Globals.SapAaMdAddIn.Application.ScreenUpdating = True
            Globals.SapAaMdAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapAaMdAddIn.Application.EnableEvents = True
            Globals.SapAaMdAddIn.Application.ScreenUpdating = True
            Globals.SapAaMdAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapAaMdRibbon_Ass.LegacyValuesPost failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Asset AddIn")
            log.Error("SapAaMdRibbon_Ass.LegacyValuesPost - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

    End Sub

    Private Sub parseHeaderLine(ByRef pWs As Excel.Worksheet, ByRef pMaxJ As Integer, Optional pMsgClmn As String = "", Optional ByRef pMsgClmnNr As Integer = 0, Optional pHdrLine As Integer = 1)
        pMaxJ = 0
        Do
            pMaxJ += 1
            If Not String.IsNullOrEmpty(pMsgClmn) And CStr(pWs.Cells(pHdrLine, pMaxJ).value) = pMsgClmn Then
                pMsgClmnNr = pMaxJ
            End If
        Loop While CStr(pWs.Cells(pHdrLine, pMaxJ + 1).value) <> ""
    End Sub

End Class

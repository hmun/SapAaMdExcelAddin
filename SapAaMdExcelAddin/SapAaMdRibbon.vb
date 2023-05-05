Imports Microsoft.Office.Tools.Ribbon
Imports SAPLogon

Public Class SapAaMdRibbon
    Private aSapCon
    Private aSapGeneral
    Private aTlPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Sub SapAaMdRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        aSapGeneral = New SapGeneral
    End Sub
    Private Function checkCon() As Integer
        Dim aSapConRet As Integer
        Dim aSapVersionRet As Integer
        checkCon = False
        log.Debug("checkCon - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            Exit Function
        End If
        log.Debug("checkCon - " & "checking Connection")
        aSapConRet = 0
        If aSapCon Is Nothing Then
            Try
                aSapCon = New SAPCon()
            Catch ex As SystemException
                log.Warn("checkCon-New SapCon - )" & ex.ToString)
            End Try
        End If
        Try
            aSapConRet = aSapCon.checkCon()
        Catch ex As SystemException
            log.Warn("checkCon-aSapCon.checkCon - )" & ex.ToString)
        End Try
        If aSapConRet = 0 Then
            log.Debug("checkCon - " & "checking version in SAP")
            Try
                aSapVersionRet = aSapGeneral.checkVersionInSAP(aSapCon)
            Catch ex As SystemException
                log.Warn("checkCon - )" & ex.ToString)
            End Try
            log.Debug("checkCon - " & "aSapVersionRet=" & CStr(aSapVersionRet))
            If aSapVersionRet = True Then
                log.Debug("checkCon - " & "checkCon = True")
                checkCon = True
            Else
                log.Debug("checkCon - " & "connection check failed")
            End If
        End If
    End Function

    Private Sub ButtonLogoff_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogoff.Click
        log.Debug("ButtonLogoff_Click - " & "starting logoff")
        If Not aSapCon Is Nothing Then
            log.Debug("ButtonLogoff_Click - " & "calling aSapCon.SAPlogoff()")
            aSapCon.SAPlogoff()
            aSapCon = Nothing
        End If
        log.Debug("ButtonLogoff_Click - " & "exit")
    End Sub

    Private Sub ButtonLogon_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogon.Click
        Dim aConRet As Integer

        log.Debug("ButtonLogon_Click - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            log.Debug("ButtonLogon_Click - " & "Version check failed")
            Exit Sub
        End If
        log.Debug("ButtonLogon_Click - " & "creating SapCon")
        If aSapCon Is Nothing Then
            aSapCon = New SapCon()
        End If
        log.Debug("ButtonLogon_Click - " & "calling SapCon.checkCon()")
        aConRet = aSapCon.checkCon()
        If aConRet = 0 Then
            log.Debug("ButtonLogon_Click - " & "connection successfull")
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap LTP")
        Else
            log.Debug("ButtonLogon_Click - " & "connection failed")
            aSapCon = Nothing
        End If
    End Sub

    Private Sub ButtonSapAssetlChange_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapAssetlChange.Click
        Dim aSapAaMdRibbon_Ass As New SapAaMdRibbon_Ass
        If checkCon() = True Then
            aSapAaMdRibbon_Ass.Change(pSapCon:=aSapCon)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonSapAssetlChange_Click")
        End If
    End Sub

    Private Sub ButtonSapAssetCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapAssetCreate.Click
        Dim aSapAaMdRibbon_Ass As New SapAaMdRibbon_Ass
        If checkCon() = True Then
            aSapAaMdRibbon_Ass.Create(pSapCon:=aSapCon)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonSapAssetCreate_Click")
        End If
    End Sub

    Private Sub ButtonSapLegValCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapLegValCheck.Click
        Dim aSapAaMdRibbon_Ass As New SapAaMdRibbon_Ass
        If checkCon() = True Then
            aSapAaMdRibbon_Ass.LegacyValuesPost(pSapCon:=aSapCon, pCheck:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonSapAssetCreate_Click")
        End If
    End Sub

    Private Sub ButtonSapLegValPost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonSapLegValPost.Click
        Dim aSapAaMdRibbon_Ass As New SapAaMdRibbon_Ass
        If checkCon() = True Then
            aSapAaMdRibbon_Ass.LegacyValuesPost(pSapCon:=aSapCon, pCheck:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonSapAssetCreate_Click")
        End If
    End Sub
End Class

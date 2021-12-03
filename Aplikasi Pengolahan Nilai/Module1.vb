Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Module Module1
    Public Conn As OleDbConnection
    Public DA As OleDbDataAdapter
    Public DS As DataSet
    Public CMD As OleDbCommand
    Public DR As OleDbDataReader

    Public cryRpt As New ReportDocument
    Public crtableLogoninfos As New TableLogOnInfos
    Public crtableLogoninfo As New TableLogOnInfo
    Public crConnectionInfo As New ConnectionInfo
    Public CrTables As Tables

    Public Sub SetingLaporan()
        With crConnectionInfo
            .ServerName = (Application.StartupPath.ToString & "\DBNilai.accdb")
            .DatabaseName = (Application.StartupPath.ToString & "\DBNilai.accdb")
            .UserID = ""
            .Password = ""
        End With

        CrTables = cryRpt.Database.Tables
        For Each CrTable In CrTables
            crtableLogoninfo = CrTable.LogOnInfo
            crtableLogoninfo.ConnectionInfo = crConnectionInfo
            CrTable.ApplyLogOnInfo(crtableLogoninfo)
        Next
    End Sub

    Public Sub Koneksi()
        Conn = New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=DBNilai.accdb")
        Conn.Open()
    End Sub
End Module

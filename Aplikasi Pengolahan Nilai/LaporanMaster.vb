Public Class LaporanMaster

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        CRV.ReportSource = Nothing
        cryRpt.Load("laporan matakuliah.rpt")
        Call SetingLaporan()
        CRV.ReportSource = cryRpt
        CRV.RefreshReport()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        CRV.ReportSource = Nothing
        cryRpt.Load("laporan mahasiswa.rpt")
        Call SetingLaporan()
        CRV.ReportSource = cryRpt
        CRV.RefreshReport()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        CRV.ReportSource = Nothing
        cryRpt.Load("laporan dosen.rpt")
        Call SetingLaporan()
        CRV.ReportSource = cryRpt
        CRV.RefreshReport()
    End Sub
End Class
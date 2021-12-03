Imports System.Data.OleDb

Public Class LaporanNilai

    Sub TampilKelas()
        CMD = New OleDbCommand("select distinct kelas from tblmahasiswa", Conn)
        DR = CMD.ExecuteReader
        ComboBox1.Items.Clear()
        Do While DR.Read
            ComboBox1.Items.Add(DR.Item("Kelas"))
        Loop
    End Sub

    Sub TampilKuliah()
        CMD = New OleDbCommand("select distinct matakuliah from tblmtkuliah", Conn)
        DR = CMD.ExecuteReader
        ComboBox2.Items.Clear()
        Do While DR.Read
            ComboBox2.Items.Add(DR.Item("matakuliah"))
        Loop
    End Sub

    Sub TampilMahasiswa()
        CMD = New OleDbCommand("select * from tblmahasiswa", Conn)
        DR = CMD.ExecuteReader
        ComboBox3.Items.Clear()
        Do While DR.Read
            ComboBox3.Items.Add(DR.Item("id_mahasiswa"))
        Loop
    End Sub

    Private Sub LaporanNilai_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TampilKelas()
        Call TampilMahasiswa()
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        CMD = New OleDbCommand("select * from TBLMTKuliah where left(id_mtkuliah,2)='" & Microsoft.VisualBasic.Left(ComboBox1.Text, 2) & "' order by 1", Conn)
        DR = CMD.ExecuteReader
        ComboBox2.Items.Clear()
        Do While DR.Read
            ComboBox2.Items.Add(DR.Item("MataKuliah"))
        Loop
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        CMD = New OleDbCommand("select * from TBLMahasiswa where ID_Mahasiswa='" & ComboBox3.Text & "'", Conn)
        DR = CMD.ExecuteReader
        DR.Read()
        If DR.HasRows Then
            TextBox2.Text = DR.Item("Nama_Mahasiswa")
            TextBox3.Text = DR.Item("Kelas")
        Else
            MsgBox("Data tidak ditemukan")
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        CRV.SelectionFormula = "({TBLMAHASISWA.KELAS})='" & ComboBox1.Text & "' AND ({TBLMTKULIAH.MATAKULIAH})='" & ComboBox2.Text & "'"
        cryRpt.Load("NILAI PER KELAS.rpt")
        Call SetingLaporan()
        CRV.ReportSource = cryRpt
        CRV.RefreshReport()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        CRV.SelectionFormula = "({TBLMAHASISWA.ID_MAHASISWA})='" & ComboBox3.Text & "'"
        cryRpt.Load("NILAI PER SISWA.rpt")
        Call SetingLaporan()
        CRV.ReportSource = cryRpt
        CRV.RefreshReport()
    End Sub
End Class
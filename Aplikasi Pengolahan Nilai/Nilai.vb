Imports System.Data.OleDb

Public Class Nilai

    Sub TampilKelas()
        CMD = New OleDbCommand("select distinct kelas from tblmahasiswa order by 1", Conn)
        DR = CMD.ExecuteReader
        ListBox1.Items.Clear()
        Do While DR.Read
            ListBox1.Items.Add(DR.Item("Kelas"))
        Loop
    End Sub

    Sub TampilsemuaMTKuliah()
        CMD = New OleDbCommand("select * from TBLMTKuliah", Conn)
        DR = CMD.ExecuteReader
        ListBox2.Items.Clear()
        Do While DR.Read
            ListBox2.Items.Add(DR.Item("ID_MTKuliah") & Space(2) & DR.Item("MataKuliah"))
        Loop
    End Sub

    Sub DataBaru()
        'seting awal kondisi grid
        DGV.Columns.Add("Absen", "Absen (15%)")
        DGV.Columns.Add("Tugas", "Tugas (15%)")
        DGV.Columns.Add("UTS", "UTS (30%)")
        DGV.Columns.Add("UAS", "UAS (40%)")

        DGV.Columns(2).Width = 50
        DGV.Columns(3).Width = 50
        DGV.Columns(4).Width = 50
        DGV.Columns(5).Width = 50

        DGV.Columns.Add("Nilai", "Nilai")
        DGV.Columns.Add("Mutu", "Mutu")
        DGV.Columns.Add("Ket", "Keterangan")

        DGV.Columns(0).Width = 100 : DGV.Columns(0).ReadOnly = True
        DGV.Columns(1).Width = 150 : DGV.Columns(1).ReadOnly = True
        DGV.Columns(6).Width = 50 : DGV.Columns(6).ReadOnly = True
        DGV.Columns(7).Width = 50 : DGV.Columns(7).ReadOnly = True
        DGV.Columns(8).Width = 150 : DGV.Columns(8).ReadOnly = True

        For baris As Integer = 0 To DGV.RowCount - 2
            DGV.Rows(baris).Cells(2).Value = 0
            DGV.Rows(baris).Cells(3).Value = 0
            DGV.Rows(baris).Cells(4).Value = 0
            DGV.Rows(baris).Cells(5).Value = 0
            DGV.Rows(baris).Cells(6).Value = 0
        Next
    End Sub

    Private Sub Nilai_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call TampilKelas()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        CMD = New OleDbCommand("select * from TBLMTKuliah where left(id_mtkuliah,2)='" & Microsoft.VisualBasic.Left(ListBox1.Text, 2) & "' order by 1", Conn)
        DR = CMD.ExecuteReader
        ListBox2.Items.Clear()
        Do While DR.Read
            ListBox2.Items.Add(DR.Item("ID_MTKuliah") & Space(2) & DR.Item("MataKuliah"))
        Loop
        ListBox3.Items.Clear()
        DGV.Columns.Clear()
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox2.SelectedIndexChanged
        CMD = New OleDbCommand("select distinct tbldosen.id_dosen,nama_dosen from tbldosen,tbldetaildosen where tbldosen.id_dosen=tbldetaildosen.id_dosen and tbldetaildosen.id_mtkuliah='" & Microsoft.VisualBasic.Left(ListBox2.Text, 4) & "'order by 1", Conn)
        DR = CMD.ExecuteReader
        ListBox3.Items.Clear()
        Do While DR.Read
            ListBox3.Items.Add(DR.Item("Id_Dosen") & Space(2) & DR.Item("nama_Dosen"))
        Loop
        DGV.Columns.Clear()
    End Sub

    Private Sub ListBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox3.SelectedIndexChanged
        DGV.Columns.Clear()
        'cari data nilai
        CMD = New OleDbCommand("select tblmahasiswa.ID_Mahasiswa,Nama_Mahasiswa,tblnilai.Absen,Tugas,UTS,UAS,Nilai,Mutu,Keterangan from TBLMahasiswa,tblnilai where tblmahasiswa.id_mahasiswa=tblnilai.id_mahasiswa and tblmahasiswa.kelas='" & ListBox1.Text & "' and tblnilai.id_mtkuliah='" & Microsoft.VisualBasic.Left(ListBox2.Text, 4) & "' and tblnilai.id_dosen='" & Microsoft.VisualBasic.Left(ListBox3.Text, 3) & "'", Conn)
        DR = CMD.ExecuteReader
        DR.Read()
        If Not DR.HasRows Then
            'jika tidak ada maka tampilkan data 'mentahan'
            DA = New OleDbDataAdapter("select ID_Mahasiswa,Nama_Mahasiswa from TBLMahasiswa where kelas='" & ListBox1.Text & "'", Conn)
            DS = New DataSet
            DA.Fill(DS)
            DGV.DataSource = DS.Tables(0)
            Call DataBaru()

        Else
            'jika data ditemukan maka tampilkan data nilai yang dicari
            DA = New OleDbDataAdapter("select tblmahasiswa.ID_Mahasiswa,Nama_Mahasiswa,tblnilai.Absen,Tugas,UTS,UAS,Nilai,Mutu,Keterangan from TBLMahasiswa,tblnilai where tblmahasiswa.id_mahasiswa=tblnilai.id_mahasiswa and tblmahasiswa.kelas='" & ListBox1.Text & "' and tblnilai.id_mtkuliah='" & Microsoft.VisualBasic.Left(ListBox2.Text, 4) & "' and tblnilai.id_dosen='" & Microsoft.VisualBasic.Left(ListBox3.Text, 3) & "'", Conn)
            DS = New DataSet
            DA.Fill(DS)
            DGV.DataSource = DS.Tables(0)

            DGV.Columns(2).Width = 50
            DGV.Columns(3).Width = 50
            DGV.Columns(4).Width = 50
            DGV.Columns(5).Width = 50

            DGV.Columns(0).Width = 100 : DGV.Columns(0).ReadOnly = True
            DGV.Columns(1).Width = 150 : DGV.Columns(1).ReadOnly = True
            DGV.Columns(6).Width = 50 : DGV.Columns(6).ReadOnly = True
            DGV.Columns(7).Width = 50 : DGV.Columns(7).ReadOnly = True
            DGV.Columns(8).Width = 150 : DGV.Columns(8).ReadOnly = True
        End If
    End Sub

    Private Sub DGV_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV.CellEndEdit
        Try
            'komposisi nilai sesuai prosedur tertentu
            If e.ColumnIndex = 2 Then
                DGV.Rows(e.RowIndex).Cells(6).Value = Val(((DGV.Rows(e.RowIndex).Cells(2).Value) * 15) / 100) + Val(((DGV.Rows(e.RowIndex).Cells(3).Value) * 15) / 100) + Val(((DGV.Rows(e.RowIndex).Cells(4).Value) * 30) / 100) + Val(((DGV.Rows(e.RowIndex).Cells(5).Value) * 40) / 100)
            End If

            If e.ColumnIndex = 3 Then
                DGV.Rows(e.RowIndex).Cells(6).Value = Val(((DGV.Rows(e.RowIndex).Cells(2).Value) * 15) / 100) + Val(((DGV.Rows(e.RowIndex).Cells(3).Value) * 15) / 100) + Val(((DGV.Rows(e.RowIndex).Cells(4).Value) * 30) / 100) + Val(((DGV.Rows(e.RowIndex).Cells(5).Value) * 40) / 100)
            End If

            If e.ColumnIndex = 4 Then
                DGV.Rows(e.RowIndex).Cells(6).Value = Val(((DGV.Rows(e.RowIndex).Cells(2).Value) * 15) / 100) + Val(((DGV.Rows(e.RowIndex).Cells(3).Value) * 15) / 100) + Val(((DGV.Rows(e.RowIndex).Cells(4).Value) * 30) / 100) + Val(((DGV.Rows(e.RowIndex).Cells(5).Value) * 40) / 100)
            End If

            If e.ColumnIndex = 5 Then
                DGV.Rows(e.RowIndex).Cells(6).Value = Val(((DGV.Rows(e.RowIndex).Cells(2).Value) * 15) / 100) + Val(((DGV.Rows(e.RowIndex).Cells(3).Value) * 15) / 100) + Val(((DGV.Rows(e.RowIndex).Cells(4).Value) * 30) / 100) + Val(((DGV.Rows(e.RowIndex).Cells(5).Value) * 40) / 100)
            End If

            'mencari nilai A,B,C,D ATAU E
            If Val(DGV.Rows(e.RowIndex).Cells(6).Value) < 40 Then
                DGV.Rows(e.RowIndex).Cells(7).Value = "E"
            ElseIf Val(DGV.Rows(e.RowIndex).Cells(6).Value) < 60 Then
                DGV.Rows(e.RowIndex).Cells(7).Value = "D"
            ElseIf Val(DGV.Rows(e.RowIndex).Cells(6).Value) < 80 Then
                DGV.Rows(e.RowIndex).Cells(7).Value = "C"
            ElseIf Val(DGV.Rows(e.RowIndex).Cells(6).Value) < 90 Then
                DGV.Rows(e.RowIndex).Cells(7).Value = "B"
            ElseIf Val(DGV.Rows(e.RowIndex).Cells(6).Value) >= 90 Then
                DGV.Rows(e.RowIndex).Cells(7).Value = "A"
            End If

            'mencari nilai LULUS, GAGAL ATAU REMEDIAL
            If DGV.Rows(e.RowIndex).Cells(7).Value = "E" Then
                DGV.Rows(e.RowIndex).Cells(8).Value = "GAGAL"
            ElseIf DGV.Rows(e.RowIndex).Cells(7).Value = "D" Then
                DGV.Rows(e.RowIndex).Cells(8).Value = "MEREDIAL"
            ElseIf DGV.Rows(e.RowIndex).Cells(7).Value = "C" Or DGV.Rows(e.RowIndex).Cells(7).Value = "B" Or DGV.Rows(e.RowIndex).Cells(7).Value = "A" Then
                DGV.Rows(e.RowIndex).Cells(8).Value = "LULUS"
            End If

        Catch ex As Exception
            MsgBox("Harus angka")
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub Bersihkan()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        DGV.Columns.Clear()
    End Sub

    
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'cari data nilai
        CMD = New OleDbCommand("select tblmahasiswa.ID_Mahasiswa,Nama_Mahasiswa,tblnilai.Absen,Tugas,UTS,UAS,Nilai,Mutu,Keterangan from TBLMahasiswa,tblnilai where tblmahasiswa.id_mahasiswa=tblnilai.id_mahasiswa and tblmahasiswa.kelas='" & ListBox1.Text & "' and tblnilai.id_mtkuliah='" & Microsoft.VisualBasic.Left(ListBox2.Text, 4) & "' and tblnilai.id_dosen='" & Microsoft.VisualBasic.Left(ListBox3.Text, 3) & "'", Conn)
        DR = CMD.ExecuteReader
        DR.Read()
        If Not DR.HasRows Then
            'jika belum ada maka simpan sebagai data baru
            For baris As Integer = 0 To DGV.RowCount - 2
                Dim simpan As String = "insert into tblnilai values ('" & DGV.Rows(baris).Cells(0).Value & "','" & Microsoft.VisualBasic.Left(ListBox2.Text, 4) & "','" & Microsoft.VisualBasic.Left(ListBox3.Text, 3) & "','" & DGV.Rows(baris).Cells(2).Value & "','" & DGV.Rows(baris).Cells(3).Value & "','" & DGV.Rows(baris).Cells(4).Value & "','" & DGV.Rows(baris).Cells(5).Value & "','" & DGV.Rows(baris).Cells(6).Value & "','" & DGV.Rows(baris).Cells(7).Value & "','" & DGV.Rows(baris).Cells(8).Value & "')"
                CMD = New OleDbCommand(simpan, Conn)
                CMD.ExecuteNonQuery()
            Next
        Else
            'jika sudah ada maka lakukan update
            For baris As Integer = 0 To DGV.RowCount - 2
                Dim edit As String = "update tblnilai set absen='" & DGV.Rows(baris).Cells(2).Value & "',tugas='" & DGV.Rows(baris).Cells(3).Value & "',uts='" & DGV.Rows(baris).Cells(4).Value & "',uas='" & DGV.Rows(baris).Cells(5).Value & "',nilai='" & DGV.Rows(baris).Cells(6).Value & "',mutu='" & DGV.Rows(baris).Cells(7).Value & "',keterangan='" & DGV.Rows(baris).Cells(8).Value & "' where id_mahasiswa='" & DGV.Rows(baris).Cells(0).Value & "' and id_mtkuliah='" & Microsoft.VisualBasic.Left(ListBox2.Text, 4) & "' and id_dosen='" & Microsoft.VisualBasic.Left(ListBox3.Text, 3) & "'"
                CMD = New OleDbCommand(edit, Conn)
                CMD.ExecuteNonQuery()
            Next
        End If
        Call Bersihkan()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call Bersihkan()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
End Class
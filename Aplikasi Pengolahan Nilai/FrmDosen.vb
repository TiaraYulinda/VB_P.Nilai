Imports System.Data.OleDb

Public Class FrmDosen

    Sub Kosongkan()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox1.Focus()
    End Sub

    Sub DataBaru()
        TextBox2.Clear()
        TextBox2.Focus()
    End Sub

    Sub Tampilgrid()
        DA = New OleDbDataAdapter("select * from TBLDosen", Conn)
        DS = New DataSet
        DA.Fill(DS)
        DGV.DataSource = DS.Tables(0)
        DGV.ReadOnly = True
    End Sub

    'nilai default pilihan = NO
    Sub BersihkanPilihan()
        For baris As Integer = 0 To DGV1.RowCount - 2
            DGV1.Rows(baris).Cells(4).Value = "NO"
        Next
    End Sub

    Sub TampilKuliah()
        On Error Resume Next
        DGV1.Columns.Clear()
        DA = New OleDbDataAdapter("select * from tblmtkuliah order by 1", Conn)
        DS = New DataSet
        DA.Fill(DS)
        DGV1.DataSource = DS.Tables(0)
        DGV1.Columns(0).ReadOnly = True
        DGV1.Columns(1).ReadOnly = True
        DGV1.Columns(2).ReadOnly = True
        DGV1.Columns(3).ReadOnly = True
        DGV1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DGV1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        'tambahkan sebuah kolom baru paling kanan berupa combobox
        Dim cbo As New DataGridViewComboBoxColumn
        DGV1.Columns.Add(cbo)
        cbo.HeaderText = "Pilih [Y/N]"

        'item dalam combobox adalah YES dan NO
        cbo.Items.Add("YES")
        cbo.Items.Add("NO")
        DGV1.Columns(4).Width = 75
        Call BersihkanPilihan()
    End Sub

    Private Sub FrmDosen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call Tampilgrid()
        Call TampilKuliah()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 3
        If e.KeyChar = Chr(13) Then
            CMD = New OleDbCommand("select * from TBLDosen where ID_Dosen='" & TextBox1.Text & "'", Conn)
            DR = CMD.ExecuteReader
            DR.Read()
            If DR.HasRows Then
                TextBox2.Text = DR.Item("Nama_Dosen")
                TextBox2.Focus()
            Else
                Call DataBaru()
            End If
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        TextBox2.MaxLength = 30
        If e.KeyChar = Chr(13) Then
            Button1.Focus()
        End If
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("data belum lengkap")
            Exit Sub
        Else
            'cari id dosen
            CMD = New OleDbCommand("select * from TBLDosen where ID_Dosen='" & TextBox1.Text & "'", Conn)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                'jika datanya tidak ada maka simpan sebagai data baru
                Dim simpan As String = "insert into TBLDosen values('" & UCase(TextBox1.Text) & "','" & TextBox2.Text & "')"
                CMD = New OleDbCommand(simpan, Conn)
                CMD.ExecuteNonQuery()

                'simpan secara berulang kode mata kuliah yang statusnya YES
                For baris As Integer = 0 To DGV1.Rows.Count - 2
                    CMD = New OleDbCommand("select * from TBLdetailDosen where ID_Dosen='" & TextBox1.Text & "' and id_mtkuliah='" & DGV1.Rows(baris).Cells(0).Value & "'", Conn)
                    DR = CMD.ExecuteReader
                    DR.Read()
                    If Not DR.HasRows Then
                        If DGV1.Rows(baris).Cells(4).Value = "YES" Then
                            Dim simpandetail As String = "insert into tbldetaildosen values ('" & TextBox1.Text & "','" & DGV1.Rows(baris).Cells(0).Value & "')"
                            CMD = New OleDbCommand(simpandetail, Conn)
                            CMD.ExecuteNonQuery()
                        End If
                    End If

                Next
                Call BersihkanPilihan()
            Else
                'jika datanya sudah ada maka lakukan update
                Dim edit As String = "update TBLDosen set nama_Dosen='" & TextBox2.Text & "' where ID_Dosen='" & TextBox1.Text & "'"
                CMD = New OleDbCommand(edit, Conn)
                CMD.ExecuteNonQuery()
                'lakukan insert secara berulang (karena data sebelumnya sudah dihapus secara otomatis ketika pilihannya NO)
                For baris As Integer = 0 To DGV1.Rows.Count - 2
                    CMD = New OleDbCommand("select * from TBLdetailDosen where ID_Dosen='" & TextBox1.Text & "' and id_mtkuliah='" & DGV1.Rows(baris).Cells(0).Value & "'", Conn)
                    DR = CMD.ExecuteReader
                    DR.Read()
                    If Not DR.HasRows Then
                        If DGV1.Rows(baris).Cells(4).Value = "YES" Then
                            Dim simpandetail As String = "insert into tbldetaildosen values ('" & TextBox1.Text & "','" & DGV1.Rows(baris).Cells(0).Value & "')"
                            CMD = New OleDbCommand(simpandetail, Conn)
                            CMD.ExecuteNonQuery()
                        End If
                    End If
                Next
                Call BersihkanPilihan()
            End If
            Call Kosongkan()
            Call Tampilgrid()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("kode user harus diisi dulu")
            TextBox1.Focus()
            Exit Sub
        Else
            If MessageBox.Show("hapus data ini...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Dim hapus As String = "delete * from TBLDosen where ID_Dosen='" & TextBox1.Text & "'"
                CMD = New OleDbCommand(hapus, Conn)
                CMD.ExecuteNonQuery()

                Dim hapusdetail As String = "delete * from TBLdetailDosen where ID_Dosen='" & TextBox1.Text & "'"
                CMD = New OleDbCommand(hapusdetail, Conn)
                CMD.ExecuteNonQuery()
                Call Kosongkan()
                Call Tampilgrid()
                Call BersihkanPilihan()
            Else
                Call Kosongkan()
                Call BersihkanPilihan()
            End If
        End If
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call Kosongkan()
        Call BersihkanPilihan()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        CMD = New OleDbCommand("select * from TBLDosen where nama_Dosen like '%" & TextBox4.Text & "%'", Conn)
        DR = CMD.ExecuteReader
        DR.Read()
        If DR.HasRows Then
            DA = New OleDbDataAdapter("select * from TBLDosen where nama_Dosen like '%" & TextBox4.Text & "%'", Conn)
            DS = New DataSet
            DA.Fill(DS)
            DGV.DataSource = DS.Tables(0)
        Else
            MsgBox("Nama Dosen tidak ditemukan")
        End If
    End Sub

    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        On Error Resume Next
        Call BersihkanPilihan()
        TextBox1.Text = DGV.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DGV.Rows(e.RowIndex).Cells(1).Value

        For baris As Integer = 0 To DGV1.Rows.Count - 2
            CMD = New OleDbCommand("select * from tbldetaildosen where id_dosen='" & TextBox1.Text & "' and id_mtkuliah='" & DGV1.Rows(baris).Cells(0).Value & "'", Conn)
            DR = CMD.ExecuteReader
            DR.Read()
            If DR.HasRows Then
                DGV1.Rows(baris).Cells(4).Value = "YES"
            End If
        Next
    End Sub

    Private Sub DGV1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV1.CellEndEdit
        'jika pilihan diubahmenjadi NO, maka hapus data tersebut di tabel detaildosen
        If DGV1.Rows(e.RowIndex).Cells(4).Value = "NO" Then
            Dim hapusdetail As String = "delete * from tbldetaildosen where id_dosen='" & TextBox1.Text & "' and id_mtkuliah='" & DGV1.Rows(e.RowIndex).Cells(0).Value & "'"
            CMD = New OleDbCommand(hapusdetail, Conn)
            CMD.ExecuteNonQuery()
        End If
    End Sub

   


    Private Sub DGV1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV1.CellContentClick

    End Sub

    Private Sub DGV_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGV.CellContentClick

    End Sub
End Class




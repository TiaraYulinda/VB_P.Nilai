Imports System.Data.OleDb

Public Class Mahasiswa

    Sub KodeOtomatis()
        CMD = New OleDbCommand("select id_mahasiswa from TBLMAHASISWA order by id_mahasiswa desc", Conn)
        DR = CMD.ExecuteReader
        DR.Read()
        If Not DR.HasRows Then
            TextBox1.Text = "00001"
        Else
            TextBox1.Text = Format(Microsoft.VisualBasic.Right(DR.Item("id_mahasiswa"), 5) + 1, "00000")
        End If
    End Sub

    Sub Kosongkan()
        Call KodeOtomatis()
        TextBox1.Enabled = False
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        ComboBox1.Text = ""
        TextBox2.Focus()
    End Sub

    Sub DataBaru()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        ComboBox1.Text = ""
        TextBox2.Focus()
    End Sub

    Sub Tampilgrid()
        DA = New OleDbDataAdapter("select * from TBLMAHASISWA", Conn)
        DS = New DataSet
        DA.Fill(DS)
        DGV.DataSource = DS.Tables(0)
        DGV.ReadOnly = True
    End Sub

    Sub TampilKelas()
        CMD = New OleDbCommand("select distinct kelas from TBLMAHASISWA", Conn)
        DR = CMD.ExecuteReader
        ComboBox1.Items.Clear()
        Do While DR.Read
            ComboBox1.Items.Add(DR.Item("kelas"))
        Loop
    End Sub

    Private Sub Mahasiswa_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call Tampilgrid()
        Call TampilKelas()
        Call Kosongkan()

    End Sub

    Sub CariKode()
        CMD = New OleDbCommand("select * from TBLMAHASISWA where ID_Mahasiswa='" & TextBox1.Text & "'", Conn)
        DR = CMD.ExecuteReader
        DR.Read()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 5
        If e.KeyChar = Chr(13) Then
            Call CariKode()
            If DR.HasRows Then
                TextBox2.Text = DR.Item("Nama_Mahasiswa")
                ComboBox1.Text = DR.Item("Kelas")
                TextBox3.Text = DR.Item("Jurusan")
                TextBox2.Focus()
            Else
                Call DataBaru()
            End If
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        TextBox2.MaxLength = 30
        If e.KeyChar = Chr(13) Then
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox1.MaxLength = 5
        ComboBox1.Text = UCase(ComboBox1.Text)
        If Microsoft.VisualBasic.Left(ComboBox1.Text, 2) = "MI" Then
            TextBox3.Text = "MANAJEMEN INFORMATIKA"
        ElseIf Microsoft.VisualBasic.Left(ComboBox1.Text, 2) = "MA" Then
            TextBox3.Text = "MANAJEMEN ADMINISTRASI"
        ElseIf Microsoft.VisualBasic.Left(ComboBox1.Text, 2) = "SK" Then
            TextBox3.Text = "SEKRETARIS"
        ElseIf Microsoft.VisualBasic.Left(ComboBox1.Text, 2) = "AK" Then
            TextBox3.Text = "AKUNTANSI"
        Else
            MsgBox("Jurusan belum terdaftar, seharusnya MI,MA,SK ATAU AK")
        End If
        TextBox3.Enabled = False
    End Sub

    Private Sub combobox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        ComboBox1.MaxLength = 5
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        TextBox3.MaxLength = 30
        If e.KeyChar = Chr(13) Then
            Button1.Focus()
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("data belum lengkap")
            Exit Sub
        Else
            Call CariKode()
            If Not DR.HasRows Then
                Dim simpan As String = "insert into TBLMAHASISWA values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & UCase(ComboBox1.Text) & "','" & TextBox3.Text & "')"
                CMD = New OleDbCommand(simpan, Conn)
                CMD.ExecuteNonQuery()
            Else
                Dim edit As String = "update TBLMAHASISWA set Nama_Mahasiswa='" & TextBox2.Text & "',Kelas='" & UCase(ComboBox1.Text) & "',Jurusan='" & TextBox3.Text & "' where ID_Mahasiswa='" & TextBox1.Text & "'"
                CMD = New OleDbCommand(edit, Conn)
                CMD.ExecuteNonQuery()
            End If
            Call Kosongkan()
            Call Tampilgrid()
            Call TampilKelas()
            Call Kosongkan()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("kode harus diisi dulu")
            TextBox1.Focus()
            Exit Sub
        End If

        Call CariKode()
        If Not DR.HasRows Then
            MsgBox("Kode tidak terdaftar")
            Exit Sub
        End If

        CMD = New OleDbCommand("select distinct id_mahasiswa from tblnilai where id_mahasiswa='" & TextBox1.Text & "'", Conn)
        DR = CMD.ExecuteReader
        DR.Read()
        If DR.HasRows Then
            MsgBox("Data tidak dapat dihapus karena sudah ada dalam transaksi")
            'Button2.PerformClick()
            Exit Sub
        End If

        If MessageBox.Show("hapus data ini...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Dim hapus As String = "delete from TBLMAHASISWA where ID_Mahasiswa='" & TextBox1.Text & "'"
            CMD = New OleDbCommand(hapus, Conn)
            CMD.ExecuteNonQuery()
            Call Kosongkan()
            Call Tampilgrid()
            Call TampilKelas()
        Else
            Call Kosongkan()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call Kosongkan()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        DA = New OleDbDataAdapter("select * from TBLMAHASISWA where Nama_Mahasiswa like '%" & TextBox4.Text & "%'", Conn)
        DS = New DataSet
        DA.Fill(DS)
        DGV.DataSource = DS.Tables(0)
        DGV.ReadOnly = True
    End Sub

    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        On Error Resume Next
        TextBox1.Text = DGV.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DGV.Rows(e.RowIndex).Cells(1).Value
        ComboBox1.Text = DGV.Rows(e.RowIndex).Cells(2).Value
        TextBox3.Text = DGV.Rows(e.RowIndex).Cells(3).Value
    End Sub
End Class


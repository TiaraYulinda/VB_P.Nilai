Imports System.Data.OleDb

Public Class Kelas

    Sub Kosongkan()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox1.Focus()
    End Sub

    Sub DataBaru()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox2.Focus()
    End Sub

    Sub Tampilgrid()
        'Call Koneksi()
        DA = New OleDbDataAdapter("select * from TBLKelas", Conn)
        DS = New DataSet
        DA.Fill(DS)
        DGV.DataSource = DS.Tables(0)
        DGV.ReadOnly = True
    End Sub

    Private Sub Kelas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call Tampilgrid()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 3
        If e.KeyChar = Chr(13) Then
            'Call Koneksi()
            CMD = New OleDbCommand("select * from TBLKelas where ID_Kelas='" & TextBox1.Text & "'", Conn)
            DR = CMD.ExecuteReader
            DR.Read()
            If DR.HasRows Then
                On Error Resume Next
                TextBox2.Text = DR.Item("Kelas")
                TextBox3.Text = DR.Item("Jurusan")
                TextBox2.Focus()
            Else
                Call DataBaru()
            End If
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        TextBox2.MaxLength = 5
        If e.KeyChar = Chr(13) Then
            If Microsoft.VisualBasic.Left(TextBox2.Text, 2) = "MI" Then
                TextBox3.Text = "MANAJEMEN INFORMATIKA"
            ElseIf Microsoft.VisualBasic.Left(TextBox2.Text, 2) = "MA" Then
                TextBox3.Text = "MANAJEMEN ADMINISTRASI"
            ElseIf Microsoft.VisualBasic.Left(TextBox2.Text, 2) = "SK" Then
                TextBox3.Text = "SEKRETARIS"
            ElseIf Microsoft.VisualBasic.Left(TextBox2.Text, 2) = "AK" Then
                TextBox3.Text = "AKUNTANSI"
            Else
                MsgBox("jurusan belum terdaftar, seharusnya MI,MA,SK ATAU AK")
            End If
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        TextBox3.MaxLength = 30
        If e.KeyChar = Chr(13) Then
            Button1.Focus()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
            MsgBox("data belum lengkap")
            Exit Sub
        Else
            'Call Koneksi()
            CMD = New OleDbCommand("select * from TBLKelas where ID_Kelas='" & TextBox1.Text & "'", Conn)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                'Call Koneksi()
                Dim simpan As String = "insert into TBLKelas values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "')"
                CMD = New OleDbCommand(simpan, Conn)
                CMD.ExecuteNonQuery()
            Else
                'Call Koneksi()
                Dim edit As String = "update TBLKelas set Kelas='" & TextBox2.Text & "',Jurusan='" & TextBox3.Text & "' where ID_Kelas='" & TextBox1.Text & "'"
                CMD = New OleDbCommand(edit, Conn)
                CMD.ExecuteNonQuery()
            End If
            Call Kosongkan()
            Call Tampilgrid()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("kode harus diisi dulu")
            TextBox1.Focus()
            Exit Sub
        Else
            If MessageBox.Show("hapus data ini...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                'Call Koneksi()
                Dim hapus As String = "delete from TBLKelas where ID_Kelas='" & TextBox1.Text & "'"
                CMD = New OleDbCommand(hapus, Conn)
                CMD.ExecuteNonQuery()
                Call Kosongkan()
                Call Tampilgrid()
            Else
                Call Kosongkan()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call Kosongkan()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        'Call Koneksi()
        CMD = New OleDbCommand("select * from TBLKelas where Kelas like '%" & TextBox4.Text & "%'", Conn)
        DR = CMD.ExecuteReader
        DR.Read()
        If DR.HasRows Then
            'Call Koneksi()
            DA = New OleDbDataAdapter("select * from TBLKelas where Kelas like '%" & TextBox4.Text & "%'", Conn)
            DS = New DataSet
            DA.Fill(DS)
            DGV.DataSource = DS.Tables(0)
        Else
            MsgBox("Nama user tidak ditemukan")
        End If
    End Sub



    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        On Error Resume Next
        TextBox1.Text = DGV.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DGV.Rows(e.RowIndex).Cells(1).Value
        TextBox3.Text = DGV.Rows(e.RowIndex).Cells(2).Value
    End Sub
End Class

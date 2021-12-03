Imports System.Data.OleDb

Public Class MataKuliah

    Sub Kosongkan()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.Text = ""
        TextBox1.Focus()
    End Sub

    Sub DataBaru()
        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.Text = ""
        TextBox2.Focus()
    End Sub

    Sub Tampilgrid()
        'Call Koneksi()
        DA = New OleDbDataAdapter("select * from TBLMTKULIAH order by 1", Conn)
        DS = New DataSet
        DA.Fill(DS)
        DGV.DataSource = DS.Tables(0)
        DGV.ReadOnly = True
    End Sub


    Private Sub MataKuliah_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call Tampilgrid()
        For smt As Integer = 1 To 6
            ComboBox1.Items.Add(smt)
        Next
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 3
        If e.KeyChar = Chr(13) Then
            'Call Koneksi()
            CMD = New OleDbCommand("select * from TBLMTKULIAH where ID_MTKuliah='" & TextBox1.Text & "'", Conn)
            DR = CMD.ExecuteReader
            DR.Read()
            If DR.HasRows Then
                TextBox2.Text = DR.Item("MataKuliah")
                TextBox3.Text = DR.Item("SKS")
                ComboBox1.Text = DR.Item("Semester")
                TextBox2.Focus()
            Else
                Call DataBaru()
            End If
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        TextBox2.MaxLength = 30
        If e.KeyChar = Chr(13) Then
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        TextBox3.MaxLength = 2
        If e.KeyChar = Chr(13) Then
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub combobox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        ComboBox1.MaxLength = 1
        If e.KeyChar = Chr(13) Then
            Button1.Focus()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("data belum lengkap")
            Exit Sub
        Else
            'Call Koneksi()
            CMD = New OleDbCommand("select * from TBLMTKULIAH where ID_MTKuliah='" & TextBox1.Text & "'", Conn)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                'Call Koneksi()
                Dim simpan As String = "insert into TBLMTKULIAH values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & UCase(ComboBox1.Text) & "')"
                CMD = New OleDbCommand(simpan, Conn)
                CMD.ExecuteNonQuery()
            Else
                'Call Koneksi()
                Dim edit As String = "update TBLMTKULIAH set MataKuliah='" & TextBox2.Text & "',SKS='" & TextBox3.Text & "',Semester='" & UCase(ComboBox1.Text) & "' where ID_MTKuliah='" & TextBox1.Text & "'"
                CMD = New OleDbCommand(edit, Conn)
                CMD.ExecuteNonQuery()
            End If
            Call Kosongkan()
            Call Tampilgrid()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("kode Pelajaran harus diisi dulu")
            TextBox1.Focus()
            Exit Sub
        Else
            If MessageBox.Show("hapus data ini...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                'Call Koneksi()
                Dim hapus As String = "delete from TBLMTKULIAH where ID_MTKuliah='" & TextBox1.Text & "'"
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
        CMD = New OleDbCommand("select * from TBLMTKULIAH where MataKuliah like '%" & TextBox4.Text & "%'", Conn)
        DR = CMD.ExecuteReader
        DR.Read()
        If DR.HasRows Then
            'Call Koneksi()
            DA = New OleDbDataAdapter("select * from TBLMTKULIAH where MataKuliah like '%" & TextBox4.Text & "%'", Conn)
            DS = New DataSet
            DA.Fill(DS)
            DGV.DataSource = DS.Tables(0)
        Else
            MsgBox("Nama Pelajaran tidak ditemukan")
        End If
    End Sub



    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        On Error Resume Next
        TextBox1.Text = DGV.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DGV.Rows(e.RowIndex).Cells(1).Value
        TextBox3.Text = DGV.Rows(e.RowIndex).Cells(2).Value
        ComboBox1.Text = DGV.Rows(e.RowIndex).Cells(3).Value
    End Sub
End Class


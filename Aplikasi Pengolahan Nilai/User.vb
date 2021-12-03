Imports System.Data.OleDb

Public Class User

    Sub KodeOtomatis()
        CMD = New OleDbCommand("select id_user from tbluser order by id_user desc", Conn)
        DR = CMD.ExecuteReader
        DR.Read()
        If Not DR.HasRows Then
            TextBox1.Text = "USR01"
        Else
            TextBox1.Text = "USR" + Format(Microsoft.VisualBasic.Right(DR.Item("id_user"), 2) + 1, "00")
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
        DA = New OleDbDataAdapter("select * from tbluser", Conn)
        ds = New DataSet
        da.Fill(ds)
        DGV.DataSource = ds.Tables(0)
        DGV.ReadOnly = True
    End Sub

    Private Sub User_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Call Tampilgrid()
        Call Kosongkan()
        Call KodeOtomatis()
    End Sub

    Sub CariKode()
        CMD = New OleDbCommand("select * from tbluser where ID_User='" & TextBox1.Text & "'", Conn)
        DR = CMD.ExecuteReader
        DR.Read()
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 5
        If e.KeyChar = Chr(13) Then
            Call CariKode()
            If DR.HasRows Then
                TextBox2.Text = DR.Item("nama_user")
                TextBox3.Text = DR.Item("pwd_user")
                ComboBox1.Text = DR.Item("Status")
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
        TextBox3.MaxLength = 10
        If e.KeyChar = Chr(13) Then
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub combobox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        ComboBox1.MaxLength = 15
        If e.KeyChar = Chr(13) Then
            Button1.Focus()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("data belum lengkap")
            Exit Sub
        Else
            Call CariKode()
            If Not DR.HasRows Then
                Dim simpan As String = "insert into tbluser values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & UCase(ComboBox1.Text) & "')"
                CMD = New OleDbCommand(simpan, Conn)
                CMD.ExecuteNonQuery()
            Else
                Dim edit As String = "update tbluser set nama_user='" & TextBox2.Text & "',pwd_user='" & TextBox3.Text & "',Status='" & UCase(ComboBox1.Text) & "' where ID_User='" & TextBox1.Text & "'"
                CMD = New OleDbCommand(edit, Conn)
                CMD.ExecuteNonQuery()
            End If
            Call Kosongkan()
            Call Tampilgrid()
            Call KodeOtomatis()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("kode user harus diisi dulu")
            TextBox1.Focus()
            Exit Sub
        End If

        Call CariKode()
        If Not DR.HasRows Then
            MsgBox("Kode user tidak terdaftar")
            Exit Sub
        End If

            If MessageBox.Show("hapus data ini...?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Dim hapus As String = "delete from tbluser where ID_User='" & TextBox1.Text & "'"
                CMD = New OleDbCommand(hapus, Conn)
                CMD.ExecuteNonQuery()
                Call Kosongkan()
                Call Tampilgrid()
                Call KodeOtomatis()
            Else
                Call Kosongkan()
                Call KodeOtomatis()
            End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call Kosongkan()
        Call KodeOtomatis()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged  
        DA = New OleDbDataAdapter("select * from tbluser where nama_user like '%" & TextBox4.Text & "%'", Conn)
        DS = New DataSet
        DA.Fill(DS)
        DGV.DataSource = DS.Tables(0)
    End Sub

    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        On Error Resume Next
        TextBox1.Text = DGV.Rows(e.RowIndex).Cells(0).Value
        TextBox2.Text = DGV.Rows(e.RowIndex).Cells(1).Value
        TextBox3.Text = DGV.Rows(e.RowIndex).Cells(2).Value
        ComboBox1.Text = DGV.Rows(e.RowIndex).Cells(3).Value
    End Sub
End Class

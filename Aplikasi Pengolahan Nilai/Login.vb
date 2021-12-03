Imports System.Data.OleDb

Public Class Login


    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            Button1.Focus()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            Call Koneksi()
            CMD = New OleDbCommand("select * from tbluser where nama_user='" & TextBox1.Text & "' and pwd_user='" & TextBox2.Text & "'", Conn)
            DR = CMD.ExecuteReader
            DR.Read()
            If Not DR.HasRows Then
                MsgBox("Login gagal")
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox1.Focus()

            Else
                Me.Visible = False
                MenuUtama.Show()
                MenuUtama.Panel1.Text = DR.Item("id_user")
                MenuUtama.Panel2.Text = DR.Item("nama_user")
                MenuUtama.Panel3.Text = DR.Item("statuS")

                If MenuUtama.Panel3.Text = "USER" And MenuUtama.Panel3.Text = "OPERATOR" Then
                    MenuUtama.Button1.Enabled = False
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        End
    End Sub

End Class
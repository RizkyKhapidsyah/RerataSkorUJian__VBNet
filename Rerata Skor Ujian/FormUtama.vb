Imports System.Windows.Forms


Public Class FormUtama

    Private Sub tombolHitung_Click(sender As Object, e As EventArgs) Handles tombolHitung.Click
        'variabel-variabel untuk memuat skor ujian dan reratanya
        Dim dblSkor1 As Double
        Dim dblSkor2 As Double
        Dim dblSkor3 As Double
        Dim dblRerata As Double
        Dim JudulPesan As String
        Dim Objek As TextBox



        'konstanta-konstanta
        Const BANYAK_SKOR As Integer = 3
        Const dblSKOR_TINGGI As Double = 95.0

        Try
            'menyalin isi TextBox ke variabel
            dblSkor1 = CDbl(teksSkor1.Text)
            dblSkor2 = CDbl(teksSkor2.Text)
            dblSkor3 = CDbl(teksSkor3.Text)

            'menghitung rerata
            dblRerata = (dblSkor1 + dblSkor2 + dblSkor3) / BANYAK_SKOR
            'menampilkan rerata, dibulatkan ke 2 dijit di belakang titik desimal
            labelRerata.Text = dblRerata.ToString("n2")
            'jika rerata tinggi, puji siswa tsb
            If dblRerata > dblSKOR_TINGGI Then
                labelPesan.Text = "Selamat! Anda berhasil!"
            End If
        Catch
            'menampilkan pesan error
            JudulPesan = "Isi Nilai!"
            For Each Objek In GroupBox1.Controls
                If Objek.Text = "" Then
                    MsgBox("Silahkan Masukkan Nilai!", vbExclamation, JudulPesan)
                Else
                    labelPesan.Text = "Silahkan Isi Hanya Data Numberik!"
                End If
            Next
        End Try
    End Sub

    Private Sub tombolBersihkan_Click(sender As Object, e As EventArgs) Handles tombolBersihkan.Click
        teksSkor1.Text = String.Empty
        teksSkor2.Text = String.Empty
        teksSkor3.Text = String.Empty
        labelRerata.Text = String.Empty
        labelPesan.Text = String.Empty
    End Sub

    Private Sub tombolKeluar_Click(sender As Object, e As EventArgs) Handles tombolKeluar.Click
        'Menutup form
        Me.Close()
    End Sub
End Class

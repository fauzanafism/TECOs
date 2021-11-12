Imports System.Math
Public Class Form1

    'Open Button
    Private Sub bOpen_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles bOpen.Click
        Using openFile As New OpenFileDialog
            With openFile
                .Filter = "Comma Separated Value|*.csv|All Files|*.*"
                If .ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
                bOpen.Text = .FileName
            End With
        End Using
    End Sub

    'Save Button
    Private Sub bSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bSave.Click
        Using saveFile As New SaveFileDialog
            With saveFile
                .Filter = "Comma Separated Value|*.csv|All Files|*.*"
                If .ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
                bSave.Text = .FileName
            End With
        End Using
    End Sub

    'Process Button
    Private Sub bProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bProcess.Click
        'Deklarasi bentuk tiap variabel
        Dim fileOpen, fileSave As String
        fileOpen = bOpen.Text
        fileSave = bSave.Text

        Dim bacafile As New System.IO.StreamReader(fileOpen)
        Dim simpanfile As New System.IO.StreamWriter(fileSave)

        Dim bacaisi As String
        Dim isi() As String
        Dim isi_txt As String
        Dim i, j As Integer
        Dim tgl_tanam As Integer

        i = 0

        While Not (bacafile.EndOfStream)
            bacaisi = bacafile.ReadLine
            isi = bacaisi.Split(",")
            i += 1

            ch(i) = isi(0)
            rad(i) = isi(1)
            suhu(i) = isi(2) + Val(inPerubahanSuhu.Text)
            rh(i) = isi(3)
            angin(i) = isi(4)

            jml_data += 1

        End While

        isi_txt = "HST,d-Fase,I-Fase,Fase,Daun,Akar,Batang,Tongkol"
        simpanfile.WriteLine(isi_txt)
        'simpanfile.WriteLine()

        'Mengambil nilai input
        'Tanggal Tanam
        tgl_tanam = Val(inJD.Text)

        'Thermal Unit dan Base Temperature
        tu1 = Val(inTU1.Text)
        tu2 = Val(inTU2.Text)
        tu3 = Val(inTU3.Text)
        tu4 = Val(inTU4.Text)

        tb = Val(inTB.Text)

        'Koefisien Tanaman
        k = Val(inK.Text)
        sla = Val(inSLA.Text)
        rue = Val(inRUE.Text)
        lai(0) = Val(inLAI.Text)


        'Neraca Air
        'kl = Val(inKL.Text) * Val(inKedalamanAkar.Text) / 100
        'tlp = Val(inTLP.Text) * Val(inKedalamanAkar.Text) / 100
        'KAT(0) = Val(inKAT.Text)
        'kemiringan = Val(inKemiringan.Text)

        'Lainnya'
        tu(0) = 0
        hari = 0
        s(0) = 0

        wdaun(0) = 0
        wbatang(0) = 0
        wakar(0) = 0
        wtongkol(0) = 0

        For j = tgl_tanam To jml_data
            If s(hari) >= 1 Then
                GoTo Keluar
            End If

            hari += 1

            ch(hari) = ch(j)
            rad(hari) = rad(j)
            suhu(hari) = suhu(j)
            rh(hari) = rh(j)
            angin(hari) = angin(j)

            tu(hari) = tu(hari - 1) + suhu(hari) - tb

            'Perkembangan
            If s(hari - 1) >= 0.55 Then
                'Fase ke 4
                ds(hari) = 0.45 * ((suhu(hari) - tb) / tu4)
                fase(hari) = "Fase ke-4"
                s(hari) = s(hari - 1) + ds(hari)
                If s(hari) >= 1 Then s(hari) = 1

            ElseIf s(hari - 1) >= 0.27 Then
                'Fase ke 3
                ds(hari) = 0.28 * ((suhu(hari) - tb) / tu3)
                fase(hari) = "Fase ke-3"
                s(hari) = s(hari - 1) + ds(hari)
                If s(hari) >= 0.55 Then s(hari) = 0.55

            ElseIf s(hari - 1) >= 0.04 Then
                'Fase ke 2
                ds(hari) = 0.23 * ((suhu(hari) - tb) / tu2)
                fase(hari) = "Fase ke-2"
                s(hari) = s(hari - 1) + ds(hari)
                If s(hari) >= 0.27 Then s(hari) = 0.27

            Else
                'Fase ke 1
                ds(hari) = 0.04 * ((suhu(hari) - tb) / tu1)
                fase(hari) = "Fase ke-1"
                s(hari) = s(hari - 1) + ds(hari)
                If s(hari) > 0.04 Then s(hari) = 0.04
            End If

            'radiasi intersepsi
            Qint(hari) = rad(hari) * (1 - Math.Exp(-k * lai(hari - 1)))
            dW(hari) = (rue * Qint(hari) * 10) '^ 4) * wdf(hari - 1)
            Q10(hari) = 2 * ((suhu(hari) - 20) / 10)
            'dwa(hari) = (1 - 0.13) * dW(hari)

            'respirasi setiap organ
            rdaun(hari) = 0.03 * wdaun(hari - 1) * Q10(hari)
            rbatang(hari) = 0.015 * wbatang(hari - 1) * Q10(hari)
            rakar(hari) = 0.01 * wakar(hari - 1) * Q10(hari)
            rtongkol(hari) = 0.01 * wtongkol(hari - 1) * Q10(hari)

            'Menghitung pertambahan biomassa aktual
            If s(hari) > 0.55 And s(hari) <= 1 Then
                pdaun(hari) = (-0.1582 * s(hari)) + 0.5381
                pbatang(hari) = (-2.6256 * s(hari)) + 2.4231
                pakar(hari) = (-0.1826 * s(hari)) + 0.1653
                ptongkol(hari) = (0.92165 * s(hari)) + 0.21165
            ElseIf s(hari) > 0.27 And s(hari) <= 0.55 Then
                pdaun(hari) = (1.3 * s(hari)) + 0.09
                pbatang(hari) = (0.694 * s(hari)) + 0.19372
                pakar(hari) = (0.1 * s(hari)) + 0.1
                ptongkol(hari) = 0
            ElseIf s(hari) > 0.04 And s(hari) <= 0.27 Then
                pdaun(hari) = (2.9 * s(hari)) + 0.5711
                pbatang(hari) = 0
                pakar(hari) = (-1.682 * s(hari)) + 0.4289
                ptongkol(hari) = 0
            Else
                pdaun(hari) = 0.44
                pbatang(hari) = 0.31
                pakar(hari) = 0.25
                ptongkol(hari) = 0
            End If

            'Menghitung pertambahan biomassa
            dwdaun(hari) = (pdaun(hari) * dW(hari)) - rdaun(hari)
            dwbatang(hari) = (pbatang(hari) * dW(hari)) - rbatang(hari)
            dwakar(hari) = (pakar(hari) * dW(hari)) - rakar(hari)
            dwtongkol(hari) = (ptongkol(hari) * dW(hari)) - rtongkol(hari)


            'Menghitung biomassa tiap organ
            wdaun(hari) = wdaun(hari - 1) + dwdaun(hari)
            wbatang(hari) = wbatang(hari - 1) + dwbatang(hari)
            wakar(hari) = wakar(hari - 1) + dwakar(hari)
            wtongkol(hari) = wtongkol(hari - 1) + dwtongkol(hari)

            'Menghitung biomassa total
            wtot(hari) = wakar(hari) + wbatang(hari) + wdaun(hari) + wtongkol(hari)

            'Menghitung LAI
            lai(hari) = lai(hari - 1) + (sla * dwdaun(hari))

            isi_txt = hari.ToString + "," + ds(hari).ToString + "," + s(hari).ToString + "," + fase(hari).ToString + "," + wdaun(hari).ToString + "," + wakar(hari).ToString + "," + wbatang(hari).ToString + "," + wtongkol(hari).ToString
            simpanfile.WriteLine(isi_txt)
            'simpanfile.WriteLine()

        Next

Keluar:

        bacafile.Close()
        simpanfile.Close()

        'Menampilkan hasil
        outProduktivitas.Text = Math.Round(wtongkol(hari), 2)
        outUmur.Text = hari

        'MessageBox.Show("Ini isi pesan", "Judul Pesan", MessageBoxButtons.YesNo)
        MessageBox.Show("Fase Perkembangan & Pertumbuhan selesai", "PESAN", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    'Pengaturan tanggal
    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        inJD.Text = DateTimePicker1.Value.DayOfYear
    End Sub

    Private Sub CloseBtn_Click(sender As Object, e As EventArgs) Handles CloseBtn.Click
        Me.Close()
    End Sub

    Private Sub MinimizeBtn_Click(sender As Object, e As EventArgs) Handles MinimizeBtn.Click
        Me.WindowState = WindowState.Minimized
    End Sub

    Private Sub PerubahanSuhu_CheckedChanged(sender As Object, e As EventArgs) Handles PerubahanSuhu.CheckedChanged
        If PerubahanSuhu.Checked = True Then
            'lbCH.Visible = True
            sSUHU.Visible = True
            inPerubahanSuhu.Visible = True
        Else
            'lbCH.Visible = False
            sSUHU.Visible = False
            inPerubahanSuhu.Visible = False
        End If
    End Sub

    Private Sub sSUHU_Scroll(sender As Object, e As ScrollEventArgs) Handles sSUHU.Scroll
        inPerubahanSuhu.Text = sSUHU.Value / 2
    End Sub


End Class


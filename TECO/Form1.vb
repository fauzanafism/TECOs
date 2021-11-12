Public Class Form1

    'Open Button
    Private Sub bOpen_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles bOpen.Click
        Using openFile As New OpenFileDialog
            With openFile
                .Filter = "Comma Separated Value|*.csv|All Files|*.*"
                If .ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
                tOpen.Text = .FileName
            End With
        End Using
    End Sub

    'Save Button
    Private Sub bSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bSave.Click
        Using saveFile As New SaveFileDialog
            With saveFile
                .Filter = "Comma Separated Value|*.csv|All Files|*.*"
                If .ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
                tSave.Text = .FileName
            End With
        End Using
    End Sub

    'Process Button
    Private Sub bProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bProcess.Click
        'Deklarasi bentuk tiap variabel
        Dim dataCSV() As String
        Dim j, jml_baris As Integer
        Dim jml_data As Double
        Dim tgl_tanam As Double
        'Dim perubahan_suhu As Double

        'Mengambil nilai input
        'Tanggal Tanam
        tgl_tanam = Val(inJD.Text)

        'Thermal Unit dan Base Temperature
        tu1 = Val(inTU1.Text)
        tu2 = Val(inTU2.Text)
        tu3 = Val(inTU3.Text)
        tu4 = Val(inTU4.Text)

        tb = Val(inTB.Text)

        'perubahan_suhu = Val(inPerubahanSuhu.Text)

        'Koefisien Tanaman
        k = Val(inK.Text)
        sla = Val(inSLA.Text)
        rue = Val(inRUE.Text)
        lai(0) = Val(inLAI.Text)


        'Neraca Air
        kl = Val(inKL.Text) * Val(inKedalamanAkar.Text) / 100
        tlp = Val(inTLP.Text) * Val(inKedalamanAkar.Text) / 100
        kat(0) = Val(inKAT.Text)
        kemiringan = Val(inKemiringan.Text)

        'Lainnya'
        i = 0
        s(0) = 0
        jml_data = 0
        wdf(0) = 1
        lai(0) = 0.1
        wdaun(0) = 0
        wbatang(0) = 0
        wakar(0) = 0
        wtongkol(0) = 0

        'Buka file
        Using file1 As New IO.StreamReader(tOpen.Text)
            While Not (file1.EndOfStream)

                'Menambah jumlah data terbaca
                jml_data += 1

                'Mengecek data dengan tanggal tanam
                If jml_data < tgl_tanam Then dataCSV = Split(file1.ReadLine, ",") : GoTo lewat

                If s(i) > 1 Then s(i) = 1 : GoTo keluar

                'Menambah hari dan memisahkan data iklim
                i += 1
                dataCSV = Split(file1.ReadLine, ",")

                'Memasukkan data iklim
                ch(i) = Val(dataCSV(0))
                rad(i) = Val(dataCSV(1))
                suhu(i) = Val(dataCSV(2))
                rh(i) = Val(dataCSV(3))
                angin(i) = Val(dataCSV(4))

                'Perubahan suhu (codingan checkbox belum, belum di define)
                'suhu(i) = suhu(i) + perubahan_suhu

                'Menghitung indeks - Fase perkembangan (ds)
                Call perkembangan()

                'Menghitung biomassa - Fase pertumbuhan
                Call pertumbuhan()

                'Menghitung neraca air
                Call neraca_air()

                'Mengitung jumlah baris
                jml_baris = i

lewat:
            End While
        End Using

keluar:
        Using file2 As New System.IO.StreamWriter(tSave.Text)
            file2.WriteLine("HST" & "," & "d-Fase" & "," & "I-Fase" & "," & "Suhu" & "," & "Fase" & "," & "Daun" & "," & "Akar" & "," & "Batang" & "," & "Tongkol" & "," & "Radiasi" & "," & "Intersepsi" & "," & "KAT")

            'Menulis isi file
            For j = tgl_tanam To jml_baris
                file2.WriteLine(j & "," & ds(j) & "," & s(j) & "," & suhu(j) & "," & fase(j) & "," & wdaun(j) & "," & wakar(j) & "," & wbatang(j) & "," & wtongkol(j) & "," & rad(j) & "," & ic(j) & "," & KAT(j))
            Next j
        End Using

        'Menampilkan hasil
        outProduktivitas.Text = Math.Round(wtongkol(jml_baris), 2)
        outUmur.Text = jml_baris

        'MessageBox.Show("Ini isi pesan", "Judul Pesan", MessageBoxButtons.YesNo)
        MessageBox.Show("Fase Perkembangan & Pertumbuhan selesai", "PESAN", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    'Pengaturan tanggal
    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        inJD.Text = DateTimePicker1.Value.DayOfYear
    End Sub

End Class


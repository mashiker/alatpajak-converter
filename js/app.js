document.addEventListener('DOMContentLoaded', function() {
    // ===== Kolom Spreadsheet =====
    const columns = [
        { type: 'text', title: 'Nomor Faktur Pajak', width: 280 },
        { type: 'text', title: 'Bulan Faktur Pajak', width: 220 }
    ];

    // ===== Inisialisasi Jspreadsheet (500 baris) =====
    const table = jspreadsheet(document.getElementById('spreadsheet'), {
        data: [['', '']], 
        columns: columns,
        minDimensions: [2, 500], // diubah menjadi 500 baris sesuai permintaan Jevi
        allowInsertRow: true,
        allowManualInsertRow: true,
        allowInsertColumn: false,
        allowDeleteRow: true,
        wordWrap: true,
    });

    // ===== Tombol Bersihkan =====
    document.getElementById('btn-clear').addEventListener('click', function() {
        if (confirm('Bersihkan semua data?')) {
            table.setData([['', '']]);
        }
    });

    // ===== Tombol Tambah Baris Kustom =====
    document.getElementById('btn-add-rows').addEventListener('click', function() {
        var inputVal = document.getElementById('input-rows').value;
        var count = parseInt(inputVal) || 500;
        if (count <= 0 || count > 10000) {
            Swal.fire({
                title: 'Jumlah tidak valid',
                text: 'Harap masukkan jumlah baris antara 1 sampai 10.000.',
                icon: 'warning',
                confirmButtonText: 'OK'
            });
            return;
        }
        
        try {
            table.insertRow(count);
            Swal.fire({
                title: 'Berhasil',
                text: count + ' baris baru telah ditambahkan ke tabel.',
                icon: 'success',
                timer: 1500,
                showConfirmButton: false
            });
        } catch(e) {
            // Fallback loop jika insertRow(count) tidak didukung
            for (var i = 0; i < count; i++) {
                table.insertRow();
            }
            Swal.fire({
                title: 'Berhasil',
                text: count + ' baris baru telah ditambahkan.',
                icon: 'success',
                timer: 1500,
                showConfirmButton: false
            });
        }
    });

    // ===== Tutorial & Penjelasan Tombol =====
    const showTutorial = function() {
        Swal.fire({
            title: 'Panduan CSV Converter E-Faktur',
            html: `
                <div style="text-align: left; font-size: 14px; line-height: 1.6; color: #4d4d4d; max-height: 400px; overflow-y: auto;">
                    <p>Selamat datang di <strong>CSV Converter E-Faktur</strong>! Alat ini mempermudah rekonsiliasi dan konversi data bukti potong sebelum diunggah ke sistem e-Faktur.</p>
                    
                    <h3 style="font-size: 15px; font-weight: 600; color: #171717; margin-top: 16px; margin-bottom: 8px;">Cara Penggunaan:</h3>
                    <ol style="padding-left: 20px; margin-bottom: 16px;">
                        <li style="margin-bottom: 8px;">
                            <strong>Copy & Paste:</strong> Copy dua kolom data dari Excel Anda (kolom <em>Nomor Faktur</em> dan <em>Bulan Faktur</em>), lalu klik sel pertama (<strong>A1</strong>) di tabel ini dan tekan <strong>Ctrl + V</strong> (atau Cmd + V di Mac).
                        </li>
                        <li style="margin-bottom: 8px;">
                            <strong>Validasi Otomatis (Merah Muda):</strong> Jika ada data yang salah format, sel akan otomatis berwarna merah muda:
                            <ul style="padding-left: 15px; margin-top: 4px; list-style-type: circle;">
                                <li><strong>Nomor Faktur:</strong> Harus tepat 17 digit angka (contoh: <code>01000223000000123</code>).</li>
                                <li><strong>Bulan Faktur:</strong> Harus ejaan bahasa Indonesia lengkap (contoh: <code>Januari</code>, <code>Februari</code>, dst. - case sensitive).</li>
                            </ul>
                        </li>
                        <li style="margin-bottom: 8px;">
                            <strong>Download CSV:</strong> Jika semua data sudah valid (tidak ada sel merah muda), klik tombol <strong>Download CSV</strong> untuk mengunduh file CSV siap upload.
                        </li>
                    </ol>

                    <h3 style="font-size: 15px; font-weight: 600; color: #171717; margin-top: 16px; margin-bottom: 8px;">Penjelasan Tombol & Fitur:</h3>
                    <ul style="padding-left: 20px; margin-bottom: 16px; list-style-type: square;">
                        <li style="margin-bottom: 6px;"><strong>📖 Cara Pakai:</strong> Membuka jendela panduan ini kembali.</li>
                        <li style="margin-bottom: 6px;"><strong>➕ Tambah Baris:</strong> Memasukkan jumlah baris kosong baru di bagian bawah tabel. Ketik jumlahnya di input box sebelah tombol (default 500).</li>
                        <li style="margin-bottom: 6px;"><strong>Bersihkan Tabel:</strong> Menghapus seluruh isi tabel untuk memulai konversi data yang baru.</li>
                        <li style="margin-bottom: 6px;"><strong>Download CSV:</strong> Melakukan validasi total dan menyimpan data sebagai file <code>.csv</code> standar dua kolom.</li>
                    </ul>
                    
                    <div style="background-color: #fafafa; border: 1px solid rgba(0,0,0,0.08); padding: 12px; border-radius: 6px; font-size: 13px; margin-top: 16px;">
                        💡 <strong>Tips:</strong> Secara default, tabel ini dimuat dengan <strong>500 baris kosong</strong> untuk mempercepat pengerjaan data dalam jumlah banyak.
                    </div>
                </div>
            `,
            width: '600px',
            confirmButtonText: 'Saya Mengerti, Mulai Kerja!',
            confirmButtonColor: '#171717'
        });
    };

    // Tampilkan otomatis saat pertama kali dibuka
    if (!localStorage.getItem('efaktur_converter_tutorial_shown')) {
        showTutorial();
        localStorage.setItem('efaktur_converter_tutorial_shown', 'true');
    }

    // Tombol cara pakai manual
    document.getElementById('btn-tutorial').addEventListener('click', showTutorial);

    // ===== Validator & Export =====
    const INDONESIAN_MONTHS = new Set([
        'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
        'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ]);

    document.getElementById('btn-export').addEventListener('click', function() {
        var rawData = table.getData();
        var alertBox = document.getElementById('validation-alert');
        var errors = [];
        var validData = [];
        
        // Reset error highlighting
        var totalRows = table.getRowCount ? table.getRowCount() : rawData.length;
        for (var r = 0; r < totalRows; r++) {
            table.setStyle('A' + (r + 1), 'background-color', '');
            table.setStyle('B' + (r + 1), 'background-color', '');
        }

        rawData.forEach(function(row, index) {
            var noFaktur = String(row[0] || '').trim();
            var bulan = String(row[1] || '').trim();
            var rowNum = index + 1;

            // Skip baris benar-benar kosong
            if (noFaktur === '' && bulan === '') return;

            var rowHasError = false;

            // Validasi Nomor Faktur: 17 digit angka
            if (!/^\d{17}$/.test(noFaktur)) {
                errors.push('Baris ' + rowNum + ': Nomor Faktur Pajak harus 17 digit angka.');
                table.setStyle('A' + rowNum, 'background-color', '#fef2f2');
                rowHasError = true;
            }
            
            // Validasi Bulan: harus nama bulan Indonesia lengkap
            if (!INDONESIAN_MONTHS.has(bulan)) {
                errors.push('Baris ' + rowNum + ': Bulan Faktur Pajak tidak valid (Harus ejaan lengkap, misal: Januari, Februari, dst.).');
                table.setStyle('B' + rowNum, 'background-color', '#fef2f2');
                rowHasError = true;
            }

            if (!rowHasError) {
                validData.push([noFaktur, bulan]);
            }
        });

        // Jika ada error: tampilkan alert, block download
        if (errors.length > 0) {
            var errorList = errors.slice(0, 5).map(function(e) {
                return '<li>' + e + '</li>';
            }).join('');
            if (errors.length > 5) {
                errorList += '<li>...dan ' + (errors.length - 5) + ' lainnya</li>';
            }
            alertBox.innerHTML = '<strong>\u26A0\uFE0F Ditemukan ' + errors.length + ' error. Perbaiki sel berwarna merah muda sebelum download:</strong><ul>' + errorList + '</ul>';
            alertBox.classList.remove('hidden');
            
            // Scroll ke alert box agar user tau ada error
            alertBox.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            return;
        }

        // Jika tidak ada data valid
        if (validData.length === 0) {
            alertBox.innerHTML = 'Tabel masih kosong. Tempel data dari Excel terlebih dahulu.';
            alertBox.classList.remove('hidden');
            return;
        }

        // Sukses: sembunyikan alert
        alertBox.classList.add('hidden');

        // Header CSV sesuai permintaan Jevi
        var csvData = [['Nomor Faktur Pajak', 'Bulan Faktur Pajak']].concat(validData);
        var csvString = Papa.unparse(csvData, { 
            quotes: false, 
            delimiter: ',', 
            newline: '\r\n' 
        });

        // Trigger download
        var blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
        var url = URL.createObjectURL(blob);
        var link = document.createElement('a');
        link.setAttribute('href', url);
        link.setAttribute('download', 'Data_Upload_EFaktur_' + new Date().getTime() + '.csv');
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });

    // Simpan reference table ke window
    window.efakturTable = table;
});

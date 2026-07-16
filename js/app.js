document.addEventListener('DOMContentLoaded', function() {
    // ===== Kolom Spreadsheet =====
    const columns = [
        { type: 'text', title: 'Nomor Faktur Pajak', width: 280 },
        { type: 'text', title: 'Bulan Faktur Pajak', width: 220 }
    ];

    const DEFAULT_ROWS = 500;

    // ===== Helper modal native (tanpa CDN tambahan) =====
    function showModal(title, html, options) {
        options = options || {};
        var existing = document.getElementById('app-modal-overlay');
        if (existing) existing.remove();

        var overlay = document.createElement('div');
        overlay.id = 'app-modal-overlay';
        overlay.className = 'app-modal-overlay';
        overlay.innerHTML = `
            <div class="app-modal" role="dialog" aria-modal="true" aria-labelledby="app-modal-title">
                <button type="button" class="app-modal-close" aria-label="Tutup">×</button>
                <h2 id="app-modal-title">${title}</h2>
                <div class="app-modal-content">${html}</div>
                <div class="app-modal-actions">
                    <button type="button" class="btn btn-primary app-modal-ok">${options.confirmText || 'OK'}</button>
                </div>
            </div>
        `;
        document.body.appendChild(overlay);

        function closeModal() {
            overlay.remove();
        }

        overlay.querySelector('.app-modal-close').addEventListener('click', closeModal);
        overlay.querySelector('.app-modal-ok').addEventListener('click', closeModal);
        overlay.addEventListener('click', function(event) {
            if (event.target === overlay) closeModal();
        });
    }

    function showNotice(title, message) {
        showModal(title, '<p>' + message + '</p>', { confirmText: 'OK' });
    }

    // ===== Inisialisasi Jspreadsheet (500 baris) =====
    const table = jspreadsheet(document.getElementById('spreadsheet'), {
        data: [['', '']],
        columns: columns,
        minDimensions: [2, DEFAULT_ROWS],
        allowInsertRow: true,
        allowManualInsertRow: true,
        allowInsertColumn: false,
        allowDeleteRow: true,
        wordWrap: true,
    });

    // Paksa pastikan 500 baris walaupun library/CDN behave beda
    function ensureMinimumRows(minRows) {
        var currentRows = table.getData().length;
        var missingRows = minRows - currentRows;
        if (missingRows > 0) {
            addRows(missingRows, false);
        }
    }

    function addRows(count, showFeedback) {
        count = parseInt(count, 10) || DEFAULT_ROWS;
        if (count <= 0 || count > 10000) {
            showNotice('Jumlah tidak valid', 'Harap masukkan jumlah baris antara 1 sampai 10.000.');
            return false;
        }

        // Jspreadsheet v4 CE lebih aman ditambah satu-satu agar tidak beda interpretasi API insertRow(count).
        for (var i = 0; i < count; i++) {
            table.insertRow();
        }

        if (showFeedback) {
            showNotice('Berhasil', count + ' baris baru sudah ditambahkan ke tabel.');
        }
        return true;
    }

    ensureMinimumRows(DEFAULT_ROWS);

    // ===== Tombol Bersihkan =====
    document.getElementById('btn-clear').addEventListener('click', function() {
        if (confirm('Bersihkan semua data?')) {
            table.setData([['', '']]);
            ensureMinimumRows(DEFAULT_ROWS);
        }
    });

    // ===== Tombol Tambah Baris Kustom =====
    document.getElementById('btn-add-rows').addEventListener('click', function() {
        var count = document.getElementById('input-rows').value;
        addRows(count, true);
    });

    // ===== Tutorial & Penjelasan Tombol =====
    const showTutorial = function() {
        showModal('Panduan CSV Converter E-Faktur', `
            <p>Selamat datang di <strong>CSV Converter E-Faktur</strong>! Alat ini membantu mengubah data dari Excel menjadi CSV dua kolom yang siap digunakan.</p>
            
            <h3>Cara Penggunaan:</h3>
            <ol>
                <li><strong>Copy & Paste:</strong> Copy dua kolom dari Excel: <em>Nomor Faktur Pajak</em> dan <em>Bulan Faktur Pajak</em>, lalu klik sel pertama (<strong>A1</strong>) dan tekan <strong>Ctrl + V</strong>.</li>
                <li><strong>Validasi otomatis:</strong> Saat download, data dicek otomatis. Sel yang salah akan ditandai merah muda.</li>
                <li><strong>Nomor Faktur Pajak:</strong> Harus tepat 17 digit angka, contoh: <code>01000223000000123</code>.</li>
                <li><strong>Bulan Faktur Pajak:</strong> Harus salah satu dari nama bulan Indonesia lengkap dengan huruf awal kapital: <code>Januari</code>, <code>Februari</code>, <code>Maret</code>, <code>April</code>, <code>Mei</code>, <code>Juni</code>, <code>Juli</code>, <code>Agustus</code>, <code>September</code>, <code>Oktober</code>, <code>November</code>, atau <code>Desember</code>.</li>
                <li><strong>Download CSV:</strong> Jika tidak ada error, klik <strong>Download CSV</strong> untuk menyimpan file.</li>
            </ol>

            <h3>Penjelasan Tombol:</h3>
            <ul>
                <li><strong>📖 Cara Pakai:</strong> Membuka panduan ini lagi.</li>
                <li><strong>➕ Tambah Baris:</strong> Menambah baris kosong sesuai angka di input <strong>Jml Baris</strong>.</li>
                <li><strong>Bersihkan Tabel:</strong> Menghapus semua data dan mengembalikan tabel minimal 500 baris.</li>
                <li><strong>Download CSV:</strong> Validasi data dan unduh file CSV.</li>
            </ul>

            <div class="modal-tip">💡 <strong>Tips:</strong> Tabel sekarang otomatis dimuat dengan <strong>500 baris kosong</strong>.</div>
        `, { confirmText: 'Saya Mengerti, Mulai Kerja!' });
    };

    // Tampilkan otomatis saat pertama kali dibuka untuk versi baru
    if (!localStorage.getItem('efaktur_converter_tutorial_shown_v2')) {
        showTutorial();
        localStorage.setItem('efaktur_converter_tutorial_shown_v2', 'true');
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
                errors.push('Baris ' + rowNum + ': Bulan Faktur Pajak tidak valid. Isi harus salah satu dari: Januari, Februari, Maret, April, Mei, Juni, Juli, Agustus, September, Oktober, November, atau Desember. Gunakan huruf awal kapital.');
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

    // Simpan reference table ke window untuk debug/verifikasi
    window.efakturTable = table;
    window.efakturAddRows = addRows;
    window.efakturShowTutorial = showTutorial;
});

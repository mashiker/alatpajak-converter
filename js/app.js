document.addEventListener('DOMContentLoaded', function() {
    // ===== Kolom Spreadsheet =====
    const columns = [
        { type: 'text', title: 'Nomor Faktur Pajak', width: 280 },
        { type: 'text', title: 'Bulan Faktur Pajak', width: 220 }
    ];

    // ===== Inisialisasi Jspreadsheet =====
    const table = jspreadsheet(document.getElementById('spreadsheet'), {
        data: [['', '']], 
        columns: columns,
        minDimensions: [2, 50],
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

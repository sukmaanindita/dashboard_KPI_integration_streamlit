# All Job Report Dashboard

Dashboard untuk mengolah workbook monitoring job, menghitung KPI, membuat pivot per group/priority, dan menghasilkan Excel report baru.

Nama group distandarkan menjadi uppercase dan simbol `&` otomatis diubah menjadi
kata `DAN`, misalnya `ASURANSI & FIDUCIA` menjadi `ASURANSI DAN FIDUCIA`.

## Menjalankan aplikasi

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Aplikasi otomatis memakai workbook terbaru di `data/source`, atau pengguna dapat mengunggah workbook lain dari sidebar.

Tersedia dua pilihan download:

- **Statis** (`_web.xlsx`): seluruh hasil kalkulasi ditulis sebagai angka tetap.
- **Dinamis** (`_web_dynamic.xlsx`): KPI detail, Daily, Monthly, dan Sum Job
  menggunakan formula Excel dan dihitung ulang saat nilai detail berubah.

## Historical analysis

Pilih bulan dan tahun report dari sidebar. Nama file download dibuat otomatis dengan
format seperti `07. Report All Job ICT July 2026_web.xlsx`.

Klik **Simpan/Update Historical Snapshot** untuk menyimpan kondisi job bulan itu ke
`data/history/job_history_YYYY-MM.csv`. Tab **Historical Analysis** membandingkan
snapshot dengan bulan tersimpan sebelumnya dan menampilkan:

- jumlah job aktif;
- job baru;
- job disable (job yang tidak lagi ditemukan pada snapshot berikutnya);
- error job (job yang memiliki retry kurang atau lebih dari 3 jam);
- tren jumlah job aktif, error job, dan success rate bulanan.

Gunakan bagian **Manual Historical Adjustment** untuk mengoreksi jumlah job aktif,
error, disable, dan baru per bulan. Adjustment disimpan terpisah di
`data/history/manual_history_metrics.csv` dan digunakan pada KPI serta bar chart;
snapshot dan tabel detail asli tidak diubah.
Tombol **Reset Job Error** menghapus adjustment error untuk bulan terpilih dan
mengembalikannya ke jumlah error yang dihitung dari snapshot.

Pada Overview, bagian **Detail Perubahan Job (+ / -)** menampilkan nama job baru
dan job disable untuk setiap perbandingan snapshot bulanan. Detail dapat mengikuti
filter group pada dropdown tren.

## Output

Tombol **Download Excel Report** menghasilkan lima sheet dengan layout workbook sumber:

- `Result Detail JOB` — detail 31 hari dan metrik hasil kalkulasi Python, tanpa
  tabel summary atau pivot tambahan.
- `Daily` — hasil average per group dan priority untuk job Daily.
- `Monthly` — hasil average per group dan priority untuk job Monthly.
- `Sum Job` — blok job summary per group dengan pivot per priority di sampingnya.
- `MSG_ERROR` — isi error dari workbook sumber; tetap kosong jika sumber tidak berisi data error.

Persentase pada pivot dihitung sebagai rata-rata sederhana persentase setiap job,
sesuai perilaku `Average of Success Percentage` pada PivotTable Excel. KPI ringkasan
di bagian atas dashboard tetap memakai total status dibagi total periodic job agar
menunjukkan tingkat keberhasilan aktual seluruh eksekusi.

## Menjalankan test

```bash
pytest -q
```

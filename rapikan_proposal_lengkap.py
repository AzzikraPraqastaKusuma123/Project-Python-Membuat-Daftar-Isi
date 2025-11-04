import re
import os
from docx import Document

def rapikan_styles_proposal_lengkap(nama_file_input, nama_file_output):
    """
    Membaca file input, menerapkan styles berdasarkan pola teks,
    dan menyimpannya sebagai file output baru.

    Fungsi ini mengenali:
    1. Judul Utama (BAB, KATA PENGANTAR, DAFTAR ISI, LAMPIRAN)
    2. Nama Bab di baris terpisah (misal "BAB I" di baris 1, "PENDAHULUAN" di baris 2)
    3. Sub-bab (misal "1.1", "2.4")
    4. Sub-sub-bab (misal "1.1.1", "3.5.2")
    """
    
    # --- Langkah 1: Validasi File Input ---
    if not os.path.exists(nama_file_input):
        print(f"Error: File '{nama_file_input}' tidak ditemukan.")
        print("Pastikan nama file sudah benar dan berada di folder yang sama.")
        return

    try:
        # --- Langkah 2: Buka Dokumen dan Siapkan Pola Regex ---
        doc = Document(nama_file_input)
        print(f"Membaca file '{nama_file_input}'...")

        # Pola 1: Untuk Judul Utama (Heading 1)
        # Mencari teks seperti "BAB I", "BAB II", "KATA PENGANTAR", "DAFTAR ISI",
        # "DAFTAR PUSTAKA", atau "LAMPIRAN".
        # re.IGNORECASE = tidak peduli huruf besar/kecil.
        pola_judul_utama = re.compile(
            r'^\s*(BAB\s+[IVXLC]+|KATA\s+PENGANTAR|DAFTAR\s+(ISI|PUSTAKA|LAMPIRAN|TABEL|GAMBAR))\s*$', 
            re.IGNORECASE
        )
        
        # Pola 2: Untuk Sub-bab (Heading 2)
        # Mencari pola angka.angka (misal "1.1", "10.2") di awal baris.
        pola_sub_bab = re.compile(r'^\s*\d+\.\d+\s+')
        
        # Pola 3: Untuk Sub-sub-bab (Heading 3)
        # Mencari pola angka.angka.angka (misal "1.1.1", "4.5.2") di awal baris.
        pola_sub_sub_bab = re.compile(r'^\s*\d+\.\d+\.\d+\s+')

        jumlah_perubahan = 0
        
        # --- Langkah 3: Iterasi Melalui Setiap Paragraf ---
        
        # Kita gunakan list paragraf agar bisa "mengintip" paragraf berikutnya (i+1)
        paragraf = doc.paragraphs
        i = 0
        while i < len(paragraf):
            para = paragraf[i]
            
            # Lewati paragraf kosong
            if not para.text.strip():
                i += 1
                continue

            nama_bab = ""
            
            # --- Pengecekan Pola Heading 1 (Judul Utama) ---
            if pola_judul_utama.match(para.text):
                nama_bab = para.text.strip()
                
                # Cek apakah baris BERIKUTNYA adalah NAMA bab-nya
                # (Contoh: baris ini "BAB I", baris berikutnya "PENDAHULUAN")
                # Logika: Apakah ada baris berikutnya? DAN baris itu tidak kosong?
                # DAN baris itu BUKAN pola heading lain?
                if (i + 1) < len(paragraf) and \
                   paragraf[i+1].text.strip() and \
                   not pola_judul_utama.match(paragraf[i+1].text) and \
                   not pola_sub_bab.match(paragraf[i+1].text) and \
                   not pola_sub_sub_bab.match(paragraf[i+1].text):
                    
                    # Ambil nama bab dari baris berikutnya
                    nama_bab_tambahan = paragraf[i+1].text.strip()
                    teks_gabungan = f"{nama_bab} {nama_bab_tambahan}" # Gabungkan
                    
                    # Terapkan style 'Heading 1' ke baris pertama
                    if para.style.name != 'Heading 1':
                        para.style = 'Heading 1'
                        jumlah_perubahan += 1
                    
                    # Update teks paragraf pertama menjadi teks gabungan
                    para.text = teks_gabungan
                    print(f"  [Heading 1 Digabung] -> {teks_gabungan[:60]}...")
                    
                    # Hapus paragraf kedua (karena sudah digabung)
                    p_hapus = paragraf[i+1]
                    p_hapus._element.getparent().remove(p_hapus._element)
                    # Kita tidak perlu increment 'i' di sini, 
                    # karena paragraf[i+1] sudah dihapus, loop akan lanjut ke
                    # paragraf yang 'baru' (yang tadinya i+2)
                    
                else:
                    # Jika hanya 1 baris (misal "KATA PENGANTAR")
                    if para.style.name != 'Heading 1':
                        para.style = 'Heading 1'
                        jumlah_perubahan += 1
                        print(f"  [Heading 1] -> {para.text[:60]}...")
            
            # --- Pengecekan Pola Heading 2 (Sub-bab) ---
            elif pola_sub_bab.match(para.text):
                if para.style.name != 'Heading 2':
                    para.style = 'Heading 2'
                    jumlah_perubahan += 1
                    print(f"  [Heading 2] -> {para.text[:60]}...")

            # --- Pengecekan Pola Heading 3 (Sub-sub-bab) ---
            elif pola_sub_sub_bab.match(para.text):
                if para.style.name != 'Heading 3':
                    para.style = 'Heading 3'
                    jumlah_perubahan += 1
                    print(f"  [Heading 3] -> {para.text[:60]}...")
            
            # --- Jika Bukan Heading (Teks Biasa) ---
            # Kita set sebagai 'Normal' jika style-nya bukan 'Normal'
            # atau 'Heading'
            elif not para.style.name.startswith('Heading'):
                 if para.style.name != 'Normal':
                    para.style = 'Normal'
                    # (Tidak perlu di-print agar log tidak penuh)
            
            i += 1 # Lanjut ke paragraf berikutnya

        # --- Langkah 4: Simpan Hasil ke File Baru ---
        doc.save(nama_file_output)
        print(f"\nSelesai! {jumlah_perubahan} style heading telah diperbarui.")
        print(f"File baru disimpan sebagai: '{nama_file_output}'")

    except Exception as e:
        print(f"Terjadi error saat memproses file: {e}")
        print("Pastikan file .docx tidak korup dan tidak sedang dibuka di Word.")

# --- UTAMA: GANTI NAMA FILE DI BAWAH INI ---
if __name__ == "__main__":
    # Ganti ini dengan nama file proposal Anda
    nama_file_input = "UTS Kelas B-403 Metopen.docx"
    
    # Ini adalah nama file hasil yang akan dibuat
    nama_file_output = "UTS Kelas B-403 Metopen-V2.docx"
    
    # Panggil fungsi utama
    rapikan_styles_proposal_lengkap(nama_file_input, nama_file_output)
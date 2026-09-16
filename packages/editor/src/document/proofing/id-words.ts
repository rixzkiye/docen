/**
 * Built-in Indonesian lexicon for the spell checker.
 * Covers frequent Indonesian root words, affixes, and functional vocabulary.
 */

const ID_WORDS_LIST = `
yang di dan itu dengan untuk ini dari dalam tidak akan pada juga ke karena
bisa ada mereka saat lebih tahun menjadi sudah orang semua telah sangat hanya
oleh hari suatu harus kita ia bagi satu dapat banyak waktu membuat jika
tentang tahu saya kami lagi setelah apa diri mereka tempat baru dua seperti
hidup kata rumah antara hal tersebut sampai mana baik pekerjaan kecil besar
masih cara selalu pertama anak mereka nama beberapa ingin sebelum keluarga
tangan bagian air jalan kepala dunia negara mata malam kota uang sekolah
matahari perempuan laki pintu negara cerita pemerintah akhir tempat uang
masalah nomor kamar sisi teman kelompok ibu ayah hati tanah air teman
kerja malam pagi hari tahun jam menit detik bulan minggu sore petang
makan minum tidur bangun pergi pulang jalan lari duduk berdiri bicara dengar
lihat baca tulis bawa beli jual bayar hitung ambil simpan kirim buka tutup
tarik dorong jatuh bangun naik turun masuk keluar ganti coba pakai bantu
mulai selesai lanjut tunggu cari dapat hilang temu ingat lupa pikir rasa
paham mengerti jelas benar salah sulit mudah mungkin pasti tentu harus wajib
bisa boleh mampu sanggup penting perlu biasa senang bahagia sedih marah takut
berani tenang cemas suka benci rindu cinta sayang hormat kagum bangga malu
panjang pendek tinggi rendah luas sempit tebal tipis berat ringan keras lunak
cepat lambat panas dingin hangat sejuk terang gelap bersih kotor baru lama
tua muda kaya miskin murah mahal bagus jelek indah cantik tampan rapi
sehat sakit hidup mati lahir tumbuh kembang rusak utuh penuh kosong rata
atas bawah depan belakang samping tengah luar kiri kanan jauh dekat timur
barat utara selatan pulau laut gunung sungai danau hutan sawah ladang desa
kampung gedung jembatan pasar toko kantor kamar ruang meja kursi buku pensil
kertas layar komputer telepon surat kabar gambar foto lukisan musik lagu tari
film mobil motor sepeda kapal pesawat kereta jalan raya gang lorong pelabuhan
bandara stasiun terminal halte halte pabrik industri bisnis modal pasar saham
harga biaya rugi laba utang piutang pajak bunga cicilan sewa belanja tabungan
hukum aturan undang perkara hakim jaksa polisi tentara tahanan penjara bukti
saksi sidang pasal denda vonis bebas adil damai aman tertib rusuh perang
bangsa rakyat warga masyarakat suku budaya adat tradisi agama ibadah doa
iman tuhan nabi rasul gereja masjid pura vihara kuil surga neraka dosa pahala
akal budi pikir nalar ilmu teknologi riset studi belajar kuliah sekolah guru
dosen murid mahasiswa siswa sarjana gelar lulus ujian nilai ijazah buku teks
bahasa sastra puisi sajak prosa cerpen novel drama teater pentas seni rupa
warna merah biru kuning hijau hitam putih abu cokelat jingga ungu emas perak
hewan binatang burung ikan kucing anjing ayam sapi kambing kuda gajah singa
harimau buaya ular nyamuk lalat semut lebah kupu serangga pohon bunga daun
buah akar batang cabang biji kelapa pisang apel mangga jeruk padi jagung gandum
sayur bayam tomat wortel kentang cabai bawang garam gula beras minyak susu
daging telur tempe tahu roti kue nasi kopi teh air sirup sup soto
tubuh badan jiwa raga darah tulang daging kulit rambut kuku gigi lidah telinga
hidung bibir pipi dagu leher pundak lengan siku pergelangan jari dada perut
pinggang pinggul paha lutut betis tumit kaki langkah nafas detak suara rupa
wajah pandang senyum tawa tangis jerit bisik teriak sapa salam jumpa temu
pisah pamit datang pergi hadir lenyap musnah cipta bina bangun susun bentuk
ubah ubahan ubah perubah pengaruh akibat dampak sebab pangkal ujung awal
mulai henti henti cegah tangkal tolak terima sambut dukung dorong ajak suruh
pinta mohon harap cemas takut waspada jaga awas hati rahasia nyata semu
benar betul salah keliru sesat tepat jitu kena sasar tuju arah haluan pedoman
patokan ukuran timbangan timbang ukur hitung jumlah total rata bagi kali kurang
tambah selisih lebih sisa lebihan genap ganjil seri ragam jenis macam corak
rupa wujud sifat perangai tabiat watak akhlak moral etika sopan santun ramah
lembut halus kasar tajam tumpul runcing lonjong bundar bulat persegi kotak
lurus bengkok lekuk sudut garis bidang ruang bentuk model pola struktur rangka
wadah isi pokok dasar inti serat sari pati getah racun obat jamu ramuan sehat
panjang lebar dalam tinggi luas volume kapasitas muatan batas pinggir tepian
antar jemput hubung kait rangkai sambung putus pisah campur aduk baur padu
ikat lepas bebas pasang tanggal susun tata atur urut kelola bina pimpin pimpin
tuntun bimbing didik latih ajar perintah titah sabda wasiat petuah petunjuk
nasihat usul saran kritik puji sanjung cela hina maki kutuk fitnah tuduh sangka
duga duga curiga ragu bimbang yakin mantap tegas bulat tekad niat janji sumpah
`;

export const indonesianWords: ReadonlySet<string> = new Set(
  ID_WORDS_LIST.split(/\s+/)
    .map((w) => w.trim().toLowerCase())
    .filter(Boolean),
);

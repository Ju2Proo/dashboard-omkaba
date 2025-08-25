import re
from collections import Counter
import pandas as pd
from typing import List, Optional

def comprehensive_clean(text, company_mode=True):
    """
    Pembersihan komprehensif untuk text data
    """
    if pd.isna(text):
        return text
        
    original = text
    text = str(text)
    
    # Step 1: Basic cleaning
    text = text.strip()
    
    # Step 2: Handle special characters
    if company_mode:
        # Untuk nama perusahaan
        text = text.upper()
        
        # Standarisasi PT/CV
        text = re.sub(r'^P\.?T\.?\s*', 'PT ', text)
        text = re.sub(r'^C\.?V\.?\s*', 'CV ', text)
        text = re.sub(r'^U\.?D\.?\s*', 'UD ', text)
        
        # Hapus PT/CV di akhir
        text = re.sub(r'\s*,?\s*P\.?T\.?$', '', text)
        text = re.sub(r'\s*,?\s*C\.?V\.?$', '', text)
    else:
        # Untuk text biasa
        text = text.title()
    
    # Step 3: Clean punctuation
    text = re.sub(r'[,.\-_()]+', ' ', text)
    
    # Step 4: Multiple spaces
    text = re.sub(r'\s+', ' ', text)
    
    # Step 5: Final cleanup
    text = text.strip()
    text = re.sub(r'[,.\s]+$', '', text)
    
    # Step 6: Add prefix if company and no prefix
    if company_mode and not text.startswith(('PT ', 'CV ', 'UD ', 'PD ')):
        text = 'PT ' + text
    
    return text



# Daftar kota/kabupaten di Indonesia (contoh - bisa diperluas)
CITIES_INDONESIA = [
    # Kota Besar
    "Ambon", "Balikpapan", "Banda Aceh", "Bandar Lampung", "Bandung", "Banjarmasin", 
    "Banjarbaru", "Batam", "Batu", "Baubau", "Bekasi", "Bengkulu", "Bima", "Binjai",
    "Bitung", "Blitar", "Bogor", "Bontang", "Bukittinggi", "Cilegon", "Cimahi",
    "Cirebon", "Denpasar", "Depok", "Dumai", "Gorontalo", "Gunungsitoli", "Jakarta",
    "Jambi", "Jayapura", "Kediri", "Kendari", "Kupang", "Langsa", "Lhokseumawe",
    "Lubuklinggau", "Madiun", "Magelang", "Makassar", "Malang", "Manado", "Mataram",
    "Medan", "Metro", "Mojokerto", "Padang", "Padangpanjang", "Padangsidimpuan",
    "Pagar Alam", "Palangka Raya", "Palembang", "Palopo", "Palu", "Pangkalpinang",
    "Parepare", "Pasuruan", "Payakumbuh", "Pekalongan", "Pekanbaru", "Pematangsiantar",
    "Pontianak", "Prabumulih", "Probolinggo", "Sabang", "Salatiga", "Samarinda",
    "Sawahlunto", "Semarang", "Serang", "Sibolga", "Singkawang", "Solok", "Sorong",
    "Subulussalam", "Sukabumi", "Sungai Penuh", "Surabaya", "Surakarta", "Tangerang",
    "Tanjungbalai", "Tanjungpinang", "Tarakan", "Tasikmalaya",
    "Tebing Tinggi", "Tegal", "Ternate", "Tidore Kepulauan", "Tomohon", "Tual",
    "Yogyakarta", "Gresik", "Sidoarjo", "Mojokerto", "Jombang", "Nganjuk", "Madiun",
    "Magetan", "Ngawi", "Bojonegoro", "Tuban", "Lamongan", "Bangkalan", "Sampang",
    "Pamekasan", "Sumenep", "Kediri", "Blitar", "Tulungagung", "Trenggalek", "Ponorogo",
    "Pacitan", "Wonogiri", "Karanganyar", "Sragen", "Grobogan", "Blora", "Rembang",
    "Pati", "Kudus", "Jepara", "Demak", "Semarang", "Temanggung", "Kendal", "Batang",
    "Pekalongan", "Pemalang", "Tegal", "Brebes", "Cilacap", "Banyumas", "Purbalingga",
    "Banjarnegara", "Kebumen", "Purworejo", "Wonosobo", "Magelang", "Boyolali",
    "Klaten", "Sukoharjo", "Wonogiri", "Karanganyar", "Sragen", "Grobogan", "Blora",
    "Rembang", "Pati", "Kudus", "Jepara", "Demak", "Semarang", "Temanggung", "Kendal",
    "Batang", "Pekalongan", "Pemalang", "Tegal", "Brebes"
]

# Tambahkan mapping area ke kota utama
AREA_TO_CITY_MAPPING = {
    # Surabaya dan sekitarnya
    'waru': 'Sidoarjo',
    'gedangan': 'Sidoarjo', 
    'taman': 'Sidoarjo',
    'buduran': 'Sidoarjo',
    'candi': 'Sidoarjo',
    'porong': 'Sidoarjo',
    'sedati': 'Sidoarjo',
    'tambakrejo': 'Sidoarjo',
    'tambak sawah': 'Sidoarjo',
    'benowo': 'Surabaya',
    'lakarsantri': 'Surabaya',
    'sambikerep': 'Surabaya',
    'asemrowo': 'Surabaya',
    'krembangan': 'Surabaya',
    'pabean cantikan': 'Surabaya',
    'semampir': 'Surabaya',
    'bubutan': 'Surabaya',
    'tandes': 'Surabaya',
    'sukomanunggal': 'Surabaya',
    'sawahan': 'Surabaya',
    'tegalsari': 'Surabaya',
    'genteng': 'Surabaya',
    'gubeng': 'Surabaya',
    'sukolilo': 'Surabaya',
    'mulyorejo': 'Surabaya',
    'rungkut': 'Surabaya',
    'tenggilis mejoyo': 'Surabaya',
    'gunung anyar': 'Surabaya',
    'jambangan': 'Surabaya',
    'gayungan': 'Surabaya',
    'wonocolo': 'Surabaya',
    'wonokromo': 'Surabaya',
    'karang pilang': 'Surabaya',
    'dukuh pakis': 'Surabaya',
    'wiyung': 'Surabaya',
    'romokalisari': 'Surabaya',
    'keputih': 'Surabaya',
    'sikatan': 'Surabaya',
    'kedamean': 'Gresik',
    'driyorejo': 'Gresik',
    'cerme': 'Gresik',
    'benjeng': 'Gresik',
    'balongpanggang': 'Gresik',
    'ujungpangkah': 'Gresik',
    'panceng': 'Gresik',
    'sidayu': 'Gresik',
    'bungah': 'Gresik',
    'manyar': 'Gresik',
    'duduksampeyan': 'Gresik',
    'sangkapura': 'Gresik',
    'menganti': 'Gresik',
    'wringinanom': 'Gresik',
    'kebomas': 'Gresik',
    'gresik': 'Gresik',
    'cokelat': 'Gresik',
    
    # Jakarta dan sekitarnya
    'kelapa gading': 'Jakarta',
    'sunter': 'Jakarta',
    'tanjung priok': 'Jakarta',
    'koja': 'Jakarta',
    'cilincing': 'Jakarta',
    'pademangan': 'Jakarta',
    'penjaringan': 'Jakarta',
    'gambir': 'Jakarta',
    'sawah besar': 'Jakarta',
    'kemayoran': 'Jakarta',
    'senen': 'Jakarta',
    'cempaka putih': 'Jakarta',
    'johar baru': 'Jakarta',
    'tanah abang': 'Jakarta',
    'menteng': 'Jakarta',
    'tebet': 'Jakarta',
    'setiabudi': 'Jakarta',
    'mampang prapatan': 'Jakarta',
    'pancoran': 'Jakarta',
    'jagakarsa': 'Jakarta',
    'pasar minggu': 'Jakarta',
    'kebayoran lama': 'Jakarta',
    'kebayoran baru': 'Jakarta',
    'pesanggrahan': 'Jakarta',
    'cilandak': 'Jakarta',
    'cakung': 'Jakarta',
    'cipayung': 'Jakarta',
    'kramat jati': 'Jakarta',
    'jatinegara': 'Jakarta',
    'duren sawit': 'Jakarta',
    'cakung': 'Jakarta',
    'pulo gadung': 'Jakarta',
    'makasar': 'Jakarta',
    'pasar rebo': 'Jakarta',
    'ciracas': 'Jakarta',
    'kembangan': 'Jakarta',
    'kebon jeruk': 'Jakarta',
    'palmerah': 'Jakarta',
    'grogol petamburan': 'Jakarta',
    'tambora': 'Jakarta',
    'taman sari': 'Jakarta',
    'cengkareng': 'Jakarta',
    'kalideres': 'Jakarta',
}

def clean_address_text(address: str) -> str:
    """
    Membersihkan teks alamat dari karakter dan kata yang tidak perlu
    """
    if pd.isna(address) or not address:
        return ""
    
    # Hapus karakter newline dan tab
    address = re.sub(r'[\n\r\t]+', ' ', address)
    
    # Hapus multiple spaces
    address = re.sub(r'\s+', ' ', address)
    
    # Hapus kata-kata yang tidak perlu di awal
    address = re.sub(r'^(jalan|jl\.?|jln\.?|gang|gg\.?|komplek|kompleks|perumahan|perum\.?)\s+', '', address, flags=re.IGNORECASE)
    
    return address.strip()

def prioritize_city(cities: List[str]) -> str:
    """
    Prioritaskan kota berdasarkan kepentingan dan spesifisitas
    """
    # Daftar prioritas kota besar
    priority_cities = [
        "Jakarta", "Surabaya", "Bandung", "Medan", "Semarang", "Makassar", "Palembang",
        "Tangerang", "Tangerang Selatan", "Depok", "Bekasi", "Bogor", "Batam", 
        "Pekanbaru", "Bandar Lampung", "Malang", "Padang", "Denpasar", "Samarinda", 
        "Tasikmalaya", "Pontianak", "Cimahi", "Balikpapan", "Jambi", "Surakarta",
        "Serang", "Cilegon", "Tegal", "Binjai", "Pematangsiantar", "Jayapura", 
        "Kediri", "Cirebon", "Yogyakarta", "Magelang", "Purwokerto", "Blitar",
        "Gresik", "Sidoarjo", "Jakarta Pusat", "Jakarta Utara", "Jakarta Selatan", 
        "Jakarta Timur", "Jakarta Barat"
    ]
    
    # Cari kota prioritas terlebih dahulu
    for priority_city in priority_cities:
        if priority_city in cities:
            return priority_city
    
    # Jika tidak ada, ambil yang terpanjang (biasanya lebih spesifik)
    return max(cities, key=len)

def clean_extracted_city(city: str) -> Optional[str]:
    """
    Membersihkan hasil ekstraksi kota dari kata-kata yang tidak perlu
    """
    if not city or city == "Tidak Diketahui":
        return city
    
    # Hapus newline dan whitespace
    city = re.sub(r'[\n\r\t]+', ' ', city).strip()
    
    # Hapus kata-kata yang tidak perlu
    unwanted_patterns = [
        r'\b(provinsi|prov\.?|propinsi)\b.*',
        r'\b(indonesia|id|idn)\b.*',
        r'\b(jawa timur|jawa tengah|jawa barat|east java|west java|central java)\b.*',
        r'\b(sumatra utara|sumatra barat|sumatra selatan|kalimantan|sulawesi|bali|ntt|ntb|maluku|papua)\b.*',
        r'\b(kecamatan|kec\.?|kelurahan|kel\.?|desa|ds\.?)\b.*',
        r'\b(rt|rw)\s*[\d/]+\b.*',
        r'\b\d{5}\b.*',  # kode pos dan setelahnya
        r'\b(jl\.?|jalan|street|st\.?|road|rd\.?|raya)\b.*',  # hapus jalan dan setelahnya
        r'\s*-\s*(indonesia|id).*',  # hapus "- INDONESIA" dan setelahnya
        r'\s*,\s*(indonesia|id).*',  # hapus ", INDONESIA" dan setelahnya
    ]
    
    for pattern in unwanted_patterns:
        city = re.sub(pattern, '', city, flags=re.IGNORECASE).strip()
        if not city:  # Jika habis terhapus, kembalikan None
            return None
    
    # Handle kasus khusus seperti "Waru-Sidoarjo" -> ambil yang kedua jika dikenal
    if '-' in city and not any(x in city.lower() for x in ['jakarta', 'tangerang']):
        parts = city.split('-')
        city_candidates = [part.strip() for part in parts if part.strip()]
        if len(city_candidates) >= 2:
            # Cek apakah bagian kedua adalah kota yang dikenal
            known_cities = [c.lower() for c in CITIES_INDONESIA]
            # Prioritaskan bagian yang merupakan kota terkenal
            for part in reversed(city_candidates):  # mulai dari belakang
                if part.lower() in known_cities:
                    city = part.title()
                    break
            else:
                city = city_candidates[-1].title()  # ambil yang terakhir jika tidak ada yang dikenal
    
    # Handle kasus "Kelapa Gading Jakarta Utara" -> prioritaskan Jakarta Utara
    jakarta_areas = ['jakarta']
    city_lower = city.lower()
    for area in jakarta_areas:
        if area in city_lower:
            return area.title()
    
    # Handle kasus umum kota Jakarta
    if 'jakarta' in city_lower and not any(area in city_lower for area in jakarta_areas):
        return 'Jakarta'
    
    # Hapus leading/trailing whitespace dan multiple spaces
    city = re.sub(r'\s+', ' ', city).strip()
    
    # Filter hasil yang terlalu pendek atau terlalu panjang
    if len(city) < 2 or len(city) > 50:
        return None
    
    # Filter jika hanya mengandung angka
    if city.isdigit():
        return None
    
    # Filter jika mengandung terlalu banyak angka
    if len(re.findall(r'\d', city)) > len(city) // 2:
        return None
    
    return city.title() if city else None

def extract_city_with_area_mapping(address: str) -> Optional[str]:
    """
    Ekstrak kota dengan mempertimbangkan mapping area ke kota utama
    """
    if pd.isna(address) or not address:
        return None
    
    address_clean = clean_address_text(address)
    address_lower = address_clean.lower()
    
    # Cek apakah ada area yang bisa di-mapping ke kota utama
    for area, main_city in AREA_TO_CITY_MAPPING.items():
        pattern = r'\b' + re.escape(area.lower()) + r'\b'
        if re.search(pattern, address_lower):
            return main_city
    
    return None

def extract_city_regex_pattern(address: str, city_list: List[str]) -> Optional[str]:
    """
    Ekstrak nama kota menggunakan regex pattern matching dengan prioritas
    """
    if pd.isna(address) or not address:
        return None
    
    address_clean = clean_address_text(address)
    address_lower = address_clean.lower()
    
    # Urutkan kota berdasarkan panjang (yang lebih spesifik dulu)
    sorted_cities = sorted(city_list, key=len, reverse=True)
    
    found_cities = []
    for city in sorted_cities:
        # Buat pattern yang lebih fleksibel
        pattern = r'\b' + re.escape(city.lower()) + r'\b'
        if re.search(pattern, address_lower):
            found_cities.append(city.title())
    
    # Prioritaskan kota yang lebih spesifik atau terkenal
    if found_cities:
        return prioritize_city(found_cities)
    
    return None

def extract_city_keyword_based(address: str) -> Optional[str]:
    """
    Ekstrak kota berdasarkan kata kunci umum dalam alamat Indonesia
    """
    if pd.isna(address) or not address:
        return None
    
    address_clean = clean_address_text(address)
    address_lower = address_clean.lower()
    
    # Pattern untuk menangkap struktur alamat Indonesia yang lebih ketat
    patterns = [
        # Pattern: "kota [nama]" atau "kabupaten [nama]"
        r'\b(?:kota|kabupaten|kab\.?)\s+([a-zA-Z\s]+?)(?:\s*[,\n]|\s*\d|\s*jawa|\s*sumatra|\s*kalimantan|\s*sulawesi|\s*bali|\s*$)',
        
        # Pattern: Jakarta dengan area spesifik
        r'\b(jakarta\s+(?:pusat|utara|selatan|timur|barat))\b',
        
        # Pattern: Tangerang Selatan
        r'\b(tangerang\s+selatan)\b',
        
        # Pattern: "[nama], [nama kota/kabupaten yang dikenal]" - ambil yang kedua
        r'\b[a-zA-Z\s]+?,\s*([a-zA-Z\s]+?)(?:\s*[,\n]|\s*jawa|\s*east|\s*west|\s*central|\s*\d{5}|\s*indonesia|\s*$)',
        
        # Pattern: mencari sebelum provinsi atau negara
        r'\b([a-zA-Z\s]+?)(?:\s*[,\n]?\s*(?:jawa\s+(?:barat|tengah|timur)|east\s+java|west\s+java|central\s+java|sumatra|kalimantan|sulawesi|bali|indonesia))\b',
        
        # Pattern untuk kode pos - ambil kata sebelum kode pos
        r'\b([a-zA-Z\s]+?)\s+\d{5}\b',
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, address_lower, re.IGNORECASE)
        if matches:
            city = matches[0] if isinstance(matches[0], str) else matches[0]
            city = city.strip()
            
            # Bersihkan hasil lebih ketat
            city = re.sub(r'\b(jalan|jl|rt|rw|no|nomor|kelurahan|kecamatan|kec|gang|gg|komplek|kompleks|perumahan|perum|raya)\b.*', '', city, flags=re.IGNORECASE).strip()
            
            if len(city) > 2 and len(city) < 50 and not city.isdigit():
                cleaned_city = clean_extracted_city(city.title())
                if cleaned_city and cleaned_city != "Tidak Diketahui":
                    return cleaned_city
    
    return None

def extract_city_last_part(address: str) -> Optional[str]:
    """
    Ekstrak kota dari bagian terakhir alamat dengan perbaikan parsing
    """
    if pd.isna(address) or not address:
        return None
    
    address_clean = clean_address_text(address)
    
    # Split berdasarkan koma dan newline
    separators = [',', '\n', '  ']  # double space juga sebagai separator
    parts = [address_clean]
    
    for sep in separators:
        new_parts = []
        for part in parts:
            new_parts.extend([p.strip() for p in part.split(sep) if p.strip()])
        parts = new_parts
    
    if len(parts) >= 1:
        # Coba dari bagian terakhir ke depan
        for i in range(len(parts)-1, -1, -1):
            part = parts[i].lower().strip()
            
            # Skip jika bagian ini adalah kode pos, negara, atau provinsi saja
            skip_patterns = [
                r'^\d{5}$',  # kode pos
                r'^indonesia$',
                r'^(jawa\s+(?:barat|tengah|timur)|east\s+java|west\s+java|central\s+java|sumatra\s+(?:utara|barat|selatan)|kalimantan|sulawesi|bali|ntt|ntb|maluku|papua)$',
                r'^(id|idn)$',
            ]
            
            should_skip = any(re.match(pattern, part, re.IGNORECASE) for pattern in skip_patterns)
            if should_skip:
                continue
                
            # Bersihkan dari kata-kata yang tidak perlu
            city = part
            city = re.sub(r'\b\d{5}\b', '', city)  # Hapus kode pos
            city = re.sub(r'\b(jawa\s+(?:barat|tengah|timur)|east\s+java|west\s+java|central\s+java|sumatra|kalimantan|sulawesi|bali|ntt|ntb|maluku|papua|indonesia)\b.*', '', city, flags=re.IGNORECASE)
            city = re.sub(r'\s*-\s*.*', '', city)  # Hapus setelah tanda "-"
            city = city.strip()
            
            if len(city) > 2 and not city.isdigit():
                cleaned_city = clean_extracted_city(city.title())
                if cleaned_city and cleaned_city != "Tidak Diketahui":
                    return cleaned_city
    
    return None

def extract_city_fallback(address: str, city_list: List[str]) -> Optional[str]:
    """
    Metode fallback untuk mencari kota di seluruh alamat
    """
    if pd.isna(address) or not address:
        return None
    
    address_clean = clean_address_text(address)
    address_lower = address_clean.lower()
    
    # Cari semua kemungkinan kota yang ada di alamat
    found_cities = []
    
    # Daftar kota besar Indonesia untuk fallback dengan prioritas
    priority_cities = [
        'Jakarta', 'Surabaya', 'Bandung', 'Medan', 'Semarang', 'Makassar', 'Palembang',
        'Tangerang', 'Tangerang Selatan', 'Depok', 'Bekasi', 'Bogor', 'Yogyakarta',
        'Malang', 'Solo', 'Surakarta', 'Denpasar', 'Batam', 'Pekanbaru', 'Padang',
        'Bandar Lampung', 'Balikpapan', 'Samarinda', 'Pontianak', 'Manado', 'Jayapura',
        'Ambon', 'Kupang', 'Mataram', 'Gresik', 'Sidoarjo',
    ]
    
    for city in priority_cities:
        city_lower = city.lower()
        if city_lower in address_lower:
            # Cek apakah ini benar-benar nama kota (dengan word boundary)
            pattern = r'\b' + re.escape(city_lower) + r'\b'
            if re.search(pattern, address_lower):
                found_cities.append(city)
    
    # Jika ada kota yang ditemukan, prioritaskan
    if found_cities:
        return prioritize_city(found_cities)
    
    return None

def extract_city_comprehensive(address: str, city_list: List[str]) -> Optional[str]:
    """
    Kombinasi semua metode ekstraksi dengan perbaikan urutan prioritas
    """
    if pd.isna(address) or not address:
        return "Tidak Diketahui"
    
    # Metode 1: Area mapping (prioritas tertinggi untuk area spesifik)
    result = extract_city_with_area_mapping(address)
    if result:
        return result
    
    # Metode 2: Pattern matching dengan daftar kota
    result = extract_city_regex_pattern(address, city_list)
    if result:
        cleaned_result = clean_extracted_city(result)
        if cleaned_result and cleaned_result != "Tidak Diketahui":
            return cleaned_result
    
    # Metode 3: Keyword-based extraction
    result = extract_city_keyword_based(address)
    if result:
        return result
    
    # Metode 4: Last part extraction
    result = extract_city_last_part(address)
    if result:
        return result
    
    # Metode 5: Fallback - cari kata yang mungkin kota di seluruh alamat
    result = extract_city_fallback(address, city_list)
    if result:
        return result
    
    return "Tidak Diketahui"

# Fungsi tambahan untuk membersihkan dan memperbaiki hasil
def clean_extracted_cities_df(df: pd.DataFrame, city_column: str = 'Kota'):
    """
    Membersihkan dan standardisasi hasil ekstraksi kota
    """
    # Mapping untuk standardisasi nama kota
    city_mapping = {
        'yogya': 'Yogyakarta',
        'jogja': 'Yogyakarta', 
        'solo': 'Surakarta',
        'jakarta pusat': 'Jakarta',
        'jakarta selatan': 'Jakarta',
        'jakarta utara': 'Jakarta',
        'jakarta barat': 'Jakarta',
        'jakarta timur': 'Jakarta',
        'tangerang selatan': 'Tangerang',
        'bogor selatan': 'Bogor',
        'bandung barat': 'Bandung',
    }
    
    df[city_column] = df[city_column].replace(city_mapping)
    return df

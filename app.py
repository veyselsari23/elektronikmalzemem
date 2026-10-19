import os
import pandas as pd
from flask import Flask, jsonify, render_template_string, request

app = Flask(__name__)
DOSYA_YOLU = "ListeOZdisan.xlsx"
SHEET_ADI = "Sipariş Detay"


def verileri_oku():
  if not os.path.exists(DOSYA_YOLU):
    return pd.DataFrame()
  try:
    df = pd.read_excel(DOSYA_YOLU, sheet_name=SHEET_ADI)
    # NaN değerleri temizle (JSON dönüşümünde sorun olmasın diye)
    df = df.fillna("")
    return df
  except Exception as e:
    print(f"Excel okuma hatası: {e}")
    return pd.DataFrame()


def verileri_kaydet(df):
  with pd.ExcelWriter(DOSYA_YOLU, engine="openpyxl", mode="w") as writer:
    df.to_excel(writer, sheet_name=SHEET_ADI, index=False)


# HTML Arayüz Tasarımı (Tek Dosya, Tailwind CSS & Alpine.js içerir)
HTML_SABLONU = """
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Özdisan Sipariş & Ürün Yönetimi</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Alpine.js for interactive UI -->
    <script defer src="https://cdn.jsdelivr.net/npm/alpinejs@3.x.x/dist/cdn.min.js"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; }
    </style>
</head>
<body class="bg-slate-50 text-slate-800 min-h-screen flex flex-col" x-data="appData()">

    <!-- Header -->
    <header class="bg-indigo-600 text-white shadow-lg">
        <div class="max-w-7xl mx-auto px-4 py-6 flex flex-col sm:flex-row justify-between items-center gap-4">
            <div class="flex items-center space-x-3">
                <div class="bg-white/10 p-3 rounded-xl backdrop-blur-md">
                    <i class="fa-solid fa-boxes-stacked text-2xl"></i>
                </div>
                <div>
                    <h1 class="text-2xl font-bold tracking-tight">Özdisan Veritabanı Yöneticisi</h1>
                    <p class="text-indigo-200 text-sm">ListeOZdisan.xlsx tabanlı modern yönetim paneli</p>
                </div>
            </div>
            <button @click="modalAc = true" class="bg-white text-indigo-600 hover:bg-indigo-50 px-5 py-2.5 rounded-xl font-semibold shadow-md transition duration-200 flex items-center gap-2">
                <i class="fa-solid fa-plus-circle"></i> Yeni Ürün Ekle
            </button>
        </div>
    </header>

    <!-- Main Container -->
    <main class="max-w-7xl mx-auto px-4 py-8 flex-grow w-full">
        
        <!-- Bildirim Alanı -->
        <template x-if="mesaj">
            <div :class="mesajTuru === 'hata' ? 'bg-red-100 text-red-700 border-red-200' : 'bg-emerald-100 text-emerald-700 border-emerald-200'" class="mb-6 p-4 rounded-xl border flex items-center justify-between shadow-sm transition-all">
                <span x-text="mesaj" class="font-medium"></span>
                <button @click="mesaj = ''" class="text-sm font-bold opacity-70 hover:opacity-100">&times;</button>
            </div>
        </template>

        <!-- Arama ve Filtre Kartı -->
        <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 mb-8">
            <div class="flex flex-col md:flex-row gap-4 items-center justify-between">
                <div class="relative w-full md:w-96">
                    <span class="absolute inset-y-0 left-0 flex items-center pl-4 pointer-events-none text-slate-400">
                        <i class="fa-solid fa-magnifying-glass"></i>
                    </span>
                    <input type="text" x-model="aramaMetni" @input="veriFiltrele" placeholder="Ürün Kodu, Üretici veya Açıklama ara..." 
                           class="w-full pl-11 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition">
                </div>
                <div class="text-sm text-slate-500 font-medium">
                    Toplam <span class="text-indigo-600 font-bold" x-text="filtrelenmisVeriler.length"></span> kayıt gösteriliyor
                </div>
            </div>
        </div>

        <!-- Tablo Görünümü -->
        <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="overflow-x-auto">
                <table class="w-full text-left border-collapse">
                    <thead>
                        <tr class="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-500 uppercase tracking-wider">
                            <th class="py-4 px-6">#</th>
                            <th class="py-4 px-6">Ürün Kodu</th>
                            <th class="py-4 px-6">Üretici</th>
                            <th class="py-4 px-6">Açıklama</th>
                            <th class="py-4 px-6 text-right">Adet</th>
                            <th class="py-4 px-6 text-right">Birim Fiyat</th>
                            <th class="py-4 px-6 text-right">Toplam Fiyat</th>
                            <th class="py-4 px-6">Lokasyon</th>
                            <th class="py-4 px-6 text-center">İşlem</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100 text-sm">
                        <template x-for="item in filtrelenmisVeriler" :key="item['#']">
                            <tr class="hover:bg-slate-50/80 transition">
                                <td class="py-4 px-6 font-medium text-slate-900" x-text="item['#']"></td>
                                <td class="py-4 px-6 font-semibold text-indigo-600" x-text="item['Ürün Kodu']"></td>
                                <td class="py-4 px-6 text-slate-700" x-text="item['Üretici']"></td>
                                <td class="py-4 px-6 text-slate-600 max-w-xs truncate" x-text="item['Ürün Açıklaması']"></td>
                                <td class="py-4 px-6 text-right font-medium" x-text="item['Adet']"></td>
                                <td class="py-4 px-6 text-right text-slate-600" x-text="item['Birim Fiyat']"></td>
                                <td class="py-4 px-6 text-right font-semibold text-slate-900" x-text="item['Toplam Fiyat (KDV Dahil)']"></td>
                                <td class="py-4 px-6 text-slate-500" x-text="item['Lokasyon']"></td>
                                <td class="py-4 px-6 text-center">
                                    <button @click="veriSil(item['#'])" class="bg-red-50 text-red-600 hover:bg-red-100 p-2 rounded-lg transition" title="Sil">
                                        <i class="fa-solid fa-trash-can"></i>
                                    </button>
                                </td>
                            </tr>
                        </template>
                        <template x-if="filtrelenmisVeriler.length === 0">
                            <tr>
                                <td colspan="9" class="text-center py-12 text-slate-400">
                                    <i class="fa-solid fa-folder-open text-4xl mb-3"></i>
                                    <p>Kayıt bulunamadı.</p>
                                </td>
                            </tr>
                        </template>
                    </tbody>
                </table>
            </div>
        </div>
    </main>

    <!-- Yeni Ürün Ekle Modal -->
    <div x-show="modalAc" class="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center p-4 z-50" style="display: none;">
        <div @click.outside="modalAc = false" class="bg-white rounded-3xl shadow-2xl max-w-lg w-full p-8 overflow-hidden transform transition-all">
            <div class="flex justify-between items-center mb-6">
                <h3 class="text-xl font-bold text-slate-800">Yeni Ürün / Sipariş Ekle</h3>
                <button @click="modalAc = false" class="text-slate-400 hover:text-slate-600 text-lg">&times;</button>
            </div>
            
            <form @submit.prevent="veriEkle" class="space-y-4">
                <div>
                    <label class="block text-xs font-semibold text-slate-600 uppercase mb-1">Ürün Kodu</label>
                    <input type="text" x-model="yeniUrun.urun_kodu" required class="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                </div>
                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-semibold text-slate-600 uppercase mb-1">Üretici</label>
                        <input type="text" x-model="yeniUrun.uretici" required class="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-slate-600 uppercase mb-1">Lokasyon</label>
                        <input type="text" x-model="yeniUrun.lokasyon" class="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                    </div>
                </div>
                <div>
                    <label class="block text-xs font-semibold text-slate-600 uppercase mb-1">Ürün Açıklaması</label>
                    <textarea x-model="yeniUrun.urun_aciklamasi" rows="2" class="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:outline-none"></textarea>
                </div>
                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-semibold text-slate-600 uppercase mb-1">Adet</label>
                        <input type="number" step="any" x-model.number="yeniUrun.adet" required class="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold text-slate-600 uppercase mb-1">Birim Fiyat</label>
                        <input type="number" step="any" x-model.number="yeniUrun.birim_fiyat" required class="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                    </div>
                </div>
                
                <div class="pt-4 flex justify-end gap-3">
                    <button type="button" @click="modalAc = false" class="px-5 py-2.5 bg-slate-100 text-slate-600 rounded-xl font-medium hover:bg-slate-200 transition">İptal</button>
                    <button type="submit" class="px-5 py-2.5 bg-indigo-600 text-white rounded-xl font-medium hover:bg-indigo-700 shadow-lg shadow-indigo-200 transition">Kaydet</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Footer -->
    <footer class="text-center py-6 text-xs text-slate-400 border-t border-slate-200 bg-white mt-auto">
        Özdisan Excel Veritabanı Yöneticisi &bull; Flask & Tailwind CSS Destekli
    </footer>

    <!-- Script Mantığı -->
    <script>
        function appData() {
            return {
                veriler: [],
                filtrelenmisVeriler: [],
                aramaMetni: '',
                modalAc: false,
                mesaj: '',
                mesajTuru: 'basarili',
                yeniUrun: {
                    urun_kodu: '',
                    uretici: '',
                    urun_aciklamasi: '',
                    adet: 1,
                    birim_fiyat: 0,
                    lokasyon: ''
                },
                init() {
                    this.veriGetir();
                },
                veriGetir() {
                    fetch('/api/veriler')
                        .then(res => res.json())
                        .then(data => {
                            this.veriler = data;
                            this.veriFiltrele();
                        });
                },
                veriFiltrele() {
                    if (!this.aramaMetni) {
                        this.filtrelenmisVeriler = this.veriler;
                    } else {
                        const arama = this.aramaMetni.toLowerCase();
                        this.filtrelenmisVeriler = this.veriler.filter(item => {
                            return String(item['Ürün Kodu']).toLowerCase().includes(arama) ||
                                   String(item['Üretici']).toLowerCase().includes(arama) ||
                                   String(item['Ürün Açıklaması']).toLowerCase().includes(arama);
                        });
                    }
                },
                veriEkle() {
                    fetch('/api/ekle', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(this.yeniUrun)
                    })
                    .then(res => res.json())
                    .then(data => {
                        if (data.hata) {
                            this.mesaj = data.hata;
                            this.mesajTuru = 'hata';
                        } else {
                            this.mesaj = data.mesaj;
                            this.mesajTuru = 'basarili';
                            this.veriGetir();
                            this.modalAc = false;
                            this.yeniUrun = { urun_kodu: '', uretici: '', urun_aciklamasi: '', adet: 1, birim_fiyat: 0, lokasyon: '' };
                        }
                    });
                },
                veriSil(id) {
                    if (!confirm("Bu ürünü silmek istediğinize emin misiniz?")) return;
                    
                    fetch('/api/sil', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ id: id })
                    })
                    .then(res => res.json())
                    .then(data => {
                        if (data.hata) {
                            this.mesaj = data.hata;
                            this.mesajTuru = 'hata';
                        } else {
                            this.mesaj = data.mesaj;
                            this.mesajTuru = 'basarili';
                            this.veriGetir();
                        }
                    });
                }
            }
        }
    </script>
</body>
</html>
"""


@app.route("/")
def anasayfa():
  return render_template_string(HTML_SABLONU)


@app.route("/api/veriler", methods=["GET"])
def api_veriler():
  df = verileri_oku()
  if df.empty:
    return jsonify([])
  return jsonify(df.to_dict(orient="records"))


@app.route("/api/ekle", methods=["POST"])
def api_ekle():
  try:
    veri = request.json
    df = verileri_oku()

    yeni_id = 1 if df.empty or "#" not in df.columns else int(df["#"].max()) + 1

    adet = float(veri.get("adet", 0))
    birim_fiyat = float(veri.get("birim_fiyat", 0))
    toplam_fiyat = adet * birim_fiyat

    yeni_satir = {
        "#": yeni_id,
        "Ürün Kodu": veri.get("urun_kodu", ""),
        "Üretici": veri.get("uretici", ""),
        "Ürün Açıklaması": veri.get("urun_aciklamasi", ""),
        "Adet": adet,
        "Birim Fiyat": birim_fiyat,
        "Toplam Fiyat (KDV Dahil)": toplam_fiyat,
        "Lokasyon": veri.get("lokasyon", ""),
    }

    df = pd.concat([df, pd.DataFrame([yeni_satir])], ignore_index=True)
    verileri_kaydet(df)

    return jsonify({"mesaj": "Ürün başarıyla eklendi ve Excel dosyasına kaydedildi."})
  except Exception as e:
    return jsonify({"hata": str(e)}), 400


@app.route("/api/sil", methods=["POST"])
def api_sil():
  try:
    veri = request.json
    silinecek_id = veri.get("id")

    df = verileri_oku()
    if df.empty or "#" not in df.columns:
      return jsonify({"hata": "Veritabanı boş veya geçersiz!"}), 400

    df = df[df["#"] != silinecek_id]
    verileri_kaydet(df)

    return jsonify({"mesaj": "Ürün başarıyla listeden çıkarıldı."})
  except Exception as e:
    return jsonify({"hata": str(e)}), 400


if __name__ == "__main__":
  print("Uygulama başlatılıyor... Tarayıcınızda http://127.0.0.1:5000 adresine gidin.")
  app.run(debug=True, port=5000)
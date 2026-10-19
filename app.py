from datetime import datetime
import os
import base64
import requests
from flask import Flask, flash, redirect, render_template_string, request, session, url_for
import pandas as pd

app = Flask(__name__)
app.secret_key = "yilmaz_elektrik_gizli_anahtar"

# GİRİŞ ŞİFRESİ
SIFRE = "4072"

# GITHUB AYARLARI
GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN")
GITHUB_REPO = os.environ.get("GITHUB_REPO")
GITHUB_BRANCH = os.environ.get("GITHUB_BRANCH", "main")
DOSYA_YOLU = "ListeOZdisan.xlsx"


def github_dosyayi_oku():
  if GITHUB_TOKEN and GITHUB_REPO:
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{DOSYA_YOLU}"
    headers = {"Authorization": f"token {GITHUB_TOKEN}"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
      file_info = response.json()
      download_url = file_info["download_url"]
      file_data = requests.get(download_url).content
      with open(DOSYA_YOLU, "wb") as f:
        f.write(file_data)

  if not os.path.exists(DOSYA_YOLU):
    df = pd.DataFrame(
        columns=[
            "#",
            "Ürün Kodu",
            "Üretici",
            "Ürün Açıklaması",
            "Adet",
            "Birim Fiyat",
            "Toplam Fiyat (KDV Dahil)",
            "Lokasyon",
        ]
    )
    with pd.ExcelWriter(DOSYA_YOLU, engine="openpyxl") as writer:
      df.to_excel(writer, sheet_name="Sipariş Detay", index=False)
  return pd.read_excel(DOSYA_YOLU, sheet_name="Sipariş Detay")


def github_dosyayi_kaydet(mesaj):
  if not (GITHUB_TOKEN and GITHUB_REPO):
    return

  url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{DOSYA_YOLU}"
  headers = {"Authorization": f"token {GITHUB_TOKEN}"}

  get_resp = requests.get(url, headers=headers)
  sha = get_resp.json().get("sha") if get_resp.status_code == 200 else None

  with open(DOSYA_YOLU, "rb") as f:
    content_bytes = f.read()
  base64_content = base64.b64encode(content_bytes).decode("utf-8")

  data = {
      "message": mesaj,
      "content": base64_content,
      "branch": GITHUB_BRANCH,
  }
  if sha:
    data["sha"] = sha

  requests.put(url, headers=headers, json=data)


HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Yılmaz Elektrik Elektronik Malzeme</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-50 text-gray-800 font-sans">
    <div class="container mx-auto px-4 py-8 max-w-7xl">
        <header class="mb-8 flex flex-col md:flex-row justify-between items-center bg-white p-6 rounded-xl shadow-md border-t-4 border-yellow-500">
            <div>
                <h1 class="text-3xl font-bold text-gray-900">Yılmaz Elektrik Elektronik</h1>
                <p class="text-sm text-gray-500 mt-1">Malzeme Stok ve Veritabanı Yönetim Sistemi</p>
            </div>
            <div class="mt-4 md:mt-0 flex gap-3 items-center">
                <button onclick="openModal('addModal')" class="bg-blue-600 hover:bg-blue-700 text-white font-medium px-4 py-2 rounded-lg shadow transition">+ Yeni Ürün Ekle</button>
                <a href="/cikis" class="bg-red-50 hover:bg-red-100 text-red-600 border border-red-200 font-medium px-4 py-2 rounded-lg transition text-sm">Çıkış Yap</a>
            </div>
        </header>

        <div class="bg-white p-4 rounded-xl shadow-md mb-6 flex flex-col md:flex-row gap-4 justify-between items-center">
            <form action="/" method="GET" class="w-full flex gap-2">
                <input type="text" name="q" value="{{ request.args.get('q', '') }}" placeholder="Ürün Kodu, Üretici veya Açıklamada ara..." class="flex-grow border border-gray-300 rounded-lg px-4 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500">
                <button type="submit" class="bg-gray-800 hover:bg-gray-900 text-white px-5 py-2 rounded-lg font-medium transition">Ara</button>
                {% if request.args.get('q') %}
                    <a href="/" class="bg-gray-200 hover:bg-gray-300 text-gray-700 px-4 py-2 rounded-lg font-medium flex items-center">Temizle</a>
                {% endif %}
            </form>
        </div>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="mb-6 p-4 rounded-lg {{ 'bg-green-100 text-green-700 border border-green-200' if category == 'success' else 'bg-red-100 text-red-700 border border-red-200' }}">
                        {{ message }}
                    </div>
                {% endfor %}
            {% endif %}
        {% endwith %}

        <div class="bg-white rounded-xl shadow-md overflow-hidden">
            <div class="overflow-x-auto">
                <table class="w-full text-left border-collapse">
                    <thead>
                        <tr class="bg-gray-100 text-gray-600 uppercase text-xs tracking-wider border-b">
                            <th class="py-3 px-4">#</th>
                            <th class="py-3 px-4">Ürün Kodu</th>
                            <th class="py-3 px-4">Üretici</th>
                            <th class="py-3 px-4">Ürün Açıklaması</th>
                            <th class="py-3 px-4">Adet</th>
                            <th class="py-3 px-4">Birim Fiyat</th>
                            <th class="py-3 px-4">Toplam Fiyat</th>
                            <th class="py-3 px-4">Lokasyon</th>
                            <th class="py-3 px-4 text-center">İşlemler</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-gray-200 text-sm">
                        {% if df.empty %}
                            <tr>
                                <td colspan="9" class="text-center py-6 text-gray-500">Kayıt bulunamadı.</td>
                            </tr>
                        {% else %}
                            {% for index, row in df.iterrows() %}
                            <tr class="hover:bg-gray-50 transition">
                                <td class="py-3 px-4 font-semibold text-gray-600">{{ row['#'] }}</td>
                                <td class="py-3 px-4 font-medium text-blue-600">{{ row['Ürün Kodu'] }}</td>
                                <td class="py-3 px-4">{{ row['Üretici'] }}</td>
                                <td class="py-3 px-4 max-w-xs truncate" title="{{ row['Ürün Açıklaması'] }}">{{ row['Ürün Açıklaması'] }}</td>
                                <td class="py-3 px-4">{{ row['Adet'] }}</td>
                                <td class="py-3 px-4">{{ row['Birim Fiyat'] }}</td>
                                <td class="py-3 px-4 font-medium">{{ "%.2f"|format(row['Toplam Fiyat (KDV Dahil)']) }}</td>
                                <td class="py-3 px-4"><span class="bg-blue-50 text-blue-700 px-2 py-1 rounded text-xs font-semibold">{{ row['Lokasyon'] }}</span></td>
                                <td class="py-3 px-4 text-center space-x-2">
                                    <button onclick="openEditModal('{{ row['#'] }}', '{{ row['Ürün Kodu'] }}', '{{ row['Üretici'] }}', `{{ row['Ürün Açıklaması'] }}`, '{{ row['Adet'] }}', '{{ row['Birim Fiyat'] }}', '{{ row['Lokasyon'] }}')" class="text-blue-600 hover:text-blue-800 font-medium text-xs bg-blue-50 hover:bg-blue-100 px-2.5 py-1.5 rounded transition">Düzenle</button>
                                    <a href="/sil/{{ row['#'] }}" onclick="return confirm('Bu ürünü silmek istediğinize emin misiniz?')" class="text-red-500 hover:text-red-700 font-medium text-xs bg-red-50 hover:bg-red-100 px-2.5 py-1.5 rounded transition">Sil</a>
                                </td>
                            </tr>
                            {% endfor %}
                        {% endif %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Ekleme Modalı -->
    <div id="addModal" class="fixed inset-0 bg-black bg-opacity-50 hidden flex items-center justify-center p-4 z-50">
        <div class="bg-white rounded-xl max-w-lg w-full p-6 shadow-2xl">
            <h2 class="text-xl font-bold mb-4 text-gray-800">Yeni Ürün Ekle</h2>
            <form action="/ekle" method="POST" class="space-y-4">
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Ürün Kodu</label>
                    <input type="text" name="urun_kodu" required class="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none">
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Üretici</label>
                    <input type="text" name="uretici" required class="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none">
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Ürün Açıklaması</label>
                    <textarea name="urun_aciklamasi" rows="2" required class="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none"></textarea>
                </div>
                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <label class="block text-sm font-medium text-gray-700 mb-1">Adet</label>
                        <input type="number" step="any" name="adet" required class="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none">
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-gray-700 mb-1">Birim Fiyat</label>
                        <input type="number" step="any" name="birim_fiyat" required class="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none">
                    </div>
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Lokasyon</label>
                    <input type="text" name="lokasyon" required class="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none">
                </div>
                <div class="flex justify-end gap-3 mt-6">
                    <button type="button" onclick="closeModal('addModal')" class="bg-gray-200 hover:bg-gray-300 text-gray-700 px-4 py-2 rounded-lg font-medium transition">İptal</button>
                    <button type="submit" class="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2 rounded-lg font-medium transition">Kaydet</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Düzenleme Modalı -->
    <div id="editModal" class="fixed inset-0 bg-black bg-opacity-50 hidden flex items-center justify-center p-4 z-50">
        <div class="bg-white rounded-xl max-w-lg w-full p-6 shadow-2xl">
            <h2 class="text-xl font-bold mb-4 text-gray-800">Ürün Bilgilerini Güncelle</h2>
            <form action="/guncelle" method="POST" class="space-y-4">
                <input type="hidden" id="edit_id" name="id">
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Ürün Kodu</label>
                    <input type="text" id="edit_urun_kodu" name="urun_kodu" required class="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none">
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Üretici</label>
                    <input type="text" id="edit_uretici" name="uretici" required class="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none">
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Ürün Açıklaması</label>
                    <textarea id="edit_urun_aciklamasi" name="urun_aciklamasi" rows="2" required class="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none"></textarea>
                </div>
                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <label class="block text-sm font-medium text-gray-700 mb-1">Adet</label>
                        <input type="number" step="any" id="edit_adet" name="adet" required class="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none">
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-gray-700 mb-1">Birim Fiyat</label>
                        <input type="number" step="any" id="edit_birim_fiyat" name="birim_fiyat" required class="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none">
                    </div>
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Lokasyon</label>
                    <input type="text" id="edit_lokasyon" name="lokasyon" required class="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:outline-none">
                </div>
                <div class="flex justify-end gap-3 mt-6">
                    <button type="button" onclick="closeModal('editModal')" class="bg-gray-200 hover:bg-gray-300 text-gray-700 px-4 py-2 rounded-lg font-medium transition">İptal</button>
                    <button type="submit" class="bg-yellow-600 hover:bg-yellow-700 text-white px-5 py-2 rounded-lg font-medium transition">Güncellemeyi Kaydet</button>
                </div>
            </form>
        </div>
    </div>

    <script>
        function openModal(id) { document.getElementById(id).classList.remove('hidden'); }
        function closeModal(id) { document.getElementById(id).classList.add('hidden'); }
        function openEditModal(id, kod, uretici, aciklama, adet, fiyat, lokasyon) {
            document.getElementById('edit_id').value = id;
            document.getElementById('edit_urun_kodu').value = kod;
            document.getElementById('edit_uretici').value = uretici;
            document.getElementById('edit_urun_aciklamasi').value = aciklama;
            document.getElementById('edit_adet').value = adet;
            document.getElementById('edit_birim_fiyat').value = fiyat;
            document.getElementById('edit_lokasyon').value = lokasyon;
            openModal('editModal');
        }
    </script>
</body>
</html>
"""

LOGIN_TEMPLATE = """
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Giriş - Yılmaz Elektrik</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-100 flex items-center justify-center h-screen">
    <div class="bg-white p-8 rounded-2xl shadow-xl max-w-sm w-full border-t-4 border-yellow-500">
        <h2 class="text-2xl font-bold text-center text-gray-900 mb-2">Yılmaz Elektrik</h2>
        <p class="text-sm text-center text-gray-500 mb-6">Lütfen 4 haneli şifrenizi girin</p>
        
        {% if error %}
            <div class="mb-4 p-3 bg-red-100 border border-red-200 text-red-600 rounded-lg text-sm text-center">
                {{ error }}
            </div>
        {% endif %}

        <form method="POST" class="space-y-4">
            <div>
                <input type="password" name="sifre" placeholder="Şifre (4072)" required autofocus class="w-full border border-gray-300 rounded-lg px-4 py-3 text-center tracking-widest text-lg focus:ring-2 focus:ring-yellow-500 focus:outline-none">
            </div>
            <button type="submit" class="w-full bg-yellow-500 hover:bg-yellow-600 text-white font-semibold py-3 rounded-lg transition shadow">Sisteme Giriş Yap</button>
        </form>
    </div>
</body>
</html>
"""


@app.route("/giris", methods=["GET", "POST"])
def giris():
  error = None
  if request.method == "POST":
    if request.form.get("sifre") == SIFRE:
      session["giris_yapildi"] = True
      return redirect(url_for("index"))
    else:
      error = "Hatalı şifre! Lütfen tekrar deneyin."
  return render_template_string(LOGIN_TEMPLATE, error=error)


@app.route("/cikis")
def cikis():
  session.pop("giris_yapildi", None)
  return redirect(url_for("giris"))


@app.route("/")
def index():
  if not session.get("giris_yapildi"):
    return redirect(url_for("giris"))

  df = github_dosyayi_oku()
  q = request.args.get("q", "").strip().lower()
  if q:
    kosul = (
        df["Ürün Kodu"].astype(str).str.lower().str.contains(q, na=False)
        | df["Üretici"].astype(str).str.lower().str.contains(q, na=False)
        | df["Ürün Açıklaması"].astype(str).str.lower().str.contains(q, na=False)
    )
    df = df[kosul]
  return render_template_string(HTML_TEMPLATE, df=df)


@app.route("/ekle", methods=["POST"])
def ekle():
  if not session.get("giris_yapildi"):
    return redirect(url_for("giris"))

  df = github_dosyayi_oku()
  yeni_id = 1 if df.empty or df["#"].dropna().empty else int(df["#"].max()) + 1

  adet = float(request.form["adet"])
  birim_fiyat = float(request.form["birim_fiyat"])
  toplam_fiyat = adet * birim_fiyat

  yeni_kayit = {
      "#": yeni_id,
      "Ürün Kodu": request.form["urun_kodu"],
      "Üretici": request.form["uretici"],
      "Ürün Açıklaması": request.form["urun_aciklamasi"],
      "Adet": adet,
      "Birim Fiyat": birim_fiyat,
      "Toplam Fiyat (KDV Dahil)": toplam_fiyat,
      "Lokasyon": request.form["lokasyon"],
  }

  df = pd.concat([df, pd.DataFrame([yeni_kayit])], ignore_index=True)
  with pd.ExcelWriter(DOSYA_YOLU, engine="openpyxl", mode="w") as writer:
    df.to_excel(writer, sheet_name="Sipariş Detay", index=False)

  github_dosyayi_kaydet(f"Yeni ürün eklendi: ID {yeni_id}")
  flash("Ürün başarıyla eklendi!", "success")
  return redirect(url_for("index"))


@app.route("/guncelle", methods=["POST"])
def guncelle():
  if not session.get("giris_yapildi"):
    return redirect(url_for("giris"))

  df = github_dosyayi_oku()
  urun_id = int(request.form["id"])

  if urun_id in df["#"].values:
    adet = float(request.form["adet"])
    birim_fiyat = float(request.form["birim_fiyat"])
    toplam_fiyat = adet * birim_fiyat

    df.loc[df["#"] == urun_id, "Ürün Kodu"] = request.form["urun_kodu"]
    df.loc[df["#"] == urun_id, "Üretici"] = request.form["uretici"]
    df.loc[df["#"] == urun_id, "Ürün Açıklaması"] = request.form[
        "urun_aciklamasi"
    ]
    df.loc[df["#"] == urun_id, "Adet"] = adet
    df.loc[df["#"] == urun_id, "Birim Fiyat"] = birim_fiyat
    df.loc[df["#"] == urun_id, "Toplam Fiyat (KDV Dahil)"] = toplam_fiyat
    df.loc[df["#"] == urun_id, "Lokasyon"] = request.form["lokasyon"]

    with pd.ExcelWriter(DOSYA_YOLU, engine="openpyxl", mode="w") as writer:
      df.to_excel(writer, sheet_name="Sipariş Detay", index=False)

    github_dosyayi_kaydet(f"Ürün güncellendi: ID {urun_id}")
    flash("Ürün başarıyla güncellendi!", "success")
  else:
    flash("Güncellenecek ürün bulunamadı!", "error")

  return redirect(url_for("index"))


@app.route("/sil/<int:urun_id>")
def sil(urun_id):
  if not session.get("giris_yapildi"):
    return redirect(url_for("giris"))

  df = github_dosyayi_oku()
  if urun_id in df["#"].values:
    df = df[df["#"] != urun_id]
    with pd.ExcelWriter(DOSYA_YOLU, engine="openpyxl", mode="w") as writer:
      df.to_excel(writer, sheet_name="Sipariş Detay", index=False)

    github_dosyayi_kaydet(f"Ürün silindi: ID {urun_id}")
    flash("Ürün başarıyla silindi!", "success")
  else:
    flash("Silinecek ürün bulunamadı!", "error")

  return redirect(url_for("index"))


if __name__ == "__main__":
  app.run(debug=True)

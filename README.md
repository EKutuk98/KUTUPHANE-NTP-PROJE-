# 📚 Dijital Kütüphane Sistemi

Kitap ve üye yönetimi, ödünç işlemleri, anomali tespiti ve çok şubeli yapıyı destekleyen profesyonel bir kütüphane yönetim platformu. PyQt5 ile geliştirilmiş, SQLite veritabanı destekli, gerçek zamanlı analog saat widget'ı ve XP tabanlı üye gamification sistemi içermektedir.

**🔐 Giriş Bilgileri:**
- Admin: `kutuphane` / `12345`




---

## 📋 Özellikler

---

### 📊 Dashboard

- Toplam kitap, aktif üye, bugünkü ödünç ve gecikmiş iade KPI kartları
- QPainter bar grafik (kategori bazlı kitap dağılımı)
- QPainter pasta grafik (kitap durum dağılımı — Mevcut / Ödünçte / Kayıp)
- Gerçek zamanlı analog saat widget (her saniye güncellenir)
- Son ödünç işlemleri feed



---

### 📚 Kitaplar

- Kitap ekleme, düzenleme, soft-delete
- Ad, yazar, kategori, yayın yılı, sayfa sayısı, dil, açıklama
- Kitap durumu: Mevcut / Ödünçte / Kayıp / Bakımda
- Kategori ve şube bazlı filtreleme
- Anlık arama (ad, yazar, kategori)



---

### 👥 Üyeler

- Üye ekleme, düzenleme, soft-delete
- Üye no (otomatik), ad, soyad, email, telefon, adres
- Email ve üye no benzersizlik kontrolü
- Üye bazlı ödünç geçmişi
- XP ve seviye gösterimi (Gamification)
- Anlık arama



---

### 📖 Ödünç İşlemleri

- Kitap ödünç verme (üye + kitap seçimi, iade tarihi)
- Kitap iade alma (gecikme ücreti otomatik hesaplama)
- Durum: Ödünçte / İade Edildi / Gecikmiş
- Gecikmiş iadeler için kırmızı vurgu
- XP kazanımı: iade edilen kitap başına otomatik XP ekleme
- Anlık arama ve durum filtresi




---

### 📄 Raporlar

- Aktif ödünç listesi raporu
- Gecikmiş iade raporu (gecikme ücreti dahil)
- En çok ödünç alınan kitaplar
- Üye aktivite raporu
- CSV dışa aktarma



---

### ⚠️ Anomali Tespiti

- Otomatik anormallik tespiti (olağandışı ödünç örüntüleri)
- Aynı kitabı defalarca ödünç alan üyeler
- Hiç iade etmemiş üyeler
- Uzun süreli gecikme tespiti
- Anomali listesi ve önem derecesi



---

### 🏢 Şubeler *(Tier 6)*

- Çoklu şube yönetimi
- Şube ekleme, düzenleme
- Şube adı, şehir, adres, telefon, yönetici
- Kitap ve üyelerin şubeye atanması
- Şube bazlı istatistikler


---

### 📋 Audit Log *(Tier 6)*

- Sistemdeki tüm işlemlerin kaydı (kim, ne zaman, ne yaptı, hangi tablo)
- İşlem türüne göre filtreleme (Ekle / Güncelle / Sil / Giriş)
- Kullanıcı bazlı log görüntüleme



---

### 🎮 Gamification (Üye XP Sistemi)

- Kitap iadesi başına XP kazanımı
- Seviye sistemi: 🥉 Bronze (0–99 XP) / 🥈 Silver (100–499 XP) / 🥇 Gold (500+ XP)
- Üye listesinde seviye gösterimi
- Seviyeye göre üye sıralaması

---

## ❓ Final Soruları

### Sistemde hangi kullanıcılar veya nesneler vardır?

**Kullanıcı Türleri:**
- **Admin (Kütüphane Yöneticisi)** — Tam yetkili; kitap, üye, ödünç, şube, kullanıcı yönetimi ve raporlar
- **Personel** — Kitap, üye ve ödünç işlemleri (rol izinlerine göre kısıtlı)

**Roller:** sistem içinde tanımlı izin tablosuna göre modüler yetkilendirme (rapor_goster, anomali_goster, sube_yonet, kullanici_yonet, audit_log_goster vb.)

**Ana Nesneler / Varlıklar:**
- **Kitap** — kitap_id, şube, ad, yazar, kategori, yayın yılı, sayfa sayısı, dil, durum, açıklama
- **Üye** — uye_id, şube, üye no, ad, soyad, email, telefon, adres, durum
- **Ödünç** — odunc_id, şube, kitap, üye, ödünç tarihi, iade tarihi, durum, gecikme ücreti
- **Kategori** — kategori_id, ad, açıklama
- **Şube** — sube_id, ad, şehir, adres, telefon, yönetici
- **Üye Gamification** — uye_id, toplam XP, seviye
- **Rol** — rol_id, rol adı, açıklama
- **Rol İzni** — rol_id, izin adı
- **Sistem Kullanıcısı** — kullanici_id, şube, rol, kullanıcı adı, şifre, ad, soyad
- **Audit Log** — log_id, kullanıcı, işlem, tablo, kayıt_id, açıklama, zaman

---

### Kullanıcı sistemde hangi işlemleri gerçekleştirebilir?

**Admin:**
- Kitap ekleyebilir, düzenleyebilir, silebilir, kategori yönetebilir
- Üye ekleyebilir, düzenleyebilir, pasife alabilir
- Kitap ödünç verebilir, iade alabilir, gecikme ücreti hesaplayabilir
- Raporları görüntüleyebilir ve CSV export edebilir
- Anomali tespiti çalıştırabilir ve sonuçları görüntüleyebilir
- Şube ekleyebilir, düzenleyebilir
- Kullanıcı ekleyebilir, rol atayabilir
- Audit log kaydlarını görüntüleyebilir
- XP ve üye seviyelerini takip edebilir

**Personel (rol izinlerine göre):**
- Kitap ekleyebilir ve güncelleyebilir
- Üye ekleyebilir ve güncelleyebilir
- Ödünç verme ve iade alma işlemi yapabilir
- Temel raporları görüntüleyebilir

---

## 🖥️ Teknolojiler

| Teknoloji | Kullanım Alanı |
|-----------|----------------|
| Python 3.9+ | Ana programlama dili |
| PyQt5 | GUI Framework |
| SQLite3 | Veritabanı yönetimi (`@contextmanager`) |
| QPainter | Grafikler (Bar, Pie) + Analog Saat Widget |
| hashlib | Şifre güvenliği |
| csv | Rapor export |

# Excel -> JSON -> FlatBuffers (.bin) Converter

## 1. Projenin Amacı

Bu proje, belirli bir formatta hazırlanmış **Excel dosyalarından**:

- Seçilen isteğe (request type) uygun **JSON** üretmek
- Bu JSON’u **FlatBuffers** kullanarak **.bin** formatına çevirmek
- Üretilen **.bin** dosyasını tekrar **JSON’a çevirerek doğrulamak**

amacıyla geliştirilmiştir.

Bu sayede:

- Konfigürasyon hataları erken aşamada yakalanır
- FlatBuffers dönüşümünün kayıpsız olduğu garanti edilir
- Cihaza gönderilecek veriler test edilebilir hale gelir

---

## 2. Desteklenen Request Türleri

Uygulama şu request türlerini destekler:

- `channel_transfer`
- `channel_configure`
- `test_add_directives`
- `test_prepare`

> Kullanıcı bir tür seçtiğinde, uygulama **beklenen Excel formatını ekranda açıkça gösterir.**

---

## 3. Proje Klasör Yapısı

excel-to-json-converter/
?
?? ExcelToJsonConverter.App/
? ?? MainWindow.axaml
? ?? MainWindow.axaml.cs
? ?? ExcelToJsonConverter.App.csproj
?
?? schemas/
? ?? rft.fbs # FlatBuffers schema
?
?? Tools/
? ?? flatbuffers/
? ?? win-x64/
? ?? flatc.exe # FlatBuffers compiler
?
?? out/ # Otomatik üretilen çıktılar
? ?? <ExcelAdı>.json
? ?? <ExcelAdı>.bin
? ?? verify/
? ?? *.json # BIN ? JSON doğrulama çıktıları
?
?? README.md

?? **out/** klasörü çıktı klasörüdür, repoya eklenmesi önerilmez.

---

## 4. Gereksinimler

- Windows
- .NET SDK (projede kullanılan sürüm)
- Excel dosyaları (`.xlsx`, `.xlsm`)
- FlatBuffers compiler (`flatc.exe`)

Proje içinde: Tools/flatbuffers/win-x64/flatc.exe

---

## 5. Uygulamanın Kullanımı (Adım Adım)

### Adım 1 – Uygulamayı Çalıştır

Visual Studio veya `dotnet run` ile uygulamayı başlat.

?? **Ekran Görüntüsü (Ana ekran)**  

---

### Adım 2 – Excel Dosyasını Seç

`Pick Excel` butonuna basarak Excel dosyasını seç.

?? **Ekran Görüntüsü (Excel seçimi)**  

---

### Adım 3 – Request Türünü Seç

ComboBox üzerinden request türünü seç:

- `channel_configure`
- `channel_transfer`
- vb.

Seçim yapıldığında, sağ tarafta **beklenen Excel formatı** otomatik gösterilir.

?? **Ekran Görüntüsü (Tür seçimi + format açıklaması)**  

---

### Adım 4 – Convert (Excel ? JSON)

`Convert` butonuna basıldığında:

- Excel okunur
- Seçilen tipe uygun JSON üretilir
- JSON, Excel ile **aynı isimle** kaydedilir

Örnek:

?? **Ekran Görüntüsü (Convert sonrası başarılı çıktı)**  

---

### Adım 5 – Update (JSON ? BIN + Doğrulama)

`JSON ? BIN (Update)` butonuna basıldığında:

- JSON `out/` klasörüne yazılır
- `flatc.exe` çağrılır
- `.bin` dosyası üretilir
- Üretilen `.bin`, tekrar JSON’a çevrilir
- Doğrulama çıktısı `out/verify/` altına yazılır

?? Örnek çıktı:
out/
?? Channel_Configure.json
?? Channel_Configure.bin
?? verify/
?? Channel_Configure.json



?? **Ekran Görüntüsü (Update + doğrulama sonucu)**  

---

## 6. Round-Trip Doğrulama Nedir?

Bu proje şu akışı doğrular: Excel ? JSON ? BIN ? JSON


Amaç:

- FlatBuffers dönüşümünde veri kaybı var mı?
- JSON schema ile birebir uyumlu mu?

Bu adım özellikle **saha ve cihaz entegrasyonu öncesi kritik öneme sahiptir.**

---

## 7. Teknik Detaylar

- Excel okuma: **ClosedXML**
- JSON üretimi: **System.Text.Json**
- FlatBuffers işlemleri: **flatc.exe**
- Proses yönetimi: **ProcessStartInfo**
- Hata yönetimi: **try/catch + kullanıcıya açıklayıcı mesajlar**

---

## 8. Bilinen Kısıtlar

- `channel_configure` için şu an yalnızca **RS-485** desteklenmektedir
- UDP desteği ileride eklenebilir
- Bazı request türleri örnek implementasyon olarak sınırlı satır sayısıyla çalışır

---

## 9. Geliştirici

**Melda Hacer Çetin**  
Software Engineering Student  
Excel ? JSON ? FlatBuffers Converter

---

## 10. Not

Bu proje:

- Konfigürasyon güvenliği
- Veri doğrulama
- Test öncesi hata yakalama

amaçlarıyla geliştirilmiştir.





# Excel ? JSON ? FlatBuffers (.bin) Converter

## 1. Projenin Amacý

Bu proje, belirli bir formatta hazýrlanmýþ **Excel dosyalarýndan**:

- Seçilen isteðe (request type) uygun **JSON** üretmek
- Bu JSON’u **FlatBuffers** kullanarak **.bin** formatýna çevirmek
- Üretilen **.bin** dosyasýný tekrar **JSON’a çevirerek doðrulamak**

amacýyla geliþtirilmiþtir.

Bu sayede:

- Konfigürasyon hatalarý erken aþamada yakalanýr
- FlatBuffers dönüþümünün kayýpsýz olduðu garanti edilir
- Cihaza gönderilecek veriler test edilebilir hale gelir

---

## 2. Desteklenen Request Türleri

Uygulama þu request türlerini destekler:

- `channel_transfer`
- `channel_configure`
- `test_add_directives`
- `test_prepare`

> Kullanýcý bir tür seçtiðinde, uygulama **beklenen Excel formatýný ekranda açýkça gösterir.**

---

## 3. Proje Klasör Yapýsý

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
?? out/ # Otomatik üretilen çýktýlar
? ?? <ExcelAdý>.json
? ?? <ExcelAdý>.bin
? ?? verify/
? ?? *.json # BIN ? JSON doðrulama çýktýlarý
?
?? README.md

?? **out/** klasörü çýktý klasörüdür, repoya eklenmesi önerilmez.

---

## 4. Gereksinimler

- Windows
- .NET SDK (projede kullanýlan sürüm)
- Excel dosyalarý (`.xlsx`, `.xlsm`)
- FlatBuffers compiler (`flatc.exe`)

Proje içinde: Tools/flatbuffers/win-x64/flatc.exe

---

## 5. Uygulamanýn Kullanýmý (Adým Adým)

### Adým 1 – Uygulamayý Çalýþtýr

Visual Studio veya `dotnet run` ile uygulamayý baþlat.

?? **Ekran Görüntüsü (Ana ekran)**  

---

### Adým 2 – Excel Dosyasýný Seç

`Pick Excel` butonuna basarak Excel dosyasýný seç.

?? **Ekran Görüntüsü (Excel seçimi)**  

---

### Adým 3 – Request Türünü Seç

ComboBox üzerinden request türünü seç:

- `channel_configure`
- `channel_transfer`
- vb.

Seçim yapýldýðýnda, sað tarafta **beklenen Excel formatý** otomatik gösterilir.

?? **Ekran Görüntüsü (Tür seçimi + format açýklamasý)**  

---

### Adým 4 – Convert (Excel ? JSON)

`Convert` butonuna basýldýðýnda:

- Excel okunur
- Seçilen tipe uygun JSON üretilir
- JSON, Excel ile **ayný isimle** kaydedilir

Örnek:

?? **Ekran Görüntüsü (Convert sonrasý baþarýlý çýktý)**  

---

### Adým 5 – Update (JSON ? BIN + Doðrulama)

`JSON ? BIN (Update)` butonuna basýldýðýnda:

- JSON `out/` klasörüne yazýlýr
- `flatc.exe` çaðrýlýr
- `.bin` dosyasý üretilir
- Üretilen `.bin`, tekrar JSON’a çevrilir
- Doðrulama çýktýsý `out/verify/` altýna yazýlýr

?? Örnek çýktý:
out/
?? Channel_Configure.json
?? Channel_Configure.bin
?? verify/
?? Channel_Configure.json



?? **Ekran Görüntüsü (Update + doðrulama sonucu)**  

---

## 6. Round-Trip Doðrulama Nedir?

Bu proje þu akýþý doðrular: Excel ? JSON ? BIN ? JSON


Amaç:

- FlatBuffers dönüþümünde veri kaybý var mý?
- JSON schema ile birebir uyumlu mu?

Bu adým özellikle **saha ve cihaz entegrasyonu öncesi kritik öneme sahiptir.**

---

## 7. Teknik Detaylar

- Excel okuma: **ClosedXML**
- JSON üretimi: **System.Text.Json**
- FlatBuffers iþlemleri: **flatc.exe**
- Proses yönetimi: **ProcessStartInfo**
- Hata yönetimi: **try/catch + kullanýcýya açýklayýcý mesajlar**

---

## 8. Bilinen Kýsýtlar

- `channel_configure` için þu an yalnýzca **RS-485** desteklenmektedir
- UDP desteði ileride eklenebilir
- Bazý request türleri örnek implementasyon olarak sýnýrlý satýr sayýsýyla çalýþýr

---

## 9. Geliþtirici

**Melda Hacer Çetin**  
Software Engineering Student  
Excel ? JSON ? FlatBuffers Converter

---

## 10. Not

Bu proje:

- Konfigürasyon güvenliði
- Veri doðrulama
- Test öncesi hata yakalama

amaçlarýyla geliþtirilmiþtir.





// MainWindow.axaml.cs
// Amaç:
// - Kullanıcıdan Excel dosyası seçimini alır.
// - Seçilen türe göre Excel -> JSON dönüşümü yapar.
// - JSON -> FlatBuffers BIN dönüşümü (flatc) ve isteğe bağlı doğrulama (BIN -> JSON) adımlarını yönetir.

using Avalonia.Controls;
using ClosedXML.Excel;           // Excel okuma (sheet/row/cell)
using System;
using System.Collections.Generic;
using System.Diagnostics;        // Process çalıştırma (flatc)
using System.IO;                 // Path, File, Directory işlemleri
using System.Linq;
using System.Text;               // Encoding
using System.Text.Json;          // JSON serialize/deserialize

namespace ExcelToJsonConverter.App;

// MainWindow:
// - Uygulamanın ana penceresidir.
// - UI (XAML) ve davranış (code-behind) kısımlarını birleştirir.
// - Kullanıcı etkileşimlerini (buton tıklamaları, seçim değişiklikleri)
//   ilgili event handler'lara yönlendirir.
public partial class MainWindow : Window
{
    public MainWindow()
    {
        // XAML tarafında tanımlanan tüm UI bileşenlerini yükler
        InitializeComponent();

        // UI event bağlamaları:
        // Kullanıcının yaptığı işlemler burada ilgili metotlara yönlendirilir.
        BtnPickExcel.Click += BtnPickExcel_Click;          // Excel dosyası seçimi
        CmbType.SelectionChanged += CmbType_SelectionChanged; // Dönüşüm tipi seçimi
        BtnConvert.Click += BtnConvert_Click;              // Excel -> JSON
        BtnUpdate.Click += BtnUpdate_Click;                // JSON -> BIN + doğrulama
    }

    // Kullanıcıdan Excel dosyası seçmesini ister.
    // async kullanılmasının nedeni, dosya seçim dialog'u açıkken
    // UI thread'in bloklanmamasını sağlamaktır.
    private async void BtnPickExcel_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        // Excel dosyası seçimi için standart dosya dialog'u
        var dlg = new OpenFileDialog
        {
            Title = "Excel dosyası seç",
            AllowMultiple = false, // Kullanıcı sadece tek bir Excel dosyası seçebilir
            Filters =
        {
            new FileDialogFilter
            {
                Name = "Excel",
                Extensions = { "xlsx", "xlsm" }
            }
        }
        };

        // Dialog'u aç ve kullanıcı seçim yapana kadar bekle
        var result = await dlg.ShowAsync(this);

        // Kullanıcı bir dosya seçtiyse, seçilen dosyanın yolunu ekranda göster
        if (result is { Length: > 0 })
            TxtExcelPath.Text = result[0];
    }


    // Dönüşüm tipi ComboBox'ında seçim değiştiğinde tetiklenir.
    // Seçilen tipe göre, kullanıcıya beklenen Excel sheet ve kolon formatını
    // açıklayıcı bir metin olarak UI üzerinde gösterir.
    // Amaç: Kullanıcının yanlış formatta Excel seçmesini önlemek.
    private void CmbType_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        // Kullanıcının seçtiği dönüşüm tipi (ComboBox içeriği)
        var selectedType =
            (CmbType.SelectedItem as ComboBoxItem)?.Content?.ToString()
            ?? "unknown";

        // Seçilen tipe göre beklenen Excel formatını ekranda göster
        TxtExpectedFormat.Text = selectedType switch
        {
            "channel_transfer" =>
                "Sheet: Channel_Transfer\n" +
                "Kolonlar:\n" +
                "- tx_channel_id\n" +
                "- rx_msg_length\n" +
                "- tx_msg (virgüllü örn.: 15, 61, 62)\n" +
                "- timeout_usec",

            "channel_configure" =>
                "Sheet: Channel_Configure\n" +
                "Kolonlar:\n" +
                "- channel_id\n" +
                "- rx_channel_id\n" +
                "- interface_config_type\n" +
                "- baud_rate\n" +
                "- stop_bit\nn" +
                "- data_bits\n" +
                "- parity\n" +
                "- termination (TRUE / FALSE)\n" +
                "- timeout_usec",

            "test_add_directives" =>
                "Sheet: Test_AddDirectives\n" +
                "Kolonlar:\n" +
                "- tx_channel_id\n" +
                "- rx_msg_length\n" +
                "- step_count\n" +
                "- tx_msg (virgüllü: 10,21,22,23)\n\n" +
                "Notlar:\n" +
                "- Her satır 1 directive tanımıdır.\n" +
                "- Aynı tx_channel_id birden fazla kez tekrar edebilir;\n" +
                "  JSON çıktısında tek başlık altında gruplanır.",

            "test_prepare" =>
                "Sheet'ler:\n\n" +

                "1) Test_Prepare_General\n" +
                "   - period_usec\n\n" +

                "2) Test_Prepare_Fields\n" +
                "   - field_source (Tx / Rx)\n" +
                "   - offset\n" +
                "   - scalar_type (U32 / F32 / U8)\n" +
                "   - big_endian (TRUE / FALSE)\n" +
                "   Not: Her satır 1 field tanımıdır.\n\n" +

                "3) Test_Prepare_Bindings\n" +
                "   - tx_channel_id\n" +
                "   - arg_name (response1, command2, ...)\n" +
                "   - arg_index\n" +
                "   Not: Aynı tx_channel_id tekrar edebilir;\n" +
                "   JSON çıktısında gruplanır.\n\n" +

                "4) Test_Prepare_Evaluations\n" +
                "   - evaluation_idx\n" +
                "   - instructions (virgüllü: 9,4,Sub)\n" +
                "   - instructions_type (virgüllü: load_field,binary_op)\n" +
                "   Not: instructions ve instructions_type\n" +
                "   eleman sayıları birebir eşleşmelidir.\n\n" +

                "5) Test_Prepare_Criteria\n" +
                "   - tx_channel_id\n" +
                "   - evaluation_idx\n" +
                "   - comparison_ops (virgüllü: Ge,Le)\n" +
                "   - comparison_values (virgüllü: 99.5,100.5)\n" +
                "   - invert_logic (TRUE / FALSE)\n" +
                "   - start_time_step\n" +
                "   - end_time_step\n\n" +
                "Not: Her satır 1 kriter tanımıdır.",

            // Beklenmeyen veya boş seçimler için kullanıcıyı yönlendir
            _ => "Lütfen bir dönüşüm tipi seçiniz."
        };
    }


    // Convert: Seçilen Excel dosyasını, seçilen tipe göre JSON'a dönüştürür.
    // Çıktı:
    // - JSON dosyası, Excel dosyasının bulunduğu klasöre <ExcelAdı>.json olarak yazılır.
    // - JSON metni önizleme alanında (TxtPreview) gösterilir.
    private void BtnConvert_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try
        {
            // UI mesajlarını temizle
            SuccessPanel.IsVisible = false;
            ErrorPanel.IsVisible = false;
            TxtResultPath.Text = "";
            TxtError.Text = "";

            // Kullanıcının seçtiği dönüşüm tipi
            var selectedType =
                (CmbType.SelectedItem as ComboBoxItem)?.Content?.ToString()
                ?? "unknown";

            // Excel dosyası seçilmiş mi?
            var excelPath = TxtExcelPath.Text ?? "";
            if (string.IsNullOrWhiteSpace(excelPath) || excelPath == "(Excel Secilmedi)")
            {
                TxtPreview.Text = "Önce Excel dosyası seçmelisin.";
                return;
            }

            // 1) Seçilen tipe göre Excel -> JSON üret
            string json = selectedType switch
            {
                "channel_transfer" => ConvertChannelTransferFromExcel(excelPath),
                "channel_configure" => ConvertChannelConfigureFromExcel(excelPath),
                "test_add_directives" => ConvertTestAddDirectivesFromExcel(excelPath),
                "test_prepare" => ConvertTestPrepareFromExcel(excelPath),
                _ => throw new Exception($"Bilinmeyen tür: {selectedType}")
            };

            // 2) JSON'u Excel dosyasının yanına, aynı adla kaydet (sadece uzantı .json olur)
            var directory = Path.GetDirectoryName(excelPath) ?? Environment.CurrentDirectory;
            var baseName = Path.GetFileNameWithoutExtension(excelPath);
            var jsonPath = Path.Combine(directory, $"{baseName}.json");

            File.WriteAllText(jsonPath, json, Encoding.UTF8);

            // 3) UI: önizleme + başarı mesajı
            TxtPreview.Text = json;
            SuccessPanel.IsVisible = true;
            ErrorPanel.IsVisible = false;

            TxtSuccessTitle.Text = "JSON başarıyla oluşturuldu:";
            TxtResultPath.Text = jsonPath;
        }
        catch (Exception ex)
        {
            // UI: hata mesajı
            TxtPreview.Text = $"Hata: {ex.Message}";
            TxtError.Text = ex.Message;

            ErrorPanel.IsVisible = true;
            SuccessPanel.IsVisible = false;
        }
    }


    // Update: Seçilen Excel dosyasını seçilen tipe göre JSON'a dönüştürür,
    // ardından flatc ile JSON -> BIN üretir.
    // Son olarak round-trip doğrulama için BIN -> JSON (verify) dönüşümünü yapar.
    //
    // Çıktılar:
    // - out/<ExcelAdı>.json
    // - out/<ExcelAdı>.bin
    // - out/verify/<...>.json   (flatc çıktı adı değişebilir)
    private void BtnUpdate_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try
        {
            // UI mesajlarını temizle
            SuccessPanel.IsVisible = false;
            ErrorPanel.IsVisible = false;
            TxtError.Text = "";
            TxtResultPath.Text = "";

            // Kullanıcının seçtiği dönüşüm tipi
            var selectedType =
                (CmbType.SelectedItem as ComboBoxItem)?.Content?.ToString()
                ?? "unknown";

            // Excel seçilmiş mi?
            var excelPath = TxtExcelPath.Text ?? "";
            if (string.IsNullOrWhiteSpace(excelPath) || excelPath == "(Excel Seçilmedi)")
                throw new Exception("Önce Excel dosyası seçmelisin.");

            // 1) Seçilen tipe göre Excel -> JSON üret
            string json = selectedType switch
            {
                "channel_transfer" => ConvertChannelTransferFromExcel(excelPath),
                "channel_configure" => ConvertChannelConfigureFromExcel(excelPath),
                "test_add_directives" => ConvertTestAddDirectivesFromExcel(excelPath),
                "test_prepare" => ConvertTestPrepareFromExcel(excelPath),
                _ => throw new Exception($"Bilinmeyen tür: {selectedType}")
            };

            // 2) Proje kökünü bul ve tüm çıktıları standart olarak out/ klasörüne yaz
            var projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
            var outDir = Path.Combine(projectRoot, "out");
            Directory.CreateDirectory(outDir);

            var baseName = Path.GetFileNameWithoutExtension(excelPath);

            // Excel adı + .json (ekstra suffix eklemeden)
            var jsonPath = Path.Combine(outDir, $"{baseName}.json");
            File.WriteAllText(jsonPath, json, Encoding.UTF8);

            // 3) flatc ve schema yollarını doğrula
            var schemaPath = Path.Combine(projectRoot, "schemas", "rft.fbs");
            if (!File.Exists(schemaPath))
                throw new Exception($"Şema bulunamadı: {schemaPath}");

            var flatcPath = Path.Combine(projectRoot, "Tools", "flatbuffers", "win-x64", "flatc.exe");
            if (!File.Exists(flatcPath))
                throw new Exception($"flatc bulunamadı: {flatcPath}");

            // 4) flatc ile JSON -> BIN üret
            // - root-type: Schema'daki root_type (RFT.Request) ile aynı olmalı
            // - strict-json: schema uyumsuzsa hata vererek güvenli dönüşüm sağlar
            var args =
                $"--binary --strict-json --root-type RFT.Request " +
                $"-o \"{outDir}\" \"{schemaPath}\" \"{jsonPath}\"";

            var (exitCode, stdout, stderr) = RunProcess(flatcPath, args, projectRoot);

            if (exitCode != 0)
            {
                throw new Exception(
                    "flatc (JSON -> BIN) hata verdi.\n\n" +
                    $"Komut:\n{flatcPath} {args}\n\n" +
                    $"STDOUT:\n{stdout}\n\nSTDERR:\n{stderr}"
                );
            }

            // flatc çıktısı beklenen yerde mi?
            var binPath = Path.Combine(outDir, $"{baseName}.bin");
            if (!File.Exists(binPath))
                throw new Exception($"flatc çalıştı ama .bin bulunamadı: {binPath}");

            // 5) Round-trip doğrulama: BIN -> JSON (verify)
            // Not: flatc JSON output dosya adını her zaman aynı isimle vermeyebilir;
            // bu nedenle verify klasöründe oluşan *.json dosyasını arıyoruz.
            var verifyDir = Path.Combine(outDir, "verify");
            Directory.CreateDirectory(verifyDir);

            var verifyArgs =
                $"--json --strict-json --raw-binary " +
                $"--root-type RFT.Request " +
                $"-o \"{verifyDir}\" " +
                $"\"{schemaPath}\" -- \"{binPath}\"";


            var (exit2, out2, err2) = RunProcess(flatcPath, verifyArgs, projectRoot);

            if (exit2 != 0)
            {
                throw new Exception(
                    "flatc (BIN -> JSON verify) hata verdi.\n\n" +
                    $"Komut:\n{flatcPath} {verifyArgs}\n\n" +
                    $"ExitCode: {exit2}\n\n" +
                    $"STDOUT:\n{out2}\n\n" +
                    $"STDERR:\n{err2}"
                );
            }

            // verify klasöründe beklenen JSON'u bul (Excel dosya adına göre)
            var expectedVerifyJsonPath = Path.Combine(verifyDir, $"{baseName}.json");

            // flatc bazen isimlendirmeyi farklı yaparsa diye fallback de ekleyelim:
            if (!File.Exists(expectedVerifyJsonPath))
            {
                var verifyJsonFiles = Directory.GetFiles(verifyDir, "*.json");
                if (verifyJsonFiles.Length == 0)
                    throw new Exception("flatc çalıştı ama verify klasöründe JSON oluşmadı.");

                // baseName geçen dosyayı ara
                var match = verifyJsonFiles.FirstOrDefault(f =>
                    Path.GetFileNameWithoutExtension(f)
                        .Equals(baseName, StringComparison.OrdinalIgnoreCase));

                if (match == null)
                {
                    // hiç eşleşme yoksa, mevcutları listeleyip hatayı anlaşılır ver
                    var list = string.Join("\n", verifyJsonFiles.Select(Path.GetFileName));
                    throw new Exception(
                        $"Doğrulama JSON bulunamadı.\nBeklenen: {baseName}.json\n\nVerify klasöründekiler:\n{list}"
                    );
                }

                expectedVerifyJsonPath = match;
            }

            var verifyJsonPath = expectedVerifyJsonPath;
            var verifyJsonText = File.ReadAllText(verifyJsonPath, Encoding.UTF8);

            // 6) JSON karşılaştırma
            bool same = JsonDeepEquals(json, verifyJsonText, out var diffMsg);

            if (!same)
            {
                throw new Exception(
                    "JSON doğrulama başarısız!\n\n" +
                    diffMsg + "\n\n" +
                    $"VERIFY JSON: {verifyJsonPath}"
                );
            }

            //Başarılı UI çıktısı
            TxtPreview.Text = json; // istersen verifyJsonText de gösterebilirsin
            SuccessPanel.IsVisible = true;
            ErrorPanel.IsVisible = false;

            TxtSuccessTitle.Text = "BIN oluşturuldu, JSON doğrulama başarılı.";
            TxtResultPath.Text = $"BIN: {binPath}\nVERIFY JSON: {verifyJsonPath}";
        }
        catch (Exception ex)
        {
            // UI: hata durumunda kullanıcıya mesaj göster
            TxtError.Text = ex.Message;
            ErrorPanel.IsVisible = true;
            SuccessPanel.IsVisible = false;
        }
    }

    private static bool JsonDeepEquals(string originalJson, string verifyJson, out string diffMessage)
    {
        using var docA = JsonDocument.Parse(originalJson);
        using var docB = JsonDocument.Parse(verifyJson);

        var diffs = new List<string>();
        CompareElements(docA.RootElement, docB.RootElement, "$", diffs);

        if (diffs.Count == 0)
        {
            diffMessage = "JSON birebir uyumlu.";
            return true;
        }

        diffMessage = "JSON farkları bulundu:\n- " + string.Join("\n- ", diffs.Take(10));
        if (diffs.Count > 10) diffMessage += $"\n... (+{diffs.Count - 10} fark daha)";
        return false;
    }

    private static bool ShouldIgnoreMissingProperty(string path, string key, JsonElement aValue)
    {
        // 1) channel_criteria objesinde defaultlar
        // Path örnek: $.r.criteria[0].channel_criteria[3]
        bool inChannelCriteriaObj =
            path.Contains(".criteria[", StringComparison.Ordinal) &&
            path.Contains(".channel_criteria[", StringComparison.Ordinal) &&
            !path.Contains(".comparisons[", StringComparison.Ordinal);

        if (inChannelCriteriaObj)
        {
            if (key == "invert_logic" && aValue.ValueKind == JsonValueKind.False) return true; // default false
            if (key == "evaluation_idx" && aValue.ValueKind == JsonValueKind.Number && aValue.TryGetInt32(out int ei) && ei == 0) return true; // default 0
        }

        // 2) comparisons objesinde default op
        // Path örnek: $.r.criteria[0].channel_criteria[3].comparisons[0]
        bool inComparisonsObj =
            path.Contains(".criteria[", StringComparison.Ordinal) &&
            path.Contains(".channel_criteria[", StringComparison.Ordinal) &&
            path.Contains(".comparisons[", StringComparison.Ordinal);

        if (inComparisonsObj)
        {
            if (key == "op" && aValue.ValueKind == JsonValueKind.String && aValue.GetString() == "Eq")
                return true; // enum default Eq varsayımı (şemaya bağlı)
        }

        return false;
    }

    private static void CompareElements(JsonElement a, JsonElement b, string path, List<string> diffs)
    {
        if (a.ValueKind != b.ValueKind)
        {
            diffs.Add($"{path}: tür farklı (A={a.ValueKind}, B={b.ValueKind})");
            return;
        }

        switch (a.ValueKind)
        {
            case JsonValueKind.Object:
                {
                    var aProps = a.EnumerateObject().ToDictionary(p => p.Name, p => p.Value);
                    var bProps = b.EnumerateObject().ToDictionary(p => p.Name, p => p.Value);

                    foreach (var key in aProps.Keys.Except(bProps.Keys))
                    {
                        if (ShouldIgnoreMissingProperty(path, key, aProps[key]))
                            continue;

                        diffs.Add($"{path}.{key}: B'de yok");
                    }

                    foreach (var key in bProps.Keys.Except(aProps.Keys))
                        diffs.Add($"{path}.{key}: A'da yok");

                    foreach (var key in aProps.Keys.Intersect(bProps.Keys))
                        CompareElements(aProps[key], bProps[key], $"{path}.{key}", diffs);

                    break;
                }

            case JsonValueKind.Array:
                {
                    var aArr = a.EnumerateArray().ToArray();
                    var bArr = b.EnumerateArray().ToArray();

                    if (aArr.Length != bArr.Length)
                    {
                        diffs.Add($"{path}: dizi uzunluğu farklı (A={aArr.Length}, B={bArr.Length})");
                        return;
                    }

                    for (int i = 0; i < aArr.Length; i++)
                        CompareElements(aArr[i], bArr[i], $"{path}[{i}]", diffs);

                    break;
                }

            case JsonValueKind.String:
                if (a.GetString() != b.GetString())
                    diffs.Add($"{path}: string farklı (A='{a.GetString()}', B='{b.GetString()}')");
                break;

            case JsonValueKind.Number:
                {
                    // int gibi mi?
                    bool aIsInt = a.TryGetInt64(out long ai);
                    bool bIsInt = b.TryGetInt64(out long bi);

                    if (aIsInt && bIsInt)
                    {
                        if (ai != bi) diffs.Add($"{path}: number farklı (A={ai}, B={bi})");
                        break;
                    }

                    double da = a.GetDouble();
                    double db = b.GetDouble();

                    const double EPS = 1e-5;
                    if (Math.Abs(da - db) > EPS)
                        diffs.Add($"{path}: number farklı (A={da}, B={db})");

                    break;
                }

            case JsonValueKind.True:
            case JsonValueKind.False:
                if (a.GetBoolean() != b.GetBoolean())
                    diffs.Add($"{path}: bool farklı (A={a.GetBoolean()}, B={b.GetBoolean()})");
                break;

            case JsonValueKind.Null:
            case JsonValueKind.Undefined:
                break;

            default:
                if (a.GetRawText() != b.GetRawText())
                    diffs.Add($"{path}: değer farklı (A={a.GetRawText()}, B={b.GetRawText()})");
                break;
        }
    }


    // Dış bir executable'ı yani flatc.exe çalıştırır ve sonucu döndürür.
    private static (int exitCode, string stdout, string stderr) RunProcess(string exe, string args, string workingDir)
    {
        var psi = new ProcessStartInfo
        {
            FileName = exe,
            Arguments = args,
            WorkingDirectory = workingDir,

            // Çıktıları yakalamak için yönlendirme (logging/debug için kritik)
            RedirectStandardOutput = true,
            RedirectStandardError = true,

            // Redirect kullanabilmek için shell kapalı olmalı
            UseShellExecute = false,

            // Konsol penceresi açmadan çalıştır
            CreateNoWindow = true
        };

        using var p = new Process { StartInfo = psi };
        p.Start();

        // Çıktıları oku (flatc hata/uyarıyı genelde stderr'a basar)
        var stdout = p.StandardOutput.ReadToEnd();
        var stderr = p.StandardError.ReadToEnd();

        // İşlem bitene kadar bekle
        p.WaitForExit();

        return (p.ExitCode, stdout, stderr);
    }

    // Excel'deki "Channel_Transfer" sheet'ini okuyarak channel_transfer isteği için JSON üretir.
    // Beklenen kolonlar (1. satır header):
    // - tx_channel_id
    // - rx_msg_length
    // - tx_msg (virgülle ayrılmış byte listesi: 15,61,62 gibi)
    // - timeout_usec

    // - Üretilen JSON, FlatBuffers schema'daki RFT.Request yapısına uygun olacak şekilde hazırlanır.
    private static string ConvertChannelTransferFromExcel(string excelPath)
    {
        using var wb = new XLWorkbook(excelPath);
        var ws = wb.Worksheet("Channel_Transfer");

        // Header satırında (Row 1) kolon başlıklarının yerini bularak
        // kullanıcı kolon sırasını değiştirse bile doğru hücreleri okumayı sağlar.
        int Col(string header)
        {
            var headerRow = ws.Row(1);

            var cell = headerRow.CellsUsed()
                .FirstOrDefault(c =>
                    string.Equals(
                        c.GetString().Trim(),
                        header,
                        StringComparison.OrdinalIgnoreCase
                    )
                );

            if (cell == null)
                throw new Exception($"Excel'de '{header}' başlığı bulunamadı (Sheet: Channel_Transfer).");

            return cell.Address.ColumnNumber;
        }

        // Gerekli kolon indekslerini bul
        int cTxChannelId = Col("tx_channel_id");
        int cRxMsgLength = Col("rx_msg_length");
        int cTxMsg = Col("tx_msg");
        int cTimeoutUsec = Col("timeout_usec");

        // Veri okuma başlangıcı (Row 2): Row 1 header kabul edilir
        int row = 2;

        // Basit veri kontrolü: 2. satırda hiç veri yoksa kullanıcıya anlamlı hata ver
        var txIdCell = ws.Cell(row, cTxChannelId);
        if (txIdCell.IsEmpty() || string.IsNullOrWhiteSpace(txIdCell.GetString()))
            throw new Exception("Channel_Transfer sheet içinde veri bulunamadı (2. satır).");

        // Hücrelerden alanları oku
        int txChannelId = txIdCell.GetValue<int>();
        int rxMsgLength = ws.Cell(row, cRxMsgLength).GetValue<int>();
        string txMsgRaw = ws.Cell(row, cTxMsg).GetString();
        int timeoutUsec = ws.Cell(row, cTimeoutUsec).GetValue<int>();

        // "15,61,62" gibi bir metni byte/int dizisine çevir
        var txMsg = txMsgRaw
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(int.Parse)
            .ToArray();

        // Schema'ya uyumlu request payload:
        // r_type + r yapısı, seçilen isteğin türünü ve içeriğini taşır.
        var payload = new
        {
            r_type = "channel_transfer",
            r = new
            {
                tx_channel_id = new { id = txChannelId },
                rx_msg_length = rxMsgLength,
                tx_msg = txMsg,
                timeout_usec = timeoutUsec
            }
        };

        // JSON'u okunabilir formatta döndür (debug ve inceleme için)
        return JsonSerializer.Serialize(
            payload,
            new JsonSerializerOptions { WriteIndented = true }
        );
    }

    // Excel'deki "Channel_Configure" sheet'ini okuyarak channel_configure isteği için JSON üretir.
    // Beklenen kolonlar (1. satır header):
    // - channel_id
    // - rx_channel_id
    // - interface_config_type (rs485 / udp)
    // - baud_rate
    // - stop_bit
    // - data_bits
    // - parity
    // - termination (TRUE / FALSE)
    // - timeout_usec
    private static string ConvertChannelConfigureFromExcel(string excelPath)
    {
        using var wb = new XLWorkbook(excelPath);
        var ws = wb.Worksheet("Channel_Configure");

        // Header (Row 1) üzerinden kolon indeksini bulur.
        // Bu sayede Excel'de kolonların sırası değişse bile doğru alanlar okunur.
        int Col(string header)
        {
            var headerRow = ws.Row(1);
            var cell = headerRow.CellsUsed()
                .FirstOrDefault(c =>
                    string.Equals(
                        c.GetString().Trim(),
                        header,
                        StringComparison.OrdinalIgnoreCase
                    )
                );

            if (cell == null)
                throw new Exception($"Excel'de '{header}' başlığı bulunamadı (Sheet: Channel_Configure).");

            return cell.Address.ColumnNumber;
        }
        // Gerekli kolonları resolve et
        int cChannelId = Col("channel_id");
        int cRxChannelId = Col("rx_channel_id");
        int cType = Col("interface_config_type"); // rs485 / udp
        int cBaud = Col("baud_rate");
        int cStop = Col("stop_bit");
        int cDataBits = Col("data_bits");
        int cParity = Col("parity");
        int cTermination = Col("termination");
        int cTimeout = Col("timeout_usec");

        static string MapParity(string p)
        {
            p = (p ?? "").Trim();

            return p switch
            {
                "EvenParity" => "Even",
                "OddParity" => "Odd",
                _ => p
            };
        }

        var configs = new List<object>();

        for (int row = 2; ; row++)
        {
            var chCell = ws.Cell(row, cChannelId);

            if (chCell.IsEmpty() || string.IsNullOrWhiteSpace(chCell.GetString()))
                break;

            int channelId = chCell.GetValue<int>();
            int rxChannelId = ws.Cell(row, cRxChannelId).GetValue<int>();

            // Interface config alanları
            string ifaceType = ws.Cell(row, cType).GetString().Trim(); // "rs485" / "udp"
            string baudRate = ws.Cell(row, cBaud).GetString().Trim();
            string stopBit = ws.Cell(row, cStop).GetString().Trim();
            string dataBits = ws.Cell(row, cDataBits).GetString().Trim();
            string parity = MapParity(ws.Cell(row, cParity).GetString());
            bool termination = ws.Cell(row, cTermination).GetValue<bool>();
            int timeoutUsec = ws.Cell(row, cTimeout).GetValue<int>();

            if (!ifaceType.Equals("rs485", StringComparison.OrdinalIgnoreCase))
                throw new Exception($"Şu an sadece rs485 destekleniyor. interface_config_type: '{ifaceType}'");

            // Schema'ya uygun config nesnesi oluştur
            var config = new
            {
                channel_id = new { id = channelId },
                rx_channel_id = new { id = rxChannelId },
                interface_config_type = "rs485",
                interface_config = new
                {
                    baud_rate = baudRate,
                    stop_bit = stopBit,
                    data_bits = dataBits,
                    parity = parity,
                    termination = termination
                },
                timeout_usec = timeoutUsec
            };

            configs.Add(config);
        }
        if (configs.Count == 0)
            throw new Exception("Channel_Configure sheet içinde hiç veri satırı bulunamadı (2. satırdan itibaren).");

        // channel_configure request payload
        var payload = new
        {
            r_type = "channel_configure",
            r = new { configs }
        };

        // JSON'u okunabilir formatta döndür (debug ve inceleme için)
        return JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
    }


    private static string ConvertTestAddDirectivesFromExcel(string excelPath)
    {
        using var wb = new XLWorkbook(excelPath);
        var ws = wb.Worksheet("Test_AddDirectives");

        int Col(string header)
        {
            var headerRow = ws.Row(1);
            var cell = headerRow.CellsUsed()
                .FirstOrDefault(c => string.Equals(c.GetString().Trim(), header, StringComparison.OrdinalIgnoreCase));

            if (cell == null)
                throw new Exception($"Excel'de '{header}' basligi  bulunamadi (Sheet: Test_AddDirectives).");

            return cell.Address.ColumnNumber;
        }
        // hangi kolon nerede?
        int cTxChannelId = Col("tx_channel_id");
        int cRxMsgLength = Col("rx_msg_length");
        int cStepCount = Col("step_count");
        int cTxMsg = Col("tx_msg");

        var groups = new Dictionary<int, List<object>>();

        for (int row = 2; ; row++)
        {
            var txCell = ws.Cell(row, cTxChannelId);
            if (txCell.IsEmpty() || string.IsNullOrWhiteSpace(txCell.GetString()))
                break;

            int txChannelId = txCell.GetValue<int>();
            int rxMsgLength = ws.Cell(row, cRxMsgLength).GetValue<int>();
            int stepCount = ws.Cell(row, cStepCount).GetValue<int>();
            string txMsgRaw = ws.Cell(row, cTxMsg).GetString();

            //tx_msg yazısını diziye çeviriyoruz
            var txMsg = txMsgRaw
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(int.Parse)
                .ToArray();

            var channelDirective = new //Tek bir directive objesi oluşturuyoruz.
            {
                tx_msg = txMsg,
                rx_msg_length = rxMsgLength,
                step_count = stepCount
            };

            if (!groups.TryGetValue(txChannelId, out var list)) //bu directive'i doğru gruba ekliyoruz..
            {
                list = new List<object>();
                groups[txChannelId] = list;
            }

            list.Add(channelDirective);
        }

        if (groups.Count == 0) //Hiç veri yoksa hata verir.
            throw new Exception("Test_AddDirectives sheet içinde hiç veri satır bulunamadı.");

        var directives = groups.Select(g => new //Gruplari JSON formatina ceviriyoruz
        {
            tx_channel_id = new { id = g.Key },
            channel_directives = g.Value
        }).ToList();

        var payload = new //En dış JSON'u oluşturuyoruz.
        {
            r_type = "test_add_directives",
            r = new
            {
                directives = directives
            }
        };
        return JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
    }

    private static string ConvertTestPrepareFromExcel(string excelPath)
    {
        using var wb = new XLWorkbook(excelPath);

        static string CellStr(IXLWorksheet ws, int row, int col) => ws.Cell(row, col).GetString().Trim(); //excel h�cresini string al�r.
        static bool IsRowEmpty(IXLWorksheet ws, int row, int keyCol) // Bir sat�r�n bitip bitmedi�ini anlamak i�in
        {
            var c = ws.Cell(row, keyCol);
            return c.IsEmpty() || string.IsNullOrWhiteSpace(c.GetString());
        }

        static object WrapX(object v) => new { x = v }; //Format�n istedi�i �u yap�y� �retmek i�in

        static string[] SplitCsv(string s) =>
            s.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        static int[] ParseIntCsv(string s) => SplitCsv(s).Select(int.Parse).ToArray();

        static (List<string> types, List<object> instr) ParseInstructions(string instructionsCsv, string typesCsv) //Excel�de evaluations i�in 2 kolon var, bu da ikisinin eleman say�s� e�it mi ona bak�yor.
        {
            var types = SplitCsv(typesCsv).ToList();
            var ins = SplitCsv(instructionsCsv);

            if (types.Count != ins.Length)
                throw new Exception("Evaluations: instructions ve instructions_type adetleri e�le�miyor.");

            var instructions = new List<object>();

            for (int i = 0; i < types.Count; i++)
            {
                var t = types[i];
                var token = ins[i];

                if (t.Equals("binary_op", StringComparison.OrdinalIgnoreCase))
                {
                    instructions.Add(new { op = token }); // { "op": "Sub" }
                }
                else
                {
                    // load_field gibi durumlarda { "x": 4 } token say� olmal�
                    if (!int.TryParse(token, out int n))
                        throw new Exception($"Evaluations: '{t}' i�in say� bekleniyordu ama '{token}' geldi.");
                    instructions.Add(new { x = n });
                }
            }

            return (types, instructions);
        }

        // 1. period_usec Değeri Bulunur Yazdırılır.
        // "period_usec": 1000
        var wsGeneral = wb.Worksheet("Test_Prepare_General");
        int periodUsec = wsGeneral.Cell(2, 1).GetValue<int>();

        // 2. Fields Okuma Alani
        var wsFields = wb.Worksheet("Test_Prepare_Fields");
        var fields = new List<object>();

        // Fields içindekiler: field_source, offset, scalar_type, big_endian
        /*{
        "field_source": "Tx",
        "offset": 60,
        "scalar_type": { "x": "U32" },
        "big_endian": true 
        }*/
        for (int row = 2; !IsRowEmpty(wsFields, row, 1); row++)
        {
            string fieldSource = CellStr(wsFields, row, 1);          // Tx/Rx
            int offset = wsFields.Cell(row, 2).GetValue<int>();
            string scalarType = CellStr(wsFields, row, 3);           // U32/F32/U8
            bool bigEndian = wsFields.Cell(row, 4).GetValue<bool>();

            fields.Add(new
            {
                field_source = fieldSource,
                offset = offset,
                scalar_type = new { x = scalarType }, // <-- hedef format: { "x": "U32" }
                big_endian = bigEndian
            });
        }

        if (fields.Count == 0)
            throw new Exception("Test_Prepare_Fields i�inde hi� sat�r yok.");

        // 3.Bindings Okuma Alani
        var wsBindings = wb.Worksheet("Test_Prepare_Bindings");
        // Kolonlari: tx_channel_id, arg_name, arg_index
        var bindGroups = new Dictionary<int, Dictionary<string, int>>();

        for (int row = 2; !IsRowEmpty(wsBindings, row, 1); row++)
        {
            int txId = wsBindings.Cell(row, 1).GetValue<int>();
            string argName = CellStr(wsBindings, row, 2);
            int argIndex = wsBindings.Cell(row, 3).GetValue<int>();

            if (!bindGroups.TryGetValue(txId, out var map))
            {
                map = new Dictionary<string, int>();
                bindGroups[txId] = map;
            }

            // Ayni isim tekrar gelirse ustune yazar
            map[argName] = argIndex;
        }

        var bindings = bindGroups.Select(g => new
        {
            tx_channel_id = new { id = g.Key },
            arg_indices = g.Value
        }).ToList();

        // 4. Evaluations Okuma Alani
        var wsEval = wb.Worksheet("Test_Prepare_Evaluations");
        // Kolonlar: evaluation_idx, instructions, instructions_type
        var evaluations = new List<object>();

        for (int row = 2; !IsRowEmpty(wsEval, row, 1); row++)
        {
            string instructionsCsv = CellStr(wsEval, row, 2);
            string typesCsv = CellStr(wsEval, row, 3);

            var (types, instructions) = ParseInstructions(instructionsCsv, typesCsv);

            evaluations.Add(new
            {
                instructions_type = types,
                instructions = instructions,
                constants_type = Array.Empty<string>(),
                constants = Array.Empty<object>(),
                variables_type = Array.Empty<string>(),
                variables = Array.Empty<object>()
            });
        }

        // 5. Criteria Okuma Alani 
        var wsCrit = wb.Worksheet("Test_Prepare_Criteria");
        // Kolonlar: tx_channel_id, evaluation_idx, comparison_ops, comparison_values, invert_logic, start_time_step, end_time_step
        var critGroups = new Dictionary<int, List<object>>();

        for (int row = 2; !IsRowEmpty(wsCrit, row, 1); row++)
        {
            int txId = wsCrit.Cell(row, 1).GetValue<int>();
            int evaluationIdx = wsCrit.Cell(row, 2).GetValue<int>();
            string opsCsv = CellStr(wsCrit, row, 3);           // "Ge,Le"
            string valuesCsv = CellStr(wsCrit, row, 4);        // "99.5,100.5"
            bool invertLogic = wsCrit.Cell(row, 5).GetValue<bool>();
            int start = wsCrit.Cell(row, 6).GetValue<int>();
            int end = wsCrit.Cell(row, 7).GetValue<int>();

            var ops = SplitCsv(opsCsv);
            var vals = SplitCsv(valuesCsv).Select(s => double.Parse(s, System.Globalization.CultureInfo.InvariantCulture)).ToArray();

            if (ops.Length != vals.Length)
                throw new Exception("Criteria: comparison_ops ve comparison_values adetleri eslesmiyor.");

            var comparisons = new List<object>();
            for (int i = 0; i < ops.Length; i++)
            {
                comparisons.Add(new
                {
                    op = ops[i],
                    value_type = "f32",
                    value = new { x = vals[i] } // { "x": 99.5 }
                });
            }

            var criterion = new
            {
                evaluation_idx = evaluationIdx,
                comparisons = comparisons,
                invert_logic = invertLogic,
                start_time_step = start,
                end_time_step = end
            };

            if (!critGroups.TryGetValue(txId, out var list))
            {
                list = new List<object>();
                critGroups[txId] = list;
            }

            list.Add(criterion);
        }

        var criteria = critGroups.Select(g => new
        {
            tx_channel_id = new { id = g.Key },
            channel_criteria = g.Value
        }).ToList();

        // Final payload
        var payload = new
        {
            r_type = "test_prepare",
            r = new
            {
                fields = fields,
                bindings = bindings,
                evaluations = evaluations,
                criteria = criteria,
                period_usec = periodUsec
            }
        };

        return JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
    }
}
using Avalonia.Controls;
using ClosedXML.Excel; //Excel dosyas�n� a��p h�cre okumak
using DocumentFormat.OpenXml.VariantTypes;
using System; //temel t�rler tipler, exception vs.
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO; //dosya yolu i�lemleri, dosyaya yazma
using System.Linq;
using System.Text;
using System.Text.Json; //C# objesini JSON�a �evirme
using System.Diagnostics;
using System.Text;

namespace ExcelToJsonConverter.App;

public partial class MainWindow : Window //bizim ana s�n�f�m�z partial olmas�n�n nedeni, hem UI hem de kod taraf� farkl� dosyalarda olup sonra birle�mesidir.
{
    public MainWindow()
    {
        InitializeComponent(); //XAML'de �izdi�im UI'y� y�kler.
        BtnPickExcel.Click += BtnPickExcel_Click; //BtnPickExcel butonuna t�klan�rsa, BtnPickExcel_Click fonksiyonunu �al��t�r.
        CmbType.SelectionChanged += CmbType_SelectionChanged;
        BtnConvert.Click += BtnConvert_Click;
        BtnUpdate.Click += BtnUpdate_Click;
    }
    private async void BtnPickExcel_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e) //async ��nk� kullan�c�dan dosya se�mesini bekleyen bir i�lem var (dialog).
    {
        var dlg = new OpenFileDialog //Dosya se�me penceresi olu�turuyor
        {
            Title = "Excel dosyasi sec",
            AllowMultiple = false, //kullan�c� sadece 1 dosya se�ebilsin.
            Filters =
        {
            new FileDialogFilter { Name = "Excel", Extensions = { "xlsx", "xlsm" } }
        }
        };

        var result = await dlg.ShowAsync(this); //Pencereyi a��p kullan�c�dan se�im bekliyor.
        if (result is { Length: > 0 })
            TxtExcelPath.Text = result[0]; //Se�ilen dosyay� ekrana yaz.
    }

    private void CmbType_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        var selected = //kullanici hangi turu secti.
            (CmbType.SelectedItem as ComboBoxItem)?.Content?.ToString()
            ?? "unknown";

        TxtExpectedFormat.Text = selected switch
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
                "- stop_bit\n" +
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
        };

    }

    private void BtnConvert_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try //try-catch: Excel dosyasi bozuk olabilir ya da kullanici yanlış seçebilir. Eğer hata olursa program çökmesin diye, hata mesajını ekrana yazdırıyor.
        {
            SuccessPanel.IsVisible = false;
            ErrorPanel.IsVisible = false;

            TxtResultPath.Text = "";
            TxtError.Text = "";
            var selectedType = //kullanıcının seçtıği türü alıyoruz. 
                (CmbType.SelectedItem as ComboBoxItem)?.Content?.ToString()
                ?? "unknown";

            var excelPath = TxtExcelPath.Text ?? ""; //excel dosyasi secilmis mi?

            if (string.IsNullOrWhiteSpace(excelPath) || excelPath == "(Excel Secilmedi)")
            {
                TxtPreview.Text = "�nce Excel dosyasi secmelisin.";
                SuccessPanel.IsVisible = false;     // onemli
                TxtResultPath.Text = "";            // temizle
                return;
            }

            string json = selectedType switch //hangi tur secildiyse onu calistir.
            {
                "channel_transfer" => ConvertChannelTransferFromExcel(excelPath),
                "channel_configure" => ConvertChannelConfigureFromExcel(excelPath),
                "test_add_directives" => ConvertTestAddDirectivesFromExcel(excelPath),
                "test_prepare" => ConvertTestPrepareFromExcel(excelPath),
                _ => throw new Exception($"Bilinmeyen t�r: {selectedType}")
            };

            var directory = Path.GetDirectoryName(excelPath) ?? Environment.CurrentDirectory;
            var baseName = Path.GetFileNameWithoutExtension(excelPath);

            var jsonPath = Path.Combine(directory, $"{baseName}.json");
            File.WriteAllText(jsonPath, json, Encoding.UTF8);
            //json metnini al ve jsonPath konumuna dosya olarak kaydet.
            TxtPreview.Text = json;   // sadece JSON

            SuccessPanel.IsVisible = true;
            ErrorPanel.IsVisible = false;

            TxtSuccessTitle.Text = "JSON başarıyla oluşturuldu:";
            TxtResultPath.Text = jsonPath;

        }
        catch (Exception ex)
        {
            TxtPreview.Text = $"Hata: {ex.Message}";

            TxtError.Text = ex.Message;
            ErrorPanel.IsVisible = true;

            SuccessPanel.IsVisible = false;
        }
    }


    private void BtnUpdate_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
     try
    {
        SuccessPanel.IsVisible = false;
        ErrorPanel.IsVisible = false;
        TxtError.Text = ""; //butona basinca eski hata ya da success mesajlari siliniyor.
        TxtResultPath.Text = "";

        var selectedType =
            (CmbType.SelectedItem as ComboBoxItem)?.Content?.ToString()
            ?? "unknown";

        var excelPath = TxtExcelPath.Text ?? "";
        if (string.IsNullOrWhiteSpace(excelPath) || excelPath == "(Excel Seçilmedi)")
            throw new Exception("Önce Excel dosyası seçmelisin.");

        // 1) JSON üret (senin mevcut fonksiyonların)
        string json = selectedType switch
        {
            "channel_transfer" => ConvertChannelTransferFromExcel(excelPath),
            "channel_configure" => ConvertChannelConfigureFromExcel(excelPath),
            "test_add_directives" => ConvertTestAddDirectivesFromExcel(excelPath),
            "test_prepare" => ConvertTestPrepareFromExcel(excelPath),
            _ => throw new Exception($"Bilinmeyen tür: {selectedType}")
        };

        // 2) Proje root + out
        // Projede her şey projectRoot/out içine düşsün diye standart klasör belirliyoruz.
        var projectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
        var outDir = Path.Combine(projectRoot, "out");
        Directory.CreateDirectory(outDir);

            var baseName = Path.GetFileNameWithoutExtension(excelPath);
            var jsonPath = Path.Combine(outDir, $"{baseName}.json");
            File.WriteAllText(jsonPath, json, Encoding.UTF8);

            // 3) flatc yolları, shema ve ftatc var mi kontrol et.
            var schemaPath = Path.Combine(projectRoot, "schemas", "rft.fbs");
        if (!File.Exists(schemaPath))
            throw new Exception($"Şema bulunamadı: {schemaPath}");

            var flatcPath = Path.Combine(projectRoot, "Tools", "flatbuffers", "win-x64", "flatc.exe");

            if (!File.Exists(flatcPath))
            throw new Exception($"flatc bulunamadı: {flatcPath}");

            // 4) flatc çalıştır: JSON -> BIN
            var args = $"--binary --strict-json --root-type RFT.Request -o \"{outDir}\" \"{schemaPath}\" \"{jsonPath}\"";

            var (exitCode, stdout, stderr) = RunProcess(flatcPath, args, projectRoot);

        if (exitCode != 0)
        {
            throw new Exception(
                "flatc hata verdi.\n\n" +
                $"Komut: {flatcPath} {args}\n\n" +
                $"STDOUT:\n{stdout}\n\nSTDERR:\n{stderr}"
            );
            }
            var binPath = Path.Combine(outDir, $"{baseName}.bin");
            // 5) BIN -> JSON (doğrulama)
            var verifyDir = Path.Combine(outDir, "verify");
            Directory.CreateDirectory(verifyDir);

            var verifyArgs =
                $"--json --strict-json --defaults-json --raw-binary " +
                $"--root-type RFT.Request " +
                $"-o \"{verifyDir}\" " +
                $"\"{schemaPath}\" -- \"{binPath}\"";

            var (exit2, out2, err2) = RunProcess(flatcPath, verifyArgs, projectRoot);

            if (exit2 != 0)
            {
                throw new Exception(
                    "BIN -> JSON doğrulama (flatc) hata verdi.\n\n" +
                    $"Komut:\n{flatcPath} {verifyArgs}\n\n" +
                    $"ExitCode: {exit2}\n\n" +
                    $"STDOUT:\n{out2}\n\n" +
                    $"STDERR:\n{err2}"
                );
            }
            var verifyJsonFiles = Directory.GetFiles(verifyDir, "*.json");

            if (verifyJsonFiles.Length == 0)
                throw new Exception("flatc çalıştı ama verify klasöründe JSON oluşmadı.");

            var verifyJsonPath = verifyJsonFiles[0]; // genelde tek dosya olur

            if (!File.Exists(binPath))
        throw new Exception($"flatc çalıştı ama .bin bulunamadı: {binPath}");

            TxtPreview.Text = json; // istersen verify json'u da gösterebilirsin

            SuccessPanel.IsVisible = true;
            ErrorPanel.IsVisible = false;

            TxtSuccessTitle.Text = "BIN oluşturuldu ve tekrar JSON'a çevrilerek doğrulandı:";
            TxtResultPath.Text = $"BIN: {binPath}\nVERIFY JSON: {verifyJsonPath}";

        }
        catch (Exception ex)
    {
        TxtError.Text = ex.Message;
        ErrorPanel.IsVisible = true;
        SuccessPanel.IsVisible = false;
    }
}

    private static (int exitCode, string stdout, string stderr) RunProcess(string exe, string args, string workingDir)
    {
    var psi = new ProcessStartInfo
    {
        FileName = exe,
        Arguments = args,
        WorkingDirectory = workingDir,
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        UseShellExecute = false,
        CreateNoWindow = true
    };

    using var p = new Process { StartInfo = psi };
    p.Start();

    var stdout = p.StandardOutput.ReadToEnd();
    var stderr = p.StandardError.ReadToEnd();

    p.WaitForExit();
    return (p.ExitCode, stdout, stderr);
    }


    private static string ConvertChannelTransferFromExcel(string excelPath)
        {
            using var wb = new XLWorkbook(excelPath);
            var ws = wb.Worksheet("Channel_Transfer");

            int Col(string header)
            {
                var headerRow = ws.Row(1);
                var cell = headerRow.CellsUsed()
                    .FirstOrDefault(c => string.Equals(c.GetString().Trim(), header, StringComparison.OrdinalIgnoreCase));
                if (cell == null)
                    throw new Exception($"Excel'de '{header}' başlığı bulunamadı (Sheet: Channel_Transfer).");
                return cell.Address.ColumnNumber;
            }

            int cTxChannelId = Col("tx_channel_id");
            int cRxMsgLength = Col("rx_msg_length");
            int cTxMsg = Col("tx_msg");
            int cTimeoutUsec = Col("timeout_usec");

            int row = 2;
            var txIdCell = ws.Cell(row, cTxChannelId);
            if (txIdCell.IsEmpty() || string.IsNullOrWhiteSpace(txIdCell.GetString()))
                throw new Exception("Channel_Transfer sheet içinde veri bulunamadı (2. satır).");

            int txChannelId = txIdCell.GetValue<int>();
            int rxMsgLength = ws.Cell(row, cRxMsgLength).GetValue<int>();
            string txMsgRaw = ws.Cell(row, cTxMsg).GetString();
            int timeoutUsec = ws.Cell(row, cTimeoutUsec).GetValue<int>();

            var txMsg = txMsgRaw
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(int.Parse)
                .ToArray();

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

            return JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
        }


    private static string ConvertChannelConfigureFromExcel(string excelPath)
    {
        using var wb = new XLWorkbook(excelPath);
        var ws = wb.Worksheet("Channel_Configure");

        int Col(string header)
        {
            var headerRow = ws.Row(1);
            var cell = headerRow.CellsUsed()
                .FirstOrDefault(c => string.Equals(c.GetString().Trim(), header, StringComparison.OrdinalIgnoreCase));
            if (cell == null)
                throw new Exception($"Excel'de '{header}' başlığı bulunamadı (Sheet: Channel_Configure).");
            return cell.Address.ColumnNumber;
        }

        int cChannelId = Col("channel_id");
        int cRxChannelId = Col("rx_channel_id");
        int cType = Col("interface_config_type");   // rs485 / udp
        int cBaud = Col("baud_rate");
        int cStop = Col("stop_bit");
        int cDataBits = Col("data_bits");
        int cParity = Col("parity");
        int cTermination = Col("termination");
        int cTimeout = Col("timeout_usec");

        static string MapParity(string p)
        {
            // Direktör örneğine uyalım: NoParity / Even / Odd / Space / Mark
            // Excel bazen EvenParity/OddParity gibi gelebilir → normalize edelim
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

            string ifaceType = ws.Cell(row, cType).GetString().Trim(); // "rs485" / "udp"
            string baudRate = ws.Cell(row, cBaud).GetString().Trim();
            string stopBit = ws.Cell(row, cStop).GetString().Trim();
            string dataBits = ws.Cell(row, cDataBits).GetString().Trim();
            string parity = MapParity(ws.Cell(row, cParity).GetString());
            bool termination = ws.Cell(row, cTermination).GetValue<bool>();
            int timeoutUsec = ws.Cell(row, cTimeout).GetValue<int>();

            if (!ifaceType.Equals("rs485", StringComparison.OrdinalIgnoreCase))
                throw new Exception($"Şu an sadece rs485 destekleniyor. interface_config_type: '{ifaceType}'");

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

        var payload = new
        {
            r_type = "channel_configure",
            r = new { configs }
        };

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

            var channelDirective = new //Tek bir directive objesi oluşturuyoruz
            {
                tx_msg = txMsg,
                rx_msg_length = rxMsgLength,
                step_count = stepCount
            };

            if (!groups.TryGetValue(txChannelId, out var list)) //bu directive'i doğru gruba ekliyruz.
            {
                list = new List<object>();
                groups[txChannelId] = list;
            }

            list.Add(channelDirective);
        }

        if (groups.Count == 0) //Hic veri yoksa hata ver
            throw new Exception("Test_AddDirectives sheet i�inde hi� veri sat�r� bulunamad� (2. sat�rdan itibaren).");

        var directives = groups.Select(g => new //Gruplari JSON formatina ceviriyoruz
        {
            tx_channel_id = new { id = g.Key },
            channel_directives = g.Value
        }).ToList();

        var payload = new //En dis JSON'u olusturuyoruz.
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

        // 2. Fields Okuma Alan�
        var wsFields = wb.Worksheet("Test_Prepare_Fields");
        var fields = new List<object>();

        // Fields i�indekiler: field_source, offset, scalar_type, big_endian
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

        // 3.Bindings Okuma Alan�
        var wsBindings = wb.Worksheet("Test_Prepare_Bindings");
        // Kolonlar�: tx_channel_id, arg_name, arg_index
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
                throw new Exception("Criteria: comparison_ops ve comparison_values adetleri e�le�miyor.");

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
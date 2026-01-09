using Avalonia.Controls;
using ClosedXML.Excel; //Excel dosyasýný açýp hücre okumak
using DocumentFormat.OpenXml.VariantTypes;
using System; //temel türler tipler, exception vs.
using System.Collections.Generic;
using System.IO; //dosya yolu iþlemleri, dosyaya yazma
using System.Linq;
using System.Text.Json; //C# objesini JSON’a çevirme

namespace ExcelToJsonConverter.App;

public partial class MainWindow : Window //bizim ana sýnýfýmýz partial olmasýnýn nedeni, hem UI hem de kod tarafý farklý dosyalarda olup sonra birleþmesidir.
{
    public MainWindow()
    {
        InitializeComponent(); //XAML'de çizdiðim UI'yý yükler.
        BtnPickExcel.Click += BtnPickExcel_Click; //BtnPickExcel butonuna týklanýrsa, BtnPickExcel_Click fonksiyonunu çalýþtýr.
        CmbType.SelectionChanged += CmbType_SelectionChanged;
        BtnConvert.Click += BtnConvert_Click;
    }
    private async void BtnPickExcel_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e) //async çünkü kullanýcýdan dosya seçmesini bekleyen bir iþlem var (dialog).
    {
        var dlg = new OpenFileDialog //Dosya seçme penceresi oluþturuyor
        {
            Title = "Excel dosyasý seç",
            AllowMultiple = false, //kullanýcý sadece 1 dosya seçebilsin.
            Filters =
        {
            new FileDialogFilter { Name = "Excel", Extensions = { "xlsx", "xlsm" } }
        }
        };

        var result = await dlg.ShowAsync(this); //Pencereyi açýp kullanýcýdan seçim bekliyor.
        if (result is { Length: > 0 })
            TxtExcelPath.Text = result[0]; //Seçilen dosyayý ekrana yaz.
    }

    private void CmbType_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        var selected = //kullanýcý hangi türü seçti.
            (CmbType.SelectedItem as ComboBoxItem)?.Content?.ToString()
            ?? "unknown";

        TxtExpectedFormat.Text = selected switch
        {
            "channel_transfer" =>
                "Sheet: Channel_Transfer\n" +
                "Kolonlar:\n" +
                "- tx_channel_id\n" +
                "- rx_msg_length\n" +
                "- tx_msg (virgüllü ör.: 15,61,62)\n" +
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
                "- tx_msg (virgüllü: 10,21,22,23)\n" +
                "Not:\n" +
                "- Her satýr 1 directive satýrýdýr.\n" +
                "- Ayný tx_channel_id tekrar edebilir; JSON'da tek baþlýk altýnda gruplanýr.",

            "test_prepare" =>
                "Sheet'ler:\n" +
                "1) Test_Prepare_General\n" +
                "   - period_usec\n\n" +

                "2) Test_Prepare_Fields\n" +
                "   - field_source (Tx / Rx)\n" +
                "   - offset\n" +
                "   - scalar_type (U32 / F32 / U8)\n" +
                "   - big_endian (TRUE / FALSE)\n" +
                "   Not: Her satýr 1 field tanýmýdýr.\n\n" +

                "3) Test_Prepare_Bindings\n" +
                "   - tx_channel_id\n" +
                "   - arg_name (response1, command2, ...)\n" +
                "   - arg_index\n" +
                "   Not: Ayný tx_channel_id tekrar edebilir; JSON'da gruplanýr.\n\n" +

                "4) Test_Prepare_Evaluations\n" +
                "   - evaluation_idx\n" +
                "   - instructions (virgüllü: 9,4,Sub)\n" +
                "   - instructions_type (virgüllü: load_field,binary_op)\n" +
                "   Not: instructions ve instructions_type sýralarý birebir eþleþmelidir.\n\n" +

                "5) Test_Prepare_Criteria\n" +
                "   - tx_channel_id\n" +
                "   - evaluation_idx\n" +
                "   - comparison_ops (virgüllü: Ge,Le)\n" +
                "   - comparison_values (virgüllü: 99.5,100.5)\n" +
                "   - invert_logic (TRUE / FALSE)\n" +
                "   - start_time_step\n" +
                "   - end_time_step\n" +
                "Not: Her satýr 1 kriterdir.",
        };

    }

    private void BtnConvert_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try //try-catch: Excel dosyasý bozuk olabilir ya da kullanýcý yanlýþ seçebilir. Eðer hata olursa program çökmesin diye, hata mesajýný ekrana yazdýrýyor.
        {
            SuccessPanel.IsVisible = false;
            ErrorPanel.IsVisible = false;

            TxtResultPath.Text = "";
            TxtError.Text = "";
            var selectedType = //kullanýcýnýn seçtiði türü alýyoruz.
                (CmbType.SelectedItem as ComboBoxItem)?.Content?.ToString()
                ?? "unknown";

            var excelPath = TxtExcelPath.Text ?? ""; //excel dosyasý seçilmiþ mi?

            if (string.IsNullOrWhiteSpace(excelPath) || excelPath == "(Excel Seçilmedi)")
            {
                TxtPreview.Text = "Önce Excel dosyasý seçmelisin.";
                SuccessPanel.IsVisible = false;     // önemli
                TxtResultPath.Text = "";            // temizle
                return;
            }

            string json = selectedType switch //hangi tür seçildiyse onu çalýþtýr.
            {
                "channel_transfer" => ConvertChannelTransferFromExcel(excelPath),
                "channel_configure" => ConvertChannelConfigureFromExcel(excelPath),
                "test_add_directives" => ConvertTestAddDirectivesFromExcel(excelPath),
                "test_prepare" => ConvertTestPrepareFromExcel(excelPath),
                _ => throw new Exception($"Bilinmeyen tür: {selectedType}")
            };

            var directory = Path.GetDirectoryName(excelPath) ?? Environment.CurrentDirectory; //excel dosyasý nerede duruyor?
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath); //excel dosyasýnýn adý ne?
            var jsonPath = Path.Combine(directory, fileNameWithoutExt + ".json"); //ayný klasöre ayný isimle ama formatý .json yap.

            File.WriteAllText(jsonPath, json); //json metnini al ve jsonPath konumuna dosya olarak kaydet.
            TxtPreview.Text = json;   // sadece JSON

            TxtResultPath.Text = jsonPath;
            SuccessPanel.IsVisible = true;
        }
        catch (Exception ex)
        {
            TxtPreview.Text = $"Hata: {ex.Message}";

            TxtError.Text = ex.Message;
            ErrorPanel.IsVisible = true;

            SuccessPanel.IsVisible = false;
        }
    }


    private static string ConvertChannelTransferFromExcel(string excelPath) // excel dosyasýný aç, channel transfer dosyasýný oku, her satýrý JSON'a çevir, hepsini liste yap ve JSON metni olarak geri döndür.
    {
        using var wb = new XLWorkbook(excelPath); // excel dosyasýný aç, iþin bitince kapat(using).
        var ws = wb.Worksheet("Channel_Transfer"); // Channel_Transfer isimli sayfayý bul ve onu okuyacaðým.

        int Col(string header) //kolonlarýn yeri sabit olmayabilir diye baþlýktan buluruz.
        {
            var headerRow = ws.Row(1); //baþlýk satýrýný al (1.satýr baþlýklar)
            var cell = headerRow.CellsUsed() //baþlýklarýn içinde aradýðýný bul.
                .FirstOrDefault(c => string.Equals(c.GetString().Trim(), header, StringComparison.OrdinalIgnoreCase));

            if (cell == null) //bulamazsan hata ver.
                throw new Exception($"Excel'de '{header}' baþlýðý bulunamadý (Sheet: Channel_Transfer).");

            return cell.Address.ColumnNumber; //bulduysa sütun numarasýný döndür.
        }

        // Kolonlarý bir kere bul
        int cTxChannelId = Col("tx_channel_id"); //excelde bu baþlýklae hangi sütunda onlarý bul ve aklýnda tut.
        int cRxMsgLength = Col("rx_msg_length");
        int cTxMsg = Col("tx_msg");
        int cTimeoutUsec = Col("timeout_usec");

        var results = new System.Collections.Generic.List<object>(); //Her satýrdan bir JSON obje oluþturacaðýz. Hepsini bu listeye atacaðýz.

        // 2. satýrdan itibaren oku çünkü 1. satýr baþlýklarý içeriyor.
        for (int row = 2; ; row++)
        {
            var txIdCell = ws.Cell(row, cTxChannelId); // tx_channel_id boþsa (veya satýr tamamen boþsa) dur
            if (txIdCell.IsEmpty() || string.IsNullOrWhiteSpace(txIdCell.GetString()))
                break;

            int txChannelId = txIdCell.GetValue<int>(); // O satýrdaki deðerleri alýyoruz
            int rxMsgLength = ws.Cell(row, cRxMsgLength).GetValue<int>();
            string txMsgRaw = ws.Cell(row, cTxMsg).GetString();
            int timeoutUsec = ws.Cell(row, cTimeoutUsec).GetValue<int>();

            var txMsg = txMsgRaw //"15,61,62" yazýsýný listeye çeviriyoruz
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(int.Parse)
                .ToArray();

            var payload = new //O satýr için JSON objesini kuruyoruz
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

            results.Add(payload); //Listeye ekliyoruz.
        }

        if (results.Count == 0) //Hiç satýr yoksa hata ver.
            throw new Exception("Channel_Transfer sheet içinde hiç veri satýrý bulunamadý (2. satýrdan itibaren).");

        //Listeyi JSON metnine çevirip geri döndür
        return JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true });
    }

    private static string ConvertChannelConfigureFromExcel(string excelPath)
    {
        using var wb = new XLWorkbook(excelPath);
        var ws = wb.Worksheet("Channel_Configure");

        int Col(string header) //kolonlarýn yerini baþlýktan bulmak için kullanýlan fonksiyon.
        {
            var headerRow = ws.Row(1); //Excel’in 1. satýrý baþlýklar satýrý kabul edilir.
            var cell = headerRow.CellsUsed()
                .FirstOrDefault(c => string.Equals(c.GetString().Trim(), header, StringComparison.OrdinalIgnoreCase));

            if (cell == null)
                throw new Exception($"Excel'de '{header}' baþlýðý bulunamadý (Sheet: Channel_Configure).");

            return cell.Address.ColumnNumber; //Bulursa sütun numarasýný döndürüyor
        }
        //excelde hangi baþlýk hangi sütunda?
        int cChannelId = Col("channel_id");
        int cRxChannelId = Col("rx_channel_id");
        int cType = Col("interface_config_type");
        int cBaud = Col("baud_rate");
        int cStop = Col("stop_bit");
        int cDataBits = Col("data_bits");
        int cParity = Col("parity");
        int cTermination = Col("termination");
        int cTimeout = Col("timeout_usec");

        var configs = new System.Collections.Generic.List<object>(); //JSON’a gidecek “config”leri toplayacaðýmýz liste

        for (int row = 2; ; row++) //2. satýrdan itibaren satýr satýr gez. ;; kýsmý sonsuz döngü demek.
        {
            var chCell = ws.Cell(row, cChannelId); //satýr boþsa dur!!
            if (chCell.IsEmpty() || string.IsNullOrWhiteSpace(chCell.GetString()))
                break;

            //satýrlardaki deðerleri tek tek deðiþkenlere kaydediyoruz.
            int channelId = chCell.GetValue<int>();
            int rxChannelId = ws.Cell(row, cRxChannelId).GetValue<int>();

            string ifaceType = ws.Cell(row, cType).GetString();
            string baudRate = ws.Cell(row, cBaud).GetString();
            string stopBit = ws.Cell(row, cStop).GetString();
            string dataBits = ws.Cell(row, cDataBits).GetString();
            string parity = ws.Cell(row, cParity).GetString();

            bool termination = ws.Cell(row, cTermination).GetValue<bool>();
            int timeoutUsec = ws.Cell(row, cTimeout).GetValue<int>();

            var config = new //C#’ta hýzlýca JSON’a benzeyen bir nesne oluþturur, hedef þablon oluþturulur.
            {
                channel_id = new { id = channelId },
                rx_channel_id = new { id = rxChannelId },
                interface_config_type = ifaceType,
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

            configs.Add(config); //Excel’de her satýr okunduðunda bir config oluþur ve liste büyür.
        }

        if (configs.Count == 0) //Excel’de hiç veri yoksa hata veriyor. Sheet boþsa ya da yanlýþ excel seçilmiþse hata verir.
            throw new Exception("Channel_Configure sheet içinde hiç veri satýrý bulunamadý (2. satýrdan itibaren).");

        var payload = new //en son dýþ JSON'u kuruyoruz. 
        {
            r_type = "channel_configure",
            r = new
            {
                configs = configs
            }
        };

        return JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true }); // JSON metnine çevirip döndürüyor.
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
                throw new Exception($"Excel'de '{header}' baþlýðý bulunamadý (Sheet: Test_AddDirectives).");

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

            //tx_msg yazýsýný diziye çeviriyoruz
            var txMsg = txMsgRaw
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(int.Parse)
                .ToArray();

            var channelDirective = new //Tek bir “directive” objesi oluþturuyoruz
            {
                tx_msg = txMsg,
                rx_msg_length = rxMsgLength,
                step_count = stepCount
            };

            if (!groups.TryGetValue(txChannelId, out var list)) //bu directive'i doðru gruba ekliyruz.
            {
                list = new List<object>();
                groups[txChannelId] = list;
            }

            list.Add(channelDirective);
        }

        if (groups.Count == 0) //Hiç veri yoksa hata ver
            throw new Exception("Test_AddDirectives sheet içinde hiç veri satýrý bulunamadý (2. satýrdan itibaren).");

        var directives = groups.Select(g => new //Gruplarý JSON formatýna çeviriyoruz
        {
            tx_channel_id = new { id = g.Key },
            channel_directives = g.Value
        }).ToList();

        var payload = new //En dýþ JSON'u oluþturuyoruz.
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

        static string CellStr(IXLWorksheet ws, int row, int col) => ws.Cell(row, col).GetString().Trim(); //excel hücresini string alýr.
        static bool IsRowEmpty(IXLWorksheet ws, int row, int keyCol) // Bir satýrýn bitip bitmediðini anlamak için
        {
            var c = ws.Cell(row, keyCol);
            return c.IsEmpty() || string.IsNullOrWhiteSpace(c.GetString());
        }

        static object WrapX(object v) => new { x = v }; //Formatýn istediði þu yapýyý üretmek için

        static string[] SplitCsv(string s) =>
            s.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        static int[] ParseIntCsv(string s) => SplitCsv(s).Select(int.Parse).ToArray();

        static (List<string> types, List<object> instr) ParseInstructions(string instructionsCsv, string typesCsv) //Excel’de evaluations için 2 kolon var, bu da ikisinin eleman sayýsý eþit mi ona bakýyor.
        {
            var types = SplitCsv(typesCsv).ToList();
            var ins = SplitCsv(instructionsCsv);

            if (types.Count != ins.Length)
                throw new Exception("Evaluations: instructions ve instructions_type adetleri eþleþmiyor.");

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
                    // load_field gibi durumlarda { "x": 4 } token sayý olmalý
                    if (!int.TryParse(token, out int n))
                        throw new Exception($"Evaluations: '{t}' için sayý bekleniyordu ama '{token}' geldi.");
                    instructions.Add(new { x = n });
                }
            }

            return (types, instructions);
        }

        // 1. period_usec Deðeri Bulunur Yazdýrýlýr
        // "period_usec": 1000
        var wsGeneral = wb.Worksheet("Test_Prepare_General");
        int periodUsec = wsGeneral.Cell(2, 1).GetValue<int>();

        // 2. Fields Okuma Alaný
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
            throw new Exception("Test_Prepare_Fields içinde hiç satýr yok.");

        // 3.Bindings Okuma Alaný
        var wsBindings = wb.Worksheet("Test_Prepare_Bindings");
        // Kolonlarý: tx_channel_id, arg_name, arg_index
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

            // Ayný isim tekrar gelirse üstüne yazar
            map[argName] = argIndex;
        }

        var bindings = bindGroups.Select(g => new
        {
            tx_channel_id = new { id = g.Key },
            arg_indices = g.Value
        }).ToList();

        // 4. Evaluations Okuma Alaný
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

        // 5. Criteria Okuma Alaný 
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
                throw new Exception("Criteria: comparison_ops ve comparison_values adetleri eþleþmiyor.");

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

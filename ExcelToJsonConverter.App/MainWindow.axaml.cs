using System.IO;
using System;
using System.Linq;
using System.Text.Json;
using ClosedXML.Excel;
using Avalonia.Controls;

namespace ExcelToJsonConverter.App;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        BtnPickExcel.Click += BtnPickExcel_Click; //BtnPickExcel butonuna týklanýrsa, BtnPickExcel_Click fonksiyonunu çalýþtýr.
        CmbType.SelectionChanged += CmbType_SelectionChanged;
        BtnConvert.Click += BtnConvert_Click;
    }
    private async void BtnPickExcel_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog //Dosya seçme penceresi oluþturuyor
        {
            Title = "Excel dosyasý seç",
            AllowMultiple = false,
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
        var selected =
            (CmbType.SelectedItem as ComboBoxItem)?.Content?.ToString()
            ?? "unknown";

        TxtSelectedType.Text = $"Seçili tür: {selected}";
        TxtExpectedFormat.Text = selected switch
        {
            "channel_transfer" =>
                "Sheet: Channel_Transfer\n" +
                "Kolonlar:\n" +
                "- tx_channel_id\n" +
                "- rx_msg_length\n" +
                "- tx_msg (virgüllü: 15,61,62)\n" +
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

            _ =>
                "(Bu tür için henüz format açýklamasý yok)"
        };

    }

    private void BtnConvert_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        try
        {
            var selectedType =
                (CmbType.SelectedItem as ComboBoxItem)?.Content?.ToString()
                ?? "unknown";

            var excelPath = TxtExcelPath.Text ?? "";

            if (string.IsNullOrWhiteSpace(excelPath))
            {
                TxtPreview.Text = "Önce Excel dosyasý seçmelisin.";
                return;
            }

            string json = selectedType switch //hangi tür seçildiyse onu çalýþtýr.
            {
                "channel_transfer" => ConvertChannelTransferFromExcel(excelPath),
                "channel_configure" => ConvertChannelConfigureFromExcel(excelPath),
                _ => throw new Exception($"Bu tür henüz desteklenmiyor: {selectedType}")
            };

            var directory = Path.GetDirectoryName(excelPath);
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
            var jsonPath = Path.Combine(directory!, fileNameWithoutExt + ".json");

            File.WriteAllText(jsonPath, json);

            TxtPreview.Text = json + "\n\nKaydedildi:\n" + jsonPath;
        }
        catch (Exception ex)
        {
            TxtPreview.Text = $"Hata: {ex.Message}";
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

        int Col(string header)
        {
            var headerRow = ws.Row(1);
            var cell = headerRow.CellsUsed()
                .FirstOrDefault(c => string.Equals(c.GetString().Trim(), header, StringComparison.OrdinalIgnoreCase));

            if (cell == null)
                throw new Exception($"Excel'de '{header}' baþlýðý bulunamadý (Sheet: Channel_Configure).");

            return cell.Address.ColumnNumber;
        }

        int cChannelId = Col("channel_id");
        int cRxChannelId = Col("rx_channel_id");
        int cType = Col("interface_config_type");
        int cBaud = Col("baud_rate");
        int cStop = Col("stop_bit");
        int cDataBits = Col("data_bits");
        int cParity = Col("parity");
        int cTermination = Col("termination");
        int cTimeout = Col("timeout_usec");

        var configs = new System.Collections.Generic.List<object>();

        for (int row = 2; ; row++)
        {
            var chCell = ws.Cell(row, cChannelId);
            if (chCell.IsEmpty() || string.IsNullOrWhiteSpace(chCell.GetString()))
                break;

            int channelId = chCell.GetValue<int>();
            int rxChannelId = ws.Cell(row, cRxChannelId).GetValue<int>();

            string ifaceType = ws.Cell(row, cType).GetString();
            string baudRate = ws.Cell(row, cBaud).GetString();
            string stopBit = ws.Cell(row, cStop).GetString();
            string dataBits = ws.Cell(row, cDataBits).GetString();
            string parity = ws.Cell(row, cParity).GetString();

            bool termination = ws.Cell(row, cTermination).GetValue<bool>();
            int timeoutUsec = ws.Cell(row, cTimeout).GetValue<int>();

            var config = new
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

            configs.Add(config);
        }

        if (configs.Count == 0)
            throw new Exception("Channel_Configure sheet içinde hiç veri satýrý bulunamadý (2. satýrdan itibaren).");

        var payload = new
        {
            r_type = "channel_configure",
            r = new
            {
                configs = configs
            }
        };

        return JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
    }

}

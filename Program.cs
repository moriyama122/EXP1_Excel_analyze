using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

class Program
{
    const double TARGET_TIME = 5.0; // R後5秒

    static void Main()
    {
        string sourceFolder =
            "/Users/moriyama_yuto/Library/CloudStorage/OneDrive-KyushuUniversity/実験/EXP1データ";
        string csvPattern = "*.csv";
        string tempXlsx =
            "/Users/moriyama_yuto/ExcelColumnExtract/converted.xlsx";
        string resultXlsx =
            "/Users/moriyama_yuto/Library/CloudStorage/OneDrive-KyushuUniversity/実験/result.xlsx";

        // subject → task → values(D列)
        var speedAt5s = new Dictionary<string, Dictionary<int, List<double>>>();

        var csvFiles = Directory.GetFiles(
            sourceFolder, csvPattern, SearchOption.AllDirectories);

        foreach (var csvPath in csvFiles)
        {
            var name = Path.GetFileNameWithoutExtension(csvPath);
            var parts = name.Split('_');
            if (parts.Length < 2) continue;

            string subject = parts[0];

            var taskRun = parts[1].Split('-');
            if (!int.TryParse(taskRun[0], out int task)) continue;

            // CSV → Excel
            using (var wbTemp = new XLWorkbook())
            {
                var wsTemp = wbTemp.Worksheets.Add("Sheet1");
                using var reader = new StreamReader(csvPath);
                int row = 1;
                while (!reader.EndOfStream)
                {
                    var vals = reader.ReadLine()?.Split(',') ?? Array.Empty<string>();
                    for (int i = 0; i < vals.Length; i++)
                        wsTemp.Cell(row, i + 1).Value = vals[i];
                    row++;
                }
                wbTemp.SaveAs(tempXlsx);
            }

            using var wb = new XLWorkbook(tempXlsx);
            var ws = wb.Worksheet("Sheet1");
            var lastRowUsed = ws.LastRowUsed();
            if (lastRowUsed == null) continue;

            int lastRow = lastRowUsed.RowNumber();

            bool rStarted = false;
            double? startTime = null;
            double? valueAt5s = null;

            for (int r = 2; r <= lastRow; r++)
            {
                if (!double.TryParse(ws.Cell(r, "C").GetString(), out double time))
                    continue;

                // R=True を検知
                if (!rStarted &&
                    bool.TryParse(ws.Cell(r, "R").GetString(), out bool rBool) &&
                    rBool)
                {
                    rStarted = true;
                    startTime = time;
                    continue;
                }

                if (!rStarted || !startTime.HasValue)
                    continue;

                // R後5秒を超えた最初の行
                if (time - startTime.Value >= TARGET_TIME)
                {
                    if (double.TryParse(ws.Cell(r, "D").GetString(), out double d))
                        valueAt5s = d;

                    break;
                }
            }

            if (!valueAt5s.HasValue)
                continue;

            speedAt5s.TryAdd(subject, new Dictionary<int, List<double>>());
            speedAt5s[subject].TryAdd(task, new List<double>());
            speedAt5s[subject][task].Add(valueAt5s.Value);
        }

        // ===== result.xlsx に反映 =====
        using var wbResult = new XLWorkbook(resultXlsx);
        var wsResult = wbResult.Worksheet("Result");

        foreach (var subject in speedAt5s.Keys)
        {
            int row = FindOrCreateRow(wsResult, subject);

            for (int task = 1; task <= 3; task++)
            {
                if (speedAt5s[subject].ContainsKey(task))
                {
                    // B,C,D 列に task1,2,3
                    wsResult.Cell(row, 1 + task).Value =
                        speedAt5s[subject][task].Average();
                }
            }
        }

        wbResult.SaveAs(resultXlsx);
        Console.WriteLine("抽出完了！");
    }

    static int FindOrCreateRow(IXLWorksheet ws, string subject)
    {
        int lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        for (int r = 2; r <= lastRow; r++)
            if (ws.Cell(r, 1).GetString() == subject)
                return r;

        ws.Cell(lastRow + 1, 1).Value = subject;
        return lastRow + 1;
    }
}
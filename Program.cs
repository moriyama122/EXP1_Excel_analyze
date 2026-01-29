using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

class Program
{
    const double WINDOW = 5.0; // 秒

    static void Main()
    {
        string sourceFolder =
            "/Users/moriyama_yuto/Library/CloudStorage/OneDrive-KyushuUniversity/実験/EXP1データ";
        string csvPattern = "*.csv";
        string tempXlsx =
            "/Users/moriyama_yuto/ExcelColumnExtract/converted.xlsx";
        string resultXlsx =
            "/Users/moriyama_yuto/Library/CloudStorage/OneDrive-KyushuUniversity/実験/result.xlsx";

        // subject → task → values
        var maxE = new Dictionary<string, Dictionary<int, List<double>>>();
        var maxF = new Dictionary<string, Dictionary<int, List<double>>>();
        var sumE = new Dictionary<string, Dictionary<int, List<double>>>();
        var sumF = new Dictionary<string, Dictionary<int, List<double>>>();

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

            double eMax = double.MinValue;
            double fMax = double.MinValue;
            double eSum = 0.0;
            double fSum = 0.0;

            for (int r = 2; r <= lastRow; r++)
            {
                if (!double.TryParse(ws.Cell(r, "C").GetString(), out double time))
                    continue;

                if (!rStarted &&
                    bool.TryParse(ws.Cell(r, "R").GetString(), out bool rBool) &&
                    rBool)
                {
                    rStarted = true;
                    startTime = time;
                    continue;
                }

                if (!rStarted || !startTime.HasValue) continue;
                if (time - startTime.Value > WINDOW) break;

                if (double.TryParse(ws.Cell(r, "E").GetString(), out double e))
                {
                    eMax = Math.Max(eMax, e);
                    eSum += e;
                }

                if (double.TryParse(ws.Cell(r, "F").GetString(), out double f))
                {
                    fMax = Math.Max(fMax, f);
                    fSum += f;
                }
            }

            if (eMax == double.MinValue || fMax == double.MinValue) continue;

            // Dictionary 初期化
            maxE.TryAdd(subject, new Dictionary<int, List<double>>());
            maxF.TryAdd(subject, new Dictionary<int, List<double>>());
            sumE.TryAdd(subject, new Dictionary<int, List<double>>());
            sumF.TryAdd(subject, new Dictionary<int, List<double>>());

            maxE[subject].TryAdd(task, new List<double>());
            maxF[subject].TryAdd(task, new List<double>());
            sumE[subject].TryAdd(task, new List<double>());
            sumF[subject].TryAdd(task, new List<double>());

            maxE[subject][task].Add(eMax);
            maxF[subject][task].Add(fMax);
            sumE[subject][task].Add(eSum);
            sumF[subject][task].Add(fSum);
        }

        // ===== result.xlsx に反映 =====
        using var wbResult = new XLWorkbook(resultXlsx);
        var wsResult = wbResult.Worksheet("Result");

        foreach (var subject in maxE.Keys)
        {
            int row = FindOrCreateRow(wsResult, subject);

            for (int task = 1; task <= 3; task++)
            {
                // 最大値
                if (maxE[subject].ContainsKey(task))
                    wsResult.Cell(row, 1 + task).Value =
                        maxE[subject][task].Average();   // B–D

                if (maxF[subject].ContainsKey(task))
                    wsResult.Cell(row, 5 + task).Value =
                        maxF[subject][task].Average();   // F–H

                // 合計値
                if (sumE[subject].ContainsKey(task))
                    wsResult.Cell(row, 9 + task).Value =
                        sumE[subject][task].Average();   // J–L

                if (sumF[subject].ContainsKey(task))
                    wsResult.Cell(row, 13 + task).Value =
                        sumF[subject][task].Average();   // N–P
            }
        }

        wbResult.SaveAs(resultXlsx);
        Console.WriteLine("集計完了！");
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
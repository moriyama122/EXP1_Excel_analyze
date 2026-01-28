using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

class Program
{
    const double FS = 15.0;
    const double DT = 1.0 / FS;
    const double MAX_TIME = 10.0;
    const int MAX_BIN = (int)(FS * MAX_TIME);

    static void Main()
    {
        string sourceFolder =
            "/Users/moriyama_yuto/Library/CloudStorage/OneDrive-KyushuUniversity/実験/EXP1データ";
        string csvPattern = "*.csv";
        string tempXlsx = "/Users/moriyama_yuto/ExcelColumnExtract/converted.xlsx";
        string resultXlsx = "/Users/moriyama_yuto/ExcelColumnExtract/result.xlsx";

        // ★ 除外する実験者
        var excludeSubjects = new HashSet<string>
        {
            "EXP104", "EXP107", "EXP113", "EXP116", "EXP119"
        };

        // task → bin → speeds
        var taskBinsAll = new Dictionary<int, Dictionary<int, List<double>>>();
        var taskBinsFiltered = new Dictionary<int, Dictionary<int, List<double>>>();

        var csvFiles = Directory.GetFiles(
            sourceFolder, csvPattern, SearchOption.AllDirectories);

        foreach (var csvPath in csvFiles)
        {
            var fname = Path.GetFileNameWithoutExtension(csvPath);
            var parts = fname.Split('_');
            if (parts.Length < 2) continue;

            string subject = parts[0];   // ★ EXP101 など
            var taskRun = parts[1].Split('-');
            if (!int.TryParse(taskRun[0], out int task)) continue;

            bool isExcluded = excludeSubjects.Contains(subject);

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

            for (int r = 2; r <= lastRow; r++)
            {
                if (!double.TryParse(ws.Cell(r, "C").GetString(), out double time)) continue;
                if (!double.TryParse(ws.Cell(r, "D").GetString(), out double speedKm)) continue;

                double speed = speedKm / 3.6;

                if (!rStarted &&
                    bool.TryParse(ws.Cell(r, "R").GetString(), out bool rBool) &&
                    rBool)
                {
                    rStarted = true;
                    startTime = time;
                    continue;
                }

                if (!rStarted || !startTime.HasValue) continue;

                double relTime = time - startTime.Value;
                if (relTime < 0 || relTime > MAX_TIME) continue;

                int bin = (int)Math.Round(relTime * FS);
                if (bin < 0 || bin >= MAX_BIN) continue;

                // ===== 全員 =====
                Add(taskBinsAll, task, bin, speed);

                // ===== 除外後 =====
                if (!isExcluded)
                    Add(taskBinsFiltered, task, bin, speed);
            }
        }

        // ===== result.xlsx に追記 =====
        using var wbResult = new XLWorkbook(resultXlsx);
        var wsResult = wbResult.Worksheet("Result");

        int startCol = (wsResult.LastColumnUsed()?.ColumnNumber() ?? 1) + 1;

        // Time列
        if (wsResult.Cell(1, 1).IsEmpty())
        {
            wsResult.Cell(1, 1).Value = "Time[s]";
            for (int b = 0; b < MAX_BIN; b++)
                wsResult.Cell(b + 2, 1).Value = b * DT;
        }

        foreach (var task in taskBinsAll.Keys.OrderBy(t => t))
        {
            // 全員
            wsResult.Cell(1, startCol).Value = $"Task{task}_All";
            WriteBins(wsResult, startCol, taskBinsAll[task]);

            // 除外
            wsResult.Cell(1, startCol + 1).Value = $"Task{task}_Excluded";
            if (taskBinsFiltered.ContainsKey(task))
                WriteBins(wsResult, startCol + 1, taskBinsFiltered[task]);

            startCol += 2;
        }

        wbResult.SaveAs(resultXlsx);
        Console.WriteLine("追記完了！");
    }

    static void Add(
        Dictionary<int, Dictionary<int, List<double>>> dict,
        int task, int bin, double value)
    {
        if (!dict.ContainsKey(task))
            dict[task] = new Dictionary<int, List<double>>();
        if (!dict[task].ContainsKey(bin))
            dict[task][bin] = new List<double>();
        dict[task][bin].Add(value);
    }

    static void WriteBins(
        IXLWorksheet ws, int col,
        Dictionary<int, List<double>> bins)
    {
        foreach (var kv in bins)
            ws.Cell(kv.Key + 2, col).Value = kv.Value.Average();
    }
}
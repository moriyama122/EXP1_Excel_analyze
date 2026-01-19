using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        string sourceFolder = "/Users/moriyama_yuto/Library/CloudStorage/OneDrive-KyushuUniversity/実験/EXP1データ";
        string csvPattern = "*.csv";
        string outputXlsxAll = "/Users/moriyama_yuto/ExcelColumnExtract/converted.xlsx";
        string resultXlsx = "/Users/moriyama_yuto/ExcelColumnExtract/result.xlsx";

        // subject -> task -> list of 最小TTC
        var dataMinTTC = new Dictionary<string, Dictionary<int, List<double>>>(StringComparer.OrdinalIgnoreCase);

        var csvFiles = Directory.GetFiles(sourceFolder, csvPattern, SearchOption.AllDirectories);
        if (csvFiles.Length == 0)
        {
            Console.WriteLine("CSVファイルが見つかりません");
            return;
        }

        foreach (var csvPath in csvFiles)
        {
            var fname = Path.GetFileNameWithoutExtension(csvPath);
            var parts = fname.Split('_');
            if (parts.Length < 2) continue;

            string subject = parts[0];
            var taskAndRun = parts[1].Split('-');
            if (!int.TryParse(taskAndRun[0], out int taskNumber)) continue;

            // CSV → Excel
            using (var wbTemp = new XLWorkbook())
            {
                var wsTemp = wbTemp.Worksheets.Add("Sheet1");
                using var reader = new StreamReader(csvPath, Encoding.UTF8);
                int row = 1;
                while (!reader.EndOfStream)
                {
                    var vals = (reader.ReadLine() ?? "").Split(',');
                    for (int i = 0; i < vals.Length; i++)
                        wsTemp.Cell(row, i + 1).Value = vals[i];
                    row++;
                }
                wbTemp.SaveAs(outputXlsxAll);
            }

            using var wb = new XLWorkbook(outputXlsxAll);
            var ws = wb.Worksheet("Sheet1");
            int lastRow = ws.LastRowUsed().RowNumber();

            bool rStarted = false;
            bool inRange = false;
            bool finished = false;
            int sCount = 0;

            double minTTC = double.MaxValue;
            const double leadSpeed = 50.0 / 3.6; // m/s

            for (int r = 2; r <= lastRow; r++)
            {
                // R
                if (!rStarted &&
                    bool.TryParse(ws.Cell(r, "R").GetString(), out bool rBool) && rBool)
                {
                    rStarted = true;
                    sCount = 0;
                }

                // S (3連続)
                if (rStarted && !inRange)
                {
                    if (ws.Cell(r, "S").GetString() == "1")
                    {
                        sCount++;
                        if (sCount >= 3)
                            inRange = true;
                    }
                    else
                    {
                        sCount = 0;
                    }
                }

                // TTC計算
                if (inRange && !finished)
                {
                    if (
                        double.TryParse(ws.Cell(r, "D").GetString(), out double egoSpeedKm) &&
                        double.TryParse(ws.Cell(r, "M").GetString(), out double egoX) &&
                        double.TryParse(ws.Cell(r, "N").GetString(), out double egoZ) &&
                        double.TryParse(ws.Cell(r, "P").GetString(), out double leadX) &&
                        double.TryParse(ws.Cell(r, "Q").GetString(), out double leadZ)
                    )
                    {
                        double egoSpeed = egoSpeedKm / 3.6;
                        double relSpeed = egoSpeed - leadSpeed;
                        if (relSpeed > 0)
                        {
                            double dist = Math.Sqrt(
                                Math.Pow(leadX - egoX, 2) +
                                Math.Pow(leadZ - egoZ, 2)
                            );
                            double ttc = dist / relSpeed;
                            if (ttc < minTTC)
                                minTTC = ttc;
                        }
                    }
                }

                // U
                if (inRange &&
                    bool.TryParse(ws.Cell(r, "U").GetString(), out bool uBool) &&
                    uBool)
                {
                    finished = true;
                    break;
                }
            }

            if (!inRange || !finished || minTTC == double.MaxValue)
            {
                Console.WriteLine($"[{Path.GetFileName(csvPath)}] TTC未算出");
                continue;
            }

            dataMinTTC.TryAdd(subject, new Dictionary<int, List<double>>());
            dataMinTTC[subject].TryAdd(taskNumber, new List<double>());
            dataMinTTC[subject][taskNumber].Add(minTTC);

            Console.WriteLine($"[{Path.GetFileName(csvPath)}] 最小TTC = {minTTC:F2}s");
        }

        // result.xlsx 出力
        using var wbResult = new XLWorkbook(resultXlsx);
        var wsResult = wbResult.Worksheet("Result");

        foreach (var subject in dataMinTTC.Keys)
        {
            int row = FindOrCreateSubjectRow(wsResult, subject);
            foreach (var task in dataMinTTC[subject].Keys)
            {
                int col = TaskStartCol(task);
                wsResult.Cell(1, col).Value = $"task{task}";
                wsResult.Cell(2, col).Value = "最小TTC[s]";
                wsResult.Cell(row, col).Value = dataMinTTC[subject][task].Average();
            }
        }

        wbResult.SaveAs(resultXlsx);
        Console.WriteLine("アップデート完了！");
    }

    static int TaskStartCol(int taskNumber) => 2 + (taskNumber - 1);

    static int FindOrCreateSubjectRow(IXLWorksheet ws, string subject)
    {
        int lastRow = ws.LastRowUsed()?.RowNumber() ?? 2;
        for (int r = 3; r <= lastRow; r++)
            if (ws.Cell(r, 1).GetString().Equals(subject, StringComparison.OrdinalIgnoreCase))
                return r;
        ws.Cell(lastRow + 1, 1).Value = subject;
        return lastRow + 1;
    }
}
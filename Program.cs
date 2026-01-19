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
        string sourceFolder = "/Users/moriyama_yuto/Library/CloudStorage/OneDrive-KyushuUniversity/実験/EXP1データ/EXP120";
        string csvPattern = "*.csv";
        string outputXlsxAll = "/Users/moriyama_yuto/ExcelColumnExtract/converted.xlsx";
        string resultXlsx = "/Users/moriyama_yuto/ExcelColumnExtract/result.xlsx";

        // subject -> task -> list of G合計値
        var dataSumG = new Dictionary<string, Dictionary<int, List<double>>>(StringComparer.OrdinalIgnoreCase);

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

            // CSV → 一時Excel
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

            using var allWb = new XLWorkbook(outputXlsxAll);
            var ws = allWb.Worksheet("Sheet1");

            var lastRowUsed = ws.LastRowUsed();
            if (lastRowUsed == null)
            {
                Console.WriteLine("シートが空です: " + csvPath);
                continue;
            }
            int lastRow = lastRowUsed.RowNumber();

            bool rStarted = false;   // R=True を通過したか
            bool inRange = false;    // Gの加算開始フラグ
            bool finished = false;

            int sCount = 0;
            int sStartRow = -1;

            double sumG = 0.0;

            for (int r = 2; r <= lastRow; r++)
            {
                // --- R列チェック ---
                if (!rStarted &&
                    bool.TryParse(ws.Cell(r, "R").GetString(), out bool rBool) &&
                    rBool)
                {
                    rStarted = true;
                    sCount = 0;
                    sStartRow = -1;
                }

                // --- S列チェック（R後のみ有効） ---
                if (rStarted && !inRange)
                {
                    if (ws.Cell(r, "S").GetString() == "1")
                    {
                        if (sCount == 0)
                            sStartRow = r;   // 最初の1

                        sCount++;

                        // ★ 3連続に到達 → 開始
                        if (sCount >= 3)
                        {
                            inRange = true;
                        }
                    }
                    else
                    {
                        sCount = 0;
                        sStartRow = -1;
                    }
                }

                // --- G列：絶対値を加算 ---
                if (inRange && !finished)
                {
                    if (double.TryParse(ws.Cell(r, "G").GetString(), out double gVal))
                        sumG += Math.Abs(gVal);
                }

                // --- U列：終了 ---
                if (inRange &&
                    bool.TryParse(ws.Cell(r, "U").GetString(), out bool uBool) &&
                    uBool)
                {
                    finished = true;
                    break;
                }
            }

            if (!inRange || !finished)
            {
                Console.WriteLine($"[{Path.GetFileName(csvPath)}] S→U 区間が見つかりません");
                continue;
            }

            // 保存
            dataSumG.TryAdd(subject, new Dictionary<int, List<double>>());
            dataSumG[subject].TryAdd(taskNumber, new List<double>());
            dataSumG[subject][taskNumber].Add(sumG);

            // ログ
            Console.WriteLine($"[{Path.GetFileName(csvPath)}]");
            Console.WriteLine($"  G合計 (R→U): {sumG:F4}");
            Console.WriteLine();
        }

        // ===== result.xlsx に書き込み =====
        using var wbResult = new XLWorkbook(resultXlsx);
        var wsResult = wbResult.Worksheet("Result");

        foreach (var subject in dataSumG.Keys)
        {
            int row = FindOrCreateSubjectRow(wsResult, subject);

            foreach (var task in dataSumG[subject].Keys)
            {
                int col = TaskStartCol(task);
                wsResult.Cell(1, col).Value = $"task{task}";
                wsResult.Cell(2, col).Value = "G合計";

                wsResult.Cell(row, col).Value = dataSumG[subject][task].Average();
            }
        }

        wbResult.SaveAs(resultXlsx);
        Console.WriteLine("アップデート完了！");
    }

    // ===== 補助メソッド =====

    static int TaskStartCol(int taskNumber)
        => 2 + (taskNumber - 1);

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
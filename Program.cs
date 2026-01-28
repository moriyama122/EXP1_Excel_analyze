using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using ClosedXML.Excel;

class Program
{
    const double JERK_THRESHOLD = 2.0; // m/s^3
    const double MIN_DURATION = 0.2;   // s

    static void Main()
    {
        string sourceFolder =
            "/Users/moriyama_yuto/Library/CloudStorage/OneDrive-KyushuUniversity/実験/EXP1データ/EXP114";
        string csvPattern = "*.csv";
        string tempXlsx = "/Users/moriyama_yuto/ExcelColumnExtract/converted.xlsx";
        string resultXlsx = "/Users/moriyama_yuto/ExcelColumnExtract/result.xlsx";

        // ★ jerk発生時刻（C列）だけを溜める
        var jerkTimes = new List<double>();

        var csvFiles = Directory.GetFiles(sourceFolder, csvPattern, SearchOption.AllDirectories);
        if (csvFiles.Length == 0)
        {
            Console.WriteLine("CSVファイルが見つかりません");
            return;
        }

        foreach (var csvPath in csvFiles)
        {
            Console.WriteLine($"処理開始: {csvPath}");

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
                wbTemp.SaveAs(tempXlsx);
            }

            using var wb = new XLWorkbook(tempXlsx);
            var ws = wb.Worksheet("Sheet1");

            var lastRowUsed = ws.LastRowUsed();
            if (lastRowUsed == null)
            {
                Console.WriteLine($"[{Path.GetFileName(csvPath)}] データなし");
                continue;
            }

            int lastRow = lastRowUsed.RowNumber();

            double? prevTime = null;
            double? prevSpeed = null;
            double? prevAcc = null;

            double currentDuration = 0.0;
            double totalDuration = 0.0;   // ★ CSVごとのJerk累積時間
            double? jerkStartTime = null;

            for (int r = 2; r <= lastRow; r++)
            {
                if (!double.TryParse(ws.Cell(r, "C").GetString(), out double time)) continue;
                if (!double.TryParse(ws.Cell(r, "D").GetString(), out double speedKm)) continue;

                double speed = speedKm / 3.6; // km/h → m/s

                if (prevTime.HasValue && prevSpeed.HasValue)
                {
                    double dt = time - prevTime.Value;
                    if (dt <= 0) goto NEXT;

                    double acc = (speed - prevSpeed.Value) / dt;

                    if (prevAcc.HasValue)
                    {
                        double jerk = (acc - prevAcc.Value) / dt;

                        if (Math.Abs(jerk) >= JERK_THRESHOLD)
                        {
                            if (currentDuration == 0)
                                jerkStartTime = time; // jerk開始時刻

                            currentDuration += dt;
                        }
                        else
                        {
                            if (currentDuration >= MIN_DURATION && jerkStartTime.HasValue)
                            {
                                jerkTimes.Add(jerkStartTime.Value);
                                totalDuration += currentDuration;

                                Console.WriteLine(
                                    $"  Jerk検知: {jerkStartTime.Value:F3}s"
                                );
                            }
                            currentDuration = 0.0;
                            jerkStartTime = null;
                        }
                    }
                    prevAcc = acc;
                }

            NEXT:
                prevTime = time;
                prevSpeed = speed;
            }

            // 末尾処理
            if (currentDuration >= MIN_DURATION && jerkStartTime.HasValue)
            {
                jerkTimes.Add(jerkStartTime.Value);
                totalDuration += currentDuration;

                Console.WriteLine(
                    $"  Jerk検知: {jerkStartTime.Value:F3}s"
                );
            }

            // ★ あなたが追加したかったログ
            if (totalDuration <= 0)
                continue;

            Console.WriteLine(
                $"[{Path.GetFileName(csvPath)}] Jerk累積時間 = {totalDuration:F3}s"
            );
        }

        // ===== Result.xlsx に書き出し =====
        using var wbResult = new XLWorkbook(resultXlsx);
        var wsResult = wbResult.Worksheet("Result");

        int writeRow = Math.Max(35, wsResult.LastRowUsed()?.RowNumber() + 1 ?? 35);

        foreach (var t in jerkTimes)
        {
            wsResult.Cell(writeRow, 2).Value = t; // B列
            writeRow++;
        }

        wbResult.SaveAs(resultXlsx);
        Console.WriteLine("アップデート完了！");
    }
}
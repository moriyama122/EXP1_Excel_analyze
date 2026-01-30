using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

class Program
{
    const double JERK_THRESHOLD = 1.5; // m/s^3
    const double WINDOW = 10.0;        // R後10秒

    static void Main()
    {
        string sourceFolder =
            "/Users/moriyama_yuto/Library/CloudStorage/OneDrive-KyushuUniversity/実験/EXP1データ";
        string csvPattern = "*.csv";
        string tempXlsx =
            "/Users/moriyama_yuto/ExcelColumnExtract/converted.xlsx";
        string resultXlsx =
            "/Users/moriyama_yuto/Library/CloudStorage/OneDrive-KyushuUniversity/実験/result.xlsx";

        // subject → task → jerkイベント回数リスト
        var jerkCounts =
            new Dictionary<string, Dictionary<int, List<int>>>();

        var csvFiles = Directory.GetFiles(
            sourceFolder, csvPattern, SearchOption.AllDirectories);

        if (csvFiles.Length == 0)
        {
            Console.WriteLine("CSVファイルが見つかりません");
            return;
        }

        foreach (var csvPath in csvFiles)
        {
            Console.WriteLine($"処理開始: {Path.GetFileName(csvPath)}");

            // ==== ファイル名から subject / task を取得 ====
            var name = Path.GetFileNameWithoutExtension(csvPath);
            var parts = name.Split('_');
            if (parts.Length < 2) continue;

            string subject = parts[0];

            var taskRun = parts[1].Split('-');
            if (!int.TryParse(taskRun[0], out int task)) continue;

            // ==== CSV → Excel ====
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
            if (lastRowUsed == null) continue;

            int lastRow = lastRowUsed.RowNumber();

            bool rStarted = false;
            double? rStartTime = null;

            double? prevTime = null;
            double? prevSpeed = null;
            double? prevAcc = null;

            int jerkEventCount = 0;
            bool inJerkEvent = false; // 連続検知防止用

            for (int r = 2; r <= lastRow; r++)
            {
                if (!double.TryParse(ws.Cell(r, "C").GetString(), out double time))
                    continue;

                if (!double.TryParse(ws.Cell(r, "D").GetString(), out double speedKm))
                    continue;

                double speed = speedKm / 3.6;

                // --- R=True 検知 ---
                if (!rStarted &&
                    bool.TryParse(ws.Cell(r, "R").GetString(), out bool rBool) &&
                    rBool)
                {
                    rStarted = true;
                    rStartTime = time;
                    prevTime = time;
                    prevSpeed = speed;
                    prevAcc = null;
                    continue;
                }

                if (!rStarted || !rStartTime.HasValue)
                    continue;

                // R後10秒で打ち切り
                if (time - rStartTime.Value > WINDOW)
                    break;

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
                            if (!inJerkEvent)
                            {
                                jerkEventCount++;
                                inJerkEvent = true;

                                Console.WriteLine(
                                    $"  Jerk検知: t={time:F3}s"
                                );
                            }
                        }
                        else
                        {
                            inJerkEvent = false;
                        }
                    }

                    prevAcc = acc;
                }

            NEXT:
                prevTime = time;
                prevSpeed = speed;
            }

            jerkCounts.TryAdd(subject,
                new Dictionary<int, List<int>>());

            jerkCounts[subject].TryAdd(task,
                new List<int>());

            jerkCounts[subject][task].Add(jerkEventCount);

            Console.WriteLine(
                $"[{subject} Task{task}] Jerkイベント回数 = {jerkEventCount}"
            );
        }

        // ==== result.xlsx に平均回数を追記 ====
        using var wbResult = new XLWorkbook(resultXlsx);
        var wsResult = wbResult.Worksheet("Result");

        foreach (var subject in jerkCounts.Keys)
        {
            int row = FindOrCreateRow(wsResult, subject);

            for (int task = 1; task <= 3; task++)
            {
                if (jerkCounts[subject].ContainsKey(task))
                {
                    wsResult.Cell(row, 1 + task).Value =
                        jerkCounts[subject][task].Average();
                }
            }
        }

        wbResult.SaveAs(resultXlsx);
        Console.WriteLine("アップデート完了！");
    }

    // A列の参加者名と一致する行を探す（なければ作る）
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
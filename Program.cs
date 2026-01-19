using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;              // ★ 追加（Average用）
using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        string sourceFolder = "/Users/moriyama_yuto/Library/CloudStorage/OneDrive-KyushuUniversity/実験/EXP1データ/EXP101";
        string csvPattern = "*.csv";
        string outputXlsxAll = "/Users/moriyama_yuto/ExcelColumnExtract/converted.xlsx";
        string resultXlsx = "/Users/moriyama_yuto/ExcelColumnExtract/result.xlsx";

        var dataR = new Dictionary<string, Dictionary<int, List<double>>>(StringComparer.OrdinalIgnoreCase);
        var dataS = new Dictionary<string, Dictionary<int, List<double>>>(StringComparer.OrdinalIgnoreCase);
        var dataT = new Dictionary<string, Dictionary<int, List<double>>>(StringComparer.OrdinalIgnoreCase);

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
            var allWs = allWb.Worksheet("Sheet1");
            int lastRow = allWs.LastRowUsed().RowNumber();

            double? foundR = null, foundS = null, foundT = null;
            bool? prevR = null;
            int sCount = 0;

            for (int r = 2; r <= lastRow; r++)
            {
                bool rBool = bool.TryParse(allWs.Cell(r, "R").GetString(), out var rb) && rb;
                if (prevR == null) prevR = rBool;
                if (foundR == null && prevR == false && rBool)
                    foundR = double.TryParse(allWs.Cell(r, "C").GetString(), out var v) ? v : null;
                prevR = rBool;

                sCount = allWs.Cell(r, "S").GetString() == "1" ? sCount + 1 : 0;
                if (foundS == null && sCount >= 3)
                    foundS = double.TryParse(allWs.Cell(r, "C").GetString(), out var sv) ? sv : null;

                if (foundT == null && bool.TryParse(allWs.Cell(r, "T").GetString(), out var tb) && tb)
                    foundT = double.TryParse(allWs.Cell(r, "C").GetString(), out var tv) ? tv : null;
            }

            AddValue(dataR, subject, taskNumber, foundR);
            AddValue(dataS, subject, taskNumber, foundS);
            AddValue(dataT, subject, taskNumber, foundT);
        }

        using var wbResult = new XLWorkbook(resultXlsx);
        var wsResult = wbResult.Worksheet("Result");

        // ===== ヘッダー =====
        foreach (var subject in dataR.Keys)
        foreach (var task in dataR[subject].Keys)
        {
            int sc = TaskStartCol(task);
            wsResult.Cell(1, sc).Value = $"task{task}";
            wsResult.Range(1, sc, 1, sc + 2).Merge();
            wsResult.Cell(2, sc).Value = "R";
            wsResult.Cell(2, sc + 1).Value = "S";
            wsResult.Cell(2, sc + 2).Value = "T";
        }

        // ===== データ =====
        foreach (var subject in dataR.Keys)
        {
            int row = FindOrCreateSubjectRow(wsResult, subject);
            foreach (var task in dataR[subject].Keys)
            {
                int sc = TaskStartCol(task);
                if (dataR[subject].ContainsKey(task))
                    wsResult.Cell(row, sc).Value = dataR[subject][task].Average();
                if (dataS.ContainsKey(subject) && dataS[subject].ContainsKey(task))
                    wsResult.Cell(row, sc + 1).Value = dataS[subject][task].Average();
                if (dataT.ContainsKey(subject) && dataT[subject].ContainsKey(task))
                    wsResult.Cell(row, sc + 2).Value = dataT[subject][task].Average();
            }
        }

        wbResult.SaveAs(resultXlsx);
        Console.WriteLine("アップデート完了！");
    }

    // ===== ここから static メソッド =====

    static int TaskStartCol(int taskNumber)
        => 2 + (taskNumber - 1) * 3;

    static int FindOrCreateSubjectRow(IXLWorksheet ws, string subject)
    {
        int lastRow = ws.LastRowUsed()?.RowNumber() ?? 2;
        for (int r = 3; r <= lastRow; r++)
            if (ws.Cell(r, 1).GetString().Equals(subject, StringComparison.OrdinalIgnoreCase))
                return r;

        ws.Cell(lastRow + 1, 1).Value = subject;
        return lastRow + 1;
    }

    static void AddValue(
        Dictionary<string, Dictionary<int, List<double>>> dict,
        string subj, int task, double? val)
    {
        if (val == null) return;
        dict.TryAdd(subj, new Dictionary<int, List<double>>());
        dict[subj].TryAdd(task, new List<double>());
        dict[subj][task].Add(val.Value);
    }
}
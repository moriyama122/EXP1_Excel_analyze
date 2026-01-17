using System;
using System.IO;
using System.Text;
using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        // --- 設定（環境に合わせて書き換えてください） ---
        string sourceFolder = "/Users/moriyama_yuto/Library/CloudStorage/OneDrive-KyushuUniversity/実験/予備実験1データ/EX101";
        string csvPattern = "EX101_task1-1.csv";  // 読みたい CSV
        string outputXlsxAll = "/Users/moriyama_yuto/ExcelColumnExtract/converted.xlsx";  // 中間Excel
        string resultXlsx = "/Users/moriyama_yuto/ExcelColumnExtract/result.xlsx";  // 最終結果
        // ----------------------------------------------------

        // 1) CSVファイルを探す
        var csvFiles = Directory.GetFiles(sourceFolder, csvPattern, SearchOption.AllDirectories);
        if (csvFiles.Length == 0)
        {
            Console.WriteLine("CSVファイルが見つかりません: " + sourceFolder);
            return;
        }

        string csvPath = csvFiles[0];
        Console.WriteLine("CSV を変換します: " + csvPath);

        // 2) CSV → Excel へ変換
        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.Worksheets.Add("Sheet1");
            using (var reader = new StreamReader(csvPath, Encoding.UTF8))
            {
                int row = 1;
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine() ?? "";
                    string[] values = line.Split(',');  // カンマ区切り

                    for (int col = 0; col < values.Length; col++)
                    {
                        ws.Cell(row, col + 1).Value = values[col];
                    }
                    row++;
                }
            }
            workbook.SaveAs(outputXlsxAll);
        }

        Console.WriteLine("変換完了: " + outputXlsxAll);

        // 3) 中間Excel から C列 と R列 を抜き出す
        using var allWb = new XLWorkbook(outputXlsxAll);
        var allWs = allWb.Worksheet("Sheet1");

        using var resultWb = new XLWorkbook();
        var resultWs = resultWb.Worksheets.Add("Result");

        // ヘッダーを設定
        resultWs.Cell(1, 1).Value = "元ファイル";
        resultWs.Cell(1, 2).Value = "C列の値";
        resultWs.Cell(1, 3).Value = "R列の値";

        int outRow = 2;

        // 最後の行番号
        var lastRow = allWs.LastRowUsed().RowNumber();

        // 2行目（ヘッダーをスキップ）からループ
        for (int r = 2; r <= lastRow; r++)
        {
            // C列（3番目）と R列（18番目）を取得
            var valueC = allWs.Cell(r, "C").GetValue<string>();
            var valueR = allWs.Cell(r, "R").GetValue<string>();

            // 空白チェック（必要なければ削除OK）
            if (string.IsNullOrEmpty(valueC) && string.IsNullOrEmpty(valueR))
                continue;

            resultWs.Cell(outRow, 1).Value = Path.GetFileName(csvPath);
            resultWs.Cell(outRow, 2).Value = valueC;
            resultWs.Cell(outRow, 3).Value = valueR;
            outRow++;
        }

        // 結果を保存
        resultWb.SaveAs(resultXlsx);

        Console.WriteLine("C列 と R列 の抽出結果を保存しました: " + resultXlsx);
    }
}
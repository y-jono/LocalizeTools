using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using ClosedXML.Excel;

/// <summary>
/// ローカライズされた HTML ファイルを生成するクラスです。
/// </summary>
public class LocalizedHtmlGenerator
{
    // プログラム全体の実行時間を計測するためのストップウォッチ
    private static Stopwatch globalStopwatch = Stopwatch.StartNew();

    /// <summary>
    /// Excel ファイルを CSV ファイルに変換します。
    /// </summary>
    /// <param name="excelPath">Excel ファイルのパス</param>
    /// <param name="csvPath">出力する CSV ファイルのパス</param>
    /// <returns>生成された CSV ファイルのパス</returns>
    private string ConvertToCsv(string excelPath, string csvPath)
    {
        Console.WriteLine("ConvertToCsv メソッドを開始します。");
        Console.WriteLine($"Excel のパス: {excelPath}, CSV のパス: {csvPath}");
        Stopwatch stopwatch = Stopwatch.StartNew(); // メソッドの実行時間を計測

        try
        {
            // Excel ファイルを読み込みます。
            using (var workbook = new XLWorkbook(excelPath))
            {
                // 最初のシートを取得します。
                var worksheet = workbook.Worksheet(1);

                // CSV ファイルに書き込みます。
                using (StreamWriter sw = new StreamWriter(csvPath, false))
                {
                    var rows = worksheet.RangeUsed().RowsUsed();
                    int rowNumber = 1;
                    foreach (var row in rows)
                    {
                        // 進行状況を表示（1000 行ごと）
                        if (rowNumber % 1000 == 0) Console.WriteLine($"Processing row {rowNumber}");
                        rowNumber++;

                        var cells = row.Cells();
                        List<string> rowData = new List<string>();
                        foreach (var cell in cells)
                        {
                            string cellValue = cell.GetValue<string>();
                            // セルの値を CSV 用にエスケープして追加
                            rowData.Add("\"" + cellValue.Replace("\"", "\"\"") + "\"");
                        }
                        // CSV の行として書き込み
                        sw.WriteLine(string.Join(",", rowData));
                    }
                }
            }

            stopwatch.Stop(); // メソッドの実行時間を記録
            Console.WriteLine("ConvertToCsv メソッドが完了しました。");
            Console.WriteLine($"ConvertToCsv の合計時間: {stopwatch.Elapsed.TotalSeconds} 秒");
            Console.WriteLine($"プログラム開始からの経過時間: {globalStopwatch.Elapsed.TotalSeconds:F2} 秒");
            return csvPath;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"エラーが発生しました: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            throw;
        }
    }

    /// <summary>
    /// ローカライズされた HTML ファイルを生成します。
    /// </summary>
    /// <param name="outputDirectory">出力する ディレクトリ名</param>
    /// <param name="outputBaseName">出力する HTML ファイルのベース名</param>
    /// <param name="cultureName">生成する言語（カルチャ）の名前（例: "English"）</param>
    /// <param name="templatePath">テンプレート HTML ファイルのパス</param>
    /// <param name="excelPath">リソースが記載された Excel ファイルのパス</param>
    /// <returns>生成された HTML ファイルのパス</returns>
    public string GenerateHtml(string outputDirectory, string outputBaseName, string cultureName, string templatePath, string excelPath)
    {
        Console.WriteLine("GenerateHtml メソッドを開始します。");
        Console.WriteLine($"カルチャ: {cultureName}, テンプレートのパス: {templatePath}, Excel のパス: {excelPath}");
        Stopwatch stopwatch = Stopwatch.StartNew(); // メソッドの実行時間を計測

        // 一時的な CSV ファイルのパスを設定
        string csvPath = Path.Combine(Path.GetDirectoryName(excelPath), "temp_resources.csv");
        Console.WriteLine("Excel を CSV に変換します...");
        ConvertToCsv(excelPath, csvPath);
        Console.WriteLine("CSV への変換が完了しました。");

        // リソースを格納する辞書を作成
        Dictionary<string, string> resources = new Dictionary<string, string>();

        // CSV ファイルを読み込みます。
        using (FileStream fs = new FileStream(csvPath, FileMode.Open, FileAccess.Read))
        using (StreamReader sr = new StreamReader(fs))
        {
            // ヘッダー行を読み込み、カルチャの列を特定します。
            string headerLine = sr.ReadLine();
            Console.WriteLine($"CSV のヘッダー: {headerLine}");
            Console.WriteLine("CSV のヘッダーを正常に読み込みました。");
            if (headerLine == null)
            {
                throw new InvalidOperationException("CSV ファイルが空です。");
            }

            // ヘッダー行をカンマで分割し、カルチャ名のインデックスを取得
            string[] headers = headerLine.Split(',');
            int cultureIndex = Array.IndexOf(headers, "\"" + cultureName + "\"");
            if (cultureIndex == -1)
            {
                Console.WriteLine("利用可能なヘッダー: " + string.Join(", ", headers));
                Console.WriteLine($"カルチャ '{cultureName}' が CSV のヘッダーに見つかりません。");
                throw new ArgumentException($"Culture '{cultureName}' not found in CSV header.");
            }

            // データ行を読み込み、リソースのキーと値を取得します。
            string line;
            int lineNumber = 1;
            while ((line = sr.ReadLine()) != null)
            {
                if (lineNumber % 1000 == 0) Console.WriteLine($"Reading CSV line {lineNumber}");
                lineNumber++;
                var values = line.Split(',');
                if (values.Length > cultureIndex)
                {
                    // キーと値を取得し、辞書に追加
                    string key = values[0].Trim('"').Replace("\"\"", "\"");
                    string value = values[cultureIndex].Trim('"').Replace("\"\"", "\"");
                    resources[key] = value;
                }
            }
        }

        // テンプレート HTML ファイルを読み込み、プレースホルダーをリソースで置換します。
        string outputHtmlPath = Path.Combine(outputDirectory, $"{outputBaseName}_{cultureName}.html");
        using (FileStream inputFs = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
        using (StreamReader templateReader = new StreamReader(inputFs))
        using (FileStream outputFs = new FileStream(outputHtmlPath, FileMode.Create, FileAccess.Write))
        using (StreamWriter outputWriter = new StreamWriter(outputFs))
        {
            string templateLine;
            while ((templateLine = templateReader.ReadLine()) != null)
            {
                foreach (var resource in resources)
                {
                    // プレースホルダーの形式は {キー} です。
                    string placeholder = "{" + resource.Key + "}";
                    if (templateLine.Contains(placeholder))
                    {
                        // プレースホルダーを対応するリソースの値で置換
                        templateLine = templateLine.Replace(placeholder, resource.Value);
                    }
                }

                // 置換後の行を書き込み
                outputWriter.WriteLine(templateLine);
            }
        }

        stopwatch.Stop(); // メソッドの実行時間を記録
        Console.WriteLine($"GenerateHtml の合計時間: {stopwatch.Elapsed.TotalSeconds} 秒");
        Console.WriteLine($"プログラム開始からの経過時間: {globalStopwatch.Elapsed.TotalSeconds:F2} 秒");

        return outputHtmlPath;
    }
}

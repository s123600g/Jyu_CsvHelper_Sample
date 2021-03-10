using CsvHelper;
using CsvHelper.Configuration;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Jyu_CsvHelper_Sample
{
    public class CsvOutputProcess
    {
        /// <summary>
        /// 輸出Csv處理
        /// </summary>
        /// <param name="setOutputPath">指定輸出路徑</param>
        /// <param name="hasHead">是否有標題</param>
        /// <param name="hasHead">每列欄位索引數字陣列，放置對應欄位需要加上頭尾雙引號用途</param>
        /// <param name="sampleModel">輸出內容</param>
        public void OutPutCsvWriter(
            string setOutputPath,
            bool hasHead,
            int[] indexForColumn,
            List<SampleModel> sampleModel
        )
        {
            // C# CsvHelper - Unable to set configuration ShouldQuote #1726
            // https://github.com/JoshClose/CsvHelper/issues/1726
            // 設定輸出條件
            var csvWriterconfig = new CsvConfiguration(CultureInfo.CurrentCulture)
            {
                ShouldQuote = (args) =>
                {
                    bool addQuote = true;

                    // 判斷是否有輸出標題列
                    if (hasHead)
                    {
                        // 取得目前讀取到第幾列
                        int getCurrentRow = args.Row.Row;

                        // 如果有開啟輸出標題列，要避開第一列加入雙引號
                        if (getCurrentRow == 1) { addQuote = false; }
                    }

                    // 如果目前讀取到列內欄位索引不在欲把欄位加上頭尾雙引號索引陣列中，就避開加入欄位頭尾雙引號
                    if (!indexForColumn.Contains(args.Row.Index)) { addQuote = false; }

                    return addQuote;
                }
            };

            // https://joshclose.github.io/CsvHelper/getting-started/
            using (var writer = new StreamWriter(setOutputPath))
            using (var csv = new CsvWriter(writer, csvWriterconfig))
            {
                // 判斷是否要輸出標題列
                if (hasHead)
                {
                    csv.WriteHeader<SampleModel>();
                    csv.NextRecord();
                }

                // 輸出內容
                foreach (var item in sampleModel)
                {
                    csv.WriteRecord(item);
                    csv.NextRecord();
                }
            }
        }
    }
}
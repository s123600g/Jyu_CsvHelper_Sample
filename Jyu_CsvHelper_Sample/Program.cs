using Microsoft.Extensions.Configuration;
using NLog;
using NLog.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;

namespace Jyu_CsvHelper_Sample
{
    internal class Program
    {
        private static IConfiguration config;

        /// <summary>
        /// 測試用輸出內容
        /// </summary>
        private static string contentSample = "This's a \"TestDoubleQuote\" sample text.";

        private static void Main(string[] args)
        {
            #region 載入組態設定檔

            // 載入appsettings json 內容
            config = new ConfigurationBuilder()
            .SetBasePath(System.IO.Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .Build();

            // NLog configuration with appsettings.json
            // https://github.com/NLog/NLog.Extensions.Logging/wiki/NLog-configuration-with-appsettings.json
            // 從組態設定檔載入NLog設定
            NLog.LogManager.Configuration = new NLogLoggingConfiguration(config.GetSection("NLog"));
            Logger logger = LogManager.GetCurrentClassLogger();

            #endregion 載入組態設定檔

            logger.Info("-------------------- Jyu_CsvHelper_Sample Start --------------------");

            try
            {
                string getOutputDirName = config.GetValue<string>("OutputDirName");
                string getOutpuFileName = config.GetValue<string>("OutputFileName");
                string GenFullOutputPath = Path.Combine(
                    Directory.GetCurrentDirectory(),
                    getOutputDirName
                );

                // 判斷輸出目錄是否有存在，未存在就建立該輸出目錄
                if (!Directory.Exists(GenFullOutputPath))
                {
                    Directory.CreateDirectory(GenFullOutputPath);
                }

                // 產生輸出路徑與檔案結合詳細路徑
                GenFullOutputPath = Path.Combine(
                    GenFullOutputPath,
                    getOutpuFileName
                );

                // 判斷檔案是否存在，存在就刪除掉
                if (File.Exists(GenFullOutputPath))
                {
                    File.Delete(GenFullOutputPath);
                }

                // 設置是否有標題列
                bool HasHead = config.GetValue<bool>("HasHead");

                // 產生要加入欄位頭尾雙引號索引陣列
                int[] indexForColumns = new int[] { 0, 1 };

                logger.Info("產生輸出內容.");

                // 產生輸出內容
                List<SampleModel> outPutSampleContents = new List<SampleModel>()
                {
                    new SampleModel{ No=1 , Content= contentSample},
                    new SampleModel{ No=2 , Content= contentSample}
                };

                logger.Info("開始輸出CSV.");

                // 輸出Csv
                CsvOutputProcess csvOutputProcess = new CsvOutputProcess();
                csvOutputProcess.OutPutCsvWriter(
                    GenFullOutputPath,
                    HasHead,
                    indexForColumns,
                    outPutSampleContents
                );

                logger.Info("完成輸出CSV.");
            }
            catch (Exception ex)
            {
                logger.Error($"{ex.Message}\n{ex.StackTrace}");
            }

            logger.Info("-------------------- Jyu_CsvHelper_Sample End --------------------");
        }
    }
}
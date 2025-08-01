using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace EpcLibrary
{
    public class CsvEpcProcessor
    {
        private const string ApiKeyHeader = "itx-apiKey";
        private const int BatchSize = 500;
        private const int SuccessStatus = 207;
        private const int SupplierId = 22;
        private const string DeviceSerialNumber = "TKZ01";

        /// <summary>
        /// 处理输入的CSV文件，仅保留每个EPC的最新记录，生成一个以"dispose"为前缀的新CSV文件，
        /// 校对原始CSV的EPC完整性，若缺失则返回缺失列表，否则将数据按批次POST到指定的URL。
        /// </summary>
        /// <param name="csvPath">处理后CSV文件路径。</param>
        /// <param name="org_csvPath">原始CSV文件路径。</param>
        /// <param name="requestUrl">POST请求的目标URL。</param>
        /// <param name="apiKeyValue">API密钥的值。</param>
        /// <returns>若校对失败，返回缺失的EPC列表字符串；若成功并发送成功，返回"发送成功"。</returns>
        public async Task<string> ProcessAsync(
            string csvPath,
            string org_csvPath,
            string requestUrl,
            string apiKeyValue)
        {
            if (!File.Exists(csvPath))
                throw new FileNotFoundException("处理后CSV文件不存在。", csvPath);
            if (!File.Exists(org_csvPath))
                throw new FileNotFoundException("原始CSV文件不存在。", org_csvPath);

            // 读取并过滤最新记录
            var records = LoadAndFilterLatestRecords(csvPath);

            // 写入处理后文件
            var disposePath = GetDisposePath(csvPath);
            WriteCsv(records, disposePath);

            // 原始和处理后EPC校对
            var originalEpcs = LoadEpcsAuto(org_csvPath);
            var processedEpcs = records.Select(r => r.Epc).ToHashSet();
            var missing = originalEpcs.Except(processedEpcs).ToList();
            if (missing.Any())
            {
                // 返回缺失的EPC列表，并停止发送
                return "缺失EPC: " + string.Join(",", missing);
            }

            // POST请求
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Add(ApiKeyHeader, apiKeyValue);

            foreach (var batch in CreateBatches(records, BatchSize))
            {
                var payload = batch.Select(r => new
                {
                    epcHex = r.Epc,
                    accessPasswordHex = r.Password,
                    encodingDate = r.Time.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ"),
                    supplierId = SupplierId,
                    deviceSerialNumber = DeviceSerialNumber,
                    tidHex = r.Tid
                });

                var content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
                var response = await client.PostAsync(requestUrl, content);
                if ((int)response.StatusCode != SuccessStatus)
                {
                    // 重试一次
                    response = await client.PostAsync(requestUrl, content);
                    if ((int)response.StatusCode != SuccessStatus)
                        throw new Exception($"批次POST失败，起始EPC: {batch.First().Epc}，状态码: {response.StatusCode}");
                }
            }

            return "发送成功";
        }

        /// <summary>
        /// 从CSV中加载所有EPC值（不去重）。
        /// </summary>
        private HashSet<string> LoadEpcsFromCsv(string path)
        {
            var lines = File.ReadAllLines(path, Encoding.UTF8);
            var set = new HashSet<string>();
            for (int i = 1; i < lines.Length; i++)
            {
                var cols = lines[i].Split(',');
                if (cols.Length >= 3 && !string.IsNullOrWhiteSpace(cols[2]))
                    set.Add(cols[2]);
            }
            return set;
        }

        /// <summary>
        /// 从Excel中加载所有EPC值（不去重）。
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private HashSet<string> LoadEpcsFromExcel(string path)
        {
            var set = new HashSet<string>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(path));
            var worksheet = package.Workbook.Worksheets[0]; // 默认第一张工作表

            int rowCount = worksheet.Dimension.End.Row;

            for (int row = 2; row <= rowCount; row++) // 从第2行开始，跳过标题
            {
                var epc = worksheet.Cells[row, 4].Text; // 第4列
                if (!string.IsNullOrWhiteSpace(epc))
                    set.Add(epc);
            }

            return set;
        }



        /// <summary>
        /// 智能判断加载所有EPC值（不去重）。
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        /// <exception cref="NotSupportedException"></exception>
        private HashSet<string> LoadEpcsAuto(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext switch
            {
                ".csv" => LoadEpcsFromCsv(path),
                ".xlsx" => LoadEpcsFromExcel(path),
                _ => throw new NotSupportedException("不支持的文件格式: " + ext)
            };
        }


        private List<EpcRecord> LoadAndFilterLatestRecords(string csvPath)
        {
            var lines = File.ReadAllLines(csvPath, Encoding.UTF8);
            if (lines.Length < 2)
                return new List<EpcRecord>();

            var records = new List<EpcRecord>(lines.Length - 1);
            for (int i = 1; i < lines.Length; i++)
            {
                var cols = lines[i].Split(',');
                if (cols.Length < 12) continue;
                if (!DateTime.TryParseExact(
                    cols[10],
                    new[] { "yyyy/M/d H:mm:ss", "yyyy/M/d HH:mm:ss", "d/M/yyyy HH:mm:ss", "yyyy-MM-dd HH:mm:ss.fffff","yyyy-MM-dd HH:mm", "yyyy/M/d H:mm", "yyyy/M/d HH:mm" },
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out var dt))
                    continue;

                records.Add(new EpcRecord
                {
                    Index = cols[0],
                    Barcode = cols[1],
                    Epc = cols[2],
                    Tid = cols[3],
                    UserArea = cols[4],
                    Password = cols[5],
                    WriteSuccess = cols[6],
                    ReadSuccess = cols[7],
                    EpcLocked = cols[8],
                    StrengthDistance = cols[9],
                    Time = dt,
                    Count = cols[11]
                });
            }

            return records
                .GroupBy(r => r.Epc)
                .Select(g => g.OrderByDescending(r => r.Time).First())
                .ToList();
        }

        private void WriteCsv(List<EpcRecord> records, string outputPath)
        {
            var header = "序号,条码,EPC,TID,用户区,密匙,写码成功,读码成功,EPC锁定,强度/读距,时间,计数";
            using var writer = new StreamWriter(outputPath, false, Encoding.UTF8);
            writer.WriteLine(header);
            foreach (var r in records)
            {
                writer.WriteLine(string.Join(",", new[]
                {
                    r.Index,
                    r.Barcode,
                    r.Epc,
                    r.Tid,
                    r.UserArea,
                    r.Password,
                    r.WriteSuccess,
                    r.ReadSuccess,
                    r.EpcLocked,
                    r.StrengthDistance,
                    r.Time.ToString("yyyy/M/d H:mm:ss", CultureInfo.InvariantCulture),
                    r.Count
                }));
            }
        }

        private IEnumerable<List<EpcRecord>> CreateBatches(List<EpcRecord> records, int size)
        {
            for (int i = 0; i < records.Count; i += size)
                yield return records.Skip(i).Take(size).ToList();
        }

        private string GetDisposePath(string originalPath)
        {
            var dir = Path.GetDirectoryName(originalPath);
            var name = Path.GetFileName(originalPath);
            return Path.Combine(dir ?? string.Empty, "dispose" + name);
        }
    }

    internal class EpcRecord
    {
        public string Index { get; set; }
        public string Barcode { get; set; }
        public string Epc { get; set; }
        public string Tid { get; set; }
        public string UserArea { get; set; }
        public string Password { get; set; }
        public string WriteSuccess { get; set; }
        public string ReadSuccess { get; set; }
        public string EpcLocked { get; set; }
        public string StrengthDistance { get; set; }
        public DateTime Time { get; set; }
        public string Count { get; set; }
    }
}

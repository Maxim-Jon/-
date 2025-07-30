using System;
using System.Threading.Tasks;
using EpcLibrary;

namespace TestConsoleApp
{
    class Program
    {
        static async Task Main(string[] args)
        {

            string inputCsv = "C:\\Users\\Sabina.jin\\Documents\\etiqueta-206211-RfidStickerVariableSmallOUTPUTsmall.Csv(1).Csv";
            string url = "https://preint-api.inditex.com/etiqrfid-rfid-provider/api/v1/product/epc/log";
            string org_csvPath = "C:\\Users\\Sabina.jin\\Documents\\test_missingdata_csv.Csv";

            try
            {
                var processor = new CsvEpcProcessor();
                var return_string = await processor.ProcessAsync(inputCsv, org_csvPath, url,"qFxIaghn606z5ZHKb0Avc5XSTgCU0tJc");
                //await processor.ProcessAsync(inputCsv, url,"qFxIaghn606z5ZHKb0Avc5XSTgCU0tJc");
                //Console.WriteLine("处理完成，结果文件已生成并成功发送请求。");
                Console.WriteLine(return_string);
                char target = ',';
                int count = return_string.Count(r => r == target)+1;
                Console.WriteLine("处理结果共有" + count + "条记录。");
            }
            catch (Exception ex)
            {
                Console.WriteLine("处理失败: " + ex.Message);
            }
            Console.ReadLine();
        }
    }
}
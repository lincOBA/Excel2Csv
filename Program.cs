using System;
using System.IO;
using System.Text;
using System.Threading;

namespace Excel2Csv
{
    class Program
    {
        static int fileIdx = 0;

        static void DoCovert(object fileName)
        {
            fileIdx++;

            Console.WriteLine("正在处理第" + fileIdx + "个文件 :" + fileName.ToString());

            ExcelConvert ct = new ExcelConvert();

            ct.ConvertCsv(fileName.ToString());
        }

        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);//注册Nuget包System.Text.Encoding.CodePages中的编码到.NET Core

            DateTime beforDT = System.DateTime.Now;

            DirectoryInfo di = new DirectoryInfo(@"xls_tmp");
            var files = di.GetFiles("*.xlsx");

            ThreadPool.SetMaxThreads(8, 8);

            foreach (var file in files)
            {
                ThreadPool.QueueUserWorkItem(DoCovert, @"xls_tmp\" + file.Name);
            }

            while (true)
            {
                Thread.Sleep(1000);//这句写着，主要是没必要循环那么多次。去掉也可以。
                int maxWorkerThreads, workerThreads;
                int portThreads;
                ThreadPool.GetMaxThreads(out maxWorkerThreads, out portThreads);
                ThreadPool.GetAvailableThreads(out workerThreads, out portThreads);
                if (maxWorkerThreads - workerThreads == 0)
                {
                    break;
                }
            }

            DateTime afterDT = System.DateTime.Now;
            TimeSpan ts = afterDT.Subtract(beforDT);
            Console.WriteLine("csv文件转换成功，耗时" + ts.Seconds.ToString() + "秒\n文件校验中，请稍候...");
        }
    }
}

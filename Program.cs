using System;
using System.IO;
using System.Text;

namespace Excel2Csv
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);//注册Nuget包System.Text.Encoding.CodePages中的编码到.NET Core
            Console.WriteLine("Hello World!");
            DateTime beforDT = System.DateTime.Now;

            DirectoryInfo di = new DirectoryInfo(@"xls_tmp");
            var files = di.GetFiles("*.xlsx");

            foreach (var file in files)
            {
                Console.WriteLine(file.Name);
                ExcelConvert ct = new ExcelConvert();
                ct.ConvertCsv(@"xls_tmp\" + file.Name, @"csv\屏蔽字库.csv");
            }

            DateTime afterDT = System.DateTime.Now;
            TimeSpan ts = afterDT.Subtract(beforDT);
            Console.WriteLine("csv文件转换成功，耗时" + ts.Seconds.ToString() + "秒\n文件校验中，请稍候...");
        }
    }
}

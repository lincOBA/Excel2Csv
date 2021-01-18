using ExcelDataReader;
using ExcelNumberFormat;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel2Csv
{
    class ExcelConvert
    {
		public string mOutPath = "";

		public const int mDoublePointKeepCnt = 5;

        public ExcelConvert(string outPath)
        {
            mOutPath = outPath;
        }

        public int getDoublePointCnt(String param)
		{
			String[] ss = param.Split(".");

			if (ss.Length <= 1)
			{
				return 0;
			}

			return ss[1].Length;
		}

		public string GetFormattedValue(IExcelDataReader reader, int columnIndex, CultureInfo culture)
        {
            var value = reader.GetValue(columnIndex);

            var formatString = reader.GetNumberFormatString(columnIndex);
            if (formatString != null)
            {
                var format = new NumberFormat(formatString);

                var forstr = format.Format(value, culture);

                if (value != null && value.GetType() == typeof(System.Double) && getDoublePointCnt(forstr) > mDoublePointKeepCnt)
				{
                    forstr = Convert.ToDouble(forstr).ToString("F" + mDoublePointKeepCnt.ToString());
                    forstr = Convert.ToDouble(forstr).ToString("G");
				}

				return forstr;
            }
            return Convert.ToString(value, culture);
        }

        public void ConvertCsv(string file)
        {
            CultureInfo culture = CultureInfo.GetCultureInfo("zh-cn");

            if (file.EndsWith(".xlsx"))
            {
                try
                {
					var stream = File.Open(file, FileMode.Open, FileAccess.Read);
					var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
					do
					{

						ConverToCSV(reader, culture);

					} while (reader.NextResult());
				}
				catch
				{
					Console.WriteLine("警告：文件已被其他线程处理 " + file + ",请截图给维护人员 ");
				}
			}
        }


        private void DealEmoji(ref string str)
        {
            str = Regex.Replace(str, @"\p{Cs}", "");
            str = Regex.Replace(str, @"\n", System.Environment.NewLine);
            str = str.Replace("\"","\"\"");
            str = str.Trim();
		}

        private bool ConverToCSV(IExcelDataReader reader, CultureInfo culture)
        {
            // sheets in excel file becomes tables in dataset
            // result.Tables[0].TableName.ToString(); // to get sheet name (table name)

            var csvCon = new System.Text.StringBuilder("");

            while (reader.Read())
            {
                bool isNullRow = true;

                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var str = GetFormattedValue(reader, i, culture);

                    if (str != "")
                    {
                        isNullRow = false;
                        break;
                    }
                }
                
                if (isNullRow)
                {
                    continue;
                }
                
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var str = GetFormattedValue(reader, i, culture);

                    DealEmoji(ref str);

                    if (str.Contains(System.Environment.NewLine) || str.Contains(","))
                    {
                        str = "\"" + str + "\"";
                    }

                    if (i < reader.FieldCount - 1)
                    {
                        csvCon.Append(str + ",");
                    }
                    else
                    {
                        csvCon.Append(str + System.Environment.NewLine);
                    }
                }
            }

            try
            {
                string output = mOutPath + @"\" + reader.Name + ".csv";
                StreamWriter csv = new StreamWriter(@output, false, Encoding.UTF8);
                csv.Write(csvCon.ToString());
                csv.Close();
            }
            catch
            {
                Console.WriteLine("INFO：跳过重复的表 " + reader.Name);
            }

            return true;
        }
    }
}

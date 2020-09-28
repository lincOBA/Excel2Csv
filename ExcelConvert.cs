﻿using ExcelDataReader;
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
        public string GetFormattedValue(IExcelDataReader reader, int columnIndex, CultureInfo culture)
        {
            var value = reader.GetValue(columnIndex);
            var formatString = reader.GetNumberFormatString(columnIndex);
            if (formatString != null)
            {
                var format = new NumberFormat(formatString);
                return format.Format(value, culture);
            }
            return Convert.ToString(value, culture);
        }

        public bool ConvertCsv(string src, string tar)
        {
            GetExcelData(src);
            return true;
        }

        private void GetExcelData(string file)
        {
            CultureInfo culture = CultureInfo.GetCultureInfo("zh-cn");

            if (file.EndsWith(".xlsx"))
            {
                var stream = File.Open(file, FileMode.Open, FileAccess.Read);
                var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                do
                {
                    ConverToCSV(reader, culture);
                } while (reader.NextResult());
            }
        }

        private void DealEmoji(ref string str)
        {
            str = Regex.Replace(str, @"\p{Cs}", "");
        }

        private bool ConverToCSV(IExcelDataReader reader, CultureInfo culture)
        {
            // sheets in excel file becomes tables in dataset
            // result.Tables[0].TableName.ToString(); // to get sheet name (table name)

            Console.WriteLine("    " + reader.Name);

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

                    if (str.Contains("\n") || str.Contains(","))
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

            string output = @"csv\" + reader.Name + ".csv";
            StreamWriter csv = new StreamWriter(@output, false);
            csv.Write(csvCon.ToString());
            csv.Close();

            return true;
        }
    }
}

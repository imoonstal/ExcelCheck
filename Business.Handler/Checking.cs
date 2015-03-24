using Excel.Handler;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Business.Handler
{
    public class Checking
    {
        public static string CheckExcel(string path)
        {
            StringBuilder sb = new StringBuilder();
            //解析规则的xml文件

            //check每一个cell
            RuleModel rules = GetRule();
            ExcelHelper helper = null;
            try
            {
                helper = new ExcelHelper(path);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                sb.AppendLine("检查失败！");
                return sb.ToString();
            }
                
            
            foreach (Rule rule in rules.Rules)
            {
                switch (rule.Type)
                {
                    case "String":
                        foreach (var l in rule.Locations)
                        {
                            string sheetName = l.loc.Split(',')[0];
                            if (l.loc.Split(',')[1].Contains("+"))
                            {
                                int startRowNum = Convert.ToInt32(l.loc.Split(',')[1].Replace("+", "")) - 1;
                                int lastRowNum = helper.GetSheetRowCount(sheetName);
                                for (int i = startRowNum; i < lastRowNum; i++)
                                {
                                    sb.AppendLine(CheckStringType(helper, rule.IsNull, l, i));
                                }
                            }
                            else
                            {
                                int rowNum = Convert.ToInt32(l.loc.Split(',')[1]) - 1;
                                sb.AppendLine(CheckStringType(helper, rule.IsNull, l, rowNum));
                            }
                        }
                        break;
                    case "Date":
                        break;
                }
            }
            if (string.IsNullOrEmpty(sb.ToString()))
            {
                sb.Append("文件内容正确！");
            }
            return sb.ToString();
        }

        private static string CheckStringType(ExcelHelper helper,string isNull, Location l, int rowNum)
        {
            string result = string.Empty;

            string sheetName = l.loc.Split(',')[0];
            int cellNum = GetCellNum(l.loc.Split(',')[2]);
            string checkContent = helper.GetStringByCell(sheetName, rowNum, cellNum);
            int len = Convert.ToInt32(l.loc.Split(',')[3]);
            int realLen = GetStringLenth(checkContent);
            if (realLen == 0)
            {
                if (isNull == "N")
                {
                    result = "“" + sheetName + "”中的" + l.loc.Split(',')[1] + "行" + l.loc.Split(',')[2] + "列内容不允许为空！";
                }
            }
            else
            {
                if (realLen > len)
                {
                    result = "“" + sheetName + "”中的" + l.loc.Split(',')[1] + "行" + l.loc.Split(',')[2] + "列内容超过规定长度！";
                }
            }
            return result;
        }

        public static bool CheckString(string inputStr, int length)
        {
            bool result = false;
            int l = Encoding.Default.GetBytes(inputStr).Length;
            if (l < length)
                result = true;
            return result;
        }

        private static RuleModel GetRule()
        {
            RuleModel result = new RuleModel();
            return result;
        }

        private static int GetCellNum(string cellNum)
        {
            int result = 0;
            switch (cellNum)
            {
                case "A":
                    result = 0;
                    break;
                case "B":
                    result = 1;
                    break;
                case "C":
                    result = 2;
                    break;
                case "D":
                    result = 3;
                    break;
                case "E":
                    result = 4;
                    break;
                case "F":
                    result = 5;
                    break;
                case "G":
                    result = 6;
                    break;
                case "H":
                    result = 7;
                    break;
                case "I":
                    result = 8;
                    break;
                case "J":
                    result = 9;
                    break;
                case "K":
                    result = 10;
                    break;
                case "L":
                    result = 11;
                    break;
                case "M":
                    result = 12;
                    break;
                case "N":
                    result = 13;
                    break;
                case "O":
                    result = 14;
                    break;
                case "P":
                    result = 15;
                    break;
                case "Q":
                    result = 16;
                    break;
                case "R":
                    result = 17;
                    break;
                case "S":
                    result = 18;
                    break;
                case "T":
                    result = 19;
                    break;
                case "U":
                    result = 20;
                    break;
                case "V":
                    result = 21;
                    break;
                case "W":
                    result = 22;
                    break;
                case "X":
                    result = 23;
                    break;
                case "Y":
                    result = 24;
                    break;
                case "Z":
                    result = 25;
                    break;
                default:
                    result = -1;
                    break;
            }
            return result;
        }

        private static int GetStringLenth(string inputStr)
        {
            if (inputStr.Length == 0)
                return 0;
            else
                return Encoding.Default.GetBytes(inputStr).Length;
        }
    }
}

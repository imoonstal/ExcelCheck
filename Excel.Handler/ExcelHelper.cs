using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Excel.Handler
{
    public class ExcelHelper
    {
        private IWorkbook wb = null;
        public ExcelHelper(string filePath)
        {
            try
            {
                using (Stream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    wb = WorkbookFactory.Create(stream);
                }
            }
            catch (FileNotFoundException ex)
            {
                throw ex;
            }
            catch (UnauthorizedAccessException ex)
            {
                throw ex;
            }
            catch (IOException ex)
            {
                throw ex;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string GetStringByCell(int sheetIndex, int rowNum, int cellNum)
        {
            string result = string.Empty;
            ISheet sheet = GetSheetBy(sheetIndex);
            result = GetContent(sheet, rowNum, cellNum);
            return result;
        }

        public string GetStringByCell(string sheetName, int rowNum, int cellNum)
        {
            string result = string.Empty;
            ISheet sheet = GetSheetBy(sheetName);
            result = GetContent(sheet, rowNum, cellNum);
            return result;
        }

        private string GetContent(ISheet sheet, int rowNum, int cellNum)
        {
            string result = string.Empty;
            if (null == sheet)
            {
                Console.WriteLine("未找到“" + sheet.SheetName + "”这一页");
                return result;
            }
            IRow row = sheet.GetRow(rowNum);
            ICell cell = row.GetCell(cellNum);
            result = cell.StringCellValue;
            return result;
        }

        private ISheet GetSheetBy(int sheetIndex)
        {

            ISheet result = wb.GetSheetAt(sheetIndex);
            return result;
        }

        private ISheet GetSheetBy(string sheetName)
        {
            ISheet result = wb.GetSheet(sheetName);
            return result;
        }

        public int GetSheetRowCount(string sheetName)
        {
            return GetSheetBy(sheetName).LastRowNum;
        }

        public int GetSheetRowCount(int sheetIndex)
        {
            return GetSheetBy(sheetIndex).LastRowNum;
        }
    }
}

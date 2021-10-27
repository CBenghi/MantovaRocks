using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace IfcDemo
{
    public class AnExcelReader
    {
        internal static ISheet GetExcelSheet(string path, string sheetName)
        {
            IWorkbook book;

            try
            {
                FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                // Try to read workbook as XLSX:
                try
                {
                    book = new XSSFWorkbook(fs);
                }
                catch
                {
                    book = null;
                }

                // If reading fails, try to read workbook as XLS:
                if (book == null)
                {
                    book = new HSSFWorkbook(fs);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Excel read error", MessageBoxButton.OK, MessageBoxImage.Error);

                return null;
            }
            return book.GetSheet(sheetName);
        }
    }
}
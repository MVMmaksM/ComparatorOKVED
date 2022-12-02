using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ComparatorOKVED.Models;
using System.Windows;
using NLog;
using System.Threading;

namespace ComparatorOKVED
{
    class FileServices
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        public static List<Model> ReadExcel(string pathExcelFile)
        {
            logger.Info("Чтение Excel");
            
            try
            {
                ExcelPackage package = new ExcelPackage(pathExcelFile);
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                List<Model> list = new List<Model>();
                int colOKPO = 0;
                int colOKVED = 0;

                for (int i = 1; i < worksheet.Dimension.End.Column; i++)
                {
                    if (worksheet.Cells[1, i].Value.ToString() == "ОКПО")
                    {
                        colOKPO = i;
                        break;
                    }
                }

                for (int i = 1; i < worksheet.Dimension.End.Column; i++)
                {
                    if (worksheet.Cells[1, i].Value.ToString() == "ОКВЭД Чистый")
                    {
                        colOKVED = i;
                        break;
                    }
                }

                for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
                {
                    if ((worksheet.Cells[i, colOKPO].Value != null) && (worksheet.Cells[i, colOKVED].Value != null))
                    {
                        Model newRow = new Model();
                        newRow.OKPO = worksheet.Cells[i, colOKPO].Value.ToString();
                        newRow.OKVED = worksheet.Cells[i, colOKVED].Value.ToString();
                        list.Add(newRow);
                    }
                }                

                return list;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                logger.Error(ex.Message + ex.StackTrace);
                return null;
            }
        }

        public static byte[] WriteToExcel(List<ModelResult> ResultCompareOKVED)
        {
            logger.Info("Записиь в Excel");

            try
            {
                var package = new ExcelPackage();
                var sheetResultCompareOKVED = package.Workbook.Worksheets.Add("Сравнение ОКВЭД");
                var sheetOkpoPredNotIn = package.Workbook.Worksheets.Add("Предыдущий текущий ОКВЭД");


                int rowsCountListCompareOKVED = ResultCompareOKVED.Count + 1;

                sheetResultCompareOKVED.Cells["A1"].Value = "ОКПО";
                sheetResultCompareOKVED.Cells["B1"].Value = "ОКВЭД в предыдущем периоде";
                sheetResultCompareOKVED.Cells["C1"].Value = "ОКВЭД в текущем периоде";

                sheetResultCompareOKVED.Cells[2, 1].LoadFromCollection<ModelResult>(ResultCompareOKVED);

                sheetResultCompareOKVED.View.FreezePanes(2, 1);
                sheetResultCompareOKVED.Cells[1, 1, 1, 3].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                sheetResultCompareOKVED.Cells[1, 1, 1, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(217, 217, 217));
                sheetResultCompareOKVED.Columns[1].Width = 32;
                sheetResultCompareOKVED.Columns[2].Width = 32;
                sheetResultCompareOKVED.Columns[3].Width = 32;
                sheetResultCompareOKVED.Columns[1, 3].Style.Font.Name = "Arial";
                sheetResultCompareOKVED.Columns[1, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheetResultCompareOKVED.Columns[1, 3].Style.Font.Size = 10;
                sheetResultCompareOKVED.Cells[1, 1, 1, 3].Style.Font.Bold = true;
                sheetResultCompareOKVED.Cells[1, 1, 1, 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                sheetResultCompareOKVED.Rows[1].Height = 30;
                sheetResultCompareOKVED.Cells[1, 1, rowsCountListCompareOKVED, 3].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                sheetResultCompareOKVED.Cells[1, 1, rowsCountListCompareOKVED, 3].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                sheetResultCompareOKVED.Cells[1, 1, rowsCountListCompareOKVED, 3].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                sheetResultCompareOKVED.Cells[1, 1, rowsCountListCompareOKVED, 3].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                //int rowsCountListOkpoNotIn = listOkpoPredNotIn.Count + 1;

                //sheetOkpoPredNotIn.Cells["A1"].Value = "ОКПО";
                //sheetOkpoPredNotIn.Cells["B1"].Value = "ОКВЭД в предыдущем периоде";
                //sheetOkpoPredNotIn.Cells["C1"].Value = "ОКВЭД в текущем периоде";

                //sheetOkpoPredNotIn.Cells[2, 1].LoadFromCollection<ModelResult>(listOkpoPredNotIn);

                //sheetOkpoPredNotIn.View.FreezePanes(2, 1);
                //sheetOkpoPredNotIn.Cells[1, 1, 1, 3].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //sheetOkpoPredNotIn.Cells[1, 1, 1, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(217, 217, 217));
                //sheetOkpoPredNotIn.Columns[1].Width = 32;
                //sheetOkpoPredNotIn.Columns[2].Width = 32;
                //sheetOkpoPredNotIn.Columns[3].Width = 32;
                //sheetOkpoPredNotIn.Columns[1, 3].Style.Font.Name = "Arial";
                //sheetOkpoPredNotIn.Columns[1, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                //sheetOkpoPredNotIn.Columns[1, 3].Style.Font.Size = 10;
                //sheetOkpoPredNotIn.Cells[1, 1, 1, 3].Style.Font.Bold = true;
                //sheetOkpoPredNotIn.Cells[1, 1, 1, 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                //sheetOkpoPredNotIn.Rows[1].Height = 30;
                //sheetOkpoPredNotIn.Cells[1, 1, rowsCountListOkpoNotIn, 3].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                //sheetOkpoPredNotIn.Cells[1, 1, rowsCountListOkpoNotIn, 3].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                //sheetOkpoPredNotIn.Cells[1, 1, rowsCountListOkpoNotIn, 3].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                //sheetOkpoPredNotIn.Cells[1, 1, rowsCountListOkpoNotIn, 3].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                return package.GetAsByteArray();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                logger.Error(ex.Message + ex.StackTrace);
                return null;
            }
        }
    }
}

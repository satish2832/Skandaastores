using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Web;
using Microsoft.Office.Interop.Excel;
using StoreSite.Models;

namespace StoreSite.Helpers
{
    public class DataHelper
    {

        public List<Product> GetProducts()
        {
            return ReadExcelFile();
        }
        private List<Product> ReadExcelFile()
        {
            string filePath = System.Web.HttpContext.Current.Server.MapPath("~/DataSources/Product_Details.xlsx");
            var lsProducts = new List<Product>();
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
                int totalRows = xlRange.Rows.Count;
                int totalColumns = xlRange.Columns.Count;
                for (int rowCount = 2; rowCount <= totalRows; rowCount++)
                {
                    if (string.IsNullOrEmpty(CheckNullFromExcel(xlRange.Cells[rowCount, 1])))
                    {
                        continue;
                    }
                    lsProducts.Add(new Product()
                    {
                        Code = CheckNullFromExcel(xlRange.Cells[rowCount, 1]),
                        Title = CheckNullFromExcel(xlRange.Cells[rowCount, 2]),
                        Description = CheckNullFromExcel(xlRange.Cells[rowCount, 3]),
                        OldValue = CheckNullFromExcel(xlRange.Cells[rowCount, 4]),
                        NewValue = CheckNullFromExcel(xlRange.Cells[rowCount, 5]),
                        Discount = CheckNullFromExcel(xlRange.Cells[rowCount, 6]),
                        Variant = CheckNullFromExcel(xlRange.Cells[rowCount, 7]),
                        Colors = CheckNullFromExcel(xlRange.Cells[rowCount, 8])
                    });
                }
                xlWorkBook.Close();
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
                return lsProducts;
            }
            catch (Exception ex)
            {
                return lsProducts;
            }

        }
        public void AddCustomerEnquiry(CustomerEnquiry customerEnquiry)
        {
            AddNewRowsToExcelFile(customerEnquiry);
        }
        private void AddNewRowsToExcelFile(CustomerEnquiry customerEnquiry)
        {
            string filePath = System.Web.HttpContext.Current.Server.MapPath("~/DataSources/Customer_Enquiry.xlsx");

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false,
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange;
            int rowNumber = xlRange.Rows.Count + 1;

            xlWorkSheet.Cells[rowNumber, 1] = customerEnquiry.CustomerName;
            xlWorkSheet.Cells[rowNumber, 2] = customerEnquiry.EmailAddress;
            xlWorkSheet.Cells[rowNumber, 3] = customerEnquiry.PhoneNumber;
            xlWorkSheet.Cells[rowNumber, 4] = customerEnquiry.Description;
            xlWorkSheet.Cells[rowNumber, 5] = customerEnquiry.ProductCode;
            xlWorkSheet.Cells[rowNumber, 6] = customerEnquiry.Quantity;

            xlWorkBook.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);

            xlWorkBook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }
        private string CheckNullFromExcel(Microsoft.Office.Interop.Excel.Range exRange)
        {
            try
            {
                if (exRange != null && !string.IsNullOrEmpty(exRange.Text))
                {
                    return Convert.ToString(exRange.Text);
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                return string.Empty;
            }
        }
    }
}
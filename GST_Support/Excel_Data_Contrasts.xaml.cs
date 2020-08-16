using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using Microsoft.Win32;
namespace GST_Support
{
    /// <summary>
    /// Interaction logic for Excel_Data_Contrasts.xaml
    /// </summary>
    public partial class Excel_Data_Contrasts : Window
    {
        DataTable dt_Tally = new DataTable();
        public Excel_Data_Contrasts()
        {
            InitializeComponent();
            dt_Tally.Columns.Add("Col_Date", typeof(string));
            dt_Tally.Columns.Add("Particulars", typeof(string));
            dt_Tally.Columns.Add("GSTIN", typeof(string));
            dt_Tally.Columns.Add("Vch_Type", typeof(string));
            dt_Tally.Columns.Add("Vch_no", typeof(string));
            dt_Tally.Columns.Add("Invoice_No", typeof(string));
            dt_Tally.Columns.Add("Invoice_Date", typeof(string));
            dt_Tally.Columns.Add("Taxable_Value", typeof(string));
            dt_Tally.Columns.Add("Integrated_Tax_Amount", typeof(string));
            dt_Tally.Columns.Add("Central_Tax_Amount", typeof(string));
            dt_Tally.Columns.Add("State_Tax_Amount", typeof(string));
            dt_Tally.Columns.Add("Cess_Amount", typeof(string));
            dt_Tally.Columns.Add("Total_Tax_Amount", typeof(string));
        }

        private void btn_Validate_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btn_GST_Browse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                txt_GST_FileLoc.Text = openFileDialog.FileName.ToString();
        }

        private void btn_Tally_Browse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                txt_Tally_FileLoc.Text = openFileDialog.FileName.ToString();
        }
        private void btn_Read_ExcelData_Click(object sender, RoutedEventArgs e)
        {
            ReadTallyData();
        }

        public void ReadTallyData()
        {
            string str_Tally_FilePath = txt_Tally_FileLoc.Text;

            int int_rownumber = 0; int int_ColNumber = 0;
            using (SpreadsheetDocument spreadsheetDocument_Tally =
                SpreadsheetDocument.Open(str_Tally_FilePath, false))
            {

                WorkbookPart bkPart = spreadsheetDocument_Tally.WorkbookPart;
                DocumentFormat.OpenXml.Spreadsheet.Workbook workbook = bkPart.Workbook;
                DocumentFormat.OpenXml.Spreadsheet.Sheet s = workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(sht => sht.Name == txt_tally_SheetName.Text).FirstOrDefault();
                WorksheetPart wsPart = (WorksheetPart)bkPart.GetPartById(s.Id);
                DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData = wsPart.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetData>().FirstOrDefault();
                SharedStringTablePart stringTablePart = spreadsheetDocument_Tally.WorkbookPart.SharedStringTablePart;


                //WorkbookPart workbookPart = spreadsheetDocument_Tally.WorkbookPart;
                //DocumentFormat.OpenXml.Spreadsheet.Sheet s = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(sht => sht.Name == "Sheet1").FirstOrDefault();
                //WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                //SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                string str_CellValue;

                foreach (Row r in sheetData.Elements<Row>())
                {
                    if (int_rownumber >= 9)
                    {
                        int_ColNumber = 0;
                        DataRow dr_Tally_Row = dt_Tally.NewRow();
                        foreach (Cell c in r.Elements<Cell>())
                        {
                            try
                            {

                                if (c.CellValue == null)
                                {
                                    if (int_ColNumber == 0) break;
                                    int_ColNumber += 1;
                                    continue;
                                }
                                str_CellValue = "";
                                str_CellValue = (c.CellValue == null) ? c.InnerText : c.CellValue.Text;

                                if (int_ColNumber == 0 || int_ColNumber == 6)
                                {
                                    str_CellValue = DateTime.FromOADate(Convert.ToDouble(str_CellValue)).ToShortDateString();
                                }
                                else
                                {
                                    if (c.DataType != null && c.DataType.Value == CellValues.SharedString)
                                    {
                                        str_CellValue = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(c.CellValue.Text)].InnerText;
                                    }
                                }

                                dr_Tally_Row[int_ColNumber] = str_CellValue;
                                int_ColNumber += 1;

                            }
                            catch (Exception ex)
                            {

                            }
                        }
                        dt_Tally.Rows.Add(dr_Tally_Row);

                    }
                    int_rownumber = int_rownumber + 1;
                }
            }
        }

    }

}

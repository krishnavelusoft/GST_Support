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
        DataTable dt_GST = new DataTable();
        string str_rowRemarks;
     

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
            dt_Tally.Columns.Add("TallyExcelRowNumber", typeof(string));
            dt_Tally.Columns.Add("GSTExcelRowNumber", typeof(string));
            dt_Tally.Columns.Add("PercentageMatching", typeof(string));
            dt_Tally.Columns.Add("Remarks", typeof(string));

            dt_GST.Columns.Add("GSTIN_of_supplier", typeof(string));
            dt_GST.Columns.Add("Legal_name_of_Supplier", typeof(string));
            dt_GST.Columns.Add("Invoice_number", typeof(string));
            dt_GST.Columns.Add("Invoice_type", typeof(string));
            dt_GST.Columns.Add("Invoice_Date", typeof(string));
            dt_GST.Columns.Add("Invoice_Value", typeof(string));
            dt_GST.Columns.Add("Place_of_supply", typeof(string));
            dt_GST.Columns.Add("Supply_Attract_Reverse_Charge", typeof(string));
            dt_GST.Columns.Add("Rate", typeof(string));
            dt_GST.Columns.Add("Taxable_Value", typeof(string));
            dt_GST.Columns.Add("Integrated_Tax", typeof(string));
            dt_GST.Columns.Add("Central_Tax", typeof(string));
            dt_GST.Columns.Add("State_Tax", typeof(string));
            dt_GST.Columns.Add("Cess", typeof(string));
            dt_GST.Columns.Add("Return_status", typeof(string));
            dt_GST.Columns.Add("GSTExcelRowNumber", typeof(string));
            dt_GST.Columns.Add("TallyExcelRowNumber", typeof(string));
            dt_GST.Columns.Add("PercentageMatching", typeof(string));
            dt_Tally.Columns.Add("Remarks", typeof(string));

        }

        public void ClearValue()
        {
            str_rowRemarks = "";
            dt_Tally.Rows.Clear();
            dt_GST.Rows.Clear();
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
            ClearValue();
            ReadTallyData();
            ReadGSTData();
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
                        dr_Tally_Row["TallyExcelRowNumber"] = r.RowIndex.Value.ToString();
                        dt_Tally.Rows.Add(dr_Tally_Row);

                    }
                    int_rownumber = int_rownumber + 1;
                }
            }
        }

        public void ReadGSTData()
        {
            string str_Tally_FilePath = txt_GST_FileLoc.Text;
            
            int int_rownumber = 0; int int_ColNumber = 0;
            using (SpreadsheetDocument spreadsheetDocument_GST =
                SpreadsheetDocument.Open(str_Tally_FilePath, false))
            {

                WorkbookPart bkPart = spreadsheetDocument_GST.WorkbookPart;
                DocumentFormat.OpenXml.Spreadsheet.Workbook workbook = bkPart.Workbook;
                DocumentFormat.OpenXml.Spreadsheet.Sheet s = workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(sht => sht.Name == txt_GST_SheetName.Text).FirstOrDefault();
                WorksheetPart wsPart = (WorksheetPart)bkPart.GetPartById(s.Id);
                DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData = wsPart.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetData>().FirstOrDefault();
                SharedStringTablePart stringTablePart = spreadsheetDocument_GST.WorkbookPart.SharedStringTablePart;

                string str_CellValue;
                bool bol_adddata = false;
                bool bol_dublicatedata = false;
                foreach (Row r in sheetData.Elements<Row>())
                {
                    if (int_rownumber >= 6)
                    {
                        int_ColNumber = 0;
                        DataRow dr_GST_Row = dt_GST.NewRow();
                        bol_adddata = false;
                        foreach (Cell c in r.Elements<Cell>())
                        {
                            try
                            {

                                if (c.CellValue == null)
                                {
                                    bol_dublicatedata = false;
                                    break;
                                }
                                else if (bol_dublicatedata)
                                {
                                    bol_dublicatedata = false;
                                    break;
                                }
                                str_CellValue = "";
                                str_CellValue = (c.CellValue == null) ? c.InnerText : c.CellValue.Text;

                                //if (int_ColNumber == 4)
                                //{
                                //    str_CellValue = DateTime.FromOADate(Convert.ToDouble(str_CellValue)).ToShortDateString();
                                //}
                                //else
                                //{
                                if (c.DataType != null && c.DataType.Value == CellValues.SharedString)
                                {
                                    str_CellValue = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(c.CellValue.Text)].InnerText;
                                }
                                //}
                                bol_adddata = true;
                                dr_GST_Row[int_ColNumber] = str_CellValue;
                                int_ColNumber += 1;

                            }
                            catch (Exception ex)
                            {

                            }
                        }
                        if (bol_adddata)
                        {
                            bol_dublicatedata = true;
                            dr_GST_Row["GSTExcelRowNumber"] = r.RowIndex.Value.ToString();
                            dt_GST.Rows.Add(dr_GST_Row);
                        }
                    }
                    int_rownumber = int_rownumber + 1;

                }
            }
        }

        private void btn_Validate_Click(object sender, RoutedEventArgs e)
        {
            ValidateGSTData();
        }

        public void ValidateGSTData()
        {

            int gstLoop = 0;
            for (gstLoop = 0; gstLoop < dt_GST.Rows.Count; gstLoop++)
            {
                try
                {
                    //if (gstLoop % 2 == 0)
                    //    continue;


                    var dr_Tally_data = (from tallyRow in dt_Tally.AsEnumerable()
                                         where tallyRow.Field<string>("GSTIN") == dt_GST.Rows[gstLoop]["GSTIN_of_supplier"].ToString()
                                         && tallyRow.Field<string>("Invoice_No") == dt_GST.Rows[gstLoop]["Invoice_number"].ToString()
                                         && tallyRow.Field<string>("Taxable_Value") == dt_GST.Rows[gstLoop]["Taxable_Value"].ToString()
                                         && tallyRow.Field<string>("Central_Tax_Amount") == dt_GST.Rows[gstLoop]["Central_Tax"].ToString()
                                         && tallyRow.Field<string>("State_Tax_Amount") == dt_GST.Rows[gstLoop]["State_Tax"].ToString()
                                         && tallyRow.Field<string>("Invoice_Date") == dt_GST.Rows[gstLoop]["Invoice_Date"].ToString()
                                         select tallyRow).ToList();
                    if (dr_Tally_data.Count() == 1)
                    {
                        dr_Tally_data[0]["GSTExcelRowNumber"] = dt_GST.Rows[gstLoop]["GSTExcelRowNumber"].ToString();
                        dt_Tally.AcceptChanges();
                        dt_GST.Rows[gstLoop]["TallyExcelRowNumber"] = dr_Tally_data[0]["TallyExcelRowNumber"].ToString();
                        dt_GST.AcceptChanges();
                        continue;
                    }

                    dr_Tally_data = (from tallyRow in dt_Tally.AsEnumerable()
                                     where tallyRow.Field<string>("GSTIN") == dt_GST.Rows[gstLoop]["GSTIN_of_supplier"].ToString()
                                     select tallyRow).ToList();

                    for (int iloop = 0; iloop < dr_Tally_data.Count(); iloop++)
                    {
                        int int_total_per = 0;
                        str_rowRemarks = "";
                        string str_GST = dt_GST.Rows[gstLoop]["Invoice_number"].ToString().Trim();
                        string str_Tally = dr_Tally_data[iloop]["Invoice_No"].ToString().Trim();
                        int int_INV_Number_Percentage = fn_GST_Tally_Data_Maching(str_GST, str_Tally,"Invoice Number");
                        int_total_per += int_INV_Number_Percentage;

                        str_GST = dt_GST.Rows[gstLoop]["Invoice_Date"].ToString().Trim();
                        str_Tally = dr_Tally_data[iloop]["Invoice_Date"].ToString().Trim();
                        int int_Inv_Date_Percentage = fn_GST_Tally_Data_Maching(str_GST, str_Tally, "State Tax");
                        int_total_per += int_Inv_Date_Percentage;


                        str_GST = dt_GST.Rows[gstLoop]["Taxable_Value"].ToString().Trim();
                        str_Tally = dr_Tally_data[iloop]["Taxable_Value"].ToString().Trim();
                        int int_TaxValue_Percentage = fn_GST_Tally_Data_Maching(str_GST, str_Tally, "Taxable Value");
                        int_total_per += int_TaxValue_Percentage;


                        str_GST = dt_GST.Rows[gstLoop]["Central_Tax"].ToString().Trim();
                        str_Tally = dr_Tally_data[iloop]["Central_Tax_Amount"].ToString().Trim();
                        int int_CentralTax_Percentage = fn_GST_Tally_Data_Maching(str_GST, str_Tally, "Central Tax");
                        int_total_per += int_CentralTax_Percentage;

                        str_GST = dt_GST.Rows[gstLoop]["State_Tax"].ToString().Trim();
                        str_Tally = dr_Tally_data[iloop]["State_Tax_Amount"].ToString().Trim();
                        int int_StateTax_Percentage = fn_GST_Tally_Data_Maching(str_GST, str_Tally, "State Tax");
                        int_total_per += int_StateTax_Percentage;


                        int int_average = int_total_per / 5;

                        if(int_average ==100)
                        {
                            dr_Tally_data[iloop]["GSTExcelRowNumber"] = dt_GST.Rows[gstLoop]["GSTExcelRowNumber"].ToString();
                            dr_Tally_data[iloop]["PercentageMatching"] = int_average.ToString();
                            dt_Tally.AcceptChanges();
                            dt_GST.Rows[gstLoop]["TallyExcelRowNumber"] = dr_Tally_data[iloop]["TallyExcelRowNumber"].ToString();
                            dt_GST.Rows[gstLoop]["PercentageMatching"] = int_average.ToString();
                            dt_GST.AcceptChanges();
                            break;
                        }
                        else if(int_average >=80)
                        {
                            dr_Tally_data[iloop]["GSTExcelRowNumber"] = dt_GST.Rows[gstLoop]["GSTExcelRowNumber"].ToString();
                            dr_Tally_data[iloop]["PercentageMatching"] = int_average.ToString();
                            dr_Tally_data[iloop]["Remarks"] = str_rowRemarks;
                            dt_Tally.AcceptChanges();

                            dt_GST.Rows[gstLoop]["TallyExcelRowNumber"] = dr_Tally_data[iloop]["TallyExcelRowNumber"].ToString();
                            dt_GST.Rows[gstLoop]["PercentageMatching"] = int_average.ToString();
                            dt_GST.Rows[gstLoop]["Remarks"] = str_rowRemarks;
                            dt_GST.AcceptChanges();
                        }
                    }
                    //&& tallyRow.Field<string>("Taxable_Value") == dt_GST.Rows[gstLoop]["Taxable_Value"].ToString()
                    //                && tallyRow.Field<string>("Central_Tax_Amount") == dt_GST.Rows[gstLoop]["Central_Tax"].ToString()
                    //                 && tallyRow.Field<string>("State_Tax_Amount") == dt_GST.Rows[gstLoop]["State_Tax"].ToString()
                }
                catch (Exception ex)
                {

                }
            }
        }

        public int fn_GST_Tally_Data_Maching(string str_GST, string str_Tally,string DataFor)
        {
            int int_PercentageMaching = 0;
            if (str_GST == str_Tally)
            {
                int_PercentageMaching = 100;
                str_rowRemarks += DataFor + "  Direct Maching  " + int_PercentageMaching.ToString() + "%" + System.Environment.NewLine ;
                return int_PercentageMaching;
            }

            decimal gst_number = 0, tally_Number = 0;
            bool canConvert = decimal.TryParse(str_GST, out gst_number);
            if (canConvert)
            {
                canConvert = decimal.TryParse(str_Tally, out tally_Number);
                if (canConvert)
                {
                    if (gst_number == tally_Number)
                    {
                        int_PercentageMaching = 100;
                        str_rowRemarks += DataFor + "  Decimal Conversion Direct Match  " + int_PercentageMaching.ToString() + "%" + System.Environment.NewLine;
                        return int_PercentageMaching;
                    }
                }
            }
            WordsMatching.MatchsMaker match = new WordsMatching.MatchsMaker(str_GST, str_Tally);
            int_PercentageMaching = Convert.ToInt32(match.Score * 100);
            str_rowRemarks += DataFor + "  MatchMaker Out   " + int_PercentageMaching.ToString() + "%" + System.Environment.NewLine;
            return int_PercentageMaching;



        }

    }

}

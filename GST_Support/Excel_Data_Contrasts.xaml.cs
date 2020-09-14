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
        DataTable dt_matching_Det = new DataTable();
        string str_rowRemarks;
        int int_Selected = 1;//1--> GST ; 2-->TALLY

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
            dt_Tally.Columns.Add("TallyExcelRowNumber", typeof(Int32));
            dt_Tally.Columns.Add("GSTExcelRowNumber", typeof(Int32));
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
            dt_GST.Columns.Add("GSTExcelRowNumber", typeof(Int32));
            dt_GST.Columns.Add("TallyExcelRowNumber", typeof(Int32));
            dt_GST.Columns.Add("PercentageMatching", typeof(string));
            dt_GST.Columns.Add("Remarks", typeof(string));

            dt_matching_Det.Columns.Add("GSTExcelRowNumber");
            dt_matching_Det.Columns.Add("TallyExcelRowNumber");
            dt_matching_Det.Columns.Add("PercentageMatching", typeof(string));
            dt_matching_Det.Columns.Add("Remarks", typeof(string));
            dt_matching_Det.Columns.Add("Invoice_number_Per", typeof(string));
            dt_matching_Det.Columns.Add("Invoice_Date_Per", typeof(string));
            dt_matching_Det.Columns.Add("Taxable_Value_Per", typeof(string));
            dt_matching_Det.Columns.Add("Integrated_Tax_Per", typeof(string));
            dt_matching_Det.Columns.Add("Central_Tax_Per", typeof(string));
            dt_matching_Det.Columns.Add("State_Tax_Per", typeof(string));

            ClearValue();
            txt_Tally_INV_Number.Text = "krishna";
            pb_inv_number.Value = 30;
        }

        public void ClearValue()
        {
            str_rowRemarks = "";
            dt_Tally.Rows.Clear();
            dt_GST.Rows.Clear();
            dg_GST.Columns[1].Visibility = Visibility.Hidden;
            dg_GST.Columns[2].Visibility = Visibility.Hidden;
            dg_GST.Columns[3].Visibility = Visibility.Hidden;

            dg_Tally.Columns[1].Visibility = Visibility.Hidden;
            dg_Tally.Columns[2].Visibility = Visibility.Hidden;
            dg_Tally.Columns[3].Visibility = Visibility.Hidden;

            //txt_GST_FileLoc.Text = @"F:\GST\Document\book3.xlsx";
            //txt_Tally_FileLoc.Text = txt_GST_FileLoc.Text;

            //btn_Read_ExcelData_Click(btn_Read_ExcelData, new RoutedEventArgs());
            //btn_Validate_Click(btn_Validate, new RoutedEventArgs());

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

            dg_GST.ItemsSource = dt_GST.DefaultView;
            dg_Tally.ItemsSource = dt_Tally.DefaultView;

            dg_GST.Columns[3].Visibility = Visibility.Visible;
            dg_Tally.Columns[3].Visibility = Visibility.Visible;

        }

        public void ReadTallyData()
        {
            string str_Tally_FilePath = txt_Tally_FileLoc.Text;

            int int_ColNumber = 0;
            using (SpreadsheetDocument spreadsheetDocument_Tally =
                SpreadsheetDocument.Open(str_Tally_FilePath, false))
            {

                WorkbookPart bkPart = spreadsheetDocument_Tally.WorkbookPart;
                DocumentFormat.OpenXml.Spreadsheet.Workbook workbook = bkPart.Workbook;
                DocumentFormat.OpenXml.Spreadsheet.Sheet s = workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Where(sht => sht.Name == txt_tally_SheetName.Text).FirstOrDefault();
                WorksheetPart wsPart = (WorksheetPart)bkPart.GetPartById(s.Id);
                DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData = wsPart.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetData>().FirstOrDefault();
                SharedStringTablePart stringTablePart = spreadsheetDocument_Tally.WorkbookPart.SharedStringTablePart;


                string str_CellValue;
                bool bol_RecordStarted = false;
                bool bol_InsertRow = false;
                string str_Ready = "NOT_READY";
                foreach (Row r in sheetData.Elements<Row>())
                {

                    bol_InsertRow = false;
                    int_ColNumber = 0;
                    DataRow dr_Tally_Row = dt_Tally.NewRow();
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        try
                        {
                            str_CellValue = "";
                            if (bol_RecordStarted == false)
                            {
                                if (c.CellValue == null && bol_RecordStarted == false)
                                {
                                    break;
                                }

                                str_CellValue = (c.CellValue == null) ? c.InnerText : c.CellValue.Text;
                                if (c.DataType != null && c.DataType.Value == CellValues.SharedString)
                                {
                                    str_CellValue = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(c.CellValue.Text)].InnerText;
                                }
                                if (str_CellValue.ToUpper() == "DATE")
                                {
                                    bol_RecordStarted = true;
                                    break;
                                }
                                break;
                            }
                            if (str_Ready == "NOT_READY")
                            {
                                str_Ready = "READY";
                                break;
                            }
                            if (c.CellValue == null && bol_RecordStarted == true)
                            {
                                if (int_ColNumber == 0) break;
                                if (int_ColNumber <= 6)
                                {
                                    int_ColNumber += 1;
                                    continue;
                                }
                            }


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
                            if(int_ColNumber>6)
                            {
                                if(str_CellValue.Trim().Length==0)
                                {
                                    str_CellValue = "0";
                                }
                            }    
                            dr_Tally_Row[int_ColNumber] = str_CellValue;
                            int_ColNumber += 1;
                            bol_InsertRow = true;
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    dr_Tally_Row["TallyExcelRowNumber"] = r.RowIndex.Value.ToString();
                    if (bol_InsertRow)
                    {
                        dt_Tally.Rows.Add(dr_Tally_Row);
                    }


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
            dg_GST.Columns[1].Visibility = Visibility.Visible;
            dg_Tally.Columns[1].Visibility = Visibility.Visible;
            dg_GST.Columns[2].Visibility = Visibility.Visible;
            dg_Tally.Columns[2].Visibility = Visibility.Visible;
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
                        dr_Tally_data[0]["PercentageMatching"] = "100";
                        dt_Tally.AcceptChanges();
                        dt_GST.Rows[gstLoop]["TallyExcelRowNumber"] = dr_Tally_data[0]["TallyExcelRowNumber"].ToString();
                        dt_GST.Rows[gstLoop]["PercentageMatching"] = "100";
                        dt_GST.AcceptChanges();
                        DataRow dr_matching = dt_matching_Det.NewRow();
                        dr_matching["GSTExcelRowNumber"] = dt_GST.Rows[gstLoop]["GSTExcelRowNumber"].ToString();
                        dr_matching["TallyExcelRowNumber"] = dr_Tally_data[0]["TallyExcelRowNumber"].ToString();
                        dr_matching["PercentageMatching"] = "100";
                        dr_matching["Remarks"] = "Direct Matching";
                        dr_matching["Invoice_number_Per"] = "100";
                        dr_matching["Invoice_Date_Per"] = "100";
                        dr_matching["Taxable_Value_Per"] = "100";
                        dr_matching["Integrated_Tax_Per"] = "100";
                        dr_matching["Central_Tax_Per"] = "100";
                        dr_matching["State_Tax_Per"] = "100";
                        dt_matching_Det.Rows.Add(dr_matching);
                        continue;
                    }

                    dr_Tally_data = (from tallyRow in dt_Tally.AsEnumerable()
                                     where tallyRow.Field<string>("GSTIN") == dt_GST.Rows[gstLoop]["GSTIN_of_supplier"].ToString()
                                     && tallyRow.Field<string>("PercentageMatching") != "100"
                                     select tallyRow).ToList();

                    for (int iloop = 0; iloop < dr_Tally_data.Count(); iloop++)
                    {
                        int int_total_per = 0;
                        str_rowRemarks = "";
                        string str_GST = dt_GST.Rows[gstLoop]["Invoice_number"].ToString().Trim();
                        string str_Tally = dr_Tally_data[iloop]["Invoice_No"].ToString().Trim();
                        int int_INV_Number_Percentage = fn_GST_Tally_Data_Maching(str_GST, str_Tally, "Invoice Number");
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

                        if (int_average >= 60)
                        {
                            dr_Tally_data[iloop]["GSTExcelRowNumber"] = dt_GST.Rows[gstLoop]["GSTExcelRowNumber"].ToString();
                            dr_Tally_data[iloop]["PercentageMatching"] = int_average.ToString();
                            dr_Tally_data[iloop]["Remarks"] = str_rowRemarks;
                            dt_Tally.AcceptChanges();

                            dt_GST.Rows[gstLoop]["TallyExcelRowNumber"] = dr_Tally_data[iloop]["TallyExcelRowNumber"].ToString();
                            dt_GST.Rows[gstLoop]["PercentageMatching"] = int_average.ToString();
                            dt_GST.Rows[gstLoop]["Remarks"] = str_rowRemarks;
                            dt_GST.AcceptChanges();

                            DataRow dr_matching = dt_matching_Det.NewRow();
                            dr_matching["GSTExcelRowNumber"] = dt_GST.Rows[gstLoop]["GSTExcelRowNumber"].ToString();
                            dr_matching["TallyExcelRowNumber"] = dr_Tally_data[iloop]["TallyExcelRowNumber"].ToString();
                            dr_matching["PercentageMatching"] = int_average.ToString();
                            dr_matching["Remarks"] = str_rowRemarks;
                            dr_matching["Invoice_number_Per"] = int_INV_Number_Percentage.ToString();
                            dr_matching["Invoice_Date_Per"] = int_Inv_Date_Percentage.ToString();
                            dr_matching["Taxable_Value_Per"] = int_TaxValue_Percentage.ToString();
                            dr_matching["Integrated_Tax_Per"] = "0";
                            dr_matching["Central_Tax_Per"] = int_CentralTax_Percentage.ToString();
                            dr_matching["State_Tax_Per"] = int_StateTax_Percentage.ToString();
                            dt_matching_Det.Rows.Add(dr_matching);
                        }
                        if (int_average == 100)
                        {
                            break;
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

        public int fn_GST_Tally_Data_Maching(string str_GST, string str_Tally, string DataFor)
        {
            int int_PercentageMaching = 0;
            if (str_GST == str_Tally)
            {
                int_PercentageMaching = 100;
                str_rowRemarks += DataFor + "  Direct Maching  " + int_PercentageMaching.ToString() + "%" + System.Environment.NewLine;
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

        private void btn_AML_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataRowView dataRowView = (DataRowView)((Button)e.Source).DataContext;
                DataTable dt_matching_Rows = new DataTable();
                string str_selectedRow_Number = "";
                if (((Button)sender).Tag.ToString() == "GST")
                {
                    int_Selected = 1;
                    str_selectedRow_Number = dataRowView["GSTExcelRowNumber"].ToString();
                    dt_matching_Rows = (from matchingRows in dt_matching_Det.AsEnumerable()
                                        where matchingRows.Field<string>("GSTExcelRowNumber") == str_selectedRow_Number
                                        select matchingRows).CopyToDataTable();
                    fn_Fill_Details(str_selectedRow_Number);
                }
                else
                {
                    int_Selected = 2;
                    dt_matching_Rows = (from matchingRows in dt_matching_Det.AsEnumerable()
                                        where matchingRows.Field<string>("TallyExcelRowNumber") == dataRowView["TallyExcelRowNumber"].ToString()
                                        select matchingRows).CopyToDataTable();
                }
                dg_matching.ItemsSource = dt_matching_Rows.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
        public void fn_Fill_Details(string str_sel_Row_number)
        {

            if (int_Selected == 1)
            {
                var dr_GST_Row = (from GST_Row in dt_GST.AsEnumerable()
                                  where GST_Row.Field<Int32>("GSTExcelRowNumber") == int.Parse(str_sel_Row_number)
                                  select GST_Row).ToList();

                txt_GST_INV_Number.Text = dr_GST_Row[0]["Invoice_number"].ToString();
                txt_GST_INV_Date.Text = dr_GST_Row[0]["Invoice_Date"].ToString();
                txt_GST_TaxableValue.Text = dr_GST_Row[0]["Taxable_Value"].ToString();
                txt_GST_IntegratedTax.Text = dr_GST_Row[0]["Integrated_Tax"].ToString();
                txt_GST_CentralTax.Text = dr_GST_Row[0]["Central_Tax"].ToString();
                txt_GST_StateTax.Text = dr_GST_Row[0]["State_Tax"].ToString();

            }
        }
        private void dg_matching_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dg_matching.SelectedValue == null)
            {
                return;
            }

            if (int_Selected == 1)
            {

                string str_select_tally_row_Number = (((System.Data.DataRowView)(dg_matching.SelectedValue)).Row)[1].ToString();
                var dr_Tally_Row = (from Tally_Row in dt_Tally.AsEnumerable()
                                    where Tally_Row.Field<Int32>("TallyExcelRowNumber") == int.Parse(str_select_tally_row_Number)
                                    select Tally_Row).ToList();

                txt_Tally_INV_Number.Text = dr_Tally_Row[0]["Invoice_No"].ToString();
                txt_Tally_INV_Date.Text = dr_Tally_Row[0]["Invoice_Date"].ToString();
                txt_Tally_TaxableValue.Text = dr_Tally_Row[0]["Taxable_Value"].ToString();
                txt_Tally_IntegratedTax.Text = dr_Tally_Row[0]["Integrated_Tax_Amount"].ToString();
                txt_Tally_CentralTax.Text = dr_Tally_Row[0]["Central_Tax_Amount"].ToString();
                txt_Tally_StateTax.Text = dr_Tally_Row[0]["State_Tax_Amount"].ToString();
            }
            fn_MatchingPercentage();
        }
        public void fn_MatchingPercentage()
        {
            str_rowRemarks = "";
            pb_inv_number.Value = fn_GST_Tally_Data_Maching(txt_GST_INV_Number.Text, txt_Tally_INV_Number.Text, "InvoiceNumber");
            pb_inv_Date.Value = fn_GST_Tally_Data_Maching(txt_GST_INV_Date.Text, txt_Tally_INV_Date.Text, "InvoiceDate");
            pb_TaxableValue.Value = fn_GST_Tally_Data_Maching(txt_GST_TaxableValue.Text, txt_Tally_TaxableValue.Text, "TaxableValue");
            pb_IntegratedTax.Value = fn_GST_Tally_Data_Maching(txt_GST_IntegratedTax.Text, txt_Tally_IntegratedTax.Text, "IntegratedTax");
            pb_CentralTax.Value = fn_GST_Tally_Data_Maching(txt_GST_CentralTax.Text, txt_Tally_CentralTax.Text, "CentralTax");
            pb_StateTax.Value = fn_GST_Tally_Data_Maching(txt_GST_StateTax.Text, txt_Tally_StateTax.Text, "StateTax");


        }
    }

}

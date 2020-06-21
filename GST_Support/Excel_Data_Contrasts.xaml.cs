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

namespace GST_Support
{
    /// <summary>
    /// Interaction logic for Excel_Data_Contrasts.xaml
    /// </summary>
    public partial class Excel_Data_Contrasts : Window
    {
        public Excel_Data_Contrasts()
        {
            InitializeComponent();
        }

        private void btn_Validate_Click(object sender, RoutedEventArgs e)
        {
            string str_GST_FilePath = @"F:\GST\Document\Book3.xlsx";
            string str_DestinationFilePath = @"F:\GST\DocumentBook3.xlsx";
            // Open the document for editing.
            using (SpreadsheetDocument spreadsheetDocument_GST =
                SpreadsheetDocument.Open(str_GST_FilePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument_GST.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                SharedStringTablePart stringTablePart = spreadsheetDocument_GST.WorkbookPart.SharedStringTablePart;
                string text;
                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        try
                        {
                            text = "";
                            text = (c.CellValue == null) ? c.InnerText : c.CellValue.Text;
                            if (c.DataType != null && c.DataType.Value == CellValues.SharedString)
                            {
                                text = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(c.CellValue.Text)].InnerText;
                            }
                            else
                            {
                                var cellText = (text ?? string.Empty).Trim();
                                if (c.StyleIndex != null)
                                {
                                    var cellFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[
                                        int.Parse(c.StyleIndex.InnerText)] as CellFormat;

                                    if (cellFormat != null)
                                    {
                                        var dateFormat = GetDateTimeFormat(cellFormat.NumberFormatId);
                                        if (!string.IsNullOrEmpty(dateFormat))
                                        {
                                            if (!string.IsNullOrEmpty(cellText))
                                            {
                                                if (double.TryParse(cellText, out var cellDouble))
                                                {
                                                    var theDate = DateTime.FromOADate(cellDouble);
                                                    text = theDate.ToString(dateFormat);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            txt_test.Text = txt_test.Text + text + "  ";
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    txt_test.Text = txt_test.Text + Environment.NewLine;
                }
            }
        }
        private string GetDateTimeFormat(UInt32Value numberFormatId)
        {
            return DateFormatDictionary.ContainsKey(numberFormatId) ? DateFormatDictionary[numberFormatId] : string.Empty;
        }

        //// https://msdn.microsoft.com/en-GB/library/documentformat.openxml.spreadsheet.numberingformat(v=office.14).aspx
        private readonly Dictionary<uint, string> DateFormatDictionary = new Dictionary<uint, string>()
        {
            [14] = "dd/MM/yyyy",
            [15] = "d-MMM-yy",
            [16] = "d-MMM",
            [17] = "MMM-yy",
            [18] = "h:mm AM/PM",
            [19] = "h:mm:ss AM/PM",
            [20] = "h:mm",
            [21] = "h:mm:ss",
            [22] = "M/d/yy h:mm",
            [30] = "M/d/yy",
            [34] = "yyyy-MM-dd",
            [45] = "mm:ss",
            [46] = "[h]:mm:ss",
            [47] = "mmss.0",
            [51] = "MM-dd",
            [52] = "yyyy-MM-dd",
            [53] = "yyyy-MM-dd",
            [55] = "yyyy-MM-dd",
            [56] = "yyyy-MM-dd",
            [58] = "MM-dd",
            [164] = "dd-MM-yyyy",
            [165] = "M/d/yy",
            [166] = "dd MMMM yyyy",
            [167] = "dd/MM/yyyy",
            [168] = "dd/MM/yy",
            [169] = "d.M.yy",
            [170] = "yyyy-MM-dd",
            [171] = "dd MMMM yyyy",
            [172] = "d MMMM yyyy",
            [173] = "M/d",
            [174] = "M/d/yy",
            [175] = "MM/dd/yy",
            [176] = "d-MMM",
            [177] = "d-MMM-yy",
            [178] = "dd-MMM-yy",
            [179] = "MMM-yy",
            [180] = "MMMM-yy",
            [181] = "MMMM d, yyyy",
            [182] = "M/d/yy hh:mm t",
            [183] = "M/d/y HH:mm",
            [184] = "MMM",
            [185] = "MMM-dd",
            [186] = "M/d/yyyy",
            [187] = "d-MMM-yyyy"
        };
    }

}

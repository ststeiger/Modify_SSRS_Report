
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace Modify_SSRS_Report
{


    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();
        }


        private void btnOK_Click(object sender, EventArgs e)
        {
            this.dgvDisplayData.DataSource = ProcessWorkbook();   
        }


        private System.Data.DataTable ProcessWorkbook()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            byte[] ba = null;
            bool bDebug = true;

            string fn = @"D:\ExampleMultitabReport.xlsx";

            if (!System.IO.File.Exists(fn) || !bDebug)
            {
                // COR_Reports.ReportFormatInfo pdfFormatInfo = new COR_Reports.ReportFormatInfo(COR_Reports.ExportFormat.PDF);
                // COR_Reports.ReportFormatInfo htmlFormatInfo = new COR_Reports.ReportFormatInfo(COR_Reports.ExportFormat.HtmlFragment);
                // COR_Reports.ReportFormatInfo excelFormatInfo = new COR_Reports.ReportFormatInfo(COR_Reports.ExportFormat.Excel);
                COR_Reports.ReportFormatInfo excelOpenXmlFormatInfo = new COR_Reports.ReportFormatInfo(COR_Reports.ExportFormat.ExcelOpenXml);
                ba = GetFooter("ExampleMultitabReport.rdl", excelOpenXmlFormatInfo);
            }
            else
                ba = System.IO.File.ReadAllBytes(fn);
            

            
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream(ba))
            {

                // Get the file we are going to process
                // Open and read the XlSX file.
                // System.IO.FileInfo existingFile = new System.IO.FileInfo(fn);
                // using (OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(existingFile))
                using (OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(ms))
                {
                    // Get the work book in the file
                    OfficeOpenXml.ExcelWorkbook workBook = package.Workbook;
                    if (workBook != null)
                    {
                        // dt = ProcessingFunction(workBook);

                        // OfficeOpenXml.ExcelWorksheet ws = workBook.Worksheets.First();
                        OfficeOpenXml.ExcelWorksheet roomsWorksheet = workBook.Worksheets["Räume"];
                        OfficeOpenXml.ExcelWorksheet UsageTypesWorksheet = workBook.Worksheets["Nutzungsarten"];

                        if (roomsWorksheet == null) 
                            return dt;

                        if (roomsWorksheet != null)
                            UsageTypesWorksheet.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;



                        int iStartRow = roomsWorksheet.Dimension.Start.Row;
                        int iEndRow = roomsWorksheet.Dimension.End.Row;

                        int iStartColumn = roomsWorksheet.Dimension.Start.Column;
                        int iEndColumn = roomsWorksheet.Dimension.End.Column;



                        // OfficeOpenXml.ExcelRange cell = roomsWorksheet.Cells["A1"];
                        OfficeOpenXml.ExcelRange cell = null; // ewbGroupPermissionWorksheet.Cells[1, 1]; // Cells[y, x]
                        for (int j = 1; j <= iEndColumn; ++j)
                        {
                            cell = roomsWorksheet.Cells[3, j]; // Cells[y, x]
                            string title = GetCellValueAsString(cell);
                            
                            // if (string.IsNullOrWhiteSpace(title)) continue;

                            dt.Columns.Add(title, typeof(string));
                        }

                        System.Console.WriteLine(dt.Columns.Count);


                        int ord = dt.Columns["NA Lang DE"].Ordinal + 1;


                        // https://stackoverflow.com/questions/29764226/add-list-validation-to-column-except-the-first-two-rows
                        // public static string GetAddress(int FromRow, int FromColumn, int ToRow, int ToColumn)
                        string range = OfficeOpenXml.ExcelRange.GetAddress(4, ord, OfficeOpenXml.ExcelPackage.MaxRows, ord);
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = roomsWorksheet.DataValidations.AddListValidation(range);

                        string char1 = "A";
                        string char2 = "A";
                        int num1 = 1;
                        int num2 = 300;

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";


                        
                        for (int i = 4; i <= iEndRow; ++i)
                        {

                            for (int j = 1; j <= iEndColumn; ++j)
                            {
                                cell = roomsWorksheet.Cells[3, j]; // Cells[y, x]

                                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.HotPink);

                                // cell.Style.Font.Strike;

                                OfficeOpenXml.Style.ExcelRichText ert = roomsWorksheet.Cells[2, 1].RichText.Add("RichText");
                                ert.Bold = true;
                                ert.Color = System.Drawing.Color.Red;
                                ert.Italic = true;
                                ert.Size = 12;


                                ert = roomsWorksheet.Cells[2, 1].RichText.Add("Test");
                                ert.Bold = true;
                                ert.Color = System.Drawing.Color.Purple;
                                ert.Italic = true;
                                ert.UnderLine = true;
                                ert.Strike = true;

                                ert = roomsWorksheet.Cells[2, 1].RichText.Add("123");
                                ert.Color = System.Drawing.Color.Peru;
                                ert.Italic = false;
                                ert.Bold = false;


                                // Can't add validation when a validation already exists
                                // OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList cellValidation = cell.DataValidation.AddListDataValidation();
                                break;

                                // https://stackoverflow.com/questions/9859610/how-to-set-column-type-when-using-epplus
                                // https://stackoverflow.com/questions/22832423/excel-date-format-using-epplus

                            } // Next j 
                            break;
                        } // Next i 

                    } // End if (workBook != null)

                    using(System.IO.FileStream fs = new System.IO.FileStream(@"D:\ModifiedExcelFile.xlsx", System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None ))
                    {
                        package.SaveAs(fs);
                    } // End using fs
                    
                } // End Using package

            } // End Using ms 

            return dt;
        }


        public static string GetCellValueAsString(OfficeOpenXml.ExcelRange cell)
        {
            object objValue = null;

            if (cell == null)
                return null;

            if (cell != null)
                objValue = cell.Value;

            if (objValue == null)
                return null;

            return objValue.ToString().Trim();
        } // End Function GetCellValueAsString


        public static string GetConnectionString()
        {
            System.Data.SqlClient.SqlConnectionStringBuilder csb = new System.Data.SqlClient.SqlConnectionStringBuilder();
            csb.DataSource = System.Environment.MachineName;
            csb.InitialCatalog = "COR_Basic_BKB";
            csb.IntegratedSecurity = true;

            return csb.ConnectionString;
        }


        public static System.Data.DataTable GetDataTable(string strSQL)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            using (System.Data.SqlClient.SqlDataAdapter da = new System.Data.SqlClient.SqlDataAdapter(strSQL, GetConnectionString()))
            {
                da.Fill(dt);
            }

            return dt;
        }

        // Depends on TFS://COR-Library\COR_Reports\COR_Reports.csproj
        // COR_Legende.ReportFooterSTZH.GetFooterPDF
        //Public Shared Function GetFooter(report As String, formatInfo As COR_Reports.ReportFormatInfo, in_aperturedwg As String, in_stylizer As String, in_sprache As String, ByRef AdjustHeight As Boolean) As Byte()
        public static byte[] GetFooter(string report, COR_Reports.ReportFormatInfo formatInfo)
        {
            byte[] baReport = null;

            try
            {
                COR_Reports.ReportTools.ReportDataCallback_t cb = (COR_Reports.ReportViewer viewer, System.Xml.XmlDocument doc) =>
                {
                    //Dim lsParameters As New System.Collections.Generic.List(Of COR_Reports.ReportParameter)()



                    //lsParameters.Add(New COR_Reports.ReportParameter("in_aperturedwg", "G00020-OG02_0000"))
                    //lsParameters.Add(New COR_Reports.ReportParameter("in_stylizer", "REM Bodenbelag"))
                    //lsParameters.Add(New COR_Reports.ReportParameter("in_sprache", "DE"))

                    //If leg.IsBKB AndAlso False Then
                    //    lsParameters.Add(New COR_Reports.ReportParameter("proc", leg.UserInfo.BE_User))
                    //End If

                    //' lsParameters.Add(new COR_Reports.ReportParameter("datastart", "dateTimePickerStartRaport.Text"))
                    //' lsParameters.Add(new COR_Reports.ReportParameter("dataStop", "dateTimePickerStopRaport.Text"))

                    //viewer.SetParameters(lsParameters)
                    //lsParameters.Clear()
                    //lsParameters = Nothing


                    string sprache = System.Globalization.CultureInfo.CurrentCulture.TwoLetterISOLanguageName.ToUpperInvariant();

                    // Add data sources
                    {
                        // This refers to the dataset name in the RDL/RDLC file
                        COR_Reports.ReportDataSource rdsDATA_Raum = new COR_Reports.ReportDataSource();
                        rdsDATA_Raum.Name = "DATA_Raum";
                        string strSQL = COR_Reports.ReportTools.GetDataSetDefinition(doc, rdsDATA_Raum.Name);
                        strSQL = strSQL.Replace("@in_sprache", "'" + sprache + "'");
                        System.Data.DataTable dt = GetDataTable(strSQL);
                        rdsDATA_Raum.Value = dt;
                        viewer.DataSources.Add(rdsDATA_Raum);


                        // This refers to the dataset name in the RDL/RDLC file
                        COR_Reports.ReportDataSource rdsDATA_Nutzungsart = new COR_Reports.ReportDataSource();
                        rdsDATA_Nutzungsart.Name = "DATA_Nutzungsart";
                        strSQL = COR_Reports.ReportTools.GetDataSetDefinition(doc, rdsDATA_Nutzungsart.Name);
                        strSQL = strSQL.Replace("@in_sprache", "'" + sprache + "'");
                        System.Data.DataTable dt2 = GetDataTable(strSQL);
                        rdsDATA_Nutzungsart.Value = dt2;
                        viewer.DataSources.Add(rdsDATA_Nutzungsart);
                    }

                };

                baReport = COR_Reports.ReportTools.RenderReport(report, formatInfo, cb);

                using (System.IO.FileStream fs = System.IO.File.Create("D:\\" + System.IO.Path.GetFileNameWithoutExtension(report) + formatInfo.Extension))
                {
                    fs.Write(baReport, 0, baReport.Length);
                } // End Using fs

            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex);
            }

            return baReport;
        } // GetFooter


    }


}

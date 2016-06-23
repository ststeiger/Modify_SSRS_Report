using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Modify_SSRS_Report
{
    static class Program
    {
        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [STAThread]
        static void Main()
        {
            ProcessWorkbook();
            if (false)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }

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


        public delegate void ExcelWorksheetCallback_t(OfficeOpenXml.ExcelWorksheet sheet);

        public static void ForAllSheets(OfficeOpenXml.ExcelWorkbook workBook, ExcelWorksheetCallback_t sub)
        {
            foreach (OfficeOpenXml.ExcelWorksheet sheet in workBook.Worksheets)
            {
                //sheet.Name
                sub(sheet);
            }

        }



        public static void ColorizeTabs(OfficeOpenXml.ExcelWorksheet sheet)
        {

            if (sheet.Name.IndexOf("Entry") != -1)
                sheet.TabColor = System.Drawing.ColorTranslator.FromHtml("#92CDDC");
            else
                sheet.TabColor = System.Drawing.ColorTranslator.FromHtml("#FABF8F"); // Orange
        }


        public static void HideReferenceTabs(OfficeOpenXml.ExcelWorksheet sheet)
        {
            string[] whiteListedWorksheetNames = new string[] { 
                 "Contract Details Review"
                ,"Rent Details Review"
                ,"Options & Tasks Review"
                ,"Contract Details Entry"
                ,"Rent Details Entry"
                ,"Options & Tasks Entry"
            };


            bool bFound = false;
            for (int i = 0; i < whiteListedWorksheetNames.Length; ++i)
            {
                if (string.Equals(sheet.Name, whiteListedWorksheetNames[i], StringComparison.OrdinalIgnoreCase))
                {
                    bFound = true;
                    break;
                }

            }

            if (!bFound)
                sheet.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        }


        public static void ColorizeDatePicker(OfficeOpenXml.ExcelWorksheet sheet)
        {
            if (string.Equals(sheet.Name, "Contract Details Review", StringComparison.OrdinalIgnoreCase))
            {
                // OfficeOpenXml.ExcelRange cellSource = sheet.Cells["E7"];
                // System.Console.WriteLine(cellSource.Style.Fill.BackgroundColor);

                // 242 220 219: F2DCDB
                // Contract Details Review F7:F13
                OfficeOpenXml.ExcelRange cell = sheet.Cells["F7:F13"];
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 220, 219));
                return;
            }

            if (string.Equals(sheet.Name, "Contract Details Entry", StringComparison.OrdinalIgnoreCase))
            {
                // Contract Details Entry G8:G14
                OfficeOpenXml.ExcelRange cell = sheet.Cells["G8:G14"];
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 220, 219));
                return;
            }

            if (string.Equals(sheet.Name, "Rent Details Entry", StringComparison.OrdinalIgnoreCase))
            {
                // Rent Details Entry J5:J28
                OfficeOpenXml.ExcelRange cell = sheet.Cells["J5:J28"];
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 220, 219));
                return;
            }
            
        }


        public static void SetCellValidation(OfficeOpenXml.ExcelWorksheet sheet)
        {
            if (string.Equals(sheet.Name, "Contract Details Review", StringComparison.OrdinalIgnoreCase))
            {
                // Fix row Height
                var cl = sheet.Cells["E15"];
                sheet.Row(cl.Start.Row).Height = 25;


                // YesNo
                {
                    string YesNoRange = "=" + sheet.Workbook.Worksheets["T_YesNo"].Cells["A1:A2"].FullAddress;
                    string[] validateCells = new string[] { "E14", "C28", "C29"};

                    foreach (string thisCell in validateCells)
                    {
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation(thisCell);

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = YesNoRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }
                }
                
                // ContractStatus
                {
                    string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_ContractStatus"].Cells["A1:A7"].FullAddress; // Contract
                    OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("C16");

                    // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                    //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                    // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                    val.Formula.ExcelFormula = ContractRange;
                    val.ShowErrorMessage = true;
                    val.Error = "Select a value from list of values ...";
                }

                // ContractType
                {
                    string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_ContractType"].Cells["A1:A7"].FullAddress; // Contract
                    OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("C10");

                    // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                    //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                    // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                    val.Formula.ExcelFormula = ContractRange;
                    val.ShowErrorMessage = true;
                    val.Error = "Select a value from list of values ...";
                }


                return;
            }


            if (string.Equals(sheet.Name, "Contract Details Entry", StringComparison.OrdinalIgnoreCase))
            {
                var cl = sheet.Cells["F16"];
                sheet.Row(cl.Start.Row).Height = 25;
                return;
            }

        }

        public static void DotLineToHairLine(OfficeOpenXml.ExcelWorksheet sheet)
        {
            int iStartRow = sheet.Dimension.Start.Row;
            int iEndRow = sheet.Dimension.End.Row;

            int iStartColumn = sheet.Dimension.Start.Column;
            int iEndColumn = sheet.Dimension.End.Column;

            OfficeOpenXml.ExcelRange cell = null; // sheet.Cells[y, x]
            // OfficeOpenXml.ExcelRange cell = roomsWorksheet.Cells["A1"];

            for (int i = 1; i <= iEndRow; ++i)
            {

                for (int j = 1; j <= iEndColumn; ++j)
                {
                    cell = sheet.Cells[i, j]; // Cells[y, x]

                    if (cell.Style.Border.Top.Style == OfficeOpenXml.Style.ExcelBorderStyle.Dotted)
                        cell.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Hair;

                    if (cell.Style.Border.Left.Style == OfficeOpenXml.Style.ExcelBorderStyle.Dotted)
                        cell.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Hair;

                    if (cell.Style.Border.Right.Style == OfficeOpenXml.Style.ExcelBorderStyle.Dotted)
                        cell.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Hair;

                    if (cell.Style.Border.Bottom.Style == OfficeOpenXml.Style.ExcelBorderStyle.Dotted)
                        cell.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Hair;

                    // cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    // cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.HotPink);

                } // Next j 

            } // Next i 

        }



        private static void ProcessWorkbook()
        {
            // System.Data.DataTable dt = new System.Data.DataTable();

            bool bDebug = true;

            string fn = @"D:\username\Downloads\LeaseContractForm.xlsx";
            byte[] ba = System.IO.File.ReadAllBytes(fn);


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

                        ForAllSheets(workBook, ColorizeTabs);
                        ForAllSheets(workBook, HideReferenceTabs);
                        ForAllSheets(workBook, ColorizeDatePicker);
                        ForAllSheets(workBook, DotLineToHairLine);
                        ForAllSheets(workBook, SetCellValidation);
                        
                        


                        // OfficeOpenXml.ExcelWorksheet roomsWorksheet = workBook.Worksheets["Contract Details Review"];
                        // if (roomsWorksheet == null) return;

                        using (System.IO.FileStream fs = new System.IO.FileStream(@"D:\ModifiedExcelFile.xlsx", System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None))
                        {
                            package.SaveAs(fs);
                        } // End using fs

                    } // End Using package

                } // End Using package 

            } // End Using ms 

        } // End Sub ProcessWorkbook 


    } // End Class Program


} // End Namespace Modify_SSRS_Report 

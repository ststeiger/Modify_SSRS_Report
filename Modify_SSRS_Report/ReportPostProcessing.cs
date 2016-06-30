
namespace Portal_Reports
{


    static class LeaseContractFormPostProcessing
    {


        public delegate void ExcelWorksheetCallback_t(OfficeOpenXml.ExcelWorksheet sheet);


        public static void ForAllSheets(OfficeOpenXml.ExcelWorkbook workBook, ExcelWorksheetCallback_t perWorksheetCallback)
        {

            foreach (OfficeOpenXml.ExcelWorksheet sheet in workBook.Worksheets)
            {
                perWorksheetCallback(sheet);
            } // Next sheet 

        } // End Sub ForAllSheets 


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


        public static int FindRowNumber(OfficeOpenXml.ExcelWorksheet sheet, int startRow, int iEndRow, int iColumn, string searchTerm)
        {
            int iRowNumber = -1;

            for (int i = startRow; i <= iEndRow; ++i)
            {
                OfficeOpenXml.ExcelRange cl = sheet.Cells[i, iColumn];
                string celVal = GetCellValueAsString(cl);

                if (string.Equals(celVal, searchTerm, System.StringComparison.OrdinalIgnoreCase))
                {
                    iRowNumber = i;
                    break;
                }

            } // Next i 

            return iRowNumber;
        }


        public static void AddRowAtPos2(OfficeOpenXml.ExcelWorksheet sheet)
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

                if (string.Equals(sheet.Name, whiteListedWorksheetNames[i], System.StringComparison.OrdinalIgnoreCase))
                {
                    bFound = true;
                    break;
                } // End if (string.Equals(sheet.Name, whiteListedWorksheetNames[i], StringComparison.OrdinalIgnoreCase)) 

            } // Next i

            if (bFound)
            {
                sheet.InsertRow(2, 1);
            }
        } // End Sub AddRowAtPos2 


        public static void ColorizeTabs(OfficeOpenXml.ExcelWorksheet sheet)
        {
            if (sheet.Name.IndexOf("Entry") != -1)
                sheet.TabColor = System.Drawing.ColorTranslator.FromHtml("#FABF8F"); // Orange
            else
                sheet.TabColor = System.Drawing.ColorTranslator.FromHtml("#92CDDC"); // SkyBlue
        } // End Sub ColorizeTabs 


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

                if (string.Equals(sheet.Name, whiteListedWorksheetNames[i], System.StringComparison.OrdinalIgnoreCase))
                {
                    bFound = true;
                    break;
                } // End if (string.Equals(sheet.Name, whiteListedWorksheetNames[i], StringComparison.OrdinalIgnoreCase)) 

            } // Next i

            if (!bFound)
                sheet.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
        } // End Sub HideReferenceTabs 



        public static void SetEntryLink(OfficeOpenXml.ExcelWorksheet sheet)
        {
            if (string.Equals(sheet.Name, "Rent Details Entry", System.StringComparison.OrdinalIgnoreCase))
            {
                OfficeOpenXml.ExcelRange cellLocation = sheet.Cells["D2"];
                OfficeOpenXml.ExcelRange cellPremise = sheet.Cells["D3"];

                cellLocation.Formula = "='Contract Details Entry'!D3";
                cellPremise.Formula = "='Contract Details Entry'!D4";

                return;
            }

            if (string.Equals(sheet.Name, "Options & Tasks Entry", System.StringComparison.OrdinalIgnoreCase))
            {
                OfficeOpenXml.ExcelRange cellLocation = sheet.Cells["F2"];
                OfficeOpenXml.ExcelRange cellPremise = sheet.Cells["F3"];

                cellLocation.Formula = "='Contract Details Entry'!D3";
                cellPremise.Formula = "='Contract Details Entry'!D4";

                return;
            }
        }

        public static void ColorizeDatePicker(OfficeOpenXml.ExcelWorksheet sheet)
        {
            if (string.Equals(sheet.Name, "Contract Details Review", System.StringComparison.OrdinalIgnoreCase))
            {
                // OfficeOpenXml.ExcelRange cellSource = sheet.Cells["E7"];
                // System.Console.WriteLine(cellSource.Style.Fill.BackgroundColor);

                // 242 220 219: F2DCDB
                // Contract Details Review F7:F13
                OfficeOpenXml.ExcelRange cell = sheet.Cells["F7:F13"];
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 220, 219));
                return;
            } // End if (string.Equals(sheet.Name, "Contract Details Review", StringComparison.OrdinalIgnoreCase)) 

            if (string.Equals(sheet.Name, "Contract Details Entry", System.StringComparison.OrdinalIgnoreCase))
            {
                // Contract Details Entry G8:G14
                OfficeOpenXml.ExcelRange cell = sheet.Cells["G8:G14"];
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 220, 219));
                return;
            } // End if (string.Equals(sheet.Name, "Contract Details Entry", StringComparison.OrdinalIgnoreCase)) 

            if (string.Equals(sheet.Name, "Rent Details Entry", System.StringComparison.OrdinalIgnoreCase))
            {
                // Rent Details Entry J5:J28
                OfficeOpenXml.ExcelRange cell = sheet.Cells["J11:J34"];
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 220, 219));
                return;
            } // End if (string.Equals(sheet.Name, "Rent Details Entry", StringComparison.OrdinalIgnoreCase)) 

        } // End Sub ColorizeDatePicker 


        public static void SetCellValidation(OfficeOpenXml.ExcelWorksheet sheet)
        {
            if (sheet.Dimension != null)
            {

                int iStartRow = sheet.Dimension.Start.Row;
                int iEndRow = sheet.Dimension.End.Row;

                int iStartColumn = sheet.Dimension.Start.Column;
                int iEndColumn = sheet.Dimension.End.Column;



                if (string.Equals(sheet.Name, "Rent Details Entry", System.StringComparison.OrdinalIgnoreCase))
                {

                    // Local Currency
                    {
                        string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_Currency"].Cells["A:A"].FullAddress; // Contract
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("C6");


                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = ContractRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }


                    // VAT Payable
                    {
                        string YesNoRange = "=" + sheet.Workbook.Worksheets["T_YesNo"].Cells["A:A"].FullAddress;

                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("F6");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = YesNoRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }

                    // RentDue
                    {
                        string CycleRange = "=" + sheet.Workbook.Worksheets["T_Ref_Cycle"].Cells["A:A"].FullAddress;

                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("C7");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = CycleRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }

                    // RentDueModality
                    {
                        string TimePointRange = "=" + sheet.Workbook.Worksheets["T_Ref_TimePoint"].Cells["A:A"].FullAddress;

                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("D7");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        val.Formula.ExcelFormula = TimePointRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }

                    // KindOfArea
                    {
                        string kindOfAreaRange = "=" + sheet.Workbook.Worksheets["T_Ref_KindOfArea"].Cells["A:A"].FullAddress;
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("A11:A34");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        val.Formula.ExcelFormula = kindOfAreaRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }

                    // Quantity
                    {
                        string kindOfAreaRange = "=" + sheet.Workbook.Worksheets["T_Ref_Unit"].Cells["A:A"].FullAddress;
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("D11:D34");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        val.Formula.ExcelFormula = kindOfAreaRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }

                } // End if (string.Equals(sheet.Name, "Rent Details Entry", System.StringComparison.OrdinalIgnoreCase))



                if (string.Equals(sheet.Name, "Contract Details Entry", System.StringComparison.OrdinalIgnoreCase))
                {
                    OfficeOpenXml.ExcelRange cl = sheet.Cells["F16"];
                    sheet.Row(cl.Start.Row).Height = 25;


                    // T_Ref_Location
                    {
                        string CreLoRange = "=" + sheet.Workbook.Worksheets["T_Ref_Location"].Cells["A:A"].FullAddress; // Contract
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("D3");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = CreLoRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }

                    // T_Premises
                    {
                        // string CreLoRange = "=" + sheet.Workbook.Worksheets["T_Premises"].Cells["A:A"].FullAddress; // Contract
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("D4");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        // val.Formula.ExcelFormula = CreLoRange; 
                        val.Formula.ExcelFormula = "=INDIRECT(VLOOKUP(D3,T_Ref_Premises_Location!B1:C65535,2,FALSE),TRUE)";

                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }


                    // YesNo
                    {
                        string YesNoRange = "=" + sheet.Workbook.Worksheets["T_YesNo"].Cells["A:A"].FullAddress;
                        string[] validateCells = new string[] { "D33", "D34", "F15" };


                        foreach (string thisCell in validateCells)
                        {
                            OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation(thisCell);

                            // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                            //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                            // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                            val.Formula.ExcelFormula = YesNoRange;
                            val.ShowErrorMessage = true;
                            val.Error = "Select a value from list of values ...";
                        } // Next thisCell 
                    }



                    // T_Ref_CRELO
                    {
                        string CreLoRange = "=" + sheet.Workbook.Worksheets["T_Ref_CRELO"].Cells["A:A"].FullAddress; // Contract
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("D8");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = CreLoRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }


                    // ContractStatus
                    {
                        string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_ContractStatus"].Cells["A:A"].FullAddress; // Contract
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("D17");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = ContractRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }

                    // ContractType
                    {
                        string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_ContractType"].Cells["A:A"].FullAddress; // Contract
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("D11");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = ContractRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }

                    return;
                } // End if (string.Equals(sheet.Name, "Contract Details Entry", StringComparison.OrdinalIgnoreCase)) 


                if (string.Equals(sheet.Name, "Contract Details Review", System.StringComparison.OrdinalIgnoreCase))
                {
                    // Fix row Height
                    OfficeOpenXml.ExcelRange cl = sheet.Cells["E15"];
                    sheet.Row(cl.Start.Row).Height = 25;



                    // T_Ref_Location
                    {
                        string CreLoRange = "=" + sheet.Workbook.Worksheets["T_Ref_Location"].Cells["A:A"].FullAddress; // Contract
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("C2");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = CreLoRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }

                    // T_Premises
                    {
                        // string CreLoRange = "=" + sheet.Workbook.Worksheets["T_Premises"].Cells["A:A"].FullAddress; // Contract
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("C3");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        // val.Formula.ExcelFormula = CreLoRange; 
                        val.Formula.ExcelFormula = "=INDIRECT(VLOOKUP(C2,T_Ref_Premises_Location!B1:C65535,2,FALSE),TRUE)";

                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }



                    // YesNo
                    {
                        string YesNoRange = "=" + sheet.Workbook.Worksheets["T_YesNo"].Cells["A:A"].FullAddress;
                        int iEstoppelRow = FindRowNumber(sheet, 19, iEndRow, 1, "Estoppel/SNDA");
                        int iBoilerplateClauseRow = FindRowNumber(sheet, 19, iEndRow, 1, "Boilerplate Clauses");

                        //string[] validateCells = new string[] { "E14", "C28", "C29" };
                        string[] validateCells = new string[] { "E14", "C" + iEstoppelRow.ToString(), "C" + iBoilerplateClauseRow.ToString() };


                        foreach (string thisCell in validateCells)
                        {
                            OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation(thisCell);

                            // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                            //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                            // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                            val.Formula.ExcelFormula = YesNoRange;
                            val.ShowErrorMessage = true;
                            val.Error = "Select a value from list of values ...";
                        } // Next thisCell 
                    }


                    // T_Ref_CRELO
                    {
                        string CreLoRange = "=" + sheet.Workbook.Worksheets["T_Ref_CRELO"].Cells["A:A"].FullAddress; // Contract
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("C7");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = CreLoRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }


                    // ContractStatus
                    {
                        string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_ContractStatus"].Cells["A:A"].FullAddress; // Contract
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
                        string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_ContractType"].Cells["A:A"].FullAddress; // Contract
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("C10");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = ContractRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }


                    // Role
                    {
                        int iSecondaryTermsRow = FindRowNumber(sheet, 19, iEndRow, 1, "Secondary terms and conditions");

                        string PartyRange = "=" + sheet.Workbook.Worksheets["T_Ref_Party"].Cells["A:A"].FullAddress; // Contract
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("E19:E" + iSecondaryTermsRow.ToString());

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = PartyRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }


                    return;
                } // End if (string.Equals(sheet.Name, "Contract Details Review", StringComparison.OrdinalIgnoreCase))



                if (string.Equals(sheet.Name, "Rent Details Review", System.StringComparison.OrdinalIgnoreCase))
                {

                    // Local Currency
                    {
                        string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_Currency"].Cells["A:A"].FullAddress; // Contract
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("C6");


                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = ContractRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }


                    // VAT Payable
                    {
                        string YesNoRange = "=" + sheet.Workbook.Worksheets["T_YesNo"].Cells["A:A"].FullAddress;

                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("F6");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = YesNoRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }

                    // RentDue
                    {
                        string CycleRange = "=" + sheet.Workbook.Worksheets["T_Ref_Cycle"].Cells["A:A"].FullAddress;

                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("C7");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                        // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                        val.Formula.ExcelFormula = CycleRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }

                    // RentDueModality
                    {
                        string TimePointRange = "=" + sheet.Workbook.Worksheets["T_Ref_TimePoint"].Cells["A:A"].FullAddress;

                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("D7");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        val.Formula.ExcelFormula = TimePointRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }

                    // KindOfArea
                    {
                        string kindOfAreaRange = "=" + sheet.Workbook.Worksheets["T_Ref_KindOfArea"].Cells["A:A"].FullAddress;
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("A12:A65536");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        val.Formula.ExcelFormula = kindOfAreaRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }


                    // Quantity
                    {
                        string kindOfAreaRange = "=" + sheet.Workbook.Worksheets["T_Ref_Unit"].Cells["A:A"].FullAddress;
                        OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("D12:D65536");

                        // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                        val.Formula.ExcelFormula = kindOfAreaRange;
                        val.ShowErrorMessage = true;
                        val.Error = "Select a value from list of values ...";
                    }

                } // End if (string.Equals(sheet.Name, "Rent Details Review", System.StringComparison.OrdinalIgnoreCase)) 


                if (string.Equals(sheet.Name, "Options & Tasks Review", System.StringComparison.OrdinalIgnoreCase))
                {

                    // ContractType
                    {
                        int iEventAndTasksRowNumber = FindRowNumber(sheet, 7, iEndRow, 1, "Events and tasks");
                        iEventAndTasksRowNumber--;

                        // Agreed rights and options foo
                        {
                            // Type
                            {
                                string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_TaskCategory"].Cells["A:A"].FullAddress; // Contract
                                OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("F8:F" + iEventAndTasksRowNumber.ToString());

                                // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                                //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                                // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                                val.Formula.ExcelFormula = ContractRange;
                                val.ShowErrorMessage = true;
                                val.Error = "Select a value from list of values ...";
                            }

                            // Status
                            {
                                string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_TaskStatus"].Cells["A:A"].FullAddress; // Contract
                                OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("H8:H" + iEventAndTasksRowNumber.ToString());

                                // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                                //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                                // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                                val.Formula.ExcelFormula = ContractRange;
                                val.ShowErrorMessage = true;
                                val.Error = "Select a value from list of values ...";
                            }
                        }


                        iEventAndTasksRowNumber++;
                        iEventAndTasksRowNumber++;

                        // Events and Tasks
                        {
                            // Type
                            {
                                string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_TaskCategory"].Cells["A:A"].FullAddress; // Contract
                                OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("B" + iEventAndTasksRowNumber.ToString() + ":B65536");

                                // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                                //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                                // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                                val.Formula.ExcelFormula = ContractRange;
                                val.ShowErrorMessage = true;
                                val.Error = "Select a value from list of values ...";
                            }


                            // Event/Activity
                            {
                                string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_TaskActivity"].Cells["A:A"].FullAddress; // Contract
                                OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("E" + iEventAndTasksRowNumber.ToString() + ":E65536");

                                // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                                //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                                // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                                val.Formula.ExcelFormula = ContractRange;
                                val.ShowErrorMessage = true;
                                val.Error = "Select a value from list of values ...";
                            }

                            // Status
                            {
                                iEventAndTasksRowNumber++;
                                string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_TaskStatus"].Cells["A:A"].FullAddress; // Contract
                                OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("S" + iEventAndTasksRowNumber.ToString() + ":S65536");

                                // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                                //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                                // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                                val.Formula.ExcelFormula = ContractRange;
                                val.ShowErrorMessage = true;
                                val.Error = "Select a value from list of values ...";
                            }
                        }
                    }

                    return;
                } // End if (string.Equals(sheet.Name, "Options & Tasks Review", System.StringComparison.OrdinalIgnoreCase)) 


                if (string.Equals(sheet.Name, "Options & Tasks Entry", System.StringComparison.OrdinalIgnoreCase))
                {
                    // Agreed rights and options
                    {
                        // Type
                        {
                            string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_TaskCategory"].Cells["A:A"].FullAddress; // Contract
                            OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("F8:F17");

                            // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                            //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                            // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                            val.Formula.ExcelFormula = ContractRange;
                            val.ShowErrorMessage = true;
                            val.Error = "Select a value from list of values ...";
                        }

                        // Status
                        {
                            string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_TaskStatus"].Cells["A:A"].FullAddress; // Contract
                            OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("H8:H17");

                            // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                            //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                            // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                            val.Formula.ExcelFormula = ContractRange;
                            val.ShowErrorMessage = true;
                            val.Error = "Select a value from list of values ...";
                        }
                    }

                    // Events and tasks
                    {
                        // Type
                        {
                            string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_TaskCategory"].Cells["A:A"].FullAddress; // Contract
                            OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("B23:B37");

                            // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                            //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                            // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                            val.Formula.ExcelFormula = ContractRange;
                            val.ShowErrorMessage = true;
                            val.Error = "Select a value from list of values ...";
                        }

                        // Event/Activity
                        {
                            string ContractRange = "=" + sheet.Workbook.Worksheets["T_Ref_TaskActivity"].Cells["A:A"].FullAddress; // Contract
                            OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("E23:E37");

                            // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                            //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                            // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                            val.Formula.ExcelFormula = ContractRange;
                            val.ShowErrorMessage = true;
                            val.Error = "Select a value from list of values ...";
                        }


                        // Status
                        {
                            string StatusRange = "=" + sheet.Workbook.Worksheets["T_Ref_TaskStatus"].Cells["A:A"].FullAddress; // Contract
                            OfficeOpenXml.DataValidation.Contracts.IExcelDataValidationList val = sheet.DataValidations.AddListValidation("S23:S37");

                            // https://stackoverflow.com/questions/20259692/epplus-number-of-drop-down-items-limitation-in-excel-file
                            //val.Formula.ExcelFormula = string.Format("=DropDownWorksheetName!${0}${1}:${2}${3}", char1, num1, char2, num2);
                            // val.Formula.ExcelFormula = string.Format("=Nutzungsarten!{0}", "H2:H72");
                            val.Formula.ExcelFormula = StatusRange;
                            val.ShowErrorMessage = true;
                            val.Error = "Select a value from list of values ...";
                        }
                    }
                    return;
                } // End if (string.Equals(sheet.Name, "Options & Tasks Review", System.StringComparison.OrdinalIgnoreCase)) 

            }
        } // End Sub SetCellValidation 


        public static void DotLineToHairLine(OfficeOpenXml.ExcelWorksheet sheet)
        {
            if (sheet.Dimension != null)
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

        } // End Sub DotLineToHairLine 



        public static byte[] ProcessWorkbook(byte[] ba)
        {
            byte[] baOutput = null;

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
                        ForAllSheets(workBook, SetEntryLink);
                        // ForAllSheets(workBook, AddRowAtPos2);

                        workBook.View.ActiveTab = 0;
                        workBook.Worksheets["Contract Details Review"].View.TabSelected = true;


                        // OfficeOpenXml.ExcelWorksheet roomsWorksheet = workBook.Worksheets["Contract Details Review"];
                        // if (roomsWorksheet == null) return;


                        using (System.IO.MemoryStream msOutput = new System.IO.MemoryStream())
                        {
                            package.SaveAs(msOutput);
                            baOutput = msOutput.ToArray();
                        } // End using fs

                    } // End Using package

                } // End Using package 

            } // End Using ms 

            return baOutput;
        } // End Sub ProcessWorkbook 


        public static byte[] ProcessWorkbook()
        {
            string fn = @"D:\username\Downloads\LeaseContractForm.xlsx";
            // fn = @"D:\stefan.steiger\Downloads\TestFile.xlsx";


            byte[] ba = System.IO.File.ReadAllBytes(fn);
            ba = ProcessWorkbook(ba);

            using (System.IO.FileStream fs = new System.IO.FileStream(@"D:\ModifiedExcelFile.xlsx", System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None))
            {
                fs.Write(ba, 0, ba.Length);
            }

            return ba;
        } // End Sub ProcessWorkbook 


    } // End Class LeaseContractFormPostProcessing


} // End Namespace Portal_Reports 

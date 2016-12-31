using System;
using System.Data;                  // Used to enable the DataTable variable type
using System.Diagnostics;           // Automatically open file with Excel
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;                    // File read/write
using Microsoft.VisualBasic.FileIO; // Text Field Parser
using OfficeOpenXml;                // Excel file handling
using OfficeOpenXml.Drawing;        // Excel file handling
using OfficeOpenXml.Style;          // Excel file handling

namespace MigrateData
{
    class Program
    {
        static void Main(string[] args)
        {
            //Command entry
            bool exit = false;
            ConsoleKeyInfo keypress; // ready the variable used for accepting user input

            Console.WriteLine("Instructions");
            Console.WriteLine("This program will take InvoiceItem-GP-Invoice-Integration-w-service-dates-Detail.csv and InvoiceItemAdjustment-GP-IIA-Integration-w-Reasoncode-Detail.csv in the local directory and write three new files in the local directory, reformatting and commenting the data.");
            Console.WriteLine("Press \"P\" to begin the process.");
            Console.WriteLine("");
            Console.WriteLine("You can press Esc to exit.");
            while (!exit) // This is the main loop and will end when the user presses the Esc key.
            {
                keypress = Console.ReadKey(true); //This is listed with true so the key press is not shown.
                if (keypress.Key == ConsoleKey.P)
                {
                    ProcessMigrationData();
                }
                if (keypress.Key == ConsoleKey.Escape) // This checks to see if the key pressed was Esc
                { exit = true; } // This variable change will end the main loop and close the program.
            }
        }

        static void ProcessMigrationData()
        {
            //string[,] IIAData = ReadFileToArray("GP IIA Integration w Reasoncode.csv", 19);
            DataTable dtInv = new DataTable(); // This is to hold the finalized invoice import data
            DataTable dtIIA = new DataTable(); // This is to hold the finalized IIA import data
            DataTable dtPay = new DataTable(); // This is to hold the finalized payment import data
            string textDate = string.Format("{0:MM-dd-yyyy}",DateTime.Now);
            Console.WriteLine(textDate);
            //DateTime localDate = DateTime.Today;
            ClearLogCSV("log.csv");
            using (ExcelPackage docp = new ExcelPackage()) // This is the ExcelPackage for the invoice data
            {
                ExcelWorksheet ZInvRaw = CreateSheet(docp, "Raw Invoice Data", 1);
                dtInv = ReadFileToArray("InvoiceItem-GP-Invoice-Integration-w-service-dates-Detail.csv", dtInv);
                ZInvRaw = CopyDataToWS(ZInvRaw, dtInv); // Copy the raw unmodified "GP Invoice Integration w service dates.csv" data to the ZInvRaw worksheet
                
                ExcelWorksheet ZInvAcc = CreateSheet(docp, "Z Invoice Account Transform",2);
                dtInv = CorrectAccNum(dtInv);
                ZInvAcc = CopyDataToWS(ZInvAcc, dtInv); // Copy the modified account data (and everything else to the ZInvAcc worksheet
                
                ExcelWorksheet ZInvQua = CreateSheet(docp, "Z Invoice Quantity Transform", 3);
                dtInv = SetQuantity(dtInv, 22); // This will set the column 22's values all to "1", except for the first row/header
                ZInvQua = CopyDataToWS(ZInvQua, dtInv); // Copy the modified quantity data
                
                ExcelWorksheet ZInvCou = CreateSheet(docp, "Z Invoice Country Transform", 4);
                dtInv = CorrectCountry(dtInv);
                ZInvQua = CopyDataToWS(ZInvCou, dtInv); // Copy the modified country data
                
                // This is the Start of the IIA worksheets
                dtIIA.TableName = "Raw IIA Data";
                ExcelWorksheet ZIIARaw = CreateSheet(docp, dtIIA.TableName.ToString(), 5);
                dtIIA = ReadFileToArray("InvoiceItemAdjustment-GP-IIA-Integration-w-Reasoncode-Detail.csv", dtIIA);
                ZIIARaw = CopyDataToWS(ZIIARaw, dtIIA); // Copy the raw unmodified "GP Invoice Integration w service dates.csv" data to the ZInvRaw worksheet

                dtIIA.TableName = "Z IIA Account Transform";
                ExcelWorksheet ZIIAAcc = CreateSheet(docp, dtIIA.TableName.ToString(), 6);
                dtIIA = CorrectAccNum(dtIIA);
                ZIIAAcc = CopyDataToWS(ZIIAAcc, dtIIA); // Copy the modified account data (and everything else to the ZIIAAcc worksheet

                dtIIA.TableName = "Z IIA Quantity Transform";
                ExcelWorksheet ZIIAQua = CreateSheet(docp, dtIIA.TableName.ToString(), 7);
                dtIIA = SetQuantity(dtIIA, 10); // This will set the values in column 10 to "1", except for the first row/header
                ZIIAQua = CopyDataToWS(ZIIAQua, dtIIA); // Copy the modified quantity data

                dtIIA.TableName = "Z IIA Charge Amount Transform";
                ExcelWorksheet ZIIACharge = CreateSheet(docp, dtIIA.TableName.ToString(), 8);
                dtIIA = AddChargeAmount(dtIIA);
                ZIIACharge = CopyDataToWS(ZIIACharge, dtIIA);
                
                // Compare IIA data with Invoice data to spot any subscription cleanup that isn't for the whole amount
                dtInv.TableName = "Z Invoice IIA Comparison";
                ExcelWorksheet ZInvCom = CreateSheet(docp, dtInv.TableName.ToString(), 9);
                dtInv = CompareToIIA(dtInv, dtIIA);
                dtInv = FlagInconsistent(dtInv);
                ZInvCom = CopyDataToWS(ZInvCom, dtInv); // Comparison with IIA, reason codes, highlights copied from IIA data to Invoice data
                ZInvCom = HighlightCells(ZInvCom, dtInv); // Highlights mismatched amount subscription cleanup, and other possible issues
                dtInv = ClearErrors(dtInv); // Clears the errors column

                // Remove Subscription Cleanup and invoices
                dtInv.TableName = "Z Invoice Remove Cleanup";
                ExcelWorksheet ZInvSub = CreateSheet(docp, dtInv.TableName.ToString(), 10);
                dtInv = RemoveEntries(dtInv, "Subscription Cleanup", 37);
                ZInvSub = CopyDataToWS(ZInvSub, dtInv);

                // Flag TBD IIA's
                dtIIA.TableName = "Z IIA Flag TBD";
                ExcelWorksheet ZIIATBD = CreateSheet(docp, dtIIA.TableName.ToString(), 11);
                dtIIA = FlagTBD(dtIIA, 9);
                ZIIATBD = CopyDataToWS(ZIIATBD, dtIIA);

                // Remove Subscription Cleanup and TBD IIA's
                dtIIA.TableName = "Z IIA Remove Cleanup and TBD";
                ExcelWorksheet ZIIASub = CreateSheet(docp, dtIIA.TableName.ToString(), 12);
                dtIIA = RemoveEntries(dtIIA, "Subscription Cleanup", 9);
                dtIIA = RemoveEntries(dtIIA, "TBD", 9);
                ZIIASub = CopyDataToWS(ZIIASub, dtIIA);

                // Remove tax detail IIA's
                dtIIA.TableName = "Z IIA Remove Tax";
                ExcelWorksheet ZIIATax = CreateSheet(docp, dtIIA.TableName.ToString(), 13);
                dtIIA = RemoveAllTaxEntries(dtIIA);
                ZIIATax = CopyDataToWS(ZIIATax, dtIIA);

                // Flag products listed with blank or N/A product codes (invoices)
                dtInv.TableName = "Z Invoice ProductCodes";
                ExcelWorksheet ZInvPro = CreateSheet(docp, dtInv.TableName.ToString(), 14);
                dtInv = FlagBlank(dtInv, 35);
                dtInv = FlagNA(dtInv, 35);
                dtInv = ReplaceCell(dtInv, 35, "ATB-ERP1-TXN", "ATB-ERP1-SVC");
                dtInv = ReplaceCell(dtInv, 35, "ATP-ERP1-TXN", "ATP-ERP1-SVC");
                dtInv = ReplaceCell(dtInv, 35, "ATP-GLOBAL-TXN", "ATP-ERP1-SVC");
                ZInvPro = CopyDataToWS(ZInvPro, dtInv);
                ZInvPro = HighlightCells(ZInvPro, dtInv); // Highlights any product code entries that were blank, or N/A
                dtInv = ClearErrors(dtInv); // Clears the errors column

                // Swap out the Net terms for GP friendly entries (invoices)
                dtInv.TableName = "Z Invoice DueUpon";
                ExcelWorksheet ZInvDue = CreateSheet(docp, dtInv.TableName.ToString(), 15);
                dtInv = ReplaceCellQuiet(dtInv, 28, "Due Upon Receipt", "OnReceipt");
                dtInv = ReplaceCellQuiet(dtInv, 29, "Due Upon Receipt", "OnReceipt");
                dtInv = ReplaceCellQuiet(dtInv, 28, "Net 7", "Net7");
                dtInv = ReplaceCellQuiet(dtInv, 29, "Net 7", "Net7");
                dtInv = ReplaceCellQuiet(dtInv, 28, "Net 30", "Net30");
                dtInv = ReplaceCellQuiet(dtInv, 29, "Net 30", "Net30");
                dtInv = ReplaceCellQuiet(dtInv, 28, "Net 45", "Net45");
                dtInv = ReplaceCellQuiet(dtInv, 29, "Net 45", "Net45");
                dtInv = ReplaceCellQuiet(dtInv, 28, "Net 60", "Net60");
                dtInv = ReplaceCellQuiet(dtInv, 29, "Net 60", "Net60");
                dtInv = ReplaceCellQuiet(dtInv, 28, "Net 90", "Net90");
                dtInv = ReplaceCellQuiet(dtInv, 29, "Net 90", "Net90");
                dtInv = ReplaceCellQuiet(dtInv, 28, "Net 120", "Net120");
                dtInv = ReplaceCellQuiet(dtInv, 29, "Net 120", "Net120");
                ZInvPro = CopyDataToWS(ZInvDue, dtInv);
                dtInv = ClearErrors(dtInv); // Clears the errors column

                // Replace header
                dtInv.TableName = "Z Invoice Header";
                ExcelWorksheet ZInvHead = CreateSheet(docp, dtInv.TableName.ToString(), 16);
                dtInv = ChangeHeader(dtInv);
                ZInvHead = CopyDataToWS(ZInvHead, dtInv);

                // Flag products listed with blank or N/A product codes (IIA's)
                dtIIA.TableName = "Z IIA ProductCodes";
                ExcelWorksheet ZIIAPro = CreateSheet(docp, dtIIA.TableName.ToString(), 17);
                dtIIA = FlagBlank(dtIIA, 14);
                dtIIA = FlagNA(dtIIA, 14);
                dtIIA = ReplaceCell(dtIIA, 14, "ATB-ERP1-TXN", "ATB-ERP1-SVC");
                dtIIA = ReplaceCell(dtIIA, 14, "ATP-ERP1-TXN", "ATP-ERP1-SVC");
                dtIIA = ReplaceCell(dtIIA, 14, "ATP-GLOBAL-TXN", "ATP-ERP1-SVC");
                ZIIAPro = CopyDataToWS(ZIIAPro, dtIIA);
                ZIIAPro = HighlightCells(ZIIAPro, dtIIA); // Highlights any product code entries that were blank, or N/A
                dtIIA = ClearErrors(dtIIA); // Clears the errors column

                // Write the file that includes all the above steps to disk
                WriteExcelFile(docp, textDate + "DocTW (WIP).xlsx");
            }
            using (ExcelPackage invp = new ExcelPackage())
            {
                dtInv.TableName = "Z Upload Final";
                ExcelWorksheet ZInvFinal = CreateSheet(invp, dtInv.TableName.ToString(), 1);
                ZInvFinal = CopyDataToWS(ZInvFinal, dtInv);
                WriteExcelFile(invp, textDate + "InvTW.xlsx");
            }
            using (ExcelPackage iiap = new ExcelPackage())
            {
                dtIIA.TableName = "Z IIA Upload Final";
                ExcelWorksheet ZIIAFinal = CreateSheet(iiap, dtIIA.TableName.ToString(), 1);
                ZIIAFinal = CopyDataToWS(ZIIAFinal, dtIIA);
                WriteExcelFile(iiap, textDate + "IIATW.xlsx");
            }
            /*using (ExcelPackage pay = new ExcelPackage()) // This is the ExcelPackage for the payment data
            {
                dtPay.TableName = "Raw Payment Data";
                ExcelWorksheet ZPayRaw = CreateSheet(pay, dtPay.TableName.ToString(), 1);
                dtPay = ReadFileToArray("GP PYMNT Integration Electronic - USD.csv", dtPay);
                ZPayRaw = CopyDataToWS(ZPayRaw, dtPay); // Copy the raw unmodified "GP Invoice Integration w service dates.csv" data to the ZInvRaw worksheet
                // Add the additional report data to the bottom of the data table
                dtPay.TableName = "Z Payment Account Transform";
                ExcelWorksheet ZPayAcc = CreateSheet(pay, dtPay.TableName.ToString(), 2);
                dtPay = CorrectAccNum(dtPay);
                ZPayAcc = CopyDataToWS(ZPayAcc, dtPay); // Copy the modified account data (and everything else to the ZInvAcc worksheet
                
                dtPay.TableName = "Z Payment Blank Check";
                ExcelWorksheet ZPayBlankCheck = CreateSheet(pay, dtPay.TableName.ToString(), 3);
                dtPay = FlagBlank(dtPay, 8); // Checks for blank payment type cells and notes the error in the log
                ZPayBlankCheck = CopyDataToWS(ZPayBlankCheck, dtPay);
                ZPayBlankCheck = HighlightCells(ZPayBlankCheck, dtPay); // Highlights blank cells
                dtPay = ClearErrors(dtPay); // Clears the errors column
                
                dtPay.TableName = "Z Payment Debit Check";
                ExcelWorksheet ZPayCheck = CreateSheet(pay, dtPay.TableName.ToString(), 4);
                dtPay = ReplaceCell(dtPay, 8, "DebitCard", "CreditCard"); // Check for no DebitCard entries and convert
                dtPay = CheckCreditType(dtPay);                 // Check for blank CreditCard types
                ZPayCheck = CopyDataToWS(ZPayCheck, dtPay);
                ZPayCheck = HighlightCells(ZPayCheck, dtPay);   // Highlights blank credit cards
                dtPay = ClearErrors(dtPay);                     // Clears the errors column
                WriteExcelFile(pay, textDate + "PymntLW (WIP).xlsx");
            }
            using (ExcelPackage payp = new ExcelPackage())
            {
                dtPay.TableName = "Z Pymnt Upload Final";
                ExcelWorksheet ZPayFinal = CreateSheet(payp, dtPay.TableName.ToString(), 1);
                ZPayFinal = CopyDataToWS(ZPayFinal, dtPay);
                WriteExcelFile(payp, textDate + "PymntLW.xlsx");
            }*/
            Console.WriteLine("This is the end of the process.");
        }

        static DataTable ReadFileToArray(string filename, DataTable dt)
        {
            int lineCount = File.ReadAllLines(filename).Length;
            Console.WriteLine("Processing " + lineCount + " lines of " + filename);
            TextFieldParser parser = new TextFieldParser(filename);
            parser.HasFieldsEnclosedInQuotes = true;
            //parser.SetDelimiters(","); This was used when the Zuora reports used commas in csv files. Now it is using tabs.
            parser.SetDelimiters("	"); // Note that the blank space inside the quotes is a tab
            string[] fields;
            int row = 0;
            int column = 0;
            while (!parser.EndOfData)
            {
                fields = parser.ReadFields();
                foreach (string field in fields)
                {
                    //Console.WriteLine(field); // This can be used to see data as it loads into the data table
                    column = column + 1;
                    if(dt.Columns.Count < column)
                    {
                        dt.Columns.Add(field); // This doesn't actually add the values, but it does create the columns in the data table
                    }
                }
                dt.Rows.Add(fields); // This adds a row worth of data at a time
                column = 0;
                row = row + 1;
            }
            parser.Close();
            Console.WriteLine("Data from " + filename + " loaded into table.");
            return dt;
        }

        private static ExcelWorksheet CopyDataToWS(ExcelWorksheet ws, DataTable dt)
        {
            int colIndex = 0;
            int rowIndex = 0;
            //Console.WriteLine("The first value in dt is: " + dt.Rows[0][0]); // This is for debugging
            foreach (DataRow dr in dt.Rows) // Adding Data into rows
            {
                colIndex = 1;
                rowIndex++;
                foreach (DataColumn dc in dt.Columns)
                {
                    var cell = ws.Cells[rowIndex, colIndex];

                    //Setting Value in cell
                    cell.Value = dr[dc.ColumnName]; // This looks like it could cause some formatting issues
                    colIndex++;
                }
            }
            Console.WriteLine("Worksheet " + ws.Name + " completed.");
            return ws;
        }

        private static DataTable ClearErrors(DataTable dt)
        {
            int lastColumn = Convert.ToInt32(dt.Columns.Count.ToString()) - 1;
            if(dt.Rows[0][lastColumn].ToString() == "Errors")
            {
                dt.Columns.RemoveAt(lastColumn);
            }
            return dt;
        }

        private static DataTable CorrectAccNum(DataTable dt)
        {
            //dt.Rows.RemoveAt(0); // This removes the first row, which only states the date range for the export out of Zuora
            int row = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (dt.Rows[row][0].ToString() != "" && dt.Rows[row][0].ToString() != "0" && row != 0) // This is roughly doing =IF(a2>0,a2,b2) on all rows
                {
                    dt.Rows[row][1] = dt.Rows[row][0];
                }
                row++;
            }
            dt.Columns.RemoveAt(0); // This removes the first column in the data table
            return dt;
        }

        private static DataTable SetQuantity(DataTable dt, int col) // This sets the quantity to "1", overwriting the existing data
        {
            int row = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if(row != 0)
                {
                    dt.Rows[row][col] = "1";
                }
                row++;
            }
            return dt;
        }

        private static DataTable CorrectCountry(DataTable dt)
        {
            string filename = "country.dat"; // Note that the country.dat file is simply a csv format, renamed to .dat to avoid it being confused with other files in the working directiory
            Console.WriteLine("Reading " + filename);
            int lineCount = File.ReadAllLines(filename).Length;
            TextFieldParser parser = new TextFieldParser(filename);
            parser.HasFieldsEnclosedInQuotes = true;
            parser.SetDelimiters(",");
            string[,] countrydata = new string[lineCount, 2];
            string[] fields;
            int row = 0;
            int column = 0;
            while (!parser.EndOfData)
            {
                fields = parser.ReadFields();
                foreach (string field in fields)
                {
                    countrydata[row, column] = field;
                    column = column + 1;
                }
                column = 0;
                row = row + 1;
            }
            parser.Close();
            Console.WriteLine("Country data from " + filename + " loaded into table.");
            row = 0;
            column = 0;
            int countryRow = 0;
            float count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if(row != 0)
                {
                    foreach (string field in countrydata)
                    {
                        if(dt.Rows[row][10].ToString() == field)
                        {
                            dt.Rows[row][10] = countrydata[countryRow, 0];
                            Console.WriteLine("Country at row " + row + " converted to ISO code: " + countrydata[countryRow,0] + ".");
                        }
                        count++;
                        countryRow = Convert.ToInt32(Math.Floor(count / 2));
                    }
                    countryRow = 0;
                    count = 0;
                }
                row++;
            }
            return dt;
        }

        private static DataTable AddChargeAmount(DataTable dt)
        {
            //----------------------------------------------------------------
            // insert column at 'H' in Excel
            // =if(f2="Credit",e2,e2*-1)
            dt = AddColumn(dt, 7);
            int lastRow = Convert.ToInt32(dt.Rows.Count.ToString()) - 1;
            int lastColumn = Convert.ToInt32(dt.Columns.Count.ToString()) - 1;
            int row = 1; // setting row to 1 to skip over the header row
            int column = 7;
            double sign = 1;
            dt.Rows[0][column] = "Charge Amount"; // Add the header for the column
            while (row <= lastRow)
            {
                if (dt.Rows[row][5].ToString() == "Credit")
                { sign = 1; }
                if (dt.Rows[row][5].ToString() == "Charge")
                { sign = -1; }
                if (row != 0)
                {
                    dt.Rows[row][column] = Convert.ToDouble(dt.Rows[row][4].ToString()) * sign;
                }
                row++;
            }
            return dt;
        }

        private static DataTable CompareToIIA(DataTable dtInv, DataTable dtIIA)
        {
            int InvRow = 0;
            int IIARow = 0;
            dtInv.Columns.Add();
            int CompareColumn = Convert.ToInt32(dtInv.Columns.Count.ToString()) - 1;
            foreach (DataRow drReasonCode in dtInv.Rows)
            { 
                dtInv.Rows[InvRow][CompareColumn] = "n/a";
                InvRow++;
            }
            dtInv.Rows[0][CompareColumn] = "Reason Code";
            InvRow = 0;
            dtInv.Columns.Add();
            foreach (DataRow drInv in dtInv.Rows)
            {
                if (InvRow != 0)
                { dtInv.Rows[InvRow][CompareColumn + 1] = 0; }
                InvRow++;
            }
            InvRow = 0;
            foreach (DataRow drInv in dtInv.Rows)
            {
                foreach (DataRow drIIA in dtIIA.Rows)
                {
                    if (dtInv.Rows[InvRow][11].ToString() == dtIIA.Rows[IIARow][6].ToString() && InvRow != 0 && IIARow != 0)
                    {
                        dtInv.Rows[InvRow][CompareColumn] = dtIIA.Rows[IIARow][9]; //Adds the IIA adjustment reason code
                        dtInv.Rows[InvRow][CompareColumn + 1] = Convert.ToDouble(dtInv.Rows[InvRow][CompareColumn + 1].ToString()) + Convert.ToDouble(dtIIA.Rows[IIARow][7].ToString()); //Adds all values up for each entry after the reason code
                    }
                    IIARow++;
                }
                InvRow++;
                IIARow = 0;
            }
            return dtInv;
        }

        private static DataTable FlagInconsistent(DataTable dt)
        {
            int row = 0;
            int searchRow = 0;
            AddColumn(dt, 39); // Adds yet another column to the DataTable
            foreach (DataRow dr in dt.Rows) // This cycles through the DataTable, populating the new column with 0's and adding a header
            {
                if (row == 0)
                { dt.Rows[row][39] = "Invoice Total"; }
                if (row != 0)
                { dt.Rows[row][39] = 0; }
                row++;
            }
            row = 0;
            foreach (DataRow drInv in dt.Rows) // Adds together all lines of invoices (including tax) next to the IIA total column
            {
                foreach (DataRow drSearchInv in dt.Rows)
                {
                    if(dt.Rows[row][11].ToString() == dt.Rows[searchRow][11].ToString() && row != 0) // Does the invoice number match the invoice number?
                    { dt.Rows[row][39] = Convert.ToDouble(dt.Rows[row][39].ToString()) + Convert.ToDouble(dt.Rows[searchRow][24].ToString())  + Convert.ToDouble(dt.Rows[searchRow][25].ToString()); } // If it does, add the amount and tax amount to the newly added column at the end.
                    searchRow++;
                }
                searchRow = 0;
                row++;
            }
            row = 0;
            foreach (DataRow drCom in dt.Rows)
            {
                if (row != 0 && (dt.Rows[row][37].ToString() == "Subscription Cleanup" || dt.Rows[row][37].ToString() == "TBD"))
                {
                    if(Convert.ToDouble(dt.Rows[row][38].ToString()) != Convert.ToDouble(dt.Rows[row][39].ToString()))
                    { FlagCell(dt, row, 37, "IIA: " + dt.Rows[row][11].ToString() + " has IIA amount mismatch"); }
                }
                row++;
            }
            return dt;
        }

        private static DataTable RemoveEntries(DataTable dt, string ToMatch, int col)
        {
            bool cont = true;
            int row = 0;
            DataRow DeleteRow = null;
            while (cont)
            {
                cont = false;
                foreach (DataRow dr in dt.Rows) // This will end up with DeleteRow equalling the last row marked as "Subscription Cleanup"
                {
                    if (dt.Rows[row][col].ToString() == ToMatch)
                    {
                        DeleteRow = dr;
                        cont = true;
                    }
                    row++;
                }
                if (DeleteRow != null)
                { dt.Rows.Remove(DeleteRow); } // If a row was found in the above foreach loop, this will delete that row
                DeleteRow = null;
                row = 0;
            }
            return dt;
        }

        private static DataTable RemoveAllTaxEntries(DataTable dt)
        {
            bool cont = true;
            int row = 0;
            DataRow DeleteRow = null;   // This just needed to be defined. Setting it to null seemed like the simplist definition I could give it.
            while (cont)
            {
                cont = false;
                foreach (DataRow dr in dt.Rows)
                {
                    if (row != 0)
                    {
                        if (dt.Rows[row][10].ToString() != "" && dt.Rows[row][10].ToString() != "0")
                        {
                            DeleteRow = dr;
                            cont = true;
                        }
                    }
                    row++;
                }
                if (DeleteRow != null)
                { dt.Rows.Remove(DeleteRow); } // If a row was found in the above foreach loop, this will delete that row
                DeleteRow = null;
                row = 0;
            }
            return dt;
        }

        private static DataTable ChangeHeader(DataTable dt)
        {
            dt.Rows[0][0] = "Account Number";
            dt.Rows[0][1] = "Account Name";
            dt.Rows[0][2] = "First Name";
            dt.Rows[0][3] = "Last Name";
            dt.Rows[0][4] = "Work Email";
            dt.Rows[0][5] = "Address1";
            dt.Rows[0][6] = "Address2";
            dt.Rows[0][7] = "City";
            dt.Rows[0][8] = "State";
            dt.Rows[0][9] = "Postal Code";
            dt.Rows[0][10] = "Country";
            dt.Rows[0][11] = "Invoice Number";
            dt.Rows[0][12] = "Invoice Date";
            dt.Rows[0][13] = "Due Date";
            dt.Rows[0][14] = "Contract Effective Date";
            dt.Rows[0][15] = "Effective Date C";
            dt.Rows[0][16] = "Subscription Start Date";
            dt.Rows[0][17] = "Accounting Code";
            dt.Rows[0][18] = "Charge Name";
            dt.Rows[0][19] = "Service Start Date";
            dt.Rows[0][20] = "Service End Date";
            dt.Rows[0][21] = "UOM";
            dt.Rows[0][22] = "Quantity";
            dt.Rows[0][23] = "Unit Price";
            dt.Rows[0][24] = "Charge Amount";
            dt.Rows[0][25] = "Tax Amount";
            dt.Rows[0][26] = "Currency";
            dt.Rows[0][27] = "Payment Method Type C";
            dt.Rows[0][28] = "Payment Terms C";
            dt.Rows[0][29] = "Payment Term";
            dt.Rows[0][30] = "Product Name";
            dt.Rows[0][31] = "Product Rate Plan Name";
            dt.Rows[0][32] = "ProductRatePlan.Id";
            dt.Rows[0][33] = "Product Rate Plan Charge Name";
            dt.Rows[0][34] = "ProductRatePlanCharge.Id";
            dt.Rows[0][35] = "Product Code  C";
            dt.Rows[0][36] = "Last Email Sent Date";
            dt.Rows[0][37] = "Reason Code";
            return dt;
        }

        private static DataTable FlagBlank(DataTable dt, int col)
        {
            int row = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (row != 0 && dt.Rows[row][col].ToString() == "")
                {
                    FlagCell(dt, row, col, "blank");
                }
                row++;
            }
            return dt;
        }

        private static DataTable FlagNA(DataTable dt, int col)
        {
            int row = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (row != 0 && dt.Rows[row][col].ToString() == "N/A")
                {
                    FlagCell(dt, row, col, "N/A");
                }
                row++;
            }
            return dt;
        }

        private static DataTable FlagTBD(DataTable dt, int col)
        {
            int row = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (row != 0 && dt.Rows[row][col].ToString() == "TBD")
                {
                    FlagCell(dt, row, col, "TBD: " + dt.Rows[row][3].ToString());
                }
                row++;
            }
            return dt;
        }

        private static DataTable FlagCell(DataTable dt, int row, int col, string error)
        {
            int lastColumn = Convert.ToInt32(dt.Columns.Count.ToString()) - 1;
            int PrintRow = row + 1;
            int PrintCol = col + 1;
            if(dt.Rows[0][lastColumn].ToString() != "Errors") // If an "Errors" column isn't already the last column, create a column with this header
            {
                dt = AddColumn(dt, lastColumn + 1);
                dt.Rows[0][lastColumn + 1] = "Errors";
            }
            lastColumn = Convert.ToInt32(dt.Columns.Count.ToString()) - 1; // Grab the last column number again, since it could have changed if a new "Errors" column was just added
            if(dt.Rows[0][lastColumn].ToString() == "Errors")
            {
                dt.Rows[row][lastColumn] = col;
                Console.WriteLine("Unexpected " + error + " at row " + PrintRow + ", column " + PrintCol + ".");
            }
            else
            { Console.WriteLine("Error in application. Last column expected to be titled Errors."); } // This else statment should never happen
            WriteLogCSV("log.csv", dt.TableName.ToString() + "," + PrintRow + "," + PrintCol + "," + error);
            return dt;
        }

        private static DataTable ReplaceCell(DataTable dt, int col, string findString, string replaceString)
        {
            int row = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (row != 0 && dt.Rows[row][col].ToString() == findString)
                {
                    FlagCell(dt, row, col, findString);
                    dt.Rows[row][col] = replaceString;
                }
                row++;
            }
            return dt;
        }

        private static DataTable ReplaceCellQuiet(DataTable dt, int col, string findString, string replaceString)
        {
            int row = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (row != 0 && dt.Rows[row][col].ToString() == findString)
                {
                    // Removed flagging to avoid extra clutter in logs
                    dt.Rows[row][col] = replaceString;
                }
                row++;
            }
            return dt;
        }

        private static ExcelWorksheet HighlightCells(ExcelWorksheet ws, DataTable dt) // This will highlight the cells that were flagged during the FlagCell function
        {
            int lastColumn = Convert.ToInt32(dt.Columns.Count.ToString());
            int row = 1;
            int ErrCol = 0;
            if (ws.Cells[1, lastColumn].Value.ToString() == "Errors")
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (ws.Cells[row, lastColumn].Value.ToString() != "" && row != 1)
                    {
                        ErrCol = Convert.ToInt32(ws.Cells[row, lastColumn].Value.ToString()) + 1;
                        ws.Cells[row, ErrCol].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        ws.Cells[row, ErrCol].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                    }
                    row++;
                }
            }
            else
            { Console.WriteLine("No errors"); }
            return ws;
        }

        private static DataTable AddColumn(DataTable dt, int loc)
        {
            int lastRow = Convert.ToInt32(dt.Rows.Count.ToString()) - 1;
            int lastColumn = Convert.ToInt32(dt.Columns.Count.ToString()) - 1;
            dt.Columns.Add();
            int row = 0;
            int column = lastColumn;
            while (column >= loc)
            {
                while (row <= lastRow)
                {
                    dt.Rows[row][column + 1] = dt.Rows[row][column];
                    row++;
                }
                row = 0;
                column--;
            }
            return dt;
        }

        private static ExcelWorksheet CreateSheet(ExcelPackage p, string sheetName, int position)
        {
            p.Workbook.Worksheets.Add(sheetName);
            ExcelWorksheet ws = p.Workbook.Worksheets[position];
            ws.Name = sheetName; //Setting Sheet's name
            ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
            ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

            return ws;
        }
        
        static void WriteExcelFile(ExcelPackage p, string filename) // Writes an Excel file from the provided ExcelPackage variable
        {
            Byte[] bin = p.GetAsByteArray();
            File.WriteAllBytes(filename, bin);
            Console.WriteLine(filename + " written to local directory.");
        }

        static void OpenFileInExcel(string file) // Open a file in Excel
        {
            ProcessStartInfo pi = new ProcessStartInfo(file);
            Process.Start(pi);
        }

        static void ClearLogCSV(string filename)
        {
            System.IO.File.WriteAllText(filename, string.Empty); // Before writing to the file, this empties the file. This way if there were previous contents with more lines than we are writing now, we will not have any of the old contents.
            try
            {
                var fs = File.Open(filename, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                var sw = new StreamWriter(fs);
                sw.Write("Error log for Zuora to GP data migration preparation"); // This creates the first line in the log.csv file
                sw.Flush();
                fs.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
            Console.WriteLine("-Error log cleared-");
        }

        static void WriteLogCSV(string filename, string data)
        {
            try
            {
                var fs = File.Open(filename, FileMode.Append, FileAccess.Write);
                var sw = new StreamWriter(fs);
                sw.Write("\n" + data);
                sw.Flush();
                fs.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
            Console.WriteLine("-Error written to log-");
        }

   /*private static DataTable CheckCreditType(DataTable dt)
    {
        int row = 0;
        foreach (DataRow dr in dt.Rows)
        {
            if (dt.Rows[row][8].ToString() == "CreditCard")
            {
                if(dt.Rows[row][9].ToString() == "")
                {
                    FlagCell(dt, row, 9, "blank");
                }
            }
            row++;
        }
        return dt;
    }*/
    }
}
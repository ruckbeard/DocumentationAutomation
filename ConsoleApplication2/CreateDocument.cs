using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Drawing;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ConsoleApplication2
{
    class CreateDocument
    {
        private Application oXL;
        private _Workbook oWB;
        private _Worksheet oSheet1;
        private _Worksheet oSheet2;
        private _Worksheet oSheet3;
        private object misvalue = System.Reflection.Missing.Value;

        public void createWorkbook(string projectNumber, string requestor, string subject, string description)
        {
            try
            {
                //Start Excel and get Application object.
                oXL = new Application();
                oXL.Visible = false;
                oXL.UserControl = false;

                //Get a new workbook.
                oWB = oXL.Workbooks.Add("");
                oSheet1 = initWorksheets(oWB, 1, "Tech Specs", "Technical Specifications");
                oSheet2 = initWorksheets(oWB, 2, "Coding Doc", "Coding Documentation");
                oSheet3 = initWorksheets(oWB, 3, "Unit Test", "Unit Testing");

                //Get information from user
                oSheet1.Cells[4, 3] = projectNumber;
                oSheet2.Cells[4, 3] = projectNumber;
                oSheet3.Cells[4, 3] = projectNumber;

                oSheet1.Cells[6, 3] = requestor;
                oSheet2.Cells[6, 3] = requestor;
                oSheet3.Cells[6, 3] = requestor;

                oSheet1.Cells[7, 3] = subject;
                oSheet2.Cells[7, 3] = subject;
                oSheet3.Cells[7, 3] = subject;

                oSheet1.Cells[8, 3] = description;
                oSheet2.Cells[8, 3] = description;
                oSheet3.Cells[8, 3] = description;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


        private _Worksheet initWorksheets(_Workbook oWB, int index, string documentShort, string documentFull)
        {
            _Worksheet worksheet = (_Worksheet)oWB.Worksheets[index];

            worksheet.Name = documentShort;

            string documentHeader = "Document:";
            string authorHeader = "Author:";
            string projectNumberHeader = "Project Number:";
            string dateHeader = "Date:";
            string requestorHeader = "Requestor:";
            string subjectHeader = "Subject:";
            string descriptionHeader = "Description:";

            //Add table headers going cell by cell.
            worksheet.Cells[2, 2] = documentHeader;
            worksheet.Cells[3, 2] = authorHeader;
            worksheet.Cells[4, 2] = projectNumberHeader;
            worksheet.Cells[5, 2] = dateHeader;
            worksheet.Cells[6, 2] = requestorHeader;
            worksheet.Cells[7, 2] = subjectHeader;
            worksheet.Cells[8, 2] = descriptionHeader;

            worksheet.Cells[2, 3] = documentFull;
            worksheet.Cells[3, 3] = "Kevin VanderWulp";
            worksheet.Cells[5, 3] = DateTime.Now.ToShortDateString();

            //Format A1:D1 as bold, vertical alignment = center.
            worksheet.get_Range("B2", "B8").Font.Bold = true;
            worksheet.get_Range("B2", "B8").Borders[XlBordersIndex.xlEdgeRight].Color = Color.Black;
            worksheet.get_Range("B2", "B3").Borders[XlBordersIndex.xlEdgeBottom].Color = Color.Black;
            worksheet.get_Range("B2", "C2").Borders[XlBordersIndex.xlEdgeBottom].Color = Color.Black;
            worksheet.get_Range("B3", "C3").Borders[XlBordersIndex.xlEdgeBottom].Color = Color.Black;
            worksheet.get_Range("B4", "C4").Borders[XlBordersIndex.xlEdgeBottom].Color = Color.Black;
            worksheet.get_Range("B5", "C5").Borders[XlBordersIndex.xlEdgeBottom].Color = Color.Black;
            worksheet.get_Range("B6", "C6").Borders[XlBordersIndex.xlEdgeBottom].Color = Color.Black;
            worksheet.get_Range("B7", "C7").Borders[XlBordersIndex.xlEdgeBottom].Color = Color.Black;

            worksheet.get_Range("B:B", Type.Missing).EntireColumn.ColumnWidth = 17;
            worksheet.get_Range("C:C", Type.Missing).EntireColumn.ColumnWidth = 75;

            worksheet.get_Range("C:C", Type.Missing).EntireColumn.WrapText = true;

            worksheet.get_Range("C4", Type.Missing).HorizontalAlignment = XlHAlign.xlHAlignLeft;
            worksheet.get_Range("C5", Type.Missing).HorizontalAlignment = XlHAlign.xlHAlignLeft;

            return worksheet;
        }

        public void saveWorkbook(string projectNumber)
        {
            string fileName = @"\\prddata\home\NCVKVW\Documentation\" + projectNumber + @"\" + projectNumber + @"-TechSpecs-CodingDoc-UnitTest-" + DateTime.Now.ToString("MMddyyyy") + @".xlsx";
            Directory.CreateDirectory(@"\\prddata\home\NCVKVW\Documentation\" + projectNumber);
            oWB.SaveAs(fileName, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            oWB.Close();
        }

        public void claimPaymentError(int payments, List<string> draftNumbers, List<string> claimNumbers, string projectNumber)
        {
            StringBuilder sb = new StringBuilder();
            StringBuilder sqlBuilder = new StringBuilder();

            sqlBuilder.Append("\\*\n");
            sqlBuilder.Append("\tSource Member: CL" + projectNumber + "\n");
            sqlBuilder.Append("\tTask: PR" + projectNumber + "\n");
            sqlBuilder.Append("\tAuthor: Kevin VanderWulp\n");
            sqlBuilder.Append("\tObjects: DRFTFILE, PMSPAP00\t\n");
            sqlBuilder.Append("*\\\n");

            for (int i = 0; i < payments; i++)
            {
                sb.Append("Delete the records where the draft number " + draftNumbers[i] + " and claim number " + claimNumbers[i] + " match in the PMSPAP00 and DRFTFILE tables.\n\n");

                sqlBuilder.Append("\n");
                sqlBuilder.Append("delete from @targetlib.pmspap00\n");
                sqlBuilder.Append("where mst_check_no = '" + draftNumbers[i] + "' and claimno = '" + claimNumbers[i] + "';\n");
                sqlBuilder.Append("\n");
                sqlBuilder.Append("delete from @targetlib.drftfile\n");
                sqlBuilder.Append("where drftnumber = '" + draftNumbers[i] + "' and claim = '" + claimNumbers[i] + "';\n");
            }

            sb.Append("Examiner has been notified that these records are going to be deleted.");

            oSheet1.Cells[9, 2] = "Problem:";
            oSheet1.Cells[10, 3] = "POINT has created a DRFTFILE and PMSPAP00 record without a proper PMSPCL50 for this check.";

            oSheet1.Cells[12, 2] = "Solution:";
            oSheet1.Cells[13, 3] = sb.ToString();

            string fileName = @"\\prddata\home\NCVKVW\Documentation\" + projectNumber + @"\" + projectNumber + @"-DeleteClaimPayments-" + DateTime.Now.ToString("MMddyyyy") + @".sql";
            Directory.CreateDirectory(@"\\prddata\home\NCVKVW\Documentation\" + projectNumber);
            StreamWriter sw = new StreamWriter(fileName);
            sw.Write(sqlBuilder.ToString());
            sw.Flush();
            sw.Close();
        }
    }
}

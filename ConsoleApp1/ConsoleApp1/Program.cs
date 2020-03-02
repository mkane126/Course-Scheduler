using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace mysheet
{
    class Program
    {
        static void Main(string[] args)
        {
            FileStream outputStream;
            StreamWriter writer;
            TextWriter oldOut = Console.Out;
            try
            {
                outputStream = new FileStream("C:\\temp\\output.txt", FileMode.OpenOrCreate, FileAccess.Write);
                writer = new StreamWriter(outputStream);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return;
            }
            Console.SetOut(writer);


            occupyExcel();
            writer.Close();
            outputStream.Close();
        }

        static void occupyExcel()
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open("C:\\temp\\actors.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;


            Excel.Range userRange = x.UsedRange;

            int countRows = userRange.Rows.Count;
            int countCols = userRange.Columns.Count;

            String[] lines = new string[countRows + 1];

            for (int i = 1; i <= countRows; i++)
            {

                lines[i] = "Actor ID: " + Convert.ToString((userRange.Cells[i, 1] as Excel.Range).Value2)
                    + ", Full Name: " + Convert.ToString((userRange.Cells[i, 2] as Excel.Range).Value2)
                    + " " + Convert.ToString((userRange.Cells[i, 3] as Excel.Range).Value2);

                Console.WriteLine(lines[i]);
            }
            x.Columns.AutoFit();


            sheet.Close(false, Type.Missing, Type.Missing);
            sheet = null;
            excel.Quit();
            excel = null;
            GC.Collect();

        }
    }
}

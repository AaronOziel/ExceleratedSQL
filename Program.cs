/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *\
|*  Author: Aaron Oziel (May 2014)
|*  Project: (Excel)eratedSQL
|*  Description:
|*      "Data needs to be moved from a spreadsheet to a database automatically. 
|*  Write a program that allows the user to select which .xls or .xlsx file to 
|*  import, grab the data from the spreadsheet, and send the data to a stored 
|*  procedure on a SQL Server. You are allowed to use whatever tools/methods 
|*  you can find on the internet to assist in the project. Use whatever 
|*  language/tool set you are comfortable with but you must tell me why a 
|*  tool was used."
|* 
|*  Tools:
|*  [C#]: Considering this project is dealing with two different Microsoft
|*  products (Excel and SQL Server) I thought it best to use Microsoft's 
|*  favorite langauge. 
|*  [Links]:
|*  http://www.codeproject.com/Tips/696864/Working-with-Excel-Using-Csharp
|*  http://www.codeproject.com/Articles/4416/Beginners-guide-to-accessing-SQL-Server-through-C
|*  
|* 
|*  Assumptions & Conditions: 
|*      - The excel file will be exactly 6 columns in width
|*      - Each cell will have approptiate data in it (ie: no strings in number cells)
|*      - ID column must be unique (it is the PK for the table)
|*      - Only the [Date Created] and [Rank] columns can be null
|*      - Empty rows will cause errors. Never enter data in a cell you don't intend to use. 
|* 
|*  
|*  Sample insert command:
|*        cmd.Parameters["@Deck"].Value = "G/W Aggro";
|*        cmd.Parameters["@Date"].Value = "10/12/2012";
|*        cmd.Parameters["@Rank"].Value = 2;
|*        cmd.Parameters["@Player"].Value = "Aaron Oziel";
|*        cmd.Parameters["@ID"].Value = 955288156;
|*        cmd.Parameters["@Pro"].Value = 0;
|*        cmd.ExecuteNonQuery();
|*
\* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExceleratedSQL
{
    class Program
    {
        public const int MAX_COL = 6; // Maximum number of colums

        static void Main(string[] args)
        {
            // Setup absolute file path
            string path = null;
            if (args.Length == 0)
                path = @"C:\Users\Aaron\Documents\Visual Studio 2013\Projects\ExceleratedSQL\ExceleratedSQL\ExceleratedSQL.xlsx";
            else
                path = args[0];

            // Create Excel variables 
            Excel.Application excel = null;
            Excel.Workbook wkb = null;
            int lastcol, lastrow;

            try // Open file and read out important information  
            {
                excel = new Excel.Application();
                excel.Visible = false;
                wkb = excel.Workbooks.Open(path);
                Excel.Worksheet sheet = (Excel.Worksheet)wkb.Sheets[1];
                // Get dimensions of relavent data
                lastcol = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                lastrow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                // Check on number of columns
                if (lastcol < MAX_COL)
                {   // Not enough? End program. 
                    Console.WriteLine("[Error  #100] Excel document did not have enough columns.");
                    return;
                }
                else if (lastcol > MAX_COL)
                {   // Too many? Show a warning.
                    Console.WriteLine("[Warning#100] Excel document had more columns than necessary, columns" + MAX_COL + 1 + "+ will be ignored.");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("[Error  #101] " + e.ToString());
                return;
            }

            // Create SQL variables
            SqlCommand cmd = null;
            SqlConnection myConnection = null;
            string stmt = "INSERT INTO Decks ([Deck Name], [Date Created], Rank, [Player Name], ID, Pro) VALUES(@Deck, @Date, @Rank, @Player, @ID, @Pro)";

            try
            {   // This part of the program is in need of help, too implimentation specific
                myConnection = new SqlConnection(Properties.Settings.Default.ConnetionString);
                myConnection.Open();

            }
            catch (Exception e)
            {   // If connection fail end program. 
                Console.WriteLine("[Error  #200] " + e.ToString());
                return;
            }

            // Setup parameters for SQL insert command
            cmd = new SqlCommand(stmt, myConnection);
            cmd.Parameters.Add("@Deck", System.Data.SqlDbType.VarChar, 50);
            cmd.Parameters.Add("@Date", System.Data.SqlDbType.DateTime);
            cmd.Parameters.Add("@Rank", System.Data.SqlDbType.SmallInt);
            cmd.Parameters.Add("@Player", System.Data.SqlDbType.VarChar, 50);
            cmd.Parameters.Add("@ID",   System.Data.SqlDbType.Int);
            cmd.Parameters.Add("@Pro",  System.Data.SqlDbType.Bit);

            try // Read every row in the spreadsheet and populate the database
            {
                for (int i = 1; i <= lastrow; i++)
                {
                    if (Excel.Range.Equals(excel.get_Range("A" + i, Type.Missing), "")) // TODO: Is a blank cell "" or null?
                    {
                        Console.WriteLine("[Warning#201] Row " + i + " seems to be blank. It was skipped.");
                    }
                    else
                    {
                        cmd.Parameters["@Deck"].Value   = excel.get_Range("A" + i, Type.Missing).Text;
                        cmd.Parameters["@Date"].Value = Convert.ToDateTime(excel.get_Range("B" + i, Type.Missing).Text);
                        cmd.Parameters["@Rank"].Value = excel.get_Range("C" + i, Type.Missing).Text;
                        cmd.Parameters["@Player"].Value = excel.get_Range("D" + i, Type.Missing).Text;
                        cmd.Parameters["@ID"].Value = excel.get_Range("E" + i, Type.Missing).Text;
                        cmd.Parameters["@Pro"].Value = excel.get_Range("F" + i, Type.Missing).Text;
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("[Error  #201] " + e.ToString());
            }
            
            try // Close everything, if this fails a lot of rogue processes will be running. 
            {
                TestSQL(myConnection);
                myConnection.Close();
                wkb.Close();
                excel.Quit();
            }
            catch (Exception e)
            {
                Console.WriteLine("[Error  #202] " + e.ToString());
            }

            Console.ReadLine();
        }

        public static void TestSQL(SqlConnection myConnection)
        {
            try
            {
                SqlDataReader myReader = null;
                SqlCommand myCommand = new SqlCommand("select * from Decks", myConnection);
                myReader = myCommand.ExecuteReader();
                while (myReader.Read())
                {
                    Console.WriteLine(myReader.GetString(0));
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        // http://social.msdn.microsoft.com/Forums/vstudio/en-US/78e75d1a-0795-4bdb-8a62-ae6faa909986/convert-number-to-alphabet
        // This method is taken from this URL by User "Manivannan.D.Sekaran" who wrote this on Monday, June 04, 2007 10:09 AM
        // Renamed: "Column To Letter"
        public static String colToLetter(int number, bool isCaps)
        {
            Char c = (Char)((isCaps ? 65 : 97) + number);
            return c.ToString();
        }
    }
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *\
|*  Author: Aaron Oziel (April/May 2014)
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
|*  favorite language. Other than that I used the obvious SQL and Excel
|*  libraries (System.Data.SqlClient & Microsoft.Office.Interop.Excel).
|* 
|*  [Links]:
|*  http://www.codeproject.com/Tips/696864/Working-with-Excel-Using-Csharp
|*  http://www.codeproject.com/Articles/4416/Beginners-guide-to-accessing-SQL-Server-through-C
|*  
|*  Assumptions & Conditions: 
|*      - The excel file will be exactly 6 columns in width
|*      - Each cell will have appropriate data in it (ie: no strings in number cells)
|*      - ID column must be unique (it is the PK for the table)
|*      - Empty rows will be caught and handled. Try to avoid them though. 
|*
\* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;

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
                path = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\ExceleratedSQL.xlsx";
            else
                path = args[0];

            // Create Excel variables 
            Excel.Application excel = null;
            Excel.Workbook wkb = null;
            Excel.Worksheet sheet = null;
            int lastcol, lastrow = 0;

            try // Open file and read out important information  
            {
                // Check for existance of file
                if (!File.Exists(path))
                    throw new FileNotFoundException();
                excel = new Excel.Application();
                excel.Visible = false;
                wkb = excel.Workbooks.Open(path);
                sheet = (Excel.Worksheet)wkb.Sheets[1];
                // Get dimensions of relavent data
                lastcol = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                lastrow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                // Check on number of columns
                if (lastcol < MAX_COL)
                {   // Not enough? End program. 
                    Console.WriteLine("[Error  #100] Excel document did not have enough columns.");
                    Console.ReadLine();
                    return;
                }
                else if (lastcol > MAX_COL)
                {   // Too many? Show a warning.
                    Console.WriteLine("[Warning#100] Excel document had more columns than necessary, columns " + (MAX_COL + 1) + "+ will be ignored.");
                }

                Console.WriteLine("Opened Excel File Successfully");
            }
            catch (Exception e)
            {
                Console.WriteLine("[Error  #101] " + e.ToString());
                if (wkb != null) wkb.Close(); // Close open sheet
                if (excel != null) excel.Quit(); // Close open excel application
                Console.ReadLine();
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
                Console.WriteLine("Connected to SQL Database Successfully");
            }
            catch (Exception e)
            {   // If connection fails end program. 
                Console.WriteLine("[Error  #200] " + e.ToString());
                if (wkb != null) wkb.Close(); // Close open sheet
                if (excel != null) excel.Quit(); // Close open excel application
                if (myConnection != null) myConnection.Close(); // Close open SQL connection
                Console.ReadLine();
                return;
            }

            // Setup parameters for SQL insert command
            cmd = new SqlCommand(stmt, myConnection);
            cmd.Parameters.Add("@Deck", System.Data.SqlDbType.VarChar, 50);
            cmd.Parameters.Add("@Date", System.Data.SqlDbType.DateTime);
            cmd.Parameters.Add("@Rank", System.Data.SqlDbType.SmallInt);
            cmd.Parameters.Add("@Player", System.Data.SqlDbType.VarChar, 50);
            cmd.Parameters.Add("@ID", System.Data.SqlDbType.Int);
            cmd.Parameters.Add("@Pro", System.Data.SqlDbType.Bit);

            try // Read every row in the spreadsheet and populate the database
            {
                int inserts = 0;
                for (int i = 1; i <= lastrow; i++)
                {
                    if (excel.get_Range("A" + i, Type.Missing).Text == "") // Weed out blank rows
                    {
                        Console.WriteLine("[Warning#201] Row " + i + " seems to be blank. It was skipped.");
                    }
                    else
					{
                        cmd.Parameters["@Deck"].Value = excel.get_Range("A" + i, Type.Missing).Text;
                        cmd.Parameters["@Date"].Value = Convert.ToDateTime(excel.get_Range("B" + i, Type.Missing).Text);
                        cmd.Parameters["@Rank"].Value = excel.get_Range("C" + i, Type.Missing).Text;
                        cmd.Parameters["@Player"].Value = excel.get_Range("D" + i, Type.Missing).Text;
                        cmd.Parameters["@ID"].Value = excel.get_Range("E" + i, Type.Missing).Text;
                        cmd.Parameters["@Pro"].Value = excel.get_Range("F" + i, Type.Missing).Text;
                        cmd.ExecuteNonQuery();
                        inserts++;
                    }
                }
                Console.WriteLine("Inserted [" + inserts + "] entries into the database");
            }
            catch (System.FormatException format)
            {
                Console.WriteLine("[Error  #201] A cell was improperly formatted. "
                + "Please check all cells to ensure no incorrect data exists.");
            }
            catch (System.Data.SqlClient.SqlException sql)
            {
                Console.WriteLine("[Error  #202] " + sql.ToString());
            }
            catch (Exception e)
            {
                Console.WriteLine("[Error  #203] " + e.ToString());
            }

            try // Close everything, if this fails a lot of rogue processes will be running. 
            {
                //TestSQL(myConnection);
                if (wkb != null) wkb.Close(); // Close open sheet
                if (excel != null) excel.Quit(); // Close open excel application
                if (myConnection != null) myConnection.Close(); // Close open SQL connection
            }
            catch (Exception e) // This should never be used. 
            {
                Console.WriteLine("[Error  #202] " + e.ToString());
            }
            Console.WriteLine("--------------------");
            Console.WriteLine("| Program Complete |");
            Console.WriteLine("--------------------");
            Console.ReadLine();
        }

        // Simple test function that will print the [Deck Name] of evrey entry in the database. 
        // Used to test and see if inserts were truely successful. 
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
    }
}
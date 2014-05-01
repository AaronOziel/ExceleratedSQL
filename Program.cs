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
        static void Main(string[] args)
        {
            string path = @"C:\Users\Aaron\Documents\Visual Studio 2013\Projects\ExceleratedSQL\ExceleratedSQL\ExceleratedSQL.xlsx";

            Excel.Application excel = null;
            Excel.Workbook wkb = null;

            try
            {
                excel = new Excel.Application();
                excel.Visible = false;
                wkb = excel.Workbooks.Open(path);
                Excel.Worksheet sheet = (Excel.Worksheet)wkb.Sheets[1]; // Explicit cast is not required here
                int lastrow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                int lastcol = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                Excel.Range range = null;

                for (int row = 0; row < lastrow; row++)
                {
                    for (int col = 0; col < lastcol; col++)
                    {
                        range = sheet.get_Range(colToLetter(col, true) + (row + 1), Type.Missing);
                        if (col < lastcol - 1)
                            Console.Write(range.Text + ",\t");//.ToString() + ",\t");
                        else
                            Console.Write(range.Text);
                    }
                    Console.WriteLine();
                }
                Console.WriteLine("--------------");
                Console.WriteLine("FINISHED!");
                //Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                excel.Workbooks.Close();
                //Console.ReadLine();
            }

            Console.WriteLine("Connecting to Database.");
            TestSQL();
            Console.WriteLine("Database Connection Complete.");

        }

        // http://social.msdn.microsoft.com/Forums/vstudio/en-US/78e75d1a-0795-4bdb-8a62-ae6faa909986/convert-number-to-alphabet
        // This method is taken from this URL by User "Manivannan.D.Sekaran" who wrote this on Monday, June 04, 2007 10:09 AM
        // Renamed: "Column To Letter"
        public static String colToLetter(int number, bool isCaps)
        {
            Char c = (Char)((isCaps ? 65 : 97) + number);
            return c.ToString();
        }

        public static void TestSQL()
        {
            Console.WriteLine("Opening Connection to Database.");
            SqlConnection myConnection = null;
            try
            {
                myConnection = new SqlConnection(Properties.Settings.Default.ConnetionString);
                myConnection.Open();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return;
            }

            string stmt = "INSERT INTO Decks ([Deck Name], [Date Created], Rank, [Player Name], ID, Pro) VALUES(@Deck, @Date, @Rank, @Player, @ID, @Pro)"; // TODO: Varify this is an acceptable statement
            //string stmt = "INSERT INTO [dbo].[Decks](deck, date, rank, player, id, pro) VALUES(badrdw, 10/12/2012, 2, aaronoziel, 955281856, true)"; // TODO: Varify this is an acceptable statement
            
            SqlCommand cmd = new SqlCommand(stmt, myConnection);
            cmd.Parameters.Add("@Deck", System.Data.SqlDbType.VarChar, 50);
            cmd.Parameters.Add("@Date", System.Data.SqlDbType.DateTime);
            cmd.Parameters.Add("@Rank", System.Data.SqlDbType.SmallInt);
            cmd.Parameters.Add("@Player", System.Data.SqlDbType.VarChar, 50);
            cmd.Parameters.Add("@ID", System.Data.SqlDbType.Int);
            cmd.Parameters.Add("@Pro", System.Data.SqlDbType.Bit);

            cmd.Parameters["@Deck"].Value = "G/W Aggro";
            cmd.Parameters["@Date"].Value = "10/12/2012";
            cmd.Parameters["@Rank"].Value = 2;
            cmd.Parameters["@Player"].Value = "Aaron Oziel";
            cmd.Parameters["@ID"].Value = 955288156;
            cmd.Parameters["@Pro"].Value = 0;

            cmd.ExecuteNonQuery();

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

            try
            {
                myConnection.Close();
                Console.WriteLine("Connection Closed!");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}

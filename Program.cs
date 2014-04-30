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
            
        }
    }
}

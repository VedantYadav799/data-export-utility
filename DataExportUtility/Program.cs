namespace DEU;
using System;
using System.Data;
using ClosedXML.Excel;


  class Program
    {
        static void Main(string[] args)
        {

            DataTable table= MyDAL.GetData();
            Export.ExportPdfToFile(table);
          
        }
    }
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace RDJ_Reports
{
    class Excel
    {
        string path = @"C:\RDJ Projects\RDJ-Projects\Reports\RDJ RW MO Production 2.0.xlsx";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public List<RDJSpecInfo> line1 = new List<RDJSpecInfo>();
        private int row = 2;
        private int prodNumCol = 1;
        private int formulaNumCol = 3;
        private int casesCol = 35;

        List<Dictionary<string, List<DateTime>>> batchInfoL1 = new List<Dictionary<string, List<DateTime>>>();

        public Excel(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
            wb.RefreshAll();
        }
        public DateTime Read(int i, int j)
        {
            //i++;
            //j++;
            //if (ws.Cells[i, j].value2 != null)
                
                return DateTime.FromOADate(ws.Cells[i, j].value2);
            //else
                //return "";
        }
        public int intRead(int i, int j)
        {
            int val;
            string value;
            value = Convert.ToString( ws.Cells[i, j].value2);
            if (int.TryParse((ws.Cells[i, j].value2).ToString(), out val))
                return val;
            else
                return 1;
            //else
            //return "";
        }
        public float floatRead(int i, int j)
        {
            float val;
           // string value;
            //value = Convert.ToString(ws.Cells[i, j].value2);
            if (float.TryParse((ws.Cells[i, j].value2).ToString(), out val))
                return val;
            else
                return 1;
            //else
            //return "";
        }
        public string stringRead(int i, int j)
        {
            //i++;
            //j++;
            //if (ws.Cells[i, j].value2 != null)

            return ws.Cells[i, j].value2;
            //else
            //return "";
        }
        public void Close_File()
        {
           
            wb.Close();
            
        }
    }
    class Report
    {

    }
}

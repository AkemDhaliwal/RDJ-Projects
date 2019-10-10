using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace RDJ_Reports
{
    public partial class Form1 : Form
    {
        public DateTime startDateTime;
        public DateTime endDateTime;

        List<int> batchList = new List<int>();
        List<int> prodList = new List<int>();
        IDictionary<int, float> usedBatches;
        IDictionary<int, float> requiredBatches;
        DataRow RowCal;



        public Form1()
        {
            InitializeComponent();
            startTime.CustomFormat = "hh:mm";
            endTime.CustomFormat = "hh:mm";

        }
      
        private void genConvReport_Click(object sender, EventArgs e)
        {
            usedBatches = new Dictionary<int, float>();
            requiredBatches = new Dictionary<int, float>();
            
            try
            {
                //Read RW Production data depnding on start and end date and store it in a Dataset object
                startDateTime = startDate.Value.Date + startTime.Value.TimeOfDay;
                endDateTime = endDate.Value.Date + endTime.Value.TimeOfDay;

                DataSet excelDataSetRW = new DataSet();
                DataSet excelDataSetSpecs = new DataSet();
                DataSet excelDataShow = new DataSet();

                string ConnectionStringRWData = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}; Extended Properties='Excel 12.0; HDR=Yes'",
                    @"C:\RDJ Projects\Reports\RDJ RW MO Production.xlsx");

                OleDbConnection connRWData = new OleDbConnection(ConnectionStringRWData);

                connRWData.Open();

                using (OleDbDataAdapter objDA1 = new System.Data.OleDb.OleDbDataAdapter("select[Item], [Status], [UM],[LP Quantity] from[Sheet1$] " +
                    "WHERE[LP Reported At] BETWEEN #" + startDateTime + "# and #" + endDateTime + "#", connRWData))
                {
                    objDA1.Fill(excelDataSetRW);
                }
                connRWData.Close();
                connRWData.Dispose();


                //Open RDJ Specs info(ConversionFactInfo) file
                string ConnectionStringRDJSpecs = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}; Extended Properties='Excel 12.0; HDR=Yes'",
                    @"C:\RDJ Projects\Reports\ConversionFactInfo.xlsx");
                OleDbConnection connRDJSpecs = new OleDbConnection(ConnectionStringRDJSpecs);

                connRDJSpecs.Open();

                //Go through each selected row of RW production Dataset
                foreach (DataRow row in excelDataSetRW.Tables[0].Rows)
                {
                    if (row[2].ToString() == "BATCH")
                    {
                        if (!usedBatches.ContainsKey((int)row[0]))   //0 intex is Item Number
                            usedBatches.Add((int)row[0], (float)row[3]); //3 index is quantity
                        else
                            usedBatches[(int)row[0]] = usedBatches[(int)row[0]] + (float)row[3];

                        if(!requiredBatches.ContainsKey((int)row[0]))
                            requiredBatches.Add((int)row[0], 0f); //3 index is quantity
                    }
                    if (row[2].ToString() == "CASE")
                    {
                        using (OleDbDataAdapter objDA2 = new System.Data.OleDb.OleDbDataAdapter("select[Batch Number], [Yield] from[info]" +
                            " WHERE [Internal Product Number] = #" + (int)row[0] + "#", connRDJSpecs))
                        {
                            objDA2.Fill(excelDataSetSpecs);
                        }
                        RowCal = excelDataSetSpecs.Tables[0].Rows[0];

                        if (!requiredBatches.ContainsKey((int)RowCal[2]))//2 is index for formula num in ConversionFactInfo
                            requiredBatches.Add((int)RowCal[2], 0f); //3 index is quantity

                        requiredBatches[(int)RowCal[2]] = requiredBatches[(int)RowCal[2]] + (((float)row[3]) / ((float)RowCal[4]));
                    }
                }
                connRWData.Close();
                connRWData.Dispose();

                int i = 0;
                DataRow showRow = excelDataShow.Tables[0].NewRow();
                showRow[0] = "Formula Number";
                showRow[1] = "Product Description";
                showRow[2] = "Conversion";
                excelDataShow.Tables[0].Rows.Add(showRow);
                foreach (KeyValuePair<int, float> item in usedBatches)
                {
                    showRow[0] = item.Key;
                    showRow[1] = "  ";
                    showRow[2] = (int)Math.Round((double)(100 * (requiredBatches[item.Key] /item.Value)));
                    excelDataShow.Tables[0].Rows.Add(showRow);
                }
                dataGridView1.DataSource = excelDataShow.Tables[0];
            }
            catch
            {
                errorMsg.Text = "Error in Data";
            }
        }
    }
}

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
        List<DisplayInfo> info = new List<DisplayInfo>();
        
        IDictionary<int, float> usedBatches;
        IDictionary<int, float> requiredBatches;
        IDictionary<int, string> batchDescription;
        IDictionary<int, string> finalProdInfo;
        DataRow RowCal;
        DataRow RowCal1;

        DataSet excelDataSetRW = new DataSet();
        DataSet excelDataSetSpecs = new DataSet();
        DataSet excelDataSetSpecs2 = new DataSet();
        DataSet excelDataSetSpecsAss = new DataSet();
        DataSet excelDataShow = new DataSet();
        
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
            batchDescription = new Dictionary<int, string>();
            finalProdInfo = new Dictionary<int, string>();
            var lines = new List<string> {"L1", "L2", "L3", "L4","MIXING" };
            
            try
            {
                //Read RW Production data depnding on start and end date and store it in a Dataset object
                startDateTime = startDate.Value.Date + startTime.Value.TimeOfDay;
                endDateTime = endDate.Value.Date + endTime.Value.TimeOfDay;

                string ConnectionStringRWData = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}; Extended Properties='Excel 12.0; HDR=Yes'",
                    @"C:\RDJ Projects\Reports\RDJ RW MO Production.xlsx");

                OleDbConnection connRWData = new OleDbConnection(ConnectionStringRWData);

                connRWData.Open();

                using (OleDbDataAdapter objDA1 = new System.Data.OleDb.OleDbDataAdapter("select[Item], [Status], [UM],[LP Quantity],[Line], [Item Description] from[Sheet1$] " +
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
                int j = 0;
                int i = 0;
                int k = 0;
                //Go through each selected row of RW production Dataset
                foreach (DataRow row in excelDataSetRW.Tables[0].Rows)
                {
                    if (row[2].ToString() == "BATCH" && lines.Contains(row[4].ToString(), StringComparer.OrdinalIgnoreCase))
                    {
                        if (!usedBatches.ContainsKey(Int32.Parse(row[0].ToString())))   //0 intex is Item Number
                        {
                            usedBatches.Add(Int32.Parse(row[0].ToString()), (float.Parse(row[3].ToString()))); //3 index is quantity
                            batchDescription.Add(Int32.Parse(row[0].ToString()), (row[5].ToString()));
                        }
                        else
                            usedBatches[int.Parse(row[0].ToString())] = usedBatches[int.Parse(row[0].ToString())] + float.Parse(row[3].ToString());

                        if(!requiredBatches.ContainsKey(int.Parse(row[0].ToString())))
                            requiredBatches.Add(int.Parse(row[0].ToString()), 0f); //3 index is quantity
                    }
                    else if (row[2].ToString() == "CASE" && lines.Contains(row[4].ToString(), StringComparer.OrdinalIgnoreCase))
                    {
                        int num = Int32.Parse(row[0].ToString());
                        if (num > 70000000)
                        {
                            using (OleDbDataAdapter objDA3 = new System.Data.OleDb.OleDbDataAdapter("select[Batch Number], [Weight per Case], [Finished Kilos per batch] from[Sheet2$]" + " WHERE [Internal Product Number] = " + num, connRDJSpecs))
                            {
                                objDA3.Fill(excelDataSetSpecs2);
                            }
                            k = excelDataSetSpecs2.Tables[0].Rows.Count-1;
                            RowCal1 = excelDataSetSpecs2.Tables[0].Rows[k];

                            if (!requiredBatches.ContainsKey(int.Parse(RowCal1[0].ToString())))//2 is index for formula num in ConversionFactInfo
                                requiredBatches.Add(int.Parse(RowCal1[0].ToString()), 0f); //3 index is quantity

                            requiredBatches[int.Parse(RowCal1[0].ToString())] = requiredBatches[int.Parse(RowCal1[0].ToString())] + (((float.Parse(row[3].ToString()))* (float.Parse(RowCal1[1].ToString()))) / (float.Parse(RowCal1[2].ToString())));
                            
                        }
                        else
                        {
                            using (OleDbDataAdapter objDA2 = new System.Data.OleDb.OleDbDataAdapter("select[Batch Number], [Yield] from[Sheet1$]" +
                                " WHERE [Internal Product Number] = " + num, connRDJSpecs))
                            {
                                objDA2.Fill(excelDataSetSpecs);
                            }
                            RowCal = excelDataSetSpecs.Tables[0].Rows[j];

                            if (!requiredBatches.ContainsKey(int.Parse(RowCal[0].ToString())))//2 is index for formula num in ConversionFactInfo
                                requiredBatches.Add(int.Parse(RowCal[0].ToString()), 0f); //3 index is quantity

                            requiredBatches[int.Parse(RowCal[0].ToString())] = requiredBatches[int.Parse(RowCal[0].ToString())] + ((float.Parse(row[3].ToString())) / (float.Parse(RowCal[1].ToString())));
                            j++;
                        }
                    }
                    else if(row[2].ToString() == "KG" && lines.Contains(row[4].ToString(), StringComparer.OrdinalIgnoreCase))
                    {
                        int num2 = Int32.Parse(row[0].ToString());
                        using (OleDbDataAdapter objDA3 = new System.Data.OleDb.OleDbDataAdapter("select[Batch Number], [Finished Kilos per batch] from[Sheet1$]" +
                            " WHERE [Internal Product Number] = " + num2, connRDJSpecs))
                        {
                            objDA3.Fill(excelDataSetSpecsAss);
                        }
                        RowCal = excelDataSetSpecsAss.Tables[0].Rows[i];

                        if (!requiredBatches.ContainsKey(int.Parse(RowCal[0].ToString())))//2 is index for formula num in ConversionFactInfo
                            requiredBatches.Add(int.Parse(RowCal[0].ToString()), 0f); //3 index is quantity

                        requiredBatches[int.Parse(RowCal[0].ToString())] = requiredBatches[int.Parse(RowCal[0].ToString())] + ((float.Parse(row[3].ToString())) / (float.Parse(RowCal[1].ToString())));
                        i++;
                    }
                }
                connRWData.Close();
                connRWData.Dispose();
                List<string> values = new List<string>();
                i = 0;
                excelDataShow.Tables.Add(new System.Data.DataTable());
                if(!excelDataShow.Tables[0].Columns.Contains("Formula Number"))
                {
                    excelDataShow.Tables[0].Columns.Add("Formula Number");
                    excelDataShow.Tables[0].Columns.Add("Description");
                    excelDataShow.Tables[0].Columns.Add("Conversion");
                    

                }
                List<DataRow> showRow = new List<DataRow>();
                // excelDataShow.Tables[0].NewRow()>;

                foreach (KeyValuePair<int, float> item in usedBatches)
                {
                    showRow.Add(excelDataShow.Tables[0].NewRow());

                    values.Add(item.Key.ToString());
                    values.Add(batchDescription[item.Key]);
                    values.Add(((int)Math.Round((double)(100 * (requiredBatches[item.Key]/item.Value)))).ToString());

                    showRow[i].ItemArray = values.ToArray();

                    excelDataShow.Tables[0].Rows.Add(showRow[i]);
                    values.RemoveRange(0, 3);
                    i++;
                }
                dataGridView1.DataSource = excelDataShow.Tables[0];
            }
            catch(Exception exception)
            {
                errorMsg.Text = "Error in Data" + exception.StackTrace;
            }
        }
    }
}

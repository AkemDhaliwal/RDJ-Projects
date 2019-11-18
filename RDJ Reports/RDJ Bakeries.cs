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
        IDictionary<int,DisplayInfo> info = new Dictionary<int,DisplayInfo>();
        
        IDictionary<int, float> usedBatches;
        IDictionary<int, float> laytimeBatches;
        IDictionary<int, float> requiredBatches;
        IDictionary<int, string> batchDescription;
        IDictionary<int, string> finalProdInfo;
        DataRow RowCal;
        DataRow RowCal1;

        DataSet excelDataSetRW = new DataSet();
        DataSet excelDataSetSpecs = new DataSet();
        DataSet excelDataDoughMade = new DataSet();
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
            laytimeBatches = new Dictionary<int, float>();
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
                    @"Z:\OPERATIONS\Analytics\RDJ_Specs.xlsx"); //@"C:\RDJ Projects\Reports\ConversionFactInfo.xlsx");
                OleDbConnection connRDJSpecs = new OleDbConnection(ConnectionStringRDJSpecs);

                connRDJSpecs.Open();

                int i = 0;
                int k = 0;
                //Go through each selected row of RW production Dataset
                foreach (DataRow row in excelDataSetRW.Tables[0].Rows)
                {
                    //if (row[2].ToString() == "BATCH" && lines.Contains(row[4].ToString(), StringComparer.OrdinalIgnoreCase))
                    //{
                    //    if (!usedBatches.ContainsKey(Int32.Parse(row[0].ToString())))   //0 intex is Item Number
                    //    {
                    //        usedBatches.Add(Int32.Parse(row[0].ToString()), (float.Parse(row[3].ToString()))); //3 index is quantity
                    //        batchDescription.Add(Int32.Parse(row[0].ToString()), (row[5].ToString()));
                    //    }
                    //    else
                    //        usedBatches[int.Parse(row[0].ToString())] = usedBatches[int.Parse(row[0].ToString())] + float.Parse(row[3].ToString());
                    //
                    //    //if(!requiredBatches.ContainsKey(int.Parse(row[0].ToString())))
                    //    //    requiredBatches.Add(int.Parse(row[0].ToString()), 0f); //3 index is quantity
                    //}
                    if (lines.Contains(row[4].ToString(), StringComparer.OrdinalIgnoreCase) && row[2].ToString() != "BATCH")    //else if (row[2].ToString() == "CASE" && lines.Contains(row[4].ToString(), StringComparer.OrdinalIgnoreCase))
                    {
                        if (info.ContainsKey(Int32.Parse(row[0].ToString())))
                        {
                            info[Int32.Parse(row[0].ToString())].finishedProdAmount = info[Int32.Parse(row[0].ToString())].finishedProdAmount + float.Parse(row[3].ToString());
                        }
                        else
                        {
                            string line = row[4].ToString();
                            switch (line)
                            {
                                case "L1":
                                    using (OleDbDataAdapter objDA3 = new System.Data.OleDb.OleDbDataAdapter("select [Product Description],[Formula Number], [Rest Time (Mins)], " +
                                        "[Batch Run Time (Mins)], [Batch Weight Kg's],[Case Weight] ,[100% Batch Yield] from[Line1$]" + " WHERE [Internal Product Number] = " + row[0], connRDJSpecs))
                                    {
                                        objDA3.Fill(excelDataSetSpecs);
                                    }
                                    break;

                                case "L2":
                                    using (OleDbDataAdapter objDA3 = new System.Data.OleDb.OleDbDataAdapter("select [Product Description],[Formula Number], [Rest Time (Mins)], " +
                                        "[Batch Run Time (Mins)], [Batch Weight Kg's],[Case Weight] ,[100% Batch Yield] from[Line2$]" + " WHERE [Internal Product Number] = " + row[0], connRDJSpecs))
                                    {
                                        objDA3.Fill(excelDataSetSpecs);
                                    }
                                    break;
                                case "L3":
                                    using (OleDbDataAdapter objDA3 = new System.Data.OleDb.OleDbDataAdapter("select [Product Description],[Formula Number], [Rest Time (Mins)], " +
                                        "[Batch Run Time (Mins)], [Batch Weight Kg's],[Case Weight] ,[100% Batch Yield] from[Line 3$]" + " WHERE [Internal Product Number] = " + row[0], connRDJSpecs))
                                    {
                                        objDA3.Fill(excelDataSetSpecs);
                                    }
                                    break;
                                case "L4":
                                    using (OleDbDataAdapter objDA3 = new System.Data.OleDb.OleDbDataAdapter("select [Product Description],[Formula Number], [Rest Time (Mins)], " +
                                        "[Batch Run Time (Mins)], [Batch Weight Kg's],[Case Weight] ,[100% Batch Yield] from[Line 4$]" + " WHERE [Internal Product Number] = " + row[0], connRDJSpecs))
                                    {
                                        objDA3.Fill(excelDataSetSpecs);
                                    }
                                    break;
                                default:
                                    break;

                            }

                            k = excelDataSetSpecs.Tables[0].Rows.Count - 1;
                            RowCal1 = excelDataSetSpecs.Tables[0].Rows[k];
                            
                            DisplayInfo displayInfo = new DisplayInfo();
                            displayInfo.productNum = Int32.Parse(row[0].ToString());
                            displayInfo.line = row[4].ToString();
                            displayInfo.prodDescription = row[5].ToString();
                            displayInfo.finishedProdAmount = float.Parse(row[3].ToString());
                            displayInfo.finProdunits = row[2].ToString();
                            displayInfo.formulaNum = int.Parse(RowCal1[1].ToString());
                            displayInfo.yield = float.Parse(RowCal1[6].ToString());
                            displayInfo.timeOnFloor = (float.Parse(RowCal1[2].ToString()) + float.Parse(RowCal1[3].ToString()));
                            if (!laytimeBatches.ContainsKey(displayInfo.formulaNum))
                                laytimeBatches.Add(displayInfo.formulaNum, displayInfo.timeOnFloor);

                            info.Add(displayInfo.productNum, displayInfo);
                        }
                    }
                    connRDJSpecs.Close();
                    connRDJSpecs.Dispose();
                    RowCal = RowCal1;
                }
                OleDbConnection connRWData1 = new OleDbConnection(ConnectionStringRWData);

                connRWData1.Open();

                foreach (KeyValuePair<int, float> item in laytimeBatches)
                {
                    RowCal[0] = item.Key;
                    startDateTime = startDateTime.AddMinutes(-item.Value);
                    endDateTime = endDateTime.AddMinutes(-item.Value);
                    using (OleDbDataAdapter objDA2 = new System.Data.OleDb.OleDbDataAdapter("select[Status],[LP Quantity],[Line] from[Sheet1$] " +
                        " WHERE [LP Reported At] BETWEEN #" + startDateTime + "# and #" + endDateTime + "#" , connRWData)) ///# [Item] = " + item.Key, connRWData))                             //
                    {

                        objDA2.Fill(excelDataDoughMade);
                    }
                    k = excelDataDoughMade.Tables[0].Rows.Count - 1;
                    RowCal1 = excelDataDoughMade.Tables[0].Rows[k];

                    if (!requiredBatches.ContainsKey(item.Key))
                        requiredBatches.Add(item.Key, float.Parse(RowCal1[1].ToString()));
                    else
                        requiredBatches[item.Key] = requiredBatches[item.Key] + float.Parse(RowCal1[1].ToString());
                }
                connRWData1.Close();
                connRWData1.Dispose();

                        //int num = Int32.Parse(row[0].ToString());
                        //if (num > 70000000)
                        //{
                        //    using (OleDbDataAdapter objDA3 = new System.Data.OleDb.OleDbDataAdapter("select[Batch Number], [Weight per Case], [Finished Kilos per batch] from[Sheet2$]" + " WHERE [Internal Product Number] = " + num, connRDJSpecs))
                        //    {
                        //        objDA3.Fill(excelDataSetSpecs2);
                        //    }
                        //    k = excelDataSetSpecs2.Tables[0].Rows.Count-1;
                        //    RowCal1 = excelDataSetSpecs2.Tables[0].Rows[k];
                        //
                        //    if (!requiredBatches.ContainsKey(int.Parse(RowCal1[0].ToString())))//2 is index for formula num in ConversionFactInfo
                        //        requiredBatches.Add(int.Parse(RowCal1[0].ToString()), 0f); //3 index is quantity
                        //
                        //    requiredBatches[int.Parse(RowCal1[0].ToString())] = requiredBatches[int.Parse(RowCal1[0].ToString())] + (((float.Parse(row[3].ToString()))* (float.Parse(RowCal1[1].ToString()))) / (float.Parse(RowCal1[2].ToString())));
                        //    
                        //}
                        //else
                        //{
                        //    using (OleDbDataAdapter objDA2 = new System.Data.OleDb.OleDbDataAdapter("select[Batch Number], [Yield] from[Sheet1$]" +
                        //        " WHERE [Internal Product Number] = " + num, connRDJSpecs))
                        //    {
                        //        objDA2.Fill(excelDataSetSpecs);
                        //    }
                        //    RowCal = excelDataSetSpecs.Tables[0].Rows[j];
                        //
                        //    if (!requiredBatches.ContainsKey(int.Parse(RowCal[0].ToString())))//2 is index for formula num in ConversionFactInfo
                        //        requiredBatches.Add(int.Parse(RowCal[0].ToString()), 0f); //3 index is quantity
                        //
                        //    requiredBatches[int.Parse(RowCal[0].ToString())] = requiredBatches[int.Parse(RowCal[0].ToString())] + ((float.Parse(row[3].ToString())) / (float.Parse(RowCal[1].ToString())));
                        //    j++;
                        //}
                    
                    //else if(row[2].ToString() == "KG" && lines.Contains(row[4].ToString(), StringComparer.OrdinalIgnoreCase))
                    //{
                    //    int num2 = Int32.Parse(row[0].ToString());
                    //    using (OleDbDataAdapter objDA3 = new System.Data.OleDb.OleDbDataAdapter("select[Batch Number], [Finished Kilos per batch] from[Sheet1$]" +
                    //        " WHERE [Internal Product Number] = " + num2, connRDJSpecs))
                    //    {
                    //        objDA3.Fill(excelDataSetSpecsAss);
                    //    }
                    //    RowCal = excelDataSetSpecsAss.Tables[0].Rows[i];
                    //
                    //    if (!requiredBatches.ContainsKey(int.Parse(RowCal[0].ToString())))//2 is index for formula num in ConversionFactInfo
                    //        requiredBatches.Add(int.Parse(RowCal[0].ToString()), 0f); //3 index is quantity
                    //
                    //    requiredBatches[int.Parse(RowCal[0].ToString())] = requiredBatches[int.Parse(RowCal[0].ToString())] + ((float.Parse(row[3].ToString())) / (float.Parse(RowCal[1].ToString())));
                    //    i++;
                    //}
                

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

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
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace RDJ_Reports
{
    public partial class Form1 : Form
    {
        Excel rwData;

        public DateTime startDateTime;
        public DateTime endDateTime;
        public ReadRDJSpecsInfo info;
        private int item = 2;
        private int status = 5;
        private int caseORbatch = 8;
        private int quantity = 10;
        private int date = 11;
        private int Line = 12;
        List<int> batchList = new List<int>();
        List<int> prodList = new List<int>();
        IDictionary<int, float> dataBatch = new Dictionary<int, float>();
        IDictionary<int, float> dataProduct = new Dictionary<int, float>();

        public Form1()
        {
            InitializeComponent();
            startTime.CustomFormat = "hh:mm";
            endTime.CustomFormat = "hh:mm";

            info = new ReadRDJSpecsInfo();
            
            foreach(RDJSpecInfo data in info.line1)
            {
                if(!batchList.Contains(data.formulaNum))
                    batchList.Add(data.formulaNum);
                if(!prodList.Contains(data.productNum))
                    prodList.Add(data.productNum);
            }

            rwData = new Excel(@"C:\RDJ Projects\RDJ-Projects\Reports\RDJ RW MO Production.xlsx", 1);
        }
      
        private void genConvReport_Click(object sender, EventArgs e)
        {
            startDateTime = startDate.Value.Date + startTime.Value.TimeOfDay;
            endDateTime = endDate.Value.Date + endTime.Value.TimeOfDay;

            int i = 5;
            bool count = true;
            while(count)
            {
                if (DateTime.Compare((DateTime)(rwData.Read(i, date)), startDateTime) < 0)
                    count = false;
                //rwData.stringRead(i, status) == "Closed" && 
                if (rwData.stringRead(i, Line) == "L1" && rwData.Read(i, date) < endDateTime)
                {
                    int test = 0;
                    test = rwData.intRead(i, item);

                    if (batchList.Contains(rwData.intRead(i, item)))
                    {
                        string val = "";
                        val = rwData.stringRead(i, status);
                        if (rwData.stringRead(i, status) == "Closed")
                        {
                            if (dataBatch.ContainsKey(rwData.intRead(i, item)))
                                dataBatch[rwData.intRead(i, item)] = dataBatch[rwData.intRead(i, item)] + rwData.floatRead(i, quantity);
                            else
                                dataBatch.Add(rwData.intRead(i, item), rwData.floatRead(i, quantity));
                        }
                    }else
                    {
                        if (dataProduct.ContainsKey(rwData.intRead(i, item)))
                            dataProduct[rwData.intRead(i, item)] = dataProduct[rwData.intRead(i, item)] + rwData.floatRead(i, quantity);
                        else
                            dataProduct.Add(rwData.intRead(i, item), rwData.floatRead(i, quantity));
                    }
                }
                i++;
            }

            
            float conversion = 0;
            float batchCases = 0;
            string result = "";
            foreach (KeyValuePair<int, float> item in dataBatch)
            {
                batchCases = 0;
                conversion = 0;
                foreach (RDJSpecInfo specinfo in info.line1)
                {
                    if(specinfo.formulaNum == item.Key && dataProduct.ContainsKey(specinfo.productNum))
                    {
                        batchCases = specinfo.casesInBatc;
                        conversion = (dataProduct[specinfo.productNum]/ (batchCases*item.Value))*100;

                        result += "Product Number   " + item.Key + "|| Conversion   " + conversion + "% \n";
                    }
                }
            }
            /*for (int i = 5; i < 15; i++)
            {
                result += "\n" + rwData.Read(i, 11).ToString("yyyyMMdd HHmm");
            }*/
            MessageBox.Show(result);
            rwData.Close_File();

        }
    }
}

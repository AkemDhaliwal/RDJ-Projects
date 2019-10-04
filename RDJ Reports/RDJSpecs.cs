using System.Collections.Generic;
using System.Web.UI.MobileControls;

namespace RDJ_Reports
{
    public class RDJSpecInfo
    {
        public int productNum;
        public int formulaNum;
        public float casesInBatc;
        public string prodDescription;
    }
    public class ReadRDJSpecsInfo
    {
        public List<RDJSpecInfo> line1 = new List<RDJSpecInfo>();
        private RDJSpecInfo info;
        private int row = 2;
        private int prodNumCol = 1;
        private int prodDesCol = 2;
        private int formulaNumCol = 3;
        private int casesCol = 35;

        private Excel rdjSpecs;

        public ReadRDJSpecsInfo()
        {
            rdjSpecs = new Excel(@"C:\RDJ Projects\RDJ-Projects\Reports\RDJ_Specs_test.xlsx", 1);

            for(row = 2; row < 219; row++)
            {
                info = new RDJSpecInfo();
                info.productNum = rdjSpecs.intRead(row, prodNumCol);
                info.formulaNum = rdjSpecs.intRead(row, formulaNumCol);
                info.casesInBatc = rdjSpecs.floatRead(row, casesCol);
                info.prodDescription = rdjSpecs.stringRead(row, prodDesCol);

                line1.Add(info);
            }
            rdjSpecs.Close_File();
        }

    }
}
using System.Collections.Generic;
using System.Web.UI.MobileControls;

namespace RDJ_Reports
{
    public class RDJSpecInfo
    {
        public int formulaNum;
        public string productNum;
        public float conversion;
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

        public ReadRDJSpecsInfo()
        {

        }

    }
}
using System.IO;
using System.Xml;
using OfficeOpenXml; //https://github.com/JanKallman/EPPlus

namespace XMLParse1
{
    class ApttusExcelEncode
    {
        private struct RecurseDetails
        {
            public int apttusRow; //addressing is 1-count!
            public int excelRow;
            public int level;
            public ExcelWorksheet ws; //remove me?

            public RecurseDetails(ExcelWorksheet init_ws)
            {
                this.ws = init_ws;
                this.apttusRow = 0;
                this.excelRow = 0;
                this.level = -1;
            }
        }
        /// <summary>
        /// Reads the sourceXML and transforms it to Excel.
        /// </summary>
        public static bool transformXMLtoExcel(string sourceXML, string outName)
        {
            bool runResult = false;

            XmlDocument doc = new XmlDocument();
            doc.Load(sourceXML);

            //Creates a blank workbook. Use the using statment, so the package is disposed when we are done.
            using (ExcelPackage p = new ExcelPackage())
            {
                //A workbook must have at least on cell, so lets add one... 
                ExcelWorksheet ws = p.Workbook.Worksheets.Add("Sheet1");

                //initialize
                RecurseDetails details = new RecurseDetails(ws);
                //Create the header
                details.ws.Cells[1, 1].Value = "category hierarchy: name";
                details.ws.Cells[1, 2].Value = "category hierarchy: id";
                details.ws.Cells[1, 3].Value = "level";
                details.ws.Cells[1, 4].Value = "left";
                details.ws.Cells[1, 5].Value = "right";

                details = travelXML(details, doc.ChildNodes);
                //recurse code here
                //travelXML(details, doc.ChildNodes);

                //Save the new workbook. We haven't specified the filename so use the Save as method.
                p.SaveAs(new FileInfo(outName));
            }

            return runResult;
        }

        private static RecurseDetails travelXML(RecurseDetails details, XmlNodeList nodes)
        {
            int currApttuscount = details.apttusRow + 1;
            int currExcelRow = -1;
            string ID, name;
            int left = currApttuscount;
            int level = details.level;

            if (nodes.Count > 0)
            {
                level = ++details.level;
            }

            foreach (XmlNode node in nodes)
            {
                currExcelRow = ++details.excelRow;
                currApttuscount = ++details.apttusRow;
                name = node.Attributes["name"].InnerText; //Needs error handling!
                ID = node.Attributes["ID"].InnerText; //Needs error handling!

                details.ws.Cells[currExcelRow + 1, 1].Value = name;
                details.ws.Cells[currExcelRow + 1, 2].Value = ID;
                details.ws.Cells[currExcelRow + 1, 3].Value = level;
                details.ws.Cells[currExcelRow + 1, 4].Value = currApttuscount; //left

                details = travelXML(details, node.ChildNodes);
                details.level = level;
                details.ws.Cells[currExcelRow + 1, 5].Value = ++details.apttusRow; //right
            }

            return details;
        }
    }
}

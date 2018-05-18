using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Diagnostics;
using OfficeOpenXml; //https://github.com/JanKallman/EPPlus

namespace XMLParse1
{
    class ApttusExcelEncode
    {
        private struct RecurseDetails
        {
            public int apttusRow; //addressing is 1-count!
            public int excelRow;
            public ExcelWorksheet ws;
            public List<headerCell> header;

            public RecurseDetails(ExcelWorksheet init_ws, List<headerCell> header)
            {
                this.ws = init_ws;
                this.apttusRow = 0;
                this.excelRow = 1;
                this.header = header;
            }
        }

        private struct headerCell
        {
            public int col;
            public string uid, text;

            public headerCell(int col, string uid, string text)
            {
                this.col = col;
                this.uid = uid;
                this.text = text;
            }
        }

        private static List<headerCell> getHeader(XmlNode doc)
        {
            List<headerCell> header = new List<headerCell>();
            int i = 1;
            foreach (XmlNode node in doc["header"].ChildNodes)
            {
                header.Add(new headerCell(i++, node.Name, node.InnerXml));
            }
            header.Add(new headerCell(i++, "h_left", "Left"));
            header.Add(new headerCell(i++, "h_right", "Right"));
            header.Add(new headerCell(i, "h_level", "Level"));

            return header;
        }

        /// <summary>
        /// Reads the sourceXML and transforms it to Excel.
        /// </summary>
        public static bool transformXMLtoExcel(string sourceXML, string outName)
        {
            bool runResult = false;
            XmlDocument xmlFile = new XmlDocument();
            XmlNode doc = new XmlDocument();
            List<headerCell> header;

            xmlFile.Load(sourceXML);
            doc = xmlFile.DocumentElement;
            header = getHeader(doc);

            //Creates a blank workbook. Use the using statment, so the package is disposed when we are done.
            using (ExcelPackage p = new ExcelPackage())
            {
                //A workbook must have at least on sheet, so lets add one... 
                ExcelWorksheet ws = p.Workbook.Worksheets.Add("Sheet1");

                //Create the header
                foreach (headerCell cell in header) ws.Cells[1, cell.col].Value = cell.text;
                RecurseDetails details = new RecurseDetails(ws, header);

                //recurse code here
                details = travelXML(ref details, doc["cat"]);

                //Save the new workbook. We haven't specified the filename so use the Save as method.
                p.SaveAs(new FileInfo(outName));
            }

            return runResult;
        }

        /// <summary>
        /// Recurses through the XML hiearchy, writing the contents to Excel.
        /// </summary>
        //"details" could be passed by value but "ref" saves memory.
        private static RecurseDetails travelXML(ref RecurseDetails details, XmlNode cat, int level = 0)
        {
            int currApttuscount = details.apttusRow + 1;
            int currExcelRow;
            //string ID, name;
            int left = currApttuscount;
            int i = 1;

            currExcelRow = ++details.excelRow;
            currApttuscount = ++details.apttusRow;

            foreach (XmlNode node in cat["values"].ChildNodes) details.ws.Cells[currExcelRow, i++].Value = node.InnerText;
            details.ws.Cells[currExcelRow, i++].Value = currApttuscount; // left

            if (cat["nest"] != null)
            {
                foreach (XmlNode nestedCat in cat["nest"].ChildNodes) details = travelXML(ref details, nestedCat, level + 1);
            }

            details.ws.Cells[currExcelRow, i++].Value = ++details.apttusRow; // right
            details.ws.Cells[currExcelRow, i++].Value = level.ToString(); // level

            return details;
        }
    }
}

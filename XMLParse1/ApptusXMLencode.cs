using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using Excel; //https://www.codeproject.com/Tips/801032/Csharp-How-To-Read-xlsx-Excel-File-With-Lines-of

namespace XMLParse1
{
    class ApptusXMLencode
    {
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

        /// <summary>
        /// Get the header the excel file.
        /// </summary>
        static private List<headerCell> getSheetHeader(worksheet targetSheet)
        {
            List<headerCell> header = new List<headerCell>();
            int i = 1;
            foreach (Cell cell in targetSheet.Rows[0].Cells)
            {
                if (!cell.Text.ToLower().Contains("left") && !cell.Text.ToLower().Contains("right") && !cell.Text.ToLower().Contains("level"))
                {
                    header.Add(new headerCell(cell.ColumnIndex, "h" + i++, cell.Text));
                }
            }
            return header;
        }

        /// <summary>
        /// Opens an Excel file and writes an XML file based on it's contents.
        /// </summary>
        /// Specifically, this function focuses on opening the file and checking it for the "level" column.
        /// Encoding(writing) is passed to a sub-function.
        static public bool transformExceltoXML(string sourceExcel,string outFile)
        {
            bool runResult = false;
            int levelCol = -1;
            IEnumerable<worksheet> targetExcel = Workbook.Worksheets(sourceExcel);
            worksheet targetSheet;

            //Get the first sheet. This program cannot handle more than one sheet.
            var e = targetExcel.FirstOrDefault();
            if (e != null)
            {
                targetSheet = e;
                //Get the column that contains "level"
                foreach (Cell cell in targetSheet.Rows[0].Cells) if (cell.Text.ToLower() == "level") levelCol = cell.ColumnIndex;

                if (levelCol == -1) MessageBox.Show("Oh no! it didn't find the level column! Cannot continue operation!");
                //Error checking complete. Crunk! Pull the lever!
                else runResult = writeTheXml(targetSheet, outFile, levelCol);
            }
            return runResult;
        }

        /// <summary>
        /// Creates an XML file from a SalesForce export.
        /// </summary>
        static private bool writeTheXml(worksheet targetSheet, string outFile, int levelCol)
        {
            bool runResult = false;
            int prevLevel = -1, levelsToClose;
            int currLevel = 0;
            List<headerCell> excelHeader;

            excelHeader = getSheetHeader(targetSheet);

            XmlWriterSettings writerSettings = new XmlWriterSettings();
            writerSettings.OmitXmlDeclaration = true;
            writerSettings.Indent = true;
            writerSettings.IndentChars = "    ";
            writerSettings.ConformanceLevel = ConformanceLevel.Document;
            writerSettings.CloseOutput = true;

            //This is the meat and potatoes
            using (XmlWriter writer = XmlWriter.Create(outFile, writerSettings))
            {
                //Setup the "table" node
                writer.WriteStartDocument(); //This isn't required with the current XmlWriterSettings
                writer.WriteStartElement("table");

                //Setup the "header" node
                writer.WriteStartElement("header");
                foreach (headerCell cell in excelHeader) writer.WriteElementString(cell.uid, cell.text);
                writer.WriteEndElement();

                writer.WriteStartElement("hierarchy");
                foreach (Row row in targetSheet.Rows.Skip(1)) //PRIMARY LOOP //Skip the first row (header)
                {
                    currLevel = Convert.ToInt32(row.Cells[levelCol].Text);

                    //only close element(s) if currLevel <= prevLevel
                    levelsToClose = prevLevel;
                    for (; currLevel <= levelsToClose; levelsToClose--) writer.WriteEndElement();

                    writer.WriteStartElement("cat");
                    //write the data of the row
                    foreach (headerCell cell in excelHeader) writer.WriteElementString(cell.uid, row.Cells[cell.col].Text);

                    prevLevel = currLevel;
                }// loop to next row

                //cleanup
                writer.WriteEndElement(); //hierarchy
                writer.WriteEndElement(); //table
                writer.WriteEndDocument();
                writer.Flush(); //clear the buffer
                writer.Close(); //free the allocation, close the file stream
                runResult = true;
            }
            return runResult;
        }
    }
}
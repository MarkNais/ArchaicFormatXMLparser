using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using Excel; //https://www.codeproject.com/Tips/801032/Csharp-How-To-Read-xlsx-Excel-File-With-Lines-of

namespace XMLParse1
{
    class ApptusXMLencode
    {
        /// <summary>
        /// Storage of the key excel columns. Init with -1.
        /// </summary>
        public struct ApptusHeaderIndexes
        {
            public int Level, Name, ID, Left, Right;
            public bool valid;

            public ApptusHeaderIndexes(int init)
            {
                this.Level = init;
                this.Name = init;
                this.ID = init;
                this.Left = init;
                this.Right = init;
                this.valid = false;
            }
        }

        /// <summary>
        /// Finds the key columns from the excel file.
        /// </summary>
        static public ApptusHeaderIndexes GatherApptusIndexes(string sourceExcel)
        {
            ApptusHeaderIndexes Indexes = new ApptusHeaderIndexes(-1);
            IEnumerable<worksheet> targetExcel = Workbook.Worksheets(sourceExcel);
            worksheet targetSheet = null;

            //Get the first sheet. This program cannot handle more than one sheet.
            var e = targetExcel.FirstOrDefault();
            if (e != null)
            {
                targetSheet = e;

                //foreach (worksheet targetSheet in targetExcel)
                if (targetSheet != null)
                {
                    foreach (Cell cell in targetSheet.Rows[0].Cells)
                    {
                        switch (cell.Text.ToLower())
                        {
                            case "level":
                                Indexes.Level = cell.ColumnIndex;
                                break;

                            case "category hierarchy: name":
                                Indexes.Name = cell.ColumnIndex;
                                break;

                            case "category hierarchy: id":
                                Indexes.ID = cell.ColumnIndex;
                                break;

                            case "left":
                                Indexes.Left = cell.ColumnIndex;
                                break;

                            case "right":
                                Indexes.Right = cell.ColumnIndex;
                                break;

                            default:
                                break;
                        }
                    }//for each cell
                }

                Indexes.valid = (
                    Indexes.Level > -1 &&
                    Indexes.Name > -1 &&
                    Indexes.ID > -1);
                    //Indexes.Left > -1 && //not required
                    //Indexes.Right > -1); //not required
            }
            return Indexes;
        }

        /// <summary>
        /// Creates an XML file from a SalesForce export.
        /// </summary>
        static public bool transformExceltoXML(string sourceExcel,string outFile, ApptusHeaderIndexes Indexes, bool verboseXML, bool verboseAttributes)
        {
            bool runResult = false;
            int prevLevel = 0;
            int currLevel = 0;
            int left = -1, right = -1;
            string name, ID;
            bool firstRowPassed = false, secondRowPassed = false;
            IEnumerable<worksheet> targetExcel = Workbook.Worksheets(sourceExcel);
            worksheet targetSheet = null;

            //Get the first sheet. This program cannot handle more than one sheet.
            var e = targetExcel.FirstOrDefault();
            if (e != null)
            {
                targetSheet = e;

                XmlWriterSettings writerSettings = new XmlWriterSettings();
                writerSettings.OmitXmlDeclaration = true;
                writerSettings.Indent = true;
                writerSettings.IndentChars = "    ";
                writerSettings.ConformanceLevel = ConformanceLevel.Document;
                writerSettings.CloseOutput = true;

                //This is the meat and potatoes
                using (XmlWriter writer = XmlWriter.Create(outFile, writerSettings))
                {
                    foreach(Row row in targetSheet.Rows) //PRIMARY LOOP
                    {
                        if (!firstRowPassed) //Skip the first row (header)
                        {
                            firstRowPassed = true;
                            continue;
                        }

                        //Collect this row's data from Excel
                        currLevel = Convert.ToInt32(row.Cells[Indexes.Level].Text);
                        name = row.Cells[Indexes.Name].Text;
                        ID = row.Cells[Indexes.ID].Text;
                        if(Indexes.Left > -1) left = Convert.ToInt32(row.Cells[Indexes.Left].Text);
                        if (Indexes.Right > -1) right = Convert.ToInt32(row.Cells[Indexes.Right].Text);
                        
                        //Skip closes for the second row (first row of data, no closes needed)
                        if (!secondRowPassed) secondRowPassed = true; 
                        else //only close element(s) if currLevel <= prevLevel
                        {
                            for (; currLevel <= prevLevel; prevLevel--) //prevLevel is not reused beyond this line and may be modified
                            {
                                if (verboseXML) writer.WriteRaw("\n"); //WriteRaw breaks the auto-indentation of the XML processor
                                writer.WriteEndElement();
                            }
                            if (verboseXML) writer.WriteRaw("\n");
                        }

                        //Start and populate this element
                        writer.WriteStartElement("cat");
                        writer.WriteAttributeString("name", name);
                        writer.WriteAttributeString("ID", ID);
                        if (verboseAttributes)
                        {
                            writer.WriteAttributeString("Level", Convert.ToString(currLevel));
                            if (Indexes.Left > -1) writer.WriteAttributeString("Left", Convert.ToString(left));
                            if (Indexes.Right > -1) writer.WriteAttributeString("Right", Convert.ToString(right));
                        }

                        prevLevel = currLevel;
                    }// loop to next row

                    writer.Flush(); //clear the buffer
                    writer.Close(); //free the allocation, close the file stream
                    runResult = true;
                }
            }
            return runResult;
        }

    }
    
}
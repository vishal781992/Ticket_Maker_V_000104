using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ExcelReadingApp
{
    class XMLParser: RootDirectoriesExplorer
    {
        //XmlDocument doc = new XmlDocument();
        public List<string> demoData = new List<string>();
        public List<string> ColumnValue = new List<string>();
        public List<string> AlternateColumnValue = new List<string>();
        public bool Flag_UseXMLLoadDataFun = true;

        public XMLParser()
        {
            XmlDocument doc = new XmlDocument();
        }

        public void xmlLoadData(string filename)
        {
            if (Flag_UseXMLLoadDataFun)
            {
                XmlTextReader reader = new XmlTextReader(filename);
                try
                {
                    while (reader.Read())
                    {
                        switch (reader.NodeType)
                        {
                            case XmlNodeType.Element: // The node is an element.
                                if (string.Equals(reader.Name, "ColumnName"))//reader.Name == "ColumnName"
                                {
                                    reader.Read();
                                    ColumnValue.Add(reader.Value);
                                }
                                if(string.Equals(reader.Name, "Heading"))
                                {
                                    reader.Read();
                                    try { AlternateColumnValue.Add(reader.Value); } catch { }
                                    
                                }
                                break;

                            case XmlNodeType.Text: //Display the text in each element.
                                break;

                            case XmlNodeType.EndElement: //Display the end of the element.
                                break;

                            default:
                                break;
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("The Problem occured in the xmlLoadData function\r\n" + e);
                }
            }
           
        }

        public string XMLRequestData(string ToSearch, [OptionalAttribute] string pathOfFile)  //helps finding the company name in the xml file
        {
            
            if (string.IsNullOrEmpty(pathOfFile)){ RootDirectoriesExplorer RE = new RootDirectoriesExplorer(); pathOfFile = RE.FilePathOfXMLtemp; }
                
            string returnString=string.Empty;
            XmlTextReader reader = new XmlTextReader(pathOfFile);
            while (reader.Read())
            {
                if (reader.Name == "ColumnName")
                {
                    //demoData.Add("<" + reader.Name + ">");
                    reader.Read();
                    if(reader.Value == ToSearch)
                    {
                        do { reader.Read(); } while (reader.Name != "Text");
                        //reader.Read();
                        if(reader.Name == "Text")// reader.Read();
                        {
                            reader.Read();
                            returnString = reader.Value;
                        }
                    }
                }
            }

            return returnString;
        }

        public void XMLCreator(string path,string CompanyName,string textBox_FolderName)
        {
            Directory.CreateDirectory(path+ textBox_FolderName);//CompanyName before
            string tempPathCombined = path + "\\" + textBox_FolderName + "\\" + CompanyName + ".xml";
            using (XmlWriter writer = XmlWriter.Create(tempPathCombined))
            {
                //writer.WriteString("\r\n");
                writer.WriteStartElement("ExcelLayoutList"); writer.WriteString("\r\n");
                for (int counter = 0; counter < ColumnValue.Count; counter++)
                {
                    writer.WriteStartElement("ExcelLayoutItem"); writer.WriteString("\r\n");
                    writer.WriteElementString("ColumnName", ColumnValue[counter]); writer.WriteString("\r\n");
                    writer.WriteElementString("Heading", string.Empty); writer.WriteString("\r\n");
                    if (string.Equals(ColumnValue[counter], "Company"))
                    {
                        writer.WriteElementString("Text", CompanyName); writer.WriteString("\r\n");
                    }
                    else
                    {
                        writer.WriteElementString("Text", string.Empty); writer.WriteString("\r\n");
                    }
                   
                    //writer.WriteElementString("Text", string.Empty); writer.WriteString("\r\n");
                    writer.WriteEndElement();

                }
                writer.WriteEndElement();
                writer.Flush();
                writer.Close();
            }
        }

        public int XMLCreatorFromExcel(string OutputFileCompletepath,string outputDir, List<string> Dataset, List<string> DatasetRow2)
        {
            if(!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);//CompanyName before
            try
            {
                if(!File.Exists(OutputFileCompletepath))
                {
                    using (XmlWriter writer = XmlWriter.Create(OutputFileCompletepath))
                    {
                        writer.WriteStartElement("ExcelLayoutList"); writer.WriteString("\r\n");
                        for (int counter = 0; counter < Dataset.Count; counter++)
                        {
                            writer.WriteStartElement("ExcelLayoutItem"); writer.WriteString("\r\n");
                            writer.WriteElementString("ColumnName", Dataset[counter]); writer.WriteString("\r\n");
                            writer.WriteElementString("Heading", string.Empty); writer.WriteString("\r\n");
                            writer.WriteElementString("Text", DatasetRow2[counter]); writer.WriteString("\r\n");
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();
                        writer.Flush();
                        writer.Close();
                    }
                    return 1;
                }
                return 0;
            }
            catch { return 0; }
        }
    }

    class FormatModifier : XMLParser
    {
        public string FormatString { get; set; }

        public FormatModifier() { }
        public FormatModifier(string formatString)
        {
            this.FormatString = formatString;
        }

        public void FormatParser()
        {
            try
            {
                if (!string.IsNullOrEmpty(FormatString) || !string.IsNullOrWhiteSpace(FormatString))
                {
                    int StartIndex = 0, StopIndex = 0; bool LocalFlag_ForLoop = true;

                    string TempString = FormatString.Replace(", ", ",");
                    TempString = TempString.Replace("\t", ",");
                    TempString = TempString.Replace(" ", ",");
                    TempString = TempString.Replace("\n", string.Empty);
                    int LenghtOfTempString = TempString.Length;

                    while (LocalFlag_ForLoop) //working on this function 
                    {
                        try
                        {
                            try { StopIndex = TempString.IndexOf(','); if (StopIndex == -1 || StopIndex < 0) { StopIndex = TempString.Length; LocalFlag_ForLoop = false; ColumnValue.Add(TempString); break; } }
                            catch { StopIndex = TempString.Length; LocalFlag_ForLoop = false; }

                            ColumnValue.Add(TempString.Substring(StartIndex, StopIndex));
                            TempString = TempString.Substring(StopIndex + 1, TempString.Length - (StopIndex + 1));
                            StartIndex = 0;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error in the formatParser function\r\nStopIndex < 0 can be a error\r\n" + ex);
                            Flag_UseXMLLoadDataFun = true;
                        }
                    }
                    Flag_UseXMLLoadDataFun = false;//setting the flag to false, as we dont need to set the format any more.
                }
            }
            catch
            {
                MessageBox.Show("Invalid format. See documentation!");
            }
        }
    }
}

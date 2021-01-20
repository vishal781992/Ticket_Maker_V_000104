//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace ExcelReadingApp
//{
//    class ExcelLayoutmanager
//    {
//    }
//}
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ExcelReadingApp
{
    public class ExcelLayoutItem
    {
        #region Properties

        public string ColumnName { get; set; }
        public string Heading { get; set; }
        public string Text { get; set; }

        public string TextInTheFields { get; set; }

        #endregion Properties

        #region Constructors

        public ExcelLayoutItem()
        {
            this.ColumnName = string.Empty;
            this.Heading = string.Empty;
            this.Text = string.Empty;
            this.TextInTheFields = string.Empty;
        }

        public ExcelLayoutItem(string columnName, string heading, string text)//, string TextInTheFields
        {
            this.ColumnName = columnName;
            this.Heading = heading;
            this.Text = text;
            //this.TextInTheFields = TextInTheFields;
        }

        #endregion Constructors

        #region Methods

        #region Copy

        public static ExcelLayoutItem Copy(ExcelLayoutItem source)
        {
            ExcelLayoutItem item = new ExcelLayoutItem();

            item.ColumnName = source.ColumnName;
            item.Heading = source.Heading;
            item.Text = source.Text;

            return item;
        }

        #endregion Copy

        #endregion Methods
    }

    [Serializable()]
    public class ExcelLayout
    {
        #region Properties

        // Serializes an ArrayList as a "ExcelLayoutList" array of XML elements of type named "ExcelLayoutItem".
        [XmlArray("ExcelLayoutList"), XmlArrayItem("ExcelLayoutItem", typeof(ExcelLayoutItem))]
        public List<ExcelLayoutItem> ExcelLayoutList { get; set; }

        #endregion Properties

        #region Constructors

        public ExcelLayout()
        {
            this.ExcelLayoutList = new List<ExcelLayoutItem>();
        }

        #endregion Constructors
    }

    public static class ExcelLayoutManager
    {
        #region Variables

        public static string excelLayoutFileName = "ExcelLayout.xml";

        #endregion Variables

        #region Methods

        #region Initialize

        public static ExcelLayout Initialize(List<string> Dataset, List<string> DatasetRow2)//string text1,string text2,string text3,
        {
            ExcelLayout layout = new ExcelLayout();
            for(int counter=0;counter<Dataset.Count;counter++)
            {
                layout.ExcelLayoutList.Add(new ExcelLayoutItem(Dataset[counter],string.Empty, DatasetRow2[counter]));
            }
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem(text1, text2, text3));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem(string.Empty, "PO#", "1053454-1234"));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("Batch", "Batch", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("FirmwareRevision", "FirmwareRevision", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("StatusCode", "StatusCode", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("MeterID", "MeterID", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("KwhUsage", "KwhUsage", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("ManufacturerType", "ManufacturerType", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("MeterTypeCode", "MeterTypeCode", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("ClassAmps", "ClassAmps", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("Form", "Form", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("ALSF", "AL SF", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("ALSL", "AL SL", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("ALSP", "AL SP", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("ALWA", "AL WA", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("Box", "Box", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem("Pallet", "Pallet", string.Empty));
            //layout.ExcelLayoutList.Add(new ExcelLayoutItem(string.Empty, "Comments", string.Empty));

            return layout;
        }

        #endregion Initialize

        #region Open

        #endregion Open

        #region Save

        public static void Save(string fileName, ExcelLayout layout, bool showMessage)
        {
            //string fileName = Path.Combine(folder, excelLayoutFileName);

            // save object to XML file using our ObjectXMLSerializer class...
            try
            {
                ObjectXMLSerializer<ExcelLayout>.Save(layout, fileName);

                //VestaDLL.Utilities.GrantAccess(fileName);

                if (showMessage)
                {
                    Console.WriteLine("Excel layout saved to: '" + fileName + "'");
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine("Unable to save mappings to: '" + fileName + "'");
                Console.WriteLine("Message=" + ex.Message);
                Console.WriteLine("StackTrace=" + ex.StackTrace);
                Console.WriteLine("Source=" + ex.Source);
            }
        }

        #endregion Save

        #region Load

        public static ExcelLayout Load(string fileName)
        {
            ExcelLayout layout = new ExcelLayout();
            //string fileName = Path.Combine(folders.VestaFolder, mapFileName);

            if (!File.Exists(fileName))
            {
                //layout = ExcelLayoutManager.Initialize();
                //ExcelLayoutManager.Save(fileName, layout, false);

                Console.WriteLine("Layout file is not present: '" + fileName + "'");

                return null;
            }

            try
            {
                // Load the mapping object from the XML file using our custom class...
                layout = ObjectXMLSerializer<ExcelLayout>.Load(fileName);

                if (layout == null)
                {
                    //layout = ExcelLayoutManager.Initialize();
                    //ExcelLayoutManager.Save(fileName, layout, false);

                    Console.WriteLine("Layout file read returned null: '" + fileName + "'");

                    return null;
                }

                return layout;
            }

            catch (Exception e)
            {
                Console.WriteLine("Unable to load mappings from file: '" + fileName + "'");
                Console.WriteLine("Message=" + e.Message);
                Console.WriteLine("StackTrace=" + e.StackTrace);
                Console.WriteLine("Source=" + e.Source);

                return null;
            }
        }

        #endregion Load

        #region Find Meter Type Code

        //public static string FindMeterTypeCode(List<CarrollMapItem> list, string code)
        //{
        //    var pair = list.Find(p => p.Code == code);

        //    return (pair != null) ? pair.MeterTypeCode : string.Empty;
        //}

        #endregion Find Meter Type Code

        #endregion Methods
    }
}
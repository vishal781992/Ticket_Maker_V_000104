using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
//using System.Windows.Forms;
//using Microsoft.VisualBasic.FileIO;
using System.IO;
using System.Diagnostics;
namespace ExcelReadingApp
{
    class RootDirectoriesExplorer
    {
        string root = @"M:\_ShipmentFiles";//@"\Netserver3\_ShipmentFiles";
        public List<string> DirNames = new List<string>();
        public List<string> FileNames = new List<string>();
        public List<string> FileDirecrtory = new List<string>();
        string strFilename1,
               FileFullPath,
               FileName;
        //Form1 F1 = new Form1();
        public void DirectoriesExplorer()
        {
            // Get a list of all subdirectories
            Form1 F1 = new Form1();
            var dirs = from dir in
                Directory.EnumerateDirectories(root)
                       select dir;
            F1.DisplayText("test", 2);
            F1.DisplayText("Subdirectories: "+ dirs.Count<string>().ToString(),2);//1 for append
            F1.DisplayText("List of Subdirectories",2);
            foreach (var dir in dirs)
            {
                //F1.DisplayText(""+ dir.Substring(dir.LastIndexOf("\\") + 1),2);
                DirNames.Add(dir.Substring(dir.LastIndexOf("\\") + 1));
                
                DateTime lastupdated = DateTime.MinValue;
                strFilename1 = root+"\\" + dir.Substring(dir.LastIndexOf("\\") + 1);
                string Folder = strFilename1;

                var files = new DirectoryInfo(Folder).GetFiles("*.xls");//*.*

                foreach (FileInfo file in files)
                {
                    if (file.LastWriteTime > lastupdated)
                    {
                        lastupdated = file.LastWriteTime;
                        FileFullPath = file.FullName;
                        FileName = file.Name;
                    }
                }
                FileNames.Add(FileName);
                FileDirecrtory.Add(FileFullPath);

            }

            // Get a list of all subdirectories starting with 'Ma'  
            //var MaDirs = from dir in
            //    Directory.EnumerateDirectories(root, "Ma*")
            //             select dir;
            //F1.DisplayText("Subdirectories: "+ MaDirs.Count<string>().ToString(),2);
            //F1.DisplayText("List of Subdirectories",1);
            //foreach (var dir in MaDirs)
            //{
            //    F1.DisplayText(""+ dir.Substring(dir.LastIndexOf("\\") + 1),1);
            //}
        }
    }
}

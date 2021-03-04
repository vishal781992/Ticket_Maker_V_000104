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
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using System.Security.Cryptography.X509Certificates;

namespace ExcelReadingApp
{
    class RootDirectoriesExplorer
    {
        #region Declaration

        DeclarationClass DC = new DeclarationClass();

        public List<string> DirNames = new List<string>();
        public List<string> FileNames = new List<string>();
        public List<string> FileDirecrtory = new List<string>();
        public List<string> FilenamesForSearch = new List<string>();
        public List<string> TempFileNames = new List<string>();
        public List<string> TempFileDirecrtory = new List<string>();

        public List<int> IndexOF = new List<int>();

        public string DirectoryFilename { get; set; }
        public string FileFullPath {get;set;}
        public string FileNameOnly { get; set; }
        public string XMLFileName { get; set; }
        public string XMLFileFullPath { get; set; }
        public string CompanyName{ get; set; }

        public string   FilePathOfXMLtemp,
                        ExportXlSXfilePath;

        #endregion Declaration

        #region Constructor Empty inputs
        public RootDirectoriesExplorer(){}
        #endregion Constructor Empty inputs

        #region GET Directory
        public void DirectoriesExplorer([Optional] string rootD, [Optional] string formatTolookUP)
        {
            if (string.IsNullOrEmpty(rootD))
                rootD = DeclarationClass.ROOTDIRFORXMLFILES;
            if (string.IsNullOrEmpty(formatTolookUP))
                formatTolookUP = "*.xml";
            // Get a list of all subdirectories
            try
            {
                var dirs = from dir in
               Directory.EnumerateDirectories(rootD)
                           select dir;
                
                foreach (var dir in dirs)
                {
                    this.DirNames.Add(dir.Substring(dir.LastIndexOf("\\") + 1));

                    DateTime lastupdated = DateTime.MinValue;
                    DirectoryFilename = rootD + "\\" + dir.Substring(dir.LastIndexOf("\\") + 1);
                    string Folder = DirectoryFilename;

                    var files = new DirectoryInfo(Folder).GetFiles(formatTolookUP);//*.*//"*.xml"
                }
            }
            catch { }
           
        }
        #endregion GET Directory

        #region CompanyFindfromUserInput
        public void CompanyFinder(string CompanyName)
        {
            IEnumerable<string> matchingList;
            FilenamesForSearch.Clear();
            try
            {

                if (!string.IsNullOrEmpty(CompanyName))
                {
                    matchingList = DirNames.Where(x => x.ToUpper().Contains(CompanyName.ToUpper()));
                    if (matchingList != null)
                    {
                        FilenamesForSearch = matchingList.ToList();
                    }
                }
            }
            //catch { }
            catch (Exception ex){ MessageBox.Show(ex+string.Empty); }
        }
        #endregion CompanyFindfromUserInput

        #region dataBaseFind
        public void DataBaseFinder(string DataBaseName, List<string> databaseList, [Optional] string AlternativeDataBasename)
        {
            if (!string.IsNullOrEmpty(AlternativeDataBasename))
                DataBaseName = AlternativeDataBasename;
            FilenamesForSearch.Clear();
            IEnumerable<string> matchingList;

            if (!string.IsNullOrEmpty(DataBaseName))
            {
                matchingList = databaseList.Where(x => x.ToUpper().Contains(DataBaseName.ToUpper()));
                if (matchingList != null)
                {
                    FilenamesForSearch = matchingList.ToList();
                }
            }
        }
        #endregion dataBaseFind

        #region FileSort
        public string ReferencefileSort(string FileInputDir)
        {
            DateTime lastupdated = DateTime.Today;
            string strFilename1 = FileInputDir;
            string Folder = strFilename1;

            var files = new DirectoryInfo(Folder).GetFiles("*.xls");

            foreach (FileInfo file in files)
            {
                if (file.LastWriteTime < lastupdated)
                {
                    if (file.Name.Contains(DC.File1NameTrimmed)) //softcode it
                    {
                        lastupdated = file.LastWriteTime;
                        DC.File2FullPath = file.FullName;
                        DC.File2Name = file.Name;
                    }
                }
            }
            return DC.File2FullPath;
        }

        public string LatestFileSort(string FileInputDir,string format)
        {
            DateTime lastupdated = DateTime.MinValue;
            string strFilename1 = FileInputDir;
            string Folder = strFilename1;

            var files = new DirectoryInfo(Folder).GetFiles("*."+format);//"*.xml"

            foreach (FileInfo file in files)
            {
                if (file.LastWriteTime > lastupdated)
                {
                    lastupdated = file.LastWriteTime;
                    DC.File1FullPath = file.FullName;
                    DC.File1Name = file.Name;
                }
            }
            try
            {
                int temp_index = DC.File1Name.IndexOf("_");
                DC.File1NameTrimmed = DC.File1Name.Substring(temp_index + 1, 15);
                int temp_index1 = DC.File1NameTrimmed.IndexOf("_");
                temp_index = temp_index + temp_index1 + 1;//additional 1 char helps to get the complete last word
                DC.File1NameTrimmed = DC.File1Name.Substring(0, temp_index);
                return DC.File1FullPath;
            }
            catch
            {
                return DC.File1FullPath;
            }
            
        }
        #endregion FileSort

        #region XML File Picker for Ticket

        public string XMLFilePicker(string NameOfFile)
        {
            FileFullPath = string.Empty;
            FilePathOfXMLtemp = string.Empty;
            FileNameOnly = string.Empty;
            ExportXlSXfilePath = string.Empty;
            try
            {
                FileFullPath = DeclarationClass.ROOTDIRFORXMLFILES + "\\" + NameOfFile + "\\";
                string FilePathComplete = LatestFileSort(FileFullPath,"xml");//latest Sort function

                FilePathOfXMLtemp = FilePathComplete;
                FileNameOnly = FilePathOfXMLtemp.Substring(FileFullPath.Length, (FilePathOfXMLtemp.Length - FileFullPath.Length));
                FileNameOnly = FileNameOnly.Substring(0, FileNameOnly.IndexOf('.'));
                ExportXlSXfilePath = DeclarationClass.VISHALSHIPMENTPATH_ + NameOfFile + "\\";//important in the future states
                string TempPath = DeclarationClass.VISHALSHIPMENTPATH_ + NameOfFile;
                if (!Directory.Exists(TempPath))
                {
                    Directory.CreateDirectory(TempPath);
                }
                return FilePathOfXMLtemp;
            }
            catch(Exception e)
            {
                MessageBox.Show(string.Empty+e);
                return "ERROR SELECTION";
            }

        }
        #endregion XML File Picker for Ticket

        #region RootDirectoryExlporerForXLS
        public void FileExplorerForXML(string ParentPath, DateTime startDate, DateTime endDate, [Optional] string CompanyName)
        {
            // Get a list of all subdirectories
            //TempFileNames
            //TempFileDirecrtory
        var dirs = from dir in
                Directory.EnumerateDirectories(ParentPath)
                       select dir;
           
            foreach (var dir in dirs)
            {
                TempFileNames.Clear();
                TempFileDirecrtory.Clear();
                //DirNames.Add(dir.Substring(dir.LastIndexOf("\\") + 1));
                DateTime lastupdated = startDate;
                DirectoryFilename = ParentPath + "\\" + dir.Substring(dir.LastIndexOf("\\") + 1);
                string Folder = DirectoryFilename;

                var files = new DirectoryInfo(Folder).GetFiles("*.xls");//*.*//We are tryig to find the  
                foreach (FileInfo file in files)              //important code segement for file extraction
                {
                    if (file.LastWriteTime >= startDate && file.LastWriteTime<= endDate)
                    {
                        XMLFileFullPath = file.FullName;
                        XMLFileName = file.Name;
                        TempFileNames.Add(XMLFileName);
                        TempFileDirecrtory.Add(XMLFileFullPath);
                    }
                }
                if(TempFileDirecrtory.Count>1)
                {
                    FileNames.Add(TempFileNames[TempFileNames.Count-1]);
                    FileDirecrtory.Add(TempFileDirecrtory[TempFileDirecrtory.Count-1]);
                    DirNames.Add(dir.Substring(dir.LastIndexOf("\\") + 1));
                }
                else if(TempFileDirecrtory.Count==1)
                {
                    FileNames.Add(TempFileNames[0]);
                    FileDirecrtory.Add(TempFileDirecrtory[0]);
                    DirNames.Add(dir.Substring(dir.LastIndexOf("\\") + 1));
                }
            }
        }

        /* if(XMLFileFullPath==null || XMLFileName == null|| TempFileNames[TempFileNames.Count - 1]==null || TempFileDirecrtory[TempFileDirecrtory.Count - 1]==null)
                {
                    XMLFileFullPath ="NoFileFound";
                    XMLFileName = "NoFileFound";
                }
                else
                {
                    FileNames.Add(TempFileNames[TempFileNames.Count-1]);
                    FileDirecrtory.Add(TempFileDirecrtory[TempFileDirecrtory.Count-1]);
                }
                int numm = files.Length - 1;
                //FileNames.Add(XMLFileName);
                //FileDirecrtory.Add(XMLFileFullPath);
         * 
         * 
         * 
         * if(file.LastWriteTime > lastupdated)
            {
                TempFileNames.Add(XMLFileName);
                TempFileDirecrtory.Add(XMLFileFullPath);
                //FileNames.Add(XMLFileName);
                //FileDirecrtory.Add(XMLFileFullPath);
                lastupdated = file.LastWriteTime;
            }
            else
            {
                XMLFileFullPath = string.Empty;
                XMLFileName = string.Empty;
            }*/
        #endregion RootDirectoryExlporerForXLS
    }
}

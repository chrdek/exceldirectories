using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Office.Core;
using System.Runtime.Remoting.Messaging;
using DSOFile;
using Spire.Xls;
using Spire.Pdf;
using Spire.Xls.Core.Spreadsheet;


namespace ExcelDirectories
{
    public partial class Form1 : Form
    {

        private string strFile="";
        private string strPath = "./";

        public Form1()
        {
            InitializeComponent();
        }

        private void ListDirectory(TreeView treeView, string path)
        {
            treeView.Nodes.Clear();
            path = strPath;
            var rootDirectoryInfo = new DirectoryInfo(path);
            treeView.Nodes.Add(CreateDirectoryNode(rootDirectoryInfo));
        }

        private static TreeNode CreateDirectoryNode(DirectoryInfo directoryInfo)
        {
            var directoryNode = new TreeNode(directoryInfo.Name);
            foreach (var directory in directoryInfo.GetDirectories())
                directoryNode.Nodes.Add(CreateDirectoryNode(directory));
            foreach (var file in directoryInfo.GetFiles())
                directoryNode.Nodes.Add(new TreeNode(file.Name));
            return directoryNode;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ListDirectory(treeView,strPath);
            treeView.NodeMouseClick += TreeView_NodeMouseClick;
        }


        private void button1_Click(object sender, EventArgs e)
        {
           DialogResult dres = folderBrowserDialog1.ShowDialog();

            if (dres == DialogResult.OK)
                strPath = folderBrowserDialog1.SelectedPath;
           ListDirectory(treeView,strPath);
        }


        private void TreeView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {

            FileInfo newxls = new FileInfo(@"./" + e.Node.Text + ".xls");
            FileStream fsInfo = newxls.Create();
            strFile = newxls.FullName;
            fsInfo.Close();

            CreateXLSView(strFile);
            if (!File.Exists(strFile)) throw new Exception();
            //Uri fileUri = new Uri("file:///" + strFile);
            spreadsheet1.LoadFromFile(strPath+"\\output.xls", true);
            GC.Collect();
        }


        public void CreateXLSView(string inputName)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            
            if (xlApp == null) return;
            Microsoft.Office.Interop.Excel.Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            xlApp.Workbooks.Open(inputName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t",
                true, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            Microsoft.Office.Interop.Excel.Range xlsRange = ws.get_Range("A1","J1");

            for (int i = 1; i <= xlsRange.Count; i++)
            {
                xlsRange.Cells[1, i].Interior.Color = Color.DarkGray;
                xlsRange.Cells[1, i].Font.Bold = true;
                xlsRange.Cells[1, i].Font.Size = 12;
            }
            xlsRange.Cells[1, 1].Value = "Dir. Name";
            xlsRange.Cells[1, 2].Value = "File Name";
            xlsRange.Cells[1, 3].Value = "File Extension";
            xlsRange.Cells[1, 4].Value = "Date Created";
            xlsRange.Cells[1, 5].Value = "Date Last Modified";
            xlsRange.Cells[1, 6].Value = "File Size";
            xlsRange.Cells[1, 7].Value = "Author Name";
            xlsRange.Cells[1, 8].Value = "Company Name";
            xlsRange.Cells[1, 9].Value = "Title";
            xlsRange.Cells[1, 10].Value = "Subject";

            List<string[]> outputxlsData = CreateDirDataList(DirSearch(strPath));

            for (int p=2; p < outputxlsData.Count();  p++)
            {
                for (int r=1; r < 10; r++)
                {
                    xlsRange.Cells[p, r].Value = outputxlsData[p][r];
                }
            }

            wb.SaveAs(strPath+"\\output.xls", XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
    Missing.Value, false, false, XlSaveAsAccessMode.xlNoChange,
    XlSaveConflictResolution.xlUserResolution, true,
    Missing.Value, Missing.Value, Missing.Value);
            xlApp.Workbooks.Close();
        }


  
        private static List<FileInfo> DirSearch(string sDir)
        {
            List<FileInfo> allfilesInfo = new List<FileInfo>();
            

            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    foreach (string f in Directory.GetFiles(d))
                    {
                        allfilesInfo.Add(new FileInfo(f));

                    }
                    DirSearch(d);
                }

            }

            catch (System.Exception excpt)
            {
                MessageBox.Show(excpt.Message);
                
            }

            return allfilesInfo;
        }




        public static List<string[]> CreateDirDataList(List<FileInfo> fInfoList) { 

        string[] propertiesdocs = new string[10];
        List<string[]> allFileinformation = new List<string[]>();

                foreach (var file in fInfoList)
                {
                    if (file.Extension == ".pdf" || file.Extension == ".doc" ||
                        file.Extension == ".docx" || file.Extension == ".xls" ||
                        file.Extension == ".xlsx" || file.Extension == ".rtf")
                    {
                        DSOFile.OleDocumentProperties dsoProps = new DSOFile.OleDocumentProperties();
                        dsoProps.Open(file.FullName);

                        propertiesdocs[0] = file.DirectoryName ?? " ";
                        propertiesdocs[1] = file.Name + file.Extension ?? " ";
                        propertiesdocs[2] = file.Extension ?? " ";
                        propertiesdocs[3] = file.CreationTime.ToString() ?? " ";
                        propertiesdocs[4] = file.LastAccessTime.ToString() ?? " ";
                        propertiesdocs[5] = (((float) file.Length/1024)/1024f).ToString("F3") + "MB" ?? " ";
                        propertiesdocs[6] = dsoProps.SummaryProperties.Author ?? " ";
                        propertiesdocs[7] = dsoProps.SummaryProperties.Subject ?? " ";
                        propertiesdocs[8] = dsoProps.SummaryProperties.Company ?? " ";
                        propertiesdocs[9] = dsoProps.SummaryProperties.Title ?? " ";
                        allFileinformation.Add(new string[]
                        {
                            propertiesdocs[0], propertiesdocs[1], propertiesdocs[2], propertiesdocs[3],
                        propertiesdocs[4], propertiesdocs[5], propertiesdocs[6], propertiesdocs[7], propertiesdocs[8], propertiesdocs[9]
                        });
                    }
                    else
                    {

                        propertiesdocs[0] = file.DirectoryName ?? " ";
                        propertiesdocs[1] = file.Name + file.Extension ?? " ";
                        propertiesdocs[2] = file.Extension ?? " ";
                        propertiesdocs[3] = file.CreationTime.ToString() ?? " ";
                        propertiesdocs[4] = file.LastAccessTime.ToString() ?? " ";
                        propertiesdocs[5] = (((float) file.Length/1024)/1024f).ToString("F3") + "MB" ?? " ";
                        propertiesdocs[6] = "N/A";
                        propertiesdocs[7] = "N/A";
                        propertiesdocs[8] = "N/A";
                        propertiesdocs[9] = "N/A";
                        allFileinformation.Add(new string[]
                        {
                            propertiesdocs[0], propertiesdocs[1], propertiesdocs[2], propertiesdocs[3],
                        propertiesdocs[4], propertiesdocs[5], propertiesdocs[6], propertiesdocs[7], propertiesdocs[8], propertiesdocs[9]
                        });

                }
                }
            return allFileinformation;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            workbook.LoadFromFile(strPath + "\\output.xls");
            workbook.ConverterSetting.SheetFitToPage = true;
            var worksheet = workbook.Worksheets[0];
            worksheet.SaveToPdf(strPath + "\\output.pdf");
            
            MessageBox.Show("Excel view exported to PDF.\r\n Click OK to continue.\r\n Files produced in directory: [output.pdf]");
        }


        // This method exports to both excel compatible-compliant csv and to UTF-8 compliant csv text.
        private void button3_Click(object sender, EventArgs e)
        {
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            workbook.LoadFromFile(strPath + "\\output.xls");
            workbook.ConverterSetting.SheetFitToPage = true;
            var worksheet = workbook.Worksheets[0];

            //Save to a running memory set..
            var ms = new MemoryStream();
            workbook.SaveToStream(ms, Spire.Xls.FileFormat.CSV);
            ms.Position = 0;
            using (FileStream fs = new FileStream(strPath + "\\output2.csv", FileMode.Create, FileAccess.Write))
                ms.CopyTo(fs);

            //output.csv = text only, output2.csv = excel compliant
            worksheet.SaveToFile(strPath + "\\output.csv", ",", System.Text.Encoding.UTF8);
            workbook.SaveToFile(strPath + "\\output2.csv");

            MessageBox.Show("Excel view exported to CSV.\r\n Click OK to continue.\r\n Files produced in directory: [output.csv,output2.csv]");
        }
    }
}

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
//using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;


namespace PdfToExcel2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "PDF Files|*.pdf";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }
            }
            button1.Enabled = false;
            var TextList = GetText(filePath);
            for (int i = 0; i < TextList.Count; i++)
            {
                string L = "";
                foreach (var x in TextList[i].Split("\n"))
                    L = L + x;
                TextList[i] = L;
                int a = (TextList[i]).Length;
            }
            string name = GetFileName(filePath);
            string path = GetFilePath(filePath);

            SaveToExcel(TextList, path, name);
            button1.Enabled = true;
        }

        private void SaveToExcel(List<string> Text, string path, string name)
        {
            name = GenerateName(path, name);
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                int N = Text.Count;
                for (int i=1; i<N+1; i++)
                {
                    oSheet.Cells[i, 1] = Text[i-1];
                    textBox1.Text = ((i + 1).ToString() + "/" + N.ToString() + ": " + Text[i-1]);
                }

                oWB.SaveAs(path+"\\"+name, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                textBox1.Text = "Записано";


            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }


        }

        private List<string> GetText(string path)
        {
            var res = new List<string>();
            using (PdfReader reader = new PdfReader(path))
            {
                StringBuilder text = new StringBuilder();

                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    res.Add(PdfTextExtractor.GetTextFromPage(reader, i));
                }

                return res;
            }
        }

        private string GetFileName(string path)
        {
            string res = "";
            res = path.Split('\\')[path.Split('\\').Length - 1];
            res = res.Replace(".pdf", "");
            return res;
        }

        private string GetFilePath(string path)
        {
            return path.Substring(0, path.LastIndexOf('\\'));
        }

        private string GenerateName(string path, string name)
        {
            string res = "";
            var directory = new DirectoryInfo(path);
            FileInfo[] files = directory.GetFiles();
            List<string> FileNames = new List<string>();
            foreach (FileInfo file in files)
            {
                FileNames.Add(file.Name);
            }
            if (!FileNames.Contains(name + ".xlsx"))
                return name;
            int n = 1;
            while (FileNames.Contains(name + "(" + n.ToString() + ").xlsx"))
                n++;
            //string s = name + "(" + n.ToString() + ").xlsx";
            return (name + "(" + n.ToString() + ")");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
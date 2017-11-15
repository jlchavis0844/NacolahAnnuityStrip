using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NacolahAnnuityStrip {
    class Program {

        static Microsoft.Office.Interop.Excel.Application oXL;
        static Microsoft.Office.Interop.Excel._Workbook oWB;
        static Microsoft.Office.Interop.Excel._Worksheet oSheet;
        static Microsoft.Office.Interop.Excel.Range oRng;
        static object misvalue = System.Reflection.Missing.Value;
        static List<string> pdfLines = new List<string>();
        static string fileName = "";
        static List<CommLine> commLines;

        [STAThread]
        static void Main(string[] args) {
            commLines = new List<CommLine>();
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = "P:\\RALFG\\Common Files\\Commissions & Insurance\\Commission Statements\\2017\\NACOLAH-Ann\\";
            ofd.Filter = "PDF files (*.pdf)|*.pdf";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;
            DialogResult result = ofd.ShowDialog();
            string pdfPath = "";

            if (result == DialogResult.OK) {
                pdfPath = ofd.FileName;
                fileName = System.IO.Path.GetFileName(pdfPath);
            }
            else Environment.Exit(0);

            try {
                StringBuilder text = new StringBuilder();
                PdfReader pdfReader = new PdfReader(pdfPath);
                for (int page = 1; page <= pdfReader.NumberOfPages; page++) {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                    text.Append(System.Environment.NewLine);
                    text.Append("\n Page Number:" + page);
                    text.Append(System.Environment.NewLine);
                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, 
                        Encoding.Default.GetBytes(currentText)));
                    text.Append(currentText);
                    //pdfReader.Close();
                    string[] lines = currentText.Split('\n');
                    foreach (string line in lines) {
                        pdfLines.Add(line);
                    }
                }
            }
            catch (Exception ex) {
                MessageBox.Show("Error: " + ex.Message, "Error");
            }

            //remove all page footers
            pdfLines.RemoveAll(item => item.StartsWith("https"));
            pdfLines.RemoveAll(item => item.EndsWith("CommissionStatement"));

            for (int i = 0; i < pdfLines.Count; i++) {

                while (!pdfLines[i].StartsWith("8000")) {
                    i++;
                }

                List<string> tokens = new List<string>();
                tokens.AddRange(pdfLines[i++].Split(' '));
                if (i == pdfLines.Count)
                    break;

                while (!pdfLines[i].StartsWith("8000")) {
                    tokens.AddRange(pdfLines[i++].Split(' '));
                    if (i == pdfLines.Count)
                        break;
                }
                i--;

                string policyNum = tokens[0];
                tokens.RemoveAt(0);

                DateTime temp;

                string issueDate = "";
                for (int j = 0; j < tokens.Count; j++) {
                    if (DateTime.TryParse(tokens[j], out temp)) {
                        issueDate = tokens[j];
                        tokens.RemoveAt(j);
                        break;
                    }
                }

                string premium = "";
                for (int j = 0; j < tokens.Count; j++) {
                    if (tokens[j].StartsWith("$")) {
                        premium = tokens[j];
                        tokens.RemoveAt(j);
                        break;
                    }
                }

                string rate = "";
                for (int j = 0; j < tokens.Count; j++) {
                    if (tokens[j].EndsWith("%")) {
                        rate = tokens[j];
                        tokens.RemoveAt(j);
                        break;
                    }
                }

                string commission = "";
                for (int j = 0; j < tokens.Count; j++) {
                    if (tokens[j].StartsWith("$")) {
                        commission = tokens[j];
                        tokens.RemoveAt(j);
                        break;
                    }
                }

                string split = "";
                for (int j = 0; j < tokens.Count; j++) {
                    if (tokens[j].EndsWith("%")) {
                        split = tokens[j];
                        tokens.RemoveAt(j);
                        break;
                    }
                }
                int nameInt = tokens.IndexOf("Name:") + 1;
                string name = "";
                while (nameInt != tokens.Count && tokens[nameInt] != "Agent:") {
                    name += (tokens[nameInt] + " ");
                    nameInt++;
                }

                string plan = "";
                while (!DateTime.TryParse(tokens[0], out temp)) {
                    plan += tokens[0] + " ";
                    tokens.RemoveAt(0);
                }
                if (Convert.ToDouble(commission.Replace("$", "")) > 0) {
                    commLines.Add(new CommLine(name, policyNum, issueDate, premium, rate, commission, split, plan));
                }
            }

            //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\testing\outPut.txt")) {
            //    foreach (CommLine line in commLines) {
            //        file.WriteLine(line);
            //        Console.WriteLine(line);
            //    }
            //}

            writeToExcel();
        }
 

        public static void writeToExcel() {
            string outFile = "";
            try {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = false;
                oXL.UserControl = false;
                oXL.DisplayAlerts = false;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "Policy";
                oSheet.Cells[1, 2] = "Fullname";
                oSheet.Cells[1, 3] = "Plan";
                oSheet.Cells[1, 4] = "Issue Date";
                oSheet.Cells[1, 5] = "Premium";
                oSheet.Cells[1, 6] = "Rate %";
                oSheet.Cells[1, 7] = "Rate";
                oSheet.Cells[1, 8] = "Commission";
                oSheet.Cells[1, 9] = "Renewal";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "I1").Font.Bold = true;
                oSheet.get_Range("A1", "I1").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                
                for (int i = 0; i < commLines.Count; i++) {
                    oSheet.get_Range("A" + (i + 2), "I" + (i + 2)).Value2 = commLines[i].GetData();
                }
                oRng = oSheet.get_Range("A1", "I1");
                oRng.EntireColumn.AutoFit();
                oXL.Visible = false;
                oXL.UserControl = false;

                outFile = GetSavePath();

                oWB.SaveAs(outFile, 
                    56, //Seems to work better than default excel 16
                    Type.Missing,
                    Type.Missing,
                    false,
                    false,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing);
                
                //System.Diagnostics.Process.Start(outFile);
            }
            catch (Exception ex) {
                MessageBox.Show("Error: " + ex.Message, "Error");
            } finally {
                if (oWB != null)
                    oWB.Close();
                if(File.Exists(outFile))
                    System.Diagnostics.Process.Start(outFile);
            }
        }

        public static string GetSavePath() {
            
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = "H:\\Desktop\\";
            saveFileDialog1.Filter = "xls|*.xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = fileName.Replace(".pdf","_out");

            if (saveFileDialog1.ShowDialog() == DialogResult.OK) {
                return saveFileDialog1.FileName;
            }
            else Application.Exit();
            return "";
            }
        
    }
}

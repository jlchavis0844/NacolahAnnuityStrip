using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Text;
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
            StringBuilder text = new StringBuilder();
            try {
                //StringBuilder text = new StringBuilder();
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

                System.IO.StreamWriter file = new StreamWriter("H:\\Desktop\\NacAnn.txt");
                file.WriteLine(text.ToString());
                file.Close();
            }
            catch (Exception ex) {
                MessageBox.Show("Error: " + ex.Message, "Error");
            }

            pdfLines.RemoveAll(item => item.StartsWith("3R"));
            pdfLines.RemoveAll(item => item.Trim().StartsWith("Page "));
            pdfLines.RemoveAll(item => item.Length < 1);
            
            for (int i = 0; i < pdfLines.Count; i++) {
                string polNum = "";
                string plan = "";
                string issueDate = "";
                string prem = "";
                string rate = "";
                string comm = "";
                string meth = "";
                string tDate = "";
                string owner = "";
                string agent = "";
                string commOpt = "";

                if (pdfLines[i].StartsWith("8000")) {
                    string[] tokens = pdfLines[i].Split(' ');

                    polNum = tokens[0];

                    //check for run on first line like : 8000276810 NA IncomeChoice 10 03/07/2016 $600.00 0.50% $3.00
                    if (tokens.Length < 7) {
                        for (int j = 1; j < tokens.Length; j++) {
                            plan += " " + tokens[j];
                        }
                        i++;

                        tokens = pdfLines[i].Split(' ');
                        for (int j = 0; j < tokens.Length; j++) {
                            plan += " " + tokens[j];
                        }
                        i++;
                        plan = plan.Trim();
                        tokens = pdfLines[i].Split(' ');

                        //check for run on line caused by page breaks like.
                        if (tokens.Length >= 8) {
                            issueDate = tokens[0];
                            prem = tokens[1];
                            rate = tokens[2];
                            comm = tokens[3];
                            tDate = tokens[4];
                            meth = tokens[6];
                            commOpt = tokens[7];
                            i++;
                            agent = pdfLines[i].Trim();
                            i += 2;
                            owner = pdfLines[i].Replace("Owner Name: ", "").Replace(" Writing Agent:", "");
                            commLines.Add(new CommLine(owner, polNum, issueDate, prem, rate, comm, "100", plan));
                            continue;
                        }
                        else {
                            issueDate = tokens[0];
                            prem = tokens[1];
                            rate = tokens[2];
                            comm = tokens[3];
                            i++;
                        }
                    } else {
                        DateTime temp;

                        int cntr = 1;
                        while(cntr < tokens.Length && !DateTime.TryParse(tokens[cntr],out temp)){
                            plan += " " + tokens[cntr];
                            cntr++;
                        }
                        issueDate = tokens[cntr++];
                        prem = tokens[cntr++];
                        rate = tokens[cntr++];
                        comm = tokens[cntr++];
                        i++;
                    }

                    owner = pdfLines[i].Replace("Owner Name: ", "").Replace(" Writing Agent:", "");
                    i++;

                    tokens = pdfLines[i].Split(' ');
                    tDate = tokens[0];
                    meth = tokens[2];
                    commOpt = tokens[3];
                    i++;

                    
                    agent = pdfLines[i];
                    if (agent.StartsWith("8000")) {
                        agent = "Skipped Agent";
                        i--;
                    }

                    commLines.Add(new CommLine(owner, polNum, issueDate, prem, rate, comm, "100", plan));
                }
            }

            commLines.RemoveAll(c => c.comm == 0);
            //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\testing\outPut.txt")) {
            //    foreach (CommLine line in commLines) {
            //        file.WriteLine(line);
            //        Console.WriteLine(line);
            //    }
            //}

            string pdfTotal = pdfLines.Find(e => e.StartsWith("EFT Amount")).Replace("EFT Amount","").Replace("$","").Trim();
            double commTotal = commLines.Sum(e => e.comm);

            if(commTotal != Convert.ToDouble(pdfTotal)) {
                MessageBox.Show("Warning, PDF total doesn't match commission total", "WARNING: TOTALS DON'T MATCH", MessageBoxButtons.OK);
            }

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

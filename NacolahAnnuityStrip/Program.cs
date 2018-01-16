using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Globalization;

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
            ofd.InitialDirectory = "P:\\RALFG\\Common Files\\Commissions & Insurance\\Commission Statements\\" 
                + DateTime.Now.Year.ToString() + "\\NACOLAH-Ann\\";
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
            } catch (Exception ex) {
                MessageBox.Show("Error: " + ex.Message, "Error");
            }

            //while(pdfLines[0] != "Regular Fixed Annuity Commissions") {
            //    pdfLines.RemoveAt(0);
            //}

            pdfLines.RemoveAll(item => item.StartsWith("3R"));
            pdfLines.RemoveAll(item => item.Trim().StartsWith("Page "));
            pdfLines.RemoveAll(item => item.Length < 1);

            for (int i = 0; i < pdfLines.Count; i++) {
                if (pdfLines[i].StartsWith("8000")) {
                    int start = i;
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
                    string split = "100";
                    string[] tokens = pdfLines[i].Split(' ');

                    polNum = tokens[0];
                    if (polNum == "8000132254")
                        Console.WriteLine(polNum);

                    //check for run on first line like : 8000276810 NA IncomeChoice 10 03/07/2016 $600.00 0.50% $3.00
                    if (tokens.Length < 7) {
                        //check for missing plan name by assuming line like "8000328389 01/09/2018 $167,006.25 2.50% $4,175.16" 
                        DateTime mpTemp;

                        if (DateTime.TryParse(tokens[1], out mpTemp)) {
                            plan = "MISSING PLAN NAME";
                            issueDate = tokens[1];
                            prem = tokens[2];
                            rate = tokens[3];
                            comm = tokens[4];
                            i++;
                            owner = pdfLines[i].Replace("Owner Name: ", "").Replace("Writing Agent:", "");
                            i += 2;
                            agent = pdfLines[i];

                            commLines.Add(new CommLine(owner, polNum, issueDate, prem, rate, comm, split, plan));
                            i = start;
                            continue;

                        }

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
                            commLines.Add(new CommLine(owner, polNum, issueDate, prem, rate, comm, split, plan));
                            i = start;
                            continue;
                        }
                        else {
                            DateTime temp;
                            //issue date will be dropped sometimes, let's find and use the next date
                            if (!DateTime.TryParse(tokens[0], out temp)) {
                                int fNum = i + 1;
                                while (issueDate == "") {
                                    string[] tTokens = pdfLines[fNum].Split(' ');
                                    foreach (string item in tTokens) {
                                        if (DateTime.TryParse(item, out temp)) {
                                            issueDate = item;
                                            break;
                                        }
                                    }
                                    fNum++;
                                }

                                prem = tokens[0];
                                rate = tokens[1];
                                comm = tokens[2];
                                commLines.Add(new CommLine(owner, polNum, issueDate, prem, rate, comm, split, plan));
                                i = start;
                                continue;
                            }

                            issueDate = tokens[0];
                            prem = tokens[1];
                            rate = tokens[2];
                            comm = tokens[3];
                            i++;
                        }
                    }
                    else {

                        DateTime temp;

                        int cntr = 1;
                        while (cntr < tokens.Length && !DateTime.TryParse(tokens[cntr], out temp)) {
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
                    tokens = pdfLines[i].Split(' ');
                    string tempSplit = tokens[tokens.Length - 1].Trim();

                    if (tempSplit.EndsWith("%")) {
                        double splitNum = Convert.ToDouble(tempSplit.Replace("%", ""));

                        if (splitNum < 100)
                            split = splitNum.ToString();
                    }
                    if (agent.StartsWith("8000")) {
                        agent = "Skipped Agent";
                        i--;
                    }
                    commLines.Add(new CommLine(owner, polNum, issueDate, prem, rate, comm, split, plan));
                    i = start;
                }
            }
            int rawCount = commLines.Count;
            commLines.RemoveAll(c => c.comm == 0);

            string pdfTotal = pdfLines.Find(e => e.StartsWith("EFT Amount")).Replace("EFT Amount", "").Trim();
            double commTotal = commLines.Sum(e => e.comm);
            string cTotal = commTotal.ToString("C", CultureInfo.CurrentCulture);

            if (cTotal != pdfTotal) {
                string message = "Warning, PDF total doesn't match commission total\n";
                message += "PDF total = " + pdfTotal + "\n";
                message += "calculated total = " + cTotal + "\n";
                message += "processed comm lines: " + rawCount + " of which " + commLines.Count + " were kept";
                MessageBox.Show(message, "WARNING: TOTALS DON'T MATCH", MessageBoxButtons.OK);
            }
            CheckIssueDates();
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
            }
            finally {
                if (oWB != null)
                    oWB.Close();
                if (File.Exists(outFile))
                    System.Diagnostics.Process.Start(outFile);
            }
        }

        public static string GetSavePath() {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = "H:\\Desktop\\";
            saveFileDialog1.Filter = "xls|*.xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.FileName = fileName.Replace(".pdf", "_out");

            if (saveFileDialog1.ShowDialog() == DialogResult.OK) {
                return saveFileDialog1.FileName;
            }
            else Application.Exit();
            return "";
        }

        public static int CheckIssueDates() {
            int cnt = 0;
            SqlConnection cs = new SqlConnection("Data Source=RALIMSQL1\\RALIM1; " +
                "Initial Catalog = CAMSRALFG; " +
                "Integrated Security = SSPI; " +
                "Persist Security Info = false; " +
                "Trusted_Connection = Yes");
            SqlCommand cmd = new SqlCommand();
            SqlDataReader reader;
            string currPol = "";

            foreach (CommLine line in commLines) {
                if (line.iDate == null || line.iDate == "") {
                    currPol = line.policy.ToString();
                    string query = @"SELECT Convert(varchar(10),MIN(Sales.IssueDate),101) FROM Sales WHERE Sales.[Policy#]='" + currPol + "';";

                    try {
                        cmd.CommandText = query;
                        cmd.CommandType = System.Data.CommandType.Text;
                        cmd.Connection = cs;
                        cs.Open();

                        reader = cmd.ExecuteReader();

                        if (reader.HasRows) {
                            if (!reader.Read()) {
                                throw new System.Exception("Problem reading results.");
                            }
                            line.iDate = reader.GetString(0);
                        }
                        else {
                            throw new System.Exception("Couldn't read data from Database or results were empty.");
                        }
                        cnt++;
                    }
                    catch (Exception eIDate) {
                        MessageBox.Show("Couldn't fetch missing issue date for " + currPol + "\n" + eIDate.ToString());
                    }
                    finally {
                        cs.Close();
                    }
                }
            }
            return cnt;
        }
    }
}

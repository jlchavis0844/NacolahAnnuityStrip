using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;

namespace NacolahAnnuityStrip {
    class CommLine {
        public int RENEWAL = 1;
        public int COMMISSION = 2;

        public string name { get; set; }
        public long policy { get; set; }
        public string iDate { get; set; }
        public double premium { get; set; }
        public double rate { get; set; }
        public double comm { get; set; }
        public double split { get; set; }
        public string plan { get; set; }
        public int age;
        public int type;

        public CommLine() {
            name = "";
            policy = 0;
            iDate = "";
            premium = 0.0;
            rate = 0.0;
            comm = 0.0;
            split = 0.0;
        }

        public CommLine(string name, string policy, string iDate, string premium, string rate, string comm, string split, string plan) {
            Regex nameReg = new Regex(@"[^a-zA-Z\s]");
            CultureInfo culture_info = Thread.CurrentThread.CurrentCulture;
            TextInfo text_info = culture_info.TextInfo;

            this.name = nameReg.Replace(name, "").Trim();
            this.name = text_info.ToTitleCase(this.name.ToLower());

            this.policy = Convert.ToInt64(policy);
            this.iDate = iDate;
            this.premium = Convert.ToDouble(premium.Replace("$",""));
            this.rate = Convert.ToDouble(rate.Replace("%", ""));
            this.comm = Convert.ToDouble(comm.Replace("$", ""));
            this.plan = plan;
            this.plan = text_info.ToTitleCase(this.plan.ToLower());

            if (split != "") {
                this.split = Convert.ToDouble(split.Replace("%", ""));
            }
            else this.split = 0;
            age = 0;

            DateTime temp;
            if(DateTime.TryParse(iDate, out temp)) {
                age = (DateTime.Today - temp).Days;
            }

            if (age > 365) {
                type = RENEWAL;
            }
            else type = COMMISSION;
        }

        public CommLine(string name, string policy, string iDate, string premium, string rate, string comm, string split, string plan, int type) {
            this.name = name;
            this.policy = Convert.ToInt64(policy);
            this.iDate = iDate;
            this.premium = Convert.ToDouble(premium.Replace("$", ""));
            this.rate = Convert.ToDouble(rate.Replace("%", ""));
            this.comm = Convert.ToDouble(comm.Replace("$", ""));
            this.plan = plan;
            if (split != "") {
                this.split = Convert.ToDouble(split.Replace("%", ""));
            }
            else this.split = 0;
            age = 0;

            DateTime temp;
            if (DateTime.TryParse(iDate, out temp)) {
                age = (DateTime.Today - temp).Days;
            }

            this.type = type;
        }

        public override string ToString() {
            return  "Name: " +  name + "\tPolicy: " + policy  + "\tPlan: " + plan + "\tIssueDate: " + iDate + "\tPremium: " + premium +
                    "\tRate: " + rate + "\tComm: " + comm + "\tSplit: " + split;
        }

        public object[] GetData() {
            object[] tArr;
            if (type == RENEWAL) {
                tArr = new object[] { policy.ToString(), name, plan, iDate, premium.ToString("C",CultureInfo.CurrentCulture),
                    rate.ToString("0.##"), split.ToString("0.##"), 0.ToString("C",CultureInfo.CurrentCulture),
                    comm.ToString("C",CultureInfo.CurrentCulture) };
            }
            else {
                tArr = new object[] { policy.ToString(), name, plan, iDate, premium.ToString("C", CultureInfo.CurrentCulture),
                    rate.ToString("0.##"), split.ToString("0.##"), comm.ToString("C", CultureInfo.CurrentCulture),
                    0.ToString("C",CultureInfo.CurrentCulture) };
            }
            return tArr;
        }

    }
}

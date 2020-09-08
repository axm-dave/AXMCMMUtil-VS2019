using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AXMCMMUtil
{
    public class charInfo
    {
        public double keyVal { get; private set; }
        public string run { get; private set; }
        public bool used { get; private set; }
        public int charRow { get; private set; }
        public string charNo { get; private set; }
        public string nominal { get; private set; }
        public string upper { get; private set; }
        public string lower { get; private set; }
        public string desc { get; private set; }
        public double actVal { get; private set; }
        public double devVal { get; private set; }
        public bool isCMM { get; private set; }
        public bool isBasic { get; private set; }
        public List<double> actualList { get; private set; }
        public List<double> devList { get; private set; }

        public charInfo(int row, string cNo, string nom, string up, string lo)
        {
            keyVal = 0.0;
            run = String.Empty;
            used = false;
            charRow = row;
            charNo = cNo;
            nominal = nom;
            upper = up;
            lower = lo;
            actualList = new List<double>();
            devList = new List<double>();
            isBasic = false;
            //actualCnt = 0;
            //actualTotals = 0.0;
            //actualStdDev = 0.0;
        }

        public charInfo(double key, int row, string cNo, string nom, string up, string lo, bool cmm)
        {
            keyVal = key;
            run = String.Empty;
            used = false;
            charRow = row;
            charNo = cNo;
            nominal = nom;
            upper = up;
            lower = lo;
            actualList = new List<double>();
            devList = new List<double>();
            isCMM = cmm;
            isBasic = false;
            //actualCnt = 0;
            //actualTotals = 0.0;
            //actualStdDev = 0.0;
        }

        public void SetBasic(bool val)
        {
            isBasic = val;
        }

        public void AddRun(string val)
        {
            run = val;
        }

        public void AddActual(double val)
        {
            actualList.Add(val);
        }

        public void AddDeviation(double val)
        {
            devList.Add(val);
        }

        public void AddAct(double val)
        {
            actVal = val;
        }

        public void AddDev(double val)
        {
            devVal = val;
        }

        public void AddDesc(string val)
        {
            desc = val;
        }

        public void MarkUsed()
        {
            used = true;
        }

        public bool IsUsed()
        {
            return used;
        }
        
        public bool IsCMM()
        {
            return isCMM;
        }
    }
}

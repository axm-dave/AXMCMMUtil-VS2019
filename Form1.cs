using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;
using System.IO.Packaging;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Reflection;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Drawing;

using Jurassic;
using Jurassic.Library;
using Newtonsoft.Json;

using Excel = Microsoft.Office.Interop.Excel;

//using LuaInterface;

//using MongoDB.Bson;
//using MongoDB.Driver;
//using MongoDB.Driver.Builders;
//using MongoDB.Driver.GridFS;
//using MongoDB.Driver.Linq;
using System.Deployment.Application;
using System.Data.OleDb;
using System.Security.Cryptography;


//using Excel=Microsoft.Office.Interop.Excel;

namespace AXMCMMUtil
{
    
    //public class Version AssemblyVersion 
    //{
    //    get
    //    {
    //        return ApplicationDeployment.CurrentDeployment.CurrentVersion;
    //    }
    //}

    public partial class Form1 : Form
    {
        //
        // this forms Globals
        //

        public Dictionary<string, charInfo> charInfos = new Dictionary<string, charInfo>();
        //public Dictionary<string, resultGrid> grids = new Dictionary<string, resultGrid>();
        public Dictionary<string, bool> resultFiles = new Dictionary<string, bool>();
        public Dictionary<string, resultGrid> fileToGrid = new Dictionary<string, resultGrid>();
        SLDocument sl = null;
        statsGrid statsGrid = null;
        bool statsIncluded = false;
        int nFirstRow = 0;

        public double actualTotals = 0.0;
        public double actualStdDev = 0.0;
        public int actualCnt = 0;

        string fontName = String.Empty;
        string fontSize = String.Empty;
        string fontFormat = String.Empty;
        float ffontSize = 9.5F;
        System.Drawing.Font newFont = null;

        string[] ignoreSet = { "gdtPosPol2d.X", "gdtPosPol2d.Y", "gdtPosPol2d.Z", "someOtherText" };



        class fieldObj
        {
            public string title { get; private set; }
            public string fieldType { get; private set; }
            public bool used { get; private set; }

            public fieldObj(string type, bool isUsed)
            {
                title = null;
                fieldType = type;
                used = isUsed;
            }

            public fieldObj(string titleIn, string type, bool isUsed)
            {
                title = titleIn;
                fieldType = type;
                used = isUsed;
            }
        }

        static Dictionary<string, fieldObj> setFields()
        {
            Dictionary<string, fieldObj> fieldSet = new Dictionary<string, fieldObj>();
            fieldSet.Add("planid", new fieldObj("string", true));
            fieldSet.Add("partnb", new fieldObj("int", false));
            fieldSet.Add("id", new fieldObj("Characteristic Designator:", "string", true));
            fieldSet.Add("type", new fieldObj("string", false));
            fieldSet.Add("idsymbol", new fieldObj("GD&T Symbol:", "picture", true));
            fieldSet.Add("actual", new fieldObj("Results:", "float", true));
            fieldSet.Add("nominal", new fieldObj("Requirements:", "float", true));
            fieldSet.Add("uppertol", new fieldObj("Upper Tol:", "float", true));
            fieldSet.Add("lowertol", new fieldObj("Lower Tol:", "float", true));
            fieldSet.Add("deviation", new fieldObj("Deviation:", "float", true));
            fieldSet.Add("exceed", new fieldObj("string", false));
            fieldSet.Add("warningLimitCF", new fieldObj("float", false));
            fieldSet.Add("featureid", new fieldObj("string", true));
            fieldSet.Add("featuresigma", new fieldObj("float", false));
            fieldSet.Add("comment", new fieldObj("string", true));
            fieldSet.Add("link", new fieldObj("string", false));
            fieldSet.Add("linkmode", new fieldObj("int", false));
            fieldSet.Add("mmc", new fieldObj("int", false));
            fieldSet.Add("useruppertol", new fieldObj("float", false));
            fieldSet.Add("userlowertol", new fieldObj("float", false));
            fieldSet.Add("fftphi", new fieldObj("picture", false));
            fieldSet.Add("fftphiunit", new fieldObj("string", false));
            fieldSet.Add("zoneroundnessangle", new fieldObj("string", false));
            fieldSet.Add("groupname", new fieldObj("string", false));
            fieldSet.Add("groupname2", new fieldObj("string", false));
            fieldSet.Add("datumAid", new fieldObj("string", true));
            fieldSet.Add("datumBid", new fieldObj("string", true));
            fieldSet.Add("datumCid", new fieldObj("string", true));
            fieldSet.Add("natuppertolid", new fieldObj("int", false));
            fieldSet.Add("natlowertolid", new fieldObj("int", false));
            fieldSet.Add("decimalplaces", new fieldObj("int", false));
            fieldSet.Add("featurePosX", new fieldObj("float", false));
            fieldSet.Add("featurePosY", new fieldObj("float", false));
            fieldSet.Add("featurePosZ", new fieldObj("float", false));
            fieldSet.Add("group1", new fieldObj("string", false));
            fieldSet.Add("group2", new fieldObj("string", false));
            fieldSet.Add("group3", new fieldObj("string", false));
            fieldSet.Add("group4", new fieldObj("string", false));
            fieldSet.Add("group5", new fieldObj("string", false));
            fieldSet.Add("group6", new fieldObj("string", false));
            fieldSet.Add("group7", new fieldObj("string", false));
            fieldSet.Add("group8", new fieldObj("string", false));
            fieldSet.Add("group9", new fieldObj("string", false));
            fieldSet.Add("group10", new fieldObj("string", false));

            return (fieldSet);
        }

        Dictionary<string, fieldObj> fields = setFields();

        //string connectionString = String.Empty;
        //MongoClient client = null;

        //public class Entity
        //{
        //    public ObjectId Id { get; set; }

        //    public string Name { get; set; }
        //}

        public Form1()
        {
            InitializeComponent();

            //fontName = "InspectionXpert GDT";
            fontName = "SOLIDWORKS GDT";
            //fontSize = "9.5";
            fontSize = "8";
            fontFormat = "#,##0.000";
            ffontSize = Convert.ToSingle(fontSize);
            newFont = new System.Drawing.Font(fontName, ffontSize);

            Version version = null;

            if (ApplicationDeployment.IsNetworkDeployed)
            {
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion;
            }
            else
            {
                version = Assembly.GetExecutingAssembly().GetName().Version;
            }

            this.label10.Text = String.Format(this.label10.Text, version.Major, version.Minor, version.Build, version.Revision);

            var engine = new Jurassic.ScriptEngine();
            //engine.EnableDebugging = true;
            var result = engine.Evaluate("5 * 10 + 2");
            //engine.SetGlobalValue("console", new Jurassic.Library.FirebugConsole(engine));

            engine.ExecuteFile(@"test.js");
            var result1 = engine.GetGlobalValue<string>("s");
            var result2 = engine.GetGlobalValue<string>("word");
            var resultx = engine.GetGlobalValue<int>("x");
            var resulty = engine.GetGlobalValue<string>("y");

            int result3 = 0;

            try
            {
                result3 = engine.CallGlobalFunction<int>("test2", 50, 6 * 2);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }

            string resultTest3 = String.Empty;

            try
            {
                resultTest3 = engine.CallGlobalFunction<string>("test3", 50, 6 * 2);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }

            var result5 = engine.GetGlobalValue("z");
            var result4 = engine.GetGlobalValue<string>("z");

            foreach (var entry in result4)
            {
                int x = 0;
                x++;
            }
            
            var json = JSONObject.Stringify(engine, result4);
            var json2 = JSONObject.Stringify(engine, result5);
            Object xxxx = JSONObject.Parse(engine, json2);
            JsonTextReader reader = new JsonTextReader(new StringReader(json2));
            while (reader.Read())
            {
                if (reader.Value != null)
                    Console.WriteLine("Token: {0}, Value: {1}", reader.TokenType, reader.Value);
                else
                    Console.WriteLine("Token: {0}", reader.TokenType);
            }

            //var x = result4["sVal"];


            //int i = 0;

            //string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AXMCMMUtil.config");

            //if (File.Exists(configPath))
            //{
            //    Lua lua = new Lua();

            //    lua.DoFile(configPath);

            //    fontName = (string)lua["fontName"];
            //    double tempSize = (double)lua["fontSize"];
            //    fontSize = tempSize.ToString();
            //    fontFormat = (string)lua["fontFormat"];

            //    ffontSize = Convert.ToSingle(fontSize);
            //    newFont = new System.Drawing.Font(fontName, ffontSize);
            //}


            //connectionString = "mongodb://192.168.1.9";
            //client = new MongoClient(connectionString);
            //var server = client.GetServer();
            //var database = server.GetDatabase("test"); // "test" is the name of the database

            //// "entities" is the name of the collection
            //var collection = database.GetCollection<Entity>("entities");

            //var entity = new Entity { Name = "Tom" };
            //collection.Insert(entity);
            //var id = entity.Id; // Insert will set the Id if necessary (as it was in this example)

            //var query = Query<Entity>.EQ(e => e.Id, id);
            //entity = collection.FindOne(query);

            //entity.Name = "Dick";
            //collection.Save(entity);

            //var update = Update<Entity>.Set(e => e.Name, "Harry");
            //collection.Update(query, update);

            //collection.Remove(query);

            //var fileName = "D:\\Dropbox\\Dev\\AXMCMMUtil\\bin\\Debug\\Output.zip";
            //var newFileName = "D:\\Dropbox\\Dev\\AXMCMMUtil\\bin\\Debug\\OutputNew.zip";
            //using (var fs = new FileStream(fileName, FileMode.Open))
            //{
            //    var gridFsInfo = database.GridFS.Upload(fs, fileName);
            //    var fileId = gridFsInfo.Id;

            //    ObjectId oid = new ObjectId(fileId);
            //    ObjectId xxx = new ObjectId();
            //    var file = database.GridFS.FindOne(Query.EQ("_id", oid));

            //    //using (var stream = file.OpenRead())
            //    //{
            //    //    var bytes = new byte[stream.Length];
            //    //    stream.Read(bytes, 0, (int)stream.Length);
            //    //    using (var newFs = new FileStream(newFileName, FileMode.Create))
            //    //    {
            //    //        newFs.Write(bytes, 0, bytes.Length);
            //    //    }
            //   //}
            //}
        }

        private void openSourceFolderMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();

            string[] files = Directory.GetFiles(fbd.SelectedPath);

            foreach (var item in files)
            {
                string fileName = Path.GetFileName(item);
                checkedListBox1.Items.Add(fileName, CheckState.Unchecked);
            }
        }

        private void openSourceTemplateMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void browseSourceFolder_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            string lastResultsPath = Properties.Settings.Default.LastResultsPath;
            openFileDialog1.InitialDirectory = lastResultsPath;
            openFileDialog1.Filter = "txt files (*.txt)|*.txt";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Multiselect = true;
            openFileDialog1.Title = "Select Result Files";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Properties.Settings.Default.LastResultsPath = Path.GetDirectoryName(openFileDialog1.FileName);
                    Properties.Settings.Default.Save();

                    foreach (String file in openFileDialog1.FileNames)
                    {
                        string fileName = Path.GetFileName(file);
                        if (!resultFiles.ContainsKey(fileName))
                        {
                            checkedListBox1.Items.Add(fileName, CheckState.Unchecked);
                            resultFiles.Add(fileName, false);
                        }
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void browseSourceTemplate_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog2 = new OpenFileDialog();

            string lastTemplatePath = Properties.Settings.Default.LastTemplatePath;
            openFileDialog2.InitialDirectory = lastTemplatePath;
            //openFileDialog2.Filter = "xlsx files (*.xlsx)|*.xlsx";
            openFileDialog2.Filter = "xls files (*.xls)|*.xls|xlsx files|*.xlsx";
            openFileDialog2.FilterIndex = 2;
            openFileDialog2.RestoreDirectory = true;
            openFileDialog2.Multiselect = false;
            openFileDialog2.Title = "Select Template Spreadsheet";

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Properties.Settings.Default.LastTemplatePath = Path.GetDirectoryName(openFileDialog2.FileName);
                    string file = openFileDialog2.FileName;
                    string fileName = Path.GetFileName(file);
                    sourceTemplateInput.Text = fileName;
                    Properties.Settings.Default.Save();

                    // set the default Report name and location
                    Properties.Settings.Default.LastReportPath = Properties.Settings.Default.LastTemplatePath;
                    string targetFilename = Path.GetFileNameWithoutExtension(fileName);
                    targetFilename += "-Merge.xlsx";
                    targetInput.Text = targetFilename;

                    //
                    // create the characteristics dictionary - only need to do this one time
                    //
                    string fullTemplateName = Path.Combine(Properties.Settings.Default.LastTemplatePath, sourceTemplateInput.Text);

                    // Use this code if you want to access spreadsheets without spreadsheetlight
                    //////////////////////////////////
                    //Microsoft.Office.Interop.Excel.Application ExcelObj = null;
                    //ExcelObj = new Microsoft.Office.Interop.Excel.Application();
                    //Excel.Workbook theWorkbook = ExcelObj.Workbooks.Open( fullTemplateName, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
                    //Excel.Sheets sheets = theWorkbook.Worksheets;
                    //Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(3);

                    //for (int i = 1; i <= 10; i++)
                    //{
                    //    Excel.Range range = worksheet.get_Range("A" + i.ToString(), "M" + i.ToString());
                    //    Array myvalues = (System.Array)range.Cells.Value;
                    //    string x = myvalues.GetValue(1,2).ToString();
                    //}

                    //////////////////////////////////

                    sl = new SLDocument(fullTemplateName, "Form3");

                    bool bGotChar = true;
                    bool bGotFirstRow = false;
                    string cellVal = String.Empty;
                    int nCol = 2;
                    int nRow = 1;
                    int nIdx = 1;

                    //Array myvalues;

                    while (bGotChar)
                    {
                        cellVal = sl.GetCellValueAsString(nRow, nCol);

                        //Excel.Range range = worksheet.get_Range("A" + nIdx.ToString(), "M" + nIdx.ToString());
                        //myvalues = (System.Array)range.Cells.Value;

                        nIdx++;

                        //if (myvalues.GetValue(1, 2) == null) continue;

                        //cellVal = myvalues.GetValue(1, 2).ToString();

                        if (bGotFirstRow == false)
                        {
                            if (cellVal.Contains("5. Char No"))
                            {
                                nRow += 2;
                                nFirstRow = nRow;
                                bGotFirstRow = true;
                            }
                            else
                            {
                                //if (cellVal.Contains("The ") || cellVal.Length <= 0)
                                if (cellVal.Contains("The "))
                                {
                                    bGotChar = false;
                                }
                                nRow++;
                            }
                            continue;
                        }

                        if (cellVal.Contains("The ") || cellVal.Length <= 0)
                        {
                            bGotChar = false;
                        }
                        else
                        {
                            //string charNo = myvalues.GetValue(1, 2).ToString();
                            //string nominal = myvalues.GetValue(1, 5).ToString();
                            //string upper = myvalues.GetValue(1, 7).ToString();
                            //string lower = myvalues.GetValue(1, 8).ToString();
                            //string tooling = myvalues.GetValue(1, 10).ToString();

                            string charNo = sl.GetCellValueAsString(nRow, 2);
                            string nominal = sl.GetCellValueAsString(nRow, 5);
                            string upper = sl.GetCellValueAsString(nRow, 7);
                            string lower = sl.GetCellValueAsString(nRow, 8);
                            string tooling = sl.GetCellValueAsString(nRow, 10);
                            bool isCMM = tooling.Contains("CMM");

                            //charInfo info = new charInfo(nRow, charNo, nominal, upper, lower);
                            charInfo info = new charInfo(Convert.ToDouble(cellVal), nRow, charNo, nominal, upper, lower, isCMM);
                            charInfos.Add(cellVal, info);
                            nRow++;
                        }
                    }

                    //sl.CloseWithoutSaving();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }

        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            // The user wants to exit the application. Close everything down.
            Application.Exit();
        }

        private void mergeButton_Click(object sender, EventArgs e)
        {
            bool bRC;
            int nInfoRow = nFirstRow - 4;
            //int itemCnt = 0;

            string fullTemplateName = Path.Combine(Properties.Settings.Default.LastTemplatePath, sourceTemplateInput.Text);

            bRC = sl.SelectWorksheet("Form1");

            string partNo = sl.GetCellValueAsString(nInfoRow, 2);
            string partName = sl.GetCellValueAsString(nInfoRow, 4);
            string reportNo = sl.GetCellValueAsString(nInfoRow, 8);
            string serialNo = String.Empty;

            //SLDocument sl = new SLDocument(fullTemplateName, "Form1");

            SLStyle styleFloat = sl.CreateStyle();
            styleFloat.FormatCode = fontFormat;
            styleFloat.SetFontColor(System.Drawing.Color.Black);
            styleFloat.Font.FontName = fontName;
            styleFloat.Font.FontSize = Convert.ToInt16(fontSize);
            styleFloat.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            styleFloat.Alignment.Vertical = VerticalAlignmentValues.Center;
            styleFloat.SetHorizontalAlignment(HorizontalAlignmentValues.Center);

            SLStyle styleFloatV2 = sl.CreateStyle();
            styleFloatV2.Font.FontName = fontName;
            styleFloatV2.Font.FontSize = Convert.ToInt16(fontSize);
            styleFloatV2.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            styleFloatV2.Alignment.Vertical = VerticalAlignmentValues.Center;
            styleFloatV2.SetHorizontalAlignment(HorizontalAlignmentValues.Center);

            SLStyle styleFloatV2LJ = sl.CreateStyle();
            styleFloatV2LJ.Font.FontName = fontName;
            styleFloatV2LJ.Font.FontSize = Convert.ToInt16(fontSize);
            styleFloatV2LJ.Alignment.Horizontal = HorizontalAlignmentValues.Left;
            styleFloatV2LJ.Alignment.Vertical = VerticalAlignmentValues.Center;
            styleFloatV2LJ.SetHorizontalAlignment(HorizontalAlignmentValues.Left);

            SLStyle styleFloatPass = sl.CreateStyle();
            styleFloatPass.FormatCode = fontFormat;
            //styleFloatPass.Fill.SetPatternBackgroundColor(System.Drawing.Color.LightGreen);
            styleFloatPass.SetFontColor(System.Drawing.Color.Black);
            styleFloatPass.Font.FontName = fontName;
            styleFloatPass.Font.FontSize = Convert.ToInt16(fontSize);
            styleFloatPass.Font.Bold = false;
            styleFloatPass.Font.Italic = false;
            styleFloatPass.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            styleFloatPass.Alignment.Vertical = VerticalAlignmentValues.Center;
            styleFloatPass.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            styleFloatPass.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.White, System.Drawing.Color.White);

            SLStyle styleFloatDev = sl.CreateStyle();
            styleFloatDev.FormatCode = fontFormat;
            styleFloatDev.SetFontColor(System.Drawing.Color.Black);
            styleFloatDev.Font.FontName = fontName;
            styleFloatDev.Font.FontSize = Convert.ToInt16(fontSize);
            styleFloatDev.Font.Bold = false;
            styleFloatDev.Font.Italic = false;
            styleFloatDev.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            styleFloatDev.Alignment.Vertical = VerticalAlignmentValues.Center;
            styleFloatDev.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            styleFloatDev.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.LightPink, System.Drawing.Color.Red);

            SLStyle styleString = sl.CreateStyle();
            styleString.SetFontColor(System.Drawing.Color.Black);
            styleString.Font.FontName = fontName;
            styleString.Font.FontSize = Convert.ToInt16(fontSize);

//            foreach (resultGrid grid in grids.Values)
            foreach (resultGrid grid in fileToGrid.Values)
            {
                int     nRow;
                TabPage tabPage = grid.tabPage;
                DataGridView gridCtl = grid.gridControl;
                string runNo = grid.runNo;
                string title = tabPage.Text;

                bRC = sl.SelectWorksheet("Form1");
                bRC = sl.CopyWorksheet("Form3", title);
                bRC = sl.SelectWorksheet(title);

                // set this sheets serial number
                serialNo = String.Format("{0}-{1}", partNo, runNo);
                sl.SetCellValue(nInfoRow, 12, serialNo);


                foreach (DataGridViewRow row in gridCtl.Rows)
                {
                    if (row.IsNewRow) continue;

                    string charNo = row.Cells["CharNo"].Value.ToString();
                    string nominal = row.Cells["Nominal"].Value.ToString();
                    string upper = row.Cells["Upper"].Value.ToString();
                    string lower = row.Cells["Lower"].Value.ToString();
                    string actual = row.Cells["Actual"].Value.ToString();
                    string deviation = row.Cells["Deviation"].Value.ToString();

                    nRow = charInfos[charNo].charRow;

                    // Actual
                    double actualVal;
                    if (actual.Length > 0)
                    {
                        actualVal = Math.Abs(Convert.ToDouble(actual));
                        sl.SetCellValue(nRow, 9, actualVal);
                        sl.SetCellStyle(nRow, 9, styleFloat);
                    }

                    sl.SetCellValue(nRow, 9, actual);
                    sl.SetCellStyle(nRow, 9, styleFloat);

                    sl.SetCellStyle(nRow, 2, styleFloatV2);
                    sl.SetCellStyle(nRow, 4, styleFloatV2LJ);
                    sl.SetCellStyle(nRow, 5, styleFloatV2LJ);
                    sl.SetCellStyle(nRow, 6, styleFloatV2);
                    sl.SetCellStyle(nRow, 7, styleFloatV2);
                    sl.SetCellStyle(nRow, 8, styleFloatV2);

                    // Exceed
                    double exceed = 0;
                    if (deviation.Length > 0)
                    {
                        exceed = Math.Abs(Convert.ToDouble(deviation));
                        sl.SetCellValue(nRow, 11, exceed);
                        sl.SetCellStyle(nRow, 11, styleFloatDev);
                    }
                    else
                    {
                        if (charInfos[charNo].isBasic)
                        {
                            string sRequirement = sl.GetCellValueAsString(nRow, 5);

                            if (sRequirement.StartsWith("("))
                            {
                                sl.SetCellValue(nRow, 11, "REFERENCE");
                            }
                            else
                            {
                                sl.SetCellValue(nRow, 11, "BASIC");
                            }
                        }
                        else
                        {
                            sl.SetCellValue(nRow, 11, "PASS");
                        }
                        
                        sl.SetCellStyle(nRow, 11, styleFloatPass);
                    }
                }
            }

            if (statsIncluded)
            {
                SLPageSettings ps = sl.GetPageSettings("CMMStats");

                bRC = sl.SelectWorksheet("CMMStats");
                if (bRC)
                {
                    sl.SetCellValue(nInfoRow, 2, partNo);
                    sl.SetCellValue(nInfoRow, 6, partName);
                    //sl.SetCellValue(nInfoRow, 12, serialNo);
                    sl.SetCellValue(nInfoRow, 13, reportNo);

                    int nRow = nFirstRow;

                    foreach (DataGridViewRow row in statsGrid.gridControl.Rows)
                    {
                        if (row.IsNewRow) continue;

                        //CharNo, Nominal, Mean, Variance, StdDev
                        string charNo = row.Cells["CharNo"].Value.ToString();
                        string nominal = row.Cells["Nominal"].Value.ToString();
                        string mean = row.Cells["Mean"].Value.ToString();
                        string variance = row.Cells["Variance"].Value.ToString();
                        string stdDev = row.Cells["StdDev"].Value.ToString();
                        string desc = charInfos[charNo].desc;

                        if (nRow > (nFirstRow  + 20))
                        {
                            sl.InsertRow(nRow, 1);
                            sl.MergeWorksheetCells(nRow, 3, nRow, 9);
                            sl.CopyRowStyle(nFirstRow, nRow);
                        }

                        sl.SetCellValue(nRow, 2, charNo);
                        sl.SetCellStyle(nRow, 2, sl.GetCellStyle(nFirstRow, 2));
                        sl.SetCellValue(nRow, 3, desc);
                        sl.SetCellStyle(nRow, 3, sl.GetCellStyle(nFirstRow, 3));
                        sl.SetCellStyle(nRow, 4, sl.GetCellStyle(nFirstRow, 4));
                        sl.SetCellStyle(nRow, 5, sl.GetCellStyle(nFirstRow, 5));
                        sl.SetCellStyle(nRow, 6, sl.GetCellStyle(nFirstRow, 6));
                        sl.SetCellStyle(nRow, 7, sl.GetCellStyle(nFirstRow, 7));
                        sl.SetCellStyle(nRow, 8, sl.GetCellStyle(nFirstRow, 8));
                        sl.SetCellStyle(nRow, 9, sl.GetCellStyle(nFirstRow, 9));
                        sl.SetCellValue(nRow, 10, nominal);
                        sl.SetCellStyle(nRow, 10, sl.GetCellStyle(nFirstRow, 10));
                        sl.SetCellValue(nRow, 11, mean);
                        sl.SetCellStyle(nRow, 11, sl.GetCellStyle(nFirstRow, 11));
                        sl.SetCellValue(nRow, 12, variance);
                        sl.SetCellStyle(nRow, 12, sl.GetCellStyle(nFirstRow, 12));
                        sl.SetCellValue(nRow, 13, stdDev);
                        sl.SetCellStyle(nRow, 13, sl.GetCellStyle(nFirstRow, 13));

                        nRow++;

                    }

                    sl.SetPageSettings(ps);
                }
            }
            else
            {
                sl.SelectWorksheet("Form1");
                sl.DeleteWorksheet("CMMStats");
            }

            
            sl.SelectWorksheet("Form1");
            sl.DeleteWorksheet("Form3");

            string xlsFilename = Path.Combine(Properties.Settings.Default.LastReportPath, targetInput.Text);

            sl.SaveAs(xlsFilename);

            MessageBox.Show(String.Format("Result data merge complete: \n\r{0}", xlsFilename), "FAI Report");
        }

        private void browseTargetLocation_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();

            Properties.Settings.Default.LastReportPath = fbd.SelectedPath;
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            string title = checkedListBox1.SelectedItem.ToString();
            string sLine;
            int lineNo = 1;
            string currentRef = null;
            string currentSheet = String.Empty;
            char[] lineDelimiter = { '\t' };
            //char[] lineDelSpace = { ' ' };
            //char[] dash = { '-' };
            char[] colon = { ':' };
            char[] period = { '.' };

            
            if (e.NewValue == CheckState.Checked)
            {
                // make sure source template name has been specified
                if (sourceTemplateInput.Text.Length <= 0)
                {
                    MessageBox.Show("Template Spreadsheet must be specified", "Template Spreadsheet Missing", 
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation,
                        MessageBoxDefaultButton.Button1);
                    e.NewValue = CheckState.Unchecked;
                    return;
                }

                //if (grids.ContainsKey(title))
                //{
                //    return;
                //}


                resultGrid grid = null;
                TabPage tabPage = null;

                if (fileToGrid.ContainsKey(title))
                {
                    grid = fileToGrid[title];
                    tabPage = grid.tabPage;
                    tabControl1.TabPages.Add(tabPage);
                    return;
                }

                // add data to grid view
                string fullItemName = Path.Combine(Properties.Settings.Default.LastResultsPath, title);
                string fileName = Path.GetFileNameWithoutExtension(fullItemName);
                StreamReader SrcInput = new StreamReader(fullItemName);
                errorList.Items.Add("=====================================================================");
                errorList.Items.Add(string.Format("Parsing: {0}", fullItemName));
                errorList.Items.Add("=====================================================================");
                //AddFileToZip("Output.zip", fullItemName);

                // read the first line and pull off the parameter names
                sLine = SrcInput.ReadLine();
                string[] parms = sLine.Split(lineDelimiter, StringSplitOptions.None);

                ///////////////////////////////////////////////////////////////////////////////////
                while ((sLine = SrcInput.ReadLine()) != null)
                {
                    string[] parts = sLine.Split(lineDelimiter, StringSplitOptions.None);

                    if (sLine == "END") break;

                    lineNo++;

                    int pos = Array.IndexOf(ignoreSet, parts[4]);
                    if (pos > -1)
                    {
                        string error = String.Format("Note: True Position coordinate skipped - Line: {0} - Description: {1}", lineNo, parts[2]);
                        errorList.Items.Add(error);
                        continue;
                    }

                    //string[] idParts = parts[2].Split(lineDelSpace, StringSplitOptions.None);
                    //string[] idParts = parts[2].Split(dash, StringSplitOptions.None);
                    var resultArray = Regex.Split(parts[2], @"[^0-9\.]+").Where(c => c != "." && c.Trim() != "");
                    string[] idParts = resultArray.ToArray();
                    int nIdParts = idParts.Count();

                    //string istic = idParts[0];

                    // make sure there are parts to parse
                    int cnt = idParts.Length;
                    //if (cnt < 2)
                    if (cnt < 1)
                    {
                        string error = String.Format("Error: Cannot parse row - Line: {0} - Description: {1}", lineNo, parts[2]);
                        errorList.Items.Add(error);
                        continue;
                    }

                    string istic = idParts[0];

                    // make sure it's a valid characteristic that we can parse
                    double dNum;
                    bool isNum = double.TryParse(idParts[0], out dNum);
                    if (!isNum)
                    {
                        string error = String.Format("Error: Calypso description missing hyphen after characteristic number - Line: {0} - Description: {1}", lineNo, parts[2]);
                        errorList.Items.Add(error);
                        continue;
                    }

                    //if (idParts[1] != "-")
                    //{
                    //    double dNum;
                    //    bool isNum = double.TryParse(idParts[0], out dNum);
                    //    if (!isNum)
                    //    {
                    //        string error = String.Format("Error: Calypso description missing hyphen after characteristic number - Line: {0} - Description: {1}", lineNo, parts[2]);
                    //        errorList.Items.Add(error);
                    //        continue;
                    //    }
                    //}

                    if (charInfos.ContainsKey(istic) == false)
                    {
                        //string error = String.Format("Error: Calypso characteristic ({2}) not in ballooned drawing - Line: {0} - Description: {1}", lineNo, parts[2], istic);
                        string error = String.Format("Error: Calypso characteristic ({2}) not in Solidworks spreadsheet - Line: {0} - Description: {1}", lineNo, parts[2], istic);
                        errorList.Items.Add(error);
                        continue;
                    }

                    //if (charInfos[istic].IsUsed())
                    //{
                    //    string error = String.Format("Warning: Calypso characteristic ({2}) already used - Line: {0} - Description: {1}", lineNo, parts[2], istic);
                    //    errorList.Items.Add(error);
                    //    continue;
                    //}

                    // save run number
                    charInfos[istic].AddRun(parts[1]);

                    // save the decription
                    charInfos[istic].AddDesc(parts[2]);

                    string actual = String.Empty;
                    double actualVal = 0.0;
                    string deviation = String.Empty;
                    double deviationVal = 0.0;

                    if (parts[7].Contains("999") && parts[8].Contains("999"))
                    {
                        // we have a basic dimension
                        string x = parts[8];
                        string y = parts[9];
                        charInfos[istic].SetBasic(true);
                    }

                    // Actual
                    if (parts[5].Length > 0)
                    {
                        actual = parts[5];

                        // determine if this is a degree/minute/sec measure
                        if (actual.Contains(':'))
                        {
                            string[] dms = actual.Split(colon, StringSplitOptions.None);
                            int deg = Convert.ToInt16(dms[0]);
                            int min = Convert.ToInt16(dms[1]);
                            int sec = Convert.ToInt16(dms[2]);

                            actualVal = Math.Abs((1.0 * deg) + (1.0 * min / 60.0) + (1.0 * sec / 3600.0));
                        }
                        else
                        {
                            actualVal = Math.Abs(Convert.ToDouble(actual));
                        }
                        charInfos[istic].AddAct(actualVal);
                        charInfos[istic].AddActual(actualVal);
                    }

                    // Exceed
                    if (parts[10].Length > 0)
                    {
                        deviation = parts[10];
                        deviationVal = Math.Abs(Convert.ToDouble(deviation));

                        charInfos[istic].AddDev(deviationVal);
                        charInfos[istic].AddDeviation(deviationVal);
                    }
                    else
                    {
                        charInfos[istic].AddDev(-999.00);
                    }

                    // Comment
                    string comment = parts[14];

                    charInfos[istic].MarkUsed();

                }

                SrcInput.Close();

                ///////////////////////////////////////////////////////////////////////////////////

                foreach (charInfo cInfo in charInfos.Values)
                {
                    if (cInfo.run.Length == 0)
                    {
                        if (cInfo.isCMM)
                        {
                            //string text = String.Format("Error - Characteristic in spreadsheet, but missing from txt data: CharNo={0}, Nominal={1}", cInfo.charNo, cInfo.nominal);
                            string text = String.Format("Error - Characteristic in Solidworks spreadsheet, but missing from Calypso chr data: CharNo={0}, Nominal={1}", cInfo.charNo, cInfo.nominal);
                            errorList.Items.Add(text);
                        }
                        continue;
                    }

                    if (currentRef == null || currentRef != cInfo.run)
                    {
                        currentRef = cInfo.run;
                        currentSheet = string.Format("Form3_{0}", currentRef);

                        grid = new resultGrid(title, currentRef);
                        tabPage = new TabPage(currentSheet);
                        grid.tabPage = tabPage; // saving a copy of this grids parent tabPage
                        tabPage.Controls.Add(grid.gridControl);
                        tabPage.Controls.Add(grid.label);
                        tabPage.Controls.Add(grid.fileName);
                        tabControl1.TabPages.Add(tabPage);
                        //fileToGrid.Add(title, grid);
                        try
                        {
                            fileToGrid.Add(title, grid);
                        }
                        catch (Exception ex)
                        {
                            errorList.Items.Add(string.Format("Error: {0}. Run# {1}, Characteristic {2}",
                                ex.Message, currentRef, cInfo.charNo));
                            continue;
                        }
                    }

                    DataGridViewRow row = (DataGridViewRow)grid.gridControl.Rows[0].Clone();

                    row.DefaultCellStyle = new DataGridViewCellStyle()
                    {
                        //Font = new System.Drawing.Font("InspectionXpert GDT", 9.5F)
                        Font = newFont
                    };

                    row.Cells[0].Value = cInfo.charNo;
                    row.Cells[1].Value = cInfo.nominal;
                    row.Cells[2].Value = cInfo.upper;
                    row.Cells[3].Value = cInfo.lower;
                    row.Cells[4].Value = cInfo.actVal.ToString(fontFormat);
                    row.Cells[5].Value = (cInfo.devVal == -999.00) ? "" : row.Cells[5].Value = cInfo.devVal;

                    grid.gridControl.Rows.Add(row);
                }
            }
            else
            {
                // unchecked
                
                resultGrid grid = fileToGrid[title];
                TabPage tabPage = grid.tabPage;
                tabControl1.TabPages.Remove(tabPage);
            }
        }

        private void statsTabBtn_Click(object sender, EventArgs e)
        {
            //statsGrid grid = new statsGrid("Statistical Results");
            statsGrid = new statsGrid("Statistical Results");
            TabPage tabPage = new TabPage("Statistical Results");
            statsGrid.tabPage = tabPage; // saving a copy of this grids parent tabPage
            tabPage.Controls.Add(statsGrid.gridControl);
            tabControl1.TabPages.Add(tabPage);

            foreach (charInfo info in charInfos.Values)
            {
                if (info.actualList.Count > 0)
                {
                    DataGridViewRow row = (DataGridViewRow)statsGrid.gridControl.Rows[0].Clone();

                    row.DefaultCellStyle = new DataGridViewCellStyle()
                    {
                        //Font = new System.Drawing.Font("InspectionXpert GDT", 9.5F)
                        Font = newFont
                    };

                    double mean = info.actualList.Mean();
                    double variance = info.actualList.Variance();
                    double sd = info.actualList.StandardDeviation();

                    row.Cells[0].Value = info.charNo;
                    row.Cells[1].Value = info.nominal;
                    row.Cells[2].Value = mean.ToString("F4");
                    row.Cells[3].Value = variance.ToString("F4");
                    row.Cells[4].Value = sd.ToString("F4");

                    statsGrid.gridControl.Rows.Add(row);
                }
                //else
                //{
                //    row.Cells[3].Value = "";
                //}

                //statsGrid.gridControl.Rows.Add(row);
            }

            statsIncluded = true;
        }


        private const long BUFFER_SIZE = 4096;

        public static void AddFileToZip(string zipFilename, string fileToAdd)
        {
            using (Package zip = System.IO.Packaging.Package.Open(zipFilename, FileMode.OpenOrCreate))
            {
                string destFilename = ".\\" + Path.GetFileName(fileToAdd);
                Uri uri = PackUriHelper.CreatePartUri(new Uri(destFilename, UriKind.Relative));
                if (zip.PartExists(uri))
                {
                    zip.DeletePart(uri);
                }
                PackagePart part = zip.CreatePart(uri, "", CompressionOption.Normal);
                using (FileStream fileStream = new FileStream(fileToAdd, FileMode.Open, FileAccess.Read))
                {
                    using (Stream dest = part.GetStream())
                    {
                        CopyStream(fileStream, dest);
                    }
                }
            }
        }

        public static void CopyStream(System.IO.FileStream inputStream, System.IO.Stream outputStream)
        {
            long bufferSize = inputStream.Length < BUFFER_SIZE ? inputStream.Length : BUFFER_SIZE;
            byte[] buffer = new byte[bufferSize];
            int bytesRead = 0;
            long bytesWritten = 0;
            while ((bytesRead = inputStream.Read(buffer, 0, buffer.Length)) != 0)
            {
                outputStream.Write(buffer, 0, bytesRead);
                bytesWritten += bufferSize;
            }
        }

        private void copyErrors_Click(object sender, EventArgs e)
        {
            string items = "";

            // first select all items
            for (int i = 0; i < errorList.Items.Count; i++)
            {
                errorList.SetSelected(i, true);
            }

            foreach (int index in errorList.SelectedIndices)
            {
                items += (errorList.Items[index] as string) + '\n';
            }

            Clipboard.SetText(items.Trim()); // trim to remove the remaining '\n'

            // now unselect
            for (int i = 0; i < errorList.Items.Count; i++)
            {
                errorList.SetSelected(i, false);
            }
        }
    }

    public static class MyListExtensions
    {
        public static double Mean(this List<double> values)
        {
            return values.Count == 0 ? 0 : values.Mean(0, values.Count);
        }

        public static double Mean(this List<double> values, int start, int end)
        {
            double s = 0;

            for (int i = start; i < end; i++)
            {
                s += values[i];
            }

            return s / (end - start);
        }

        public static double Variance(this List<double> values)
        {
            return values.Variance(values.Mean(), 0, values.Count);
        }

        public static double Variance(this List<double> values, double mean)
        {
            return values.Variance(mean, 0, values.Count);
        }

        public static double Variance(this List<double> values, double mean, int start, int end)
        {
            double variance = 0;

            for (int i = start; i < end; i++)
            {
                variance += Math.Pow((values[i] - mean), 2);
            }

            int n = end - start;
            if (start > 0) n -= 1;

            return variance / (n);
        }

        public static double StandardDeviation(this List<double> values)
        {
            return values.Count == 0 ? 0 : values.StandardDeviation(0, values.Count);
        }

        public static double StandardDeviation(this List<double> values, int start, int end)
        {
            double mean = values.Mean(start, end);
            double variance = values.Variance(mean, start, end);

            return Math.Sqrt(variance);
        }

    }
}

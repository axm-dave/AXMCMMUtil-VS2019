using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Drawing;

//using Excel=Microsoft.Office.Interop.Excel;

namespace AXMCMMUtil
{
    public partial class Form1 : Form
    {
        //
        // this forms Globals
        //

        public Dictionary<string, charInfo> charInfos = new Dictionary<string, charInfo>();
        public Dictionary<string, resultGrid> grids = new Dictionary<string, resultGrid>();
        public Dictionary<string, bool> resultFiles = new Dictionary<string, bool>();
        SLDocument sl = null;
        statsGrid statsGrid = null;
        bool statsIncluded = false;

        public double actualTotals = 0.0;
        public double actualStdDev = 0.0;
        public int actualCnt = 0;

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

        public Form1()
        {
            InitializeComponent();

            fontName.Text = "InspectionXpert GDT";
            fontSize.Text = "8";
            formatCode.Text = "#,##0.000";
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
            openFileDialog2.Filter = "xlsx files (*.xlsx)|*.xlsx";
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
                    //Excel.Workbook theWorkbook = ExcelObj.Workbooks.Open(
                    //    fullTemplateName, 0, true, 5,
                    //"", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
                    //0, true); 
                    //Excel.Sheets sheets = theWorkbook.Worksheets;
                    //Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(3);
                    //for (int i = 1; i <= 10; i++)
                    //{
                    //    Excel.Range range = worksheet.get_Range("A" + i.ToString(), "J" + i.ToString());
                    //    System.Array myvalues = (System.Array)range.Cells.Value;
                    //    //string[] strArray = ConvertToStringArray(myvalues);
                    //}

                    //////////////////////////////////



                    //SLDocument sl = new SLDocument(fullTemplateName, "Form3");
                    sl = new SLDocument(fullTemplateName, "Form3");

                    bool bGotChar = true;
                    string cellVal = String.Empty;
                    int nCol = 2;
                    int nRow = 8;

                    while (bGotChar)
                    {
                        cellVal = sl.GetCellValueAsString(nRow, nCol);
                        if (cellVal.Contains("The ") || cellVal.Length <= 0)
                        {
                            bGotChar = false;
                        }
                        else
                        {
                            string charNo = sl.GetCellValueAsString(nRow, 2);
                            string nominal = sl.GetCellValueAsString(nRow, 5);
                            string upper = sl.GetCellValueAsString(nRow, 7);
                            string lower = sl.GetCellValueAsString(nRow, 8);
                            //charInfo info = new charInfo(nRow, charNo, nominal, upper, lower);
                            charInfo info = new charInfo(Convert.ToDouble(cellVal), nRow, charNo, nominal, upper, lower);
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
            //int itemCnt = 0;

            string fullTemplateName = Path.Combine(Properties.Settings.Default.LastTemplatePath, sourceTemplateInput.Text);

            bRC = sl.SelectWorksheet("Form1");
            string partNo = sl.GetCellValueAsString(4, 2);
            string partName = sl.GetCellValueAsString(4, 4);
            string reportNo = sl.GetCellValueAsString(4, 8);
            string serialNo = String.Empty;

            //SLDocument sl = new SLDocument(fullTemplateName, "Form1");

            SLStyle styleFloat = sl.CreateStyle();
            styleFloat.FormatCode = formatCode.Text;
            styleFloat.SetFontColor(System.Drawing.Color.Black);
            styleFloat.Font.FontName = fontName.Text;
            styleFloat.Font.FontSize = Convert.ToInt16(fontSize.Text);

            SLStyle styleString = sl.CreateStyle();
            styleString.SetFontColor(System.Drawing.Color.Black);
            styleString.Font.FontName = fontName.Text;
            styleString.Font.FontSize = Convert.ToInt16(fontSize.Text);

            foreach (resultGrid grid in grids.Values)
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
                sl.SetCellValue(4, 12, serialNo);


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

                    // Exceed
                    double exceed = 0;
                    if (deviation.Length > 0)
                    {
                        exceed = Math.Abs(Convert.ToDouble(deviation));
                        sl.SetCellValue(nRow, 11, exceed);
                        sl.SetCellStyle(nRow, 11, styleFloat);
                    }
                }
            }

            if (statsIncluded)
            {
                SLPageSettings ps = sl.GetPageSettings("CMMStats");

                bRC = sl.SelectWorksheet("CMMStats");
                if (bRC)
                {
                    sl.SetCellValue(4, 2, partNo);
                    sl.SetCellValue(4, 6, partName);
                    sl.SetCellValue(4, 12, serialNo);
                    sl.SetCellValue(4, 13, reportNo);

                    //styleFloat.Border.LeftBorder.BorderStyle = BorderStyleValues.Medium;
                    //styleFloat.Border.RightBorder.BorderStyle = BorderStyleValues.Medium;
                    //styleFloat.Border.BottomBorder.BorderStyle = BorderStyleValues.Medium;

                    int nRow = 8;

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

                        if (nRow > 28)
                        {
                            sl.InsertRow(nRow, 1);
                            sl.MergeWorksheetCells(nRow, 3, nRow, 9);
                            sl.CopyRowStyle(8, nRow);
                        }

                        sl.SetCellValue(nRow, 2, charNo);
                        sl.SetCellStyle(nRow, 2, sl.GetCellStyle("B8"));
                        sl.SetCellValue(nRow, 3, desc);
                        sl.SetCellStyle(nRow, 3, sl.GetCellStyle("C8"));
                        sl.SetCellStyle(nRow, 4, sl.GetCellStyle("D8"));
                        sl.SetCellStyle(nRow, 5, sl.GetCellStyle("E8"));
                        sl.SetCellStyle(nRow, 6, sl.GetCellStyle("F8"));
                        sl.SetCellStyle(nRow, 7, sl.GetCellStyle("G8"));
                        sl.SetCellStyle(nRow, 8, sl.GetCellStyle("H8"));
                        sl.SetCellStyle(nRow, 9, sl.GetCellStyle("I8"));
                        sl.SetCellValue(nRow, 10, nominal);
                        sl.SetCellStyle(nRow, 10, sl.GetCellStyle("J8"));
                        sl.SetCellValue(nRow, 11, mean);
                        sl.SetCellStyle(nRow, 11, sl.GetCellStyle("K8"));
                        sl.SetCellValue(nRow, 12, variance);
                        sl.SetCellStyle(nRow, 12, sl.GetCellStyle("L8"));
                        sl.SetCellValue(nRow, 13, stdDev);
                        sl.SetCellStyle(nRow, 13, sl.GetCellStyle("M8"));

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
            int lineNo = 0;
            string currentRef = null;
            string currentSheet = String.Empty;
            char[] lineDelimiter = { '\t' };
            char[] lineDelSpace = { ' ' };
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

                if (grids.ContainsKey(title))
                {
                    return;
                }

                resultGrid grid = null;
                TabPage tabPage = null;

                // add data to grid view
                string fullItemName = Path.Combine(Properties.Settings.Default.LastResultsPath, title);
                string fileName = Path.GetFileNameWithoutExtension(fullItemName);
                StreamReader SrcInput = new StreamReader(fullItemName);

                //string sLine;
                //int lineNo = 0;
                //string currentRef = null;
                //string currentSheet = String.Empty;
                //char[] lineDelimiter = { '\t' };
                //char[] lineDelSpace = { ' ' };
                //char[] colon = { ':' };
                //char[] period = { '.' };

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

                    string[] idParts = parts[2].Split(lineDelSpace, StringSplitOptions.None);
                    string istic = idParts[0];

                    // make sure there are parts to parse
                    int cnt = idParts.Length;
                    if (cnt < 2)
                    {
                        string error = String.Format("Error: Can not parse row - Line: {0} - Description: {1}", lineNo, parts[2]);
                        errorList.Items.Add(error);
                        continue;
                    }

                    // make sure it's a valid characteristic that we can parse
                    if (idParts[1] != "-")
                    {
                        string error = String.Format("Error: Calypso description missing hypen after characteristic number - Line: {0} - Description: {1}", lineNo, parts[2]);
                        errorList.Items.Add(error);
                        continue;
                    }
                    if (charInfos.ContainsKey(istic) == false)
                    {
                        string error = String.Format("Error: Calypso characteristic ({2}) not in ballooned drawing - Line: {0} - Description: {1}", lineNo, parts[2], istic);
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
                        string text = String.Format("Error - Characteristic in spreadsheet, but missing from txt data: CharNo={0}, Nominal={1}", cInfo.charNo, cInfo.nominal);
                        errorList.Items.Add(text);
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
                        grids.Add(currentSheet, grid);
                    }

                    DataGridViewRow row = (DataGridViewRow)grid.gridControl.Rows[0].Clone();

                    row.DefaultCellStyle = new DataGridViewCellStyle()
                    {
                        Font = new System.Drawing.Font("InspectionXpert GDT", 9.5F)
                    };

                    row.Cells[0].Value = cInfo.charNo;
                    row.Cells[1].Value = cInfo.nominal;
                    row.Cells[2].Value = cInfo.upper;
                    row.Cells[3].Value = cInfo.lower;
                    row.Cells[4].Value = cInfo.actVal.ToString(formatCode.Text);
                    row.Cells[5].Value = (cInfo.devVal == -999.00) ? "" : row.Cells[5].Value = cInfo.devVal;

                    grid.gridControl.Rows.Add(row);
                }

                //while ((sLine = SrcInput.ReadLine()) != null)
                //{
                //    string[] parts = sLine.Split(lineDelimiter, StringSplitOptions.None);

                //    if (sLine == "END") break;

                //    lineNo++;

                //    if (currentRef == null || currentRef != parts[1])
                //    {
                //        // first time here
                //        currentRef = parts[1];
                //        currentSheet = string.Format("Form3_{0}", currentRef);

                //        grid = new resultGrid(title, currentRef);
                //        tabPage = new TabPage(currentSheet);
                //        grid.tabPage = tabPage; // saving a copy of this grids parent tabPage
                //        tabPage.Controls.Add(grid.gridControl);
                //        tabPage.Controls.Add(grid.label);
                //        tabPage.Controls.Add(grid.fileName);
                //        tabControl1.TabPages.Add(tabPage);
                //        grids.Add(currentSheet, grid);
                       
                //    }

                //    //if (parts[4] == "gdtPosPol2d.X" || parts[4] == "gdtPosPol2d.Y" || parts[4] == "gdtPosPol2d.Z")
                //    //{
                //    //    string error = String.Format("Note: True Position coordinate skipped - Line: {0} - Description: {1}", lineNo, parts[2]);
                //    //    errorList.Items.Add(error);
                //    //    continue;
                //    //}

                //    int pos = Array.IndexOf(ignoreSet, parts[4]);
                //    if (pos > -1)
                //    {
                //        string error = String.Format("Note: True Position coordinate skipped - Line: {0} - Description: {1}", lineNo, parts[2]);
                //        errorList.Items.Add(error);
                //        continue;
                //    }

                //    string[] idParts = parts[2].Split(lineDelSpace, StringSplitOptions.None);
                //    string istic = idParts[0];

                //    // make sure it's a valid characteristic that we can parse
                //    if (idParts[1] != "-")
                //    {
                //        string error = String.Format("Error: Calypso description missing hypen after characteristic number - Line: {0} - Description: {1}", lineNo, parts[2]);
                //        errorList.Items.Add(error);
                //        continue;
                //    }
                //    if (charInfos.ContainsKey(istic) == false)
                //    {
                //        string error = String.Format("Error: Calypso characteristic ({2}) not in ballooned drawing - Line: {0} - Description: {1}", lineNo, parts[2], istic);
                //        errorList.Items.Add(error);
                //        continue;
                //    }

                //    if (charInfos[istic].IsUsed())
                //    {
                //        string error = String.Format("Warning: Calypso characteristic ({2}) already used - Line: {0} - Description: {1}", lineNo, parts[2], istic);
                //        errorList.Items.Add(error);
                //        continue;
                //    }

                //    // save the decription
                //    charInfos[istic].AddDesc(parts[2]);

                //    int nRow = charInfos[istic].charRow;
                //    string charNo = charInfos[istic].charNo;
                //    string nominal = charInfos[istic].nominal;
                //    string upper = charInfos[istic].upper;
                //    string lower = charInfos[istic].lower;
                //    string actual = String.Empty;
                //    double actualVal = 0.0;
                //    string deviation = String.Empty;
                //    double deviationVal = 0.0;

                //    // Actual
                //    if (parts[5].Length > 0)
                //    {
                //        actual = parts[5];

                //        // determine if this is a degree/minute/sec measure
                //        if (actual.Contains(':'))
                //        {
                //            string[] dms = actual.Split(colon, StringSplitOptions.None);
                //            int deg = Convert.ToInt16(dms[0]);
                //            int min = Convert.ToInt16(dms[1]);
                //            int sec = Convert.ToInt16(dms[2]);

                //            actualVal = Math.Abs((1.0 * deg) + (1.0 * min / 60.0) + (1.0 * sec / 3600.0));
                //        }
                //        else
                //        {
                //            actualVal = Math.Abs(Convert.ToDouble(actual));
                //        }
                //        charInfos[istic].AddActual(actualVal);
                //    }

                //    // Exceed
                //    if (parts[10].Length > 0)
                //    {
                //        deviation = parts[10];
                //        deviationVal = Math.Abs(Convert.ToDouble(deviation));
                //    }

                //    // Comment
                //    string comment = parts[14];

                //    DataGridViewRow row = (DataGridViewRow)grid.gridControl.Rows[0].Clone();

                //    row.DefaultCellStyle = new DataGridViewCellStyle()
                //    {
                //        Font = new System.Drawing.Font("InspectionXpert GDT", 9.5F)
                //    };

                //    if (charNo.Contains('.'))
                //    {
                //        string[] decVals = charNo.Split(period, StringSplitOptions.None);
                //        string fraction = string.Format("{0,2:D2}", Convert.ToInt16(decVals[1]));
                //        row.Cells[0].Value = decVals[0] + '.' + fraction;
                //    }
                //    else
                //    {
                //        row.Cells[0].Value = charNo + ".00";
                //    }

                //    //row.Cells[0].Value = charNo;
                //    //double x = Convert.ToDouble(charNo);
                //    row.Cells[1].Value = nominal;
                //    //row.Cells[1].Style = new DataGridViewCellStyle()
                //    //{
                //    //    Font = new System.Drawing.Font("InspectionXpert GDT", 8F),
                //    //    BackColor = (nominal.Length <= 0) ? System.Drawing.Color.LightYellow : System.Drawing.Color.White
                //    //};
                //    row.Cells[2].Value = upper;
                //    row.Cells[3].Value = lower;
                //    row.Cells[4].Value = actualVal.ToString(formatCode.Text);
                //    row.Cells[5].Value = deviation;

                //    grid.gridControl.Rows.Add(row);
                //    int cnt = grid.gridControl.Rows.Count;
                //    charInfos[istic].MarkUsed();

                //}


                //SrcInput.Close();
            }
            else
            {
                //MessageBox.Show("un-checked");
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
                        Font = new System.Drawing.Font("InspectionXpert GDT", 9.5F)
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

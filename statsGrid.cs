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

namespace AXMCMMUtil
{
    class statsGrid
    {
        public DataGridView gridControl { get; private set; }
        public DataGridViewTextBoxColumn CharNo { get; private set; }
        public DataGridViewTextBoxColumn Nominal { get; private set; }
        public DataGridViewTextBoxColumn Mean { get; private set; }
        public DataGridViewTextBoxColumn Variance { get; private set; }
        public DataGridViewTextBoxColumn StdDev { get; private set; }
        public TabPage tabPage { get; set; }

        public statsGrid(string title)
        {
            gridControl = new System.Windows.Forms.DataGridView();
            CharNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            Nominal = new System.Windows.Forms.DataGridViewTextBoxColumn();
            Mean = new System.Windows.Forms.DataGridViewTextBoxColumn();
            Variance = new System.Windows.Forms.DataGridViewTextBoxColumn();
            StdDev = new System.Windows.Forms.DataGridViewTextBoxColumn();
            tabPage = null;

            // 
            // dataGridView1
            // 
            gridControl.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
                CharNo, Nominal, Mean, Variance, StdDev });
            gridControl.Location = new System.Drawing.Point(0, 0);
            gridControl.Name = "dataGridView2";
            gridControl.Size = new System.Drawing.Size(643, 567);
            //gridControl.Size = new System.Drawing.Size(600, 500);
            //gridControl.Width = 450;
            //gridControl.Height = 450;
            gridControl.TabIndex = 13;
            gridControl.ScrollBars = ScrollBars.Both;
            //gridControl.AutoSize = true;
            // 
            // CharNo
            // 
            //CharNo.Frozen = true;
            CharNo.HeaderText = "CharNo";
            CharNo.Name = "CharNo";
            CharNo.ReadOnly = true;
            CharNo.Width = 70;
            // 
            // Nominal
            // 
            //Nominal.Frozen = true;
            Nominal.HeaderText = "Nominal";
            Nominal.Name = "Nominal";
            Nominal.ReadOnly = true;
            Nominal.Width = 120;
            // 
            // Mean
            // 
            Mean.HeaderText = "Mean";
            Mean.Name = "Mean";
            Mean.Width = 70;
            // 
            // Variance
            // 
            Variance.HeaderText = "Variance";
            Variance.Name = "Variance";
            Variance.Width = 70;
            // 
            // StdDeviation
            // 
            StdDev.HeaderText = "StdDev";
            StdDev.Name = "StdDev";
            StdDev.Width = 70;
        }

    }
}

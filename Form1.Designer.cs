namespace AXMCMMUtil
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.openSourceFolderDialog = new System.Windows.Forms.OpenFileDialog();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.sourceTemplateInput = new System.Windows.Forms.TextBox();
            this.browseSourceFolder = new System.Windows.Forms.Button();
            this.browseSourceTemplate = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.closeButton = new System.Windows.Forms.Button();
            this.mergeButton = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.targetInput = new System.Windows.Forms.TextBox();
            this.browseTargetLocation = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.statsTabBtn = new System.Windows.Forms.Button();
            this.errorList = new System.Windows.Forms.ListBox();
            this.label7 = new System.Windows.Forms.Label();
            this.copyErrors = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // openSourceFolderDialog
            // 
            this.openSourceFolderDialog.FileName = "openFileDialog1";
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.CheckOnClick = true;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.HorizontalScrollbar = true;
            this.checkedListBox1.Location = new System.Drawing.Point(12, 79);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(398, 319);
            this.checkedListBox1.TabIndex = 1;
            this.checkedListBox1.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkedListBox1_ItemCheck);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(13, 60);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(212, 16);
            this.label1.TabIndex = 2;
            this.label1.Text = "Source Characteristic Results";
            // 
            // sourceTemplateInput
            // 
            this.sourceTemplateInput.Location = new System.Drawing.Point(14, 35);
            this.sourceTemplateInput.Name = "sourceTemplateInput";
            this.sourceTemplateInput.Size = new System.Drawing.Size(397, 20);
            this.sourceTemplateInput.TabIndex = 3;
            // 
            // browseSourceFolder
            // 
            this.browseSourceFolder.Location = new System.Drawing.Point(426, 82);
            this.browseSourceFolder.Name = "browseSourceFolder";
            this.browseSourceFolder.Size = new System.Drawing.Size(95, 31);
            this.browseSourceFolder.TabIndex = 4;
            this.browseSourceFolder.Text = "Open...";
            this.browseSourceFolder.UseVisualStyleBackColor = true;
            this.browseSourceFolder.Click += new System.EventHandler(this.browseSourceFolder_Click);
            // 
            // browseSourceTemplate
            // 
            this.browseSourceTemplate.Location = new System.Drawing.Point(426, 29);
            this.browseSourceTemplate.Name = "browseSourceTemplate";
            this.browseSourceTemplate.Size = new System.Drawing.Size(95, 31);
            this.browseSourceTemplate.TabIndex = 5;
            this.browseSourceTemplate.Text = "Open...";
            this.browseSourceTemplate.UseVisualStyleBackColor = true;
            this.browseSourceTemplate.Click += new System.EventHandler(this.browseSourceTemplate_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(15, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(167, 16);
            this.label2.TabIndex = 6;
            this.label2.Text = "Template Spreadsheet";
            // 
            // closeButton
            // 
            this.closeButton.Location = new System.Drawing.Point(1098, 773);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(95, 31);
            this.closeButton.TabIndex = 7;
            this.closeButton.Text = "Close";
            this.closeButton.UseVisualStyleBackColor = true;
            this.closeButton.Click += new System.EventHandler(this.closeButton_Click);
            // 
            // mergeButton
            // 
            this.mergeButton.Location = new System.Drawing.Point(997, 773);
            this.mergeButton.Name = "mergeButton";
            this.mergeButton.Size = new System.Drawing.Size(95, 31);
            this.mergeButton.TabIndex = 8;
            this.mergeButton.Text = "Create FAI";
            this.mergeButton.UseVisualStyleBackColor = true;
            this.mergeButton.Click += new System.EventHandler(this.mergeButton_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(15, 408);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(148, 16);
            this.label6.TabIndex = 11;
            this.label6.Text = "Report Spreadsheet";
            // 
            // targetInput
            // 
            this.targetInput.Location = new System.Drawing.Point(11, 427);
            this.targetInput.Name = "targetInput";
            this.targetInput.Size = new System.Drawing.Size(397, 20);
            this.targetInput.TabIndex = 10;
            // 
            // browseTargetLocation
            // 
            this.browseTargetLocation.Location = new System.Drawing.Point(426, 421);
            this.browseTargetLocation.Name = "browseTargetLocation";
            this.browseTargetLocation.Size = new System.Drawing.Size(95, 31);
            this.browseTargetLocation.TabIndex = 12;
            this.browseTargetLocation.Text = "Folder...";
            this.browseTargetLocation.UseVisualStyleBackColor = true;
            this.browseTargetLocation.Click += new System.EventHandler(this.browseTargetLocation_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Location = new System.Drawing.Point(542, 16);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(651, 593);
            this.tabControl1.TabIndex = 14;
            // 
            // statsTabBtn
            // 
            this.statsTabBtn.Location = new System.Drawing.Point(426, 134);
            this.statsTabBtn.Name = "statsTabBtn";
            this.statsTabBtn.Size = new System.Drawing.Size(95, 31);
            this.statsTabBtn.TabIndex = 15;
            this.statsTabBtn.Text = "Stats Tab";
            this.statsTabBtn.UseVisualStyleBackColor = true;
            this.statsTabBtn.Click += new System.EventHandler(this.statsTabBtn_Click);
            // 
            // errorList
            // 
            this.errorList.FormattingEnabled = true;
            this.errorList.Location = new System.Drawing.Point(10, 615);
            this.errorList.Name = "errorList";
            this.errorList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.errorList.Size = new System.Drawing.Size(1183, 147);
            this.errorList.TabIndex = 16;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(15, 593);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(99, 16);
            this.label7.TabIndex = 17;
            this.label7.Text = "Status Output";
            // 
            // copyErrors
            // 
            this.copyErrors.Location = new System.Drawing.Point(421, 578);
            this.copyErrors.Name = "copyErrors";
            this.copyErrors.Size = new System.Drawing.Size(100, 31);
            this.copyErrors.TabIndex = 18;
            this.copyErrors.Text = "Copy Status";
            this.copyErrors.UseVisualStyleBackColor = true;
            this.copyErrors.Click += new System.EventHandler(this.copyErrors_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(7, 773);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(113, 13);
            this.label10.TabIndex = 21;
            this.label10.Text = "Version: {0}.{1}.{2}.{3}";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1205, 811);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.copyErrors);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.errorList);
            this.Controls.Add(this.statsTabBtn);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.browseTargetLocation);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.targetInput);
            this.Controls.Add(this.mergeButton);
            this.Controls.Add(this.closeButton);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.browseSourceTemplate);
            this.Controls.Add(this.browseSourceFolder);
            this.Controls.Add(this.sourceTemplateInput);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.checkedListBox1);
            this.Name = "Form1";
            this.Text = "AXM CMM Merge Results Utility";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openSourceFolderDialog;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox sourceTemplateInput;
        private System.Windows.Forms.Button browseSourceFolder;
        private System.Windows.Forms.Button browseSourceTemplate;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button closeButton;
        private System.Windows.Forms.Button mergeButton;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox targetInput;
        private System.Windows.Forms.Button browseTargetLocation;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.Button statsTabBtn;
        private System.Windows.Forms.ListBox errorList;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button copyErrors;
        private System.Windows.Forms.Label label10;
    }
}


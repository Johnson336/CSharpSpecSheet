namespace CSharpSpecSheet
{
    partial class AdminMenu
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
            this.frameSQLStmt = new System.Windows.Forms.GroupBox();
            this.txtSQLStmt = new System.Windows.Forms.TextBox();
            this.buttonSQLExecute = new System.Windows.Forms.Button();
            this.frameSQLResults = new System.Windows.Forms.GroupBox();
            this.txtSQLResults = new System.Windows.Forms.TextBox();
            this.frameSQLStmt.SuspendLayout();
            this.frameSQLResults.SuspendLayout();
            this.SuspendLayout();
            // 
            // frameSQLStmt
            // 
            this.frameSQLStmt.Controls.Add(this.buttonSQLExecute);
            this.frameSQLStmt.Controls.Add(this.txtSQLStmt);
            this.frameSQLStmt.Location = new System.Drawing.Point(12, 12);
            this.frameSQLStmt.Name = "frameSQLStmt";
            this.frameSQLStmt.Size = new System.Drawing.Size(560, 46);
            this.frameSQLStmt.TabIndex = 0;
            this.frameSQLStmt.TabStop = false;
            this.frameSQLStmt.Text = "Enter SQL Statement";
            // 
            // txtSQLStmt
            // 
            this.txtSQLStmt.Location = new System.Drawing.Point(3, 16);
            this.txtSQLStmt.Name = "txtSQLStmt";
            this.txtSQLStmt.Size = new System.Drawing.Size(470, 20);
            this.txtSQLStmt.TabIndex = 0;
            // 
            // buttonSQLExecute
            // 
            this.buttonSQLExecute.Location = new System.Drawing.Point(479, 14);
            this.buttonSQLExecute.Name = "buttonSQLExecute";
            this.buttonSQLExecute.Size = new System.Drawing.Size(75, 23);
            this.buttonSQLExecute.TabIndex = 1;
            this.buttonSQLExecute.Text = "Execute";
            this.buttonSQLExecute.UseVisualStyleBackColor = true;
            this.buttonSQLExecute.Click += new System.EventHandler(this.buttonSQLExecute_Click);
            // 
            // frameSQLResults
            // 
            this.frameSQLResults.Controls.Add(this.txtSQLResults);
            this.frameSQLResults.Location = new System.Drawing.Point(15, 64);
            this.frameSQLResults.Name = "frameSQLResults";
            this.frameSQLResults.Size = new System.Drawing.Size(557, 185);
            this.frameSQLResults.TabIndex = 1;
            this.frameSQLResults.TabStop = false;
            this.frameSQLResults.Text = "SQL Results";
            // 
            // txtSQLResults
            // 
            this.txtSQLResults.AcceptsReturn = true;
            this.txtSQLResults.Location = new System.Drawing.Point(3, 16);
            this.txtSQLResults.Multiline = true;
            this.txtSQLResults.Name = "txtSQLResults";
            this.txtSQLResults.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtSQLResults.Size = new System.Drawing.Size(548, 163);
            this.txtSQLResults.TabIndex = 0;
            this.txtSQLResults.WordWrap = false;
            // 
            // AdminMenu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 261);
            this.Controls.Add(this.frameSQLResults);
            this.Controls.Add(this.frameSQLStmt);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(600, 300);
            this.Name = "AdminMenu";
            this.Text = "AdminMenu";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.AdminMenu_FormClosing);
            this.frameSQLStmt.ResumeLayout(false);
            this.frameSQLStmt.PerformLayout();
            this.frameSQLResults.ResumeLayout(false);
            this.frameSQLResults.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox frameSQLStmt;
        private System.Windows.Forms.Button buttonSQLExecute;
        private System.Windows.Forms.TextBox txtSQLStmt;
        private System.Windows.Forms.GroupBox frameSQLResults;
        private System.Windows.Forms.TextBox txtSQLResults;
    }
}
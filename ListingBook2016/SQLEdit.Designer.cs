namespace ListingBook2016
{
    partial class SQLEdit
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SQLEdit));
            this.panel1 = new System.Windows.Forms.Panel();
            this.progressBarGetData = new System.Windows.Forms.ProgressBar();
            this.buttonCloseTp = new System.Windows.Forms.Button();
            this.buttonGetData = new System.Windows.Forms.Button();
            this.richTextBoxSQLEdit = new System.Windows.Forms.RichTextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.textBoxCS = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.progressBarGetData);
            this.panel1.Controls.Add(this.buttonCloseTp);
            this.panel1.Controls.Add(this.buttonGetData);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 106);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(606, 30);
            this.panel1.TabIndex = 1;
            // 
            // progressBarGetData
            // 
            this.progressBarGetData.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progressBarGetData.Location = new System.Drawing.Point(130, 0);
            this.progressBarGetData.Name = "progressBarGetData";
            this.progressBarGetData.Size = new System.Drawing.Size(346, 30);
            this.progressBarGetData.TabIndex = 7;
            // 
            // buttonCloseTp
            // 
            this.buttonCloseTp.Dock = System.Windows.Forms.DockStyle.Left;
            this.buttonCloseTp.Location = new System.Drawing.Point(0, 0);
            this.buttonCloseTp.Name = "buttonCloseTp";
            this.buttonCloseTp.Size = new System.Drawing.Size(130, 30);
            this.buttonCloseTp.TabIndex = 6;
            this.buttonCloseTp.Text = "Close Task Pane";
            this.buttonCloseTp.UseVisualStyleBackColor = true;
            // 
            // buttonGetData
            // 
            this.buttonGetData.Dock = System.Windows.Forms.DockStyle.Right;
            this.buttonGetData.Location = new System.Drawing.Point(476, 0);
            this.buttonGetData.Name = "buttonGetData";
            this.buttonGetData.Size = new System.Drawing.Size(130, 30);
            this.buttonGetData.TabIndex = 4;
            this.buttonGetData.Text = "Get Data from Server";
            this.buttonGetData.UseVisualStyleBackColor = true;
            this.buttonGetData.Click += new System.EventHandler(this.buttonGetData_Click);
            // 
            // richTextBoxSQLEdit
            // 
            this.richTextBoxSQLEdit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBoxSQLEdit.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBoxSQLEdit.Location = new System.Drawing.Point(0, 0);
            this.richTextBoxSQLEdit.Name = "richTextBoxSQLEdit";
            this.richTextBoxSQLEdit.Size = new System.Drawing.Size(606, 106);
            this.richTextBoxSQLEdit.TabIndex = 2;
            this.richTextBoxSQLEdit.Text = resources.GetString("richTextBoxSQLEdit.Text");
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.textBoxCS);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 84);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(606, 22);
            this.panel2.TabIndex = 3;
            // 
            // textBoxCS
            // 
            this.textBoxCS.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxCS.Location = new System.Drawing.Point(176, 0);
            this.textBoxCS.Name = "textBoxCS";
            this.textBoxCS.Size = new System.Drawing.Size(430, 20);
            this.textBoxCS.TabIndex = 1;
            this.textBoxCS.Text = "Data Source=PQ-WORKSTATION;Initial Catalog=Northwind;Integrated Security=True";
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Left;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(176, 22);
            this.label1.TabIndex = 0;
            this.label1.Text = "Type your Connection String :";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // SQLEdit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.richTextBoxSQLEdit);
            this.Controls.Add(this.panel1);
            this.Name = "SQLEdit";
            this.Size = new System.Drawing.Size(606, 136);
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button buttonCloseTp;
        private System.Windows.Forms.Button buttonGetData;
        private System.Windows.Forms.RichTextBox richTextBoxSQLEdit;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox textBoxCS;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar progressBarGetData;
    }
}

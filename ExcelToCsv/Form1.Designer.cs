namespace ExcelToCsv
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
            this.Convert = new System.Windows.Forms.Button();
            this.BatchConvert = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Convert
            // 
            this.Convert.Location = new System.Drawing.Point(12, 12);
            this.Convert.Name = "Convert";
            this.Convert.Size = new System.Drawing.Size(112, 28);
            this.Convert.TabIndex = 0;
            this.Convert.Text = "Convert";
            this.Convert.UseVisualStyleBackColor = true;
            this.Convert.Click += new System.EventHandler(this.Convert_Click);
            // 
            // BatchConvert
            // 
            this.BatchConvert.Location = new System.Drawing.Point(12, 46);
            this.BatchConvert.Name = "BatchConvert";
            this.BatchConvert.Size = new System.Drawing.Size(112, 31);
            this.BatchConvert.TabIndex = 1;
            this.BatchConvert.Text = "Batch";
            this.BatchConvert.UseVisualStyleBackColor = true;
            this.BatchConvert.Click += new System.EventHandler(this.BatchConvert_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(135, 84);
            this.Controls.Add(this.BatchConvert);
            this.Controls.Add(this.Convert);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Convert;
        private System.Windows.Forms.Button BatchConvert;
    }
}


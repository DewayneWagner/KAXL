namespace DKARibbon
{
    partial class frmCurrencyConvert
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
            this.btn_ConvertData = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txt_ExRateCADtoUSD = new System.Windows.Forms.TextBox();
            this.btn_Exit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn_ConvertData
            // 
            this.btn_ConvertData.Location = new System.Drawing.Point(302, 8);
            this.btn_ConvertData.Name = "btn_ConvertData";
            this.btn_ConvertData.Size = new System.Drawing.Size(118, 23);
            this.btn_ConvertData.TabIndex = 0;
            this.btn_ConvertData.Text = "Convert Selection";
            this.btn_ConvertData.UseVisualStyleBackColor = true;
            this.btn_ConvertData.Click += new System.EventHandler(this.btn_ConvertData_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(201, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Enter today Exchange Rate, CAD - USD:";
            // 
            // txt_ExRateCADtoUSD
            // 
            this.txt_ExRateCADtoUSD.Location = new System.Drawing.Point(220, 10);
            this.txt_ExRateCADtoUSD.Name = "txt_ExRateCADtoUSD";
            this.txt_ExRateCADtoUSD.Size = new System.Drawing.Size(63, 20);
            this.txt_ExRateCADtoUSD.TabIndex = 2;
            // 
            // btn_Exit
            // 
            this.btn_Exit.Location = new System.Drawing.Point(426, 8);
            this.btn_Exit.Name = "btn_Exit";
            this.btn_Exit.Size = new System.Drawing.Size(75, 23);
            this.btn_Exit.TabIndex = 3;
            this.btn_Exit.Text = "Exit";
            this.btn_Exit.UseVisualStyleBackColor = true;
            this.btn_Exit.Click += new System.EventHandler(this.btn_Exit_Click);
            // 
            // frmCurrencyConvert
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(514, 40);
            this.Controls.Add(this.btn_Exit);
            this.Controls.Add(this.txt_ExRateCADtoUSD);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_ConvertData);
            this.Name = "frmCurrencyConvert";
            this.Text = "Currency Conversion - CAD to USD";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_ConvertData;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txt_ExRateCADtoUSD;
        private System.Windows.Forms.Button btn_Exit;
    }
}
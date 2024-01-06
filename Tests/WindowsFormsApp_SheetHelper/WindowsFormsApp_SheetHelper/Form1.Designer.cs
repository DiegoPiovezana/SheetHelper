namespace WindowsFormsApp_SheetHelper
{
    partial class SheetHelper_Menu
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SheetHelper_Menu));
            this.pgBarConvert = new System.Windows.Forms.ProgressBar();
            this.BtnConverter = new System.Windows.Forms.Button();
            this.lblConvertendo = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // pgBarConvert
            // 
            this.pgBarConvert.ForeColor = System.Drawing.Color.Indigo;
            this.pgBarConvert.Location = new System.Drawing.Point(95, 199);
            this.pgBarConvert.Name = "pgBarConvert";
            this.pgBarConvert.Size = new System.Drawing.Size(611, 46);
            this.pgBarConvert.TabIndex = 0;
            // 
            // BtnConverter
            // 
            this.BtnConverter.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnConverter.Location = new System.Drawing.Point(300, 59);
            this.BtnConverter.Name = "BtnConverter";
            this.BtnConverter.Size = new System.Drawing.Size(201, 68);
            this.BtnConverter.TabIndex = 2;
            this.BtnConverter.Text = "CONVERTER";
            this.BtnConverter.UseVisualStyleBackColor = true;
            this.BtnConverter.Click += new System.EventHandler(this.Button1_Click);
            // 
            // lblConvertendo
            // 
            this.lblConvertendo.AutoSize = true;
            this.lblConvertendo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblConvertendo.Location = new System.Drawing.Point(296, 264);
            this.lblConvertendo.Name = "lblConvertendo";
            this.lblConvertendo.Size = new System.Drawing.Size(209, 25);
            this.lblConvertendo.TabIndex = 4;
            this.lblConvertendo.Text = "Convertendo arquivo...";
            this.lblConvertendo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblConvertendo.Visible = false;
            // 
            // SheetHelper_Menu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.lblConvertendo);
            this.Controls.Add(this.BtnConverter);
            this.Controls.Add(this.pgBarConvert);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SheetHelper_Menu";
            this.Text = "SheetHelper Menu";
            this.Load += new System.EventHandler(this.SheetHelper_Menu_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar pgBarConvert;
        private System.Windows.Forms.Button BtnConverter;
        private System.Windows.Forms.Label lblConvertendo;
    }
}
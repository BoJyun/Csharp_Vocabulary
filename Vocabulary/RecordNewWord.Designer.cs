namespace Vocabulary
{
    partial class RecordNewWord
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
            this.InputEntxt = new System.Windows.Forms.TextBox();
            this.InputCntxt = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // InputEntxt
            // 
            this.InputEntxt.Font = new System.Drawing.Font("微軟正黑體", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.InputEntxt.Location = new System.Drawing.Point(162, 49);
            this.InputEntxt.Multiline = true;
            this.InputEntxt.Name = "InputEntxt";
            this.InputEntxt.Size = new System.Drawing.Size(190, 45);
            this.InputEntxt.TabIndex = 3;
            // 
            // InputCntxt
            // 
            this.InputCntxt.Font = new System.Drawing.Font("微軟正黑體", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.InputCntxt.Location = new System.Drawing.Point(162, 134);
            this.InputCntxt.Multiline = true;
            this.InputCntxt.Name = "InputCntxt";
            this.InputCntxt.Size = new System.Drawing.Size(190, 45);
            this.InputCntxt.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(12, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(144, 42);
            this.label1.TabIndex = 4;
            this.label1.Text = "EnWord :";
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(12, 134);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(144, 42);
            this.label2.TabIndex = 4;
            this.label2.Text = "CnWord :";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("微軟正黑體", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button1.Location = new System.Drawing.Point(374, 73);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(129, 76);
            this.button1.TabIndex = 5;
            this.button1.Text = "Record";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // RecordNewWord
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(515, 237);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.InputCntxt);
            this.Controls.Add(this.InputEntxt);
            this.Name = "RecordNewWord";
            this.Text = "RecordNewWord";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox InputEntxt;
        private System.Windows.Forms.TextBox InputCntxt;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
    }
}
namespace Yield_Query_Tool
{
    partial class Form_InformBox
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
            this.Label_Infor_Msg = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Label_Infor_Msg
            // 
            this.Label_Infor_Msg.AutoSize = true;
            this.Label_Infor_Msg.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label_Infor_Msg.Location = new System.Drawing.Point(32, 30);
            this.Label_Infor_Msg.Name = "Label_Infor_Msg";
            this.Label_Infor_Msg.Size = new System.Drawing.Size(45, 16);
            this.Label_Infor_Msg.TabIndex = 0;
            this.Label_Infor_Msg.Text = "label1";
            // 
            // Form_InformBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(367, 113);
            this.Controls.Add(this.Label_Infor_Msg);
            this.Name = "Form_InformBox";
            this.Text = "Query_Time_Remaining";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label Label_Infor_Msg;
    }
}
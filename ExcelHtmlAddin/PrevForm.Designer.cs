namespace ExcelHtmlAddin
{
    partial class PrevForm
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.reportText = new System.Windows.Forms.TextBox();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.doCopyAndCloseButton = new System.Windows.Forms.Button();
            this.doCancelButton = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.reportText, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.flowLayoutPanel1, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 89.01734F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.98266F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(508, 346);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // reportText
            // 
            this.reportText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.reportText.Location = new System.Drawing.Point(3, 3);
            this.reportText.MaxLength = 0;
            this.reportText.Multiline = true;
            this.reportText.Name = "reportText";
            this.reportText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.reportText.Size = new System.Drawing.Size(502, 302);
            this.reportText.TabIndex = 0;
            this.reportText.KeyDown += new System.Windows.Forms.KeyEventHandler(this.reportText_KeyDown);
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.doCopyAndCloseButton);
            this.flowLayoutPanel1.Controls.Add(this.doCancelButton);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Right;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(160, 311);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(345, 32);
            this.flowLayoutPanel1.TabIndex = 1;
            // 
            // doCopyAndCloseButton
            // 
            this.doCopyAndCloseButton.Location = new System.Drawing.Point(3, 3);
            this.doCopyAndCloseButton.Name = "doCopyAndCloseButton";
            this.doCopyAndCloseButton.Size = new System.Drawing.Size(99, 23);
            this.doCopyAndCloseButton.TabIndex = 0;
            this.doCopyAndCloseButton.Text = "コピーして閉じる";
            this.doCopyAndCloseButton.UseVisualStyleBackColor = true;
            this.doCopyAndCloseButton.Click += new System.EventHandler(this.doCopyAndCloseButton_Click);
            // 
            // doCancelButton
            // 
            this.doCancelButton.Location = new System.Drawing.Point(108, 3);
            this.doCancelButton.Name = "doCancelButton";
            this.doCancelButton.Size = new System.Drawing.Size(75, 23);
            this.doCancelButton.TabIndex = 1;
            this.doCancelButton.Text = "キャンセル";
            this.doCancelButton.UseVisualStyleBackColor = true;
            this.doCancelButton.Click += new System.EventHandler(this.doCancelButton_Click);
            // 
            // PrevForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(508, 346);
            this.ControlBox = false;
            this.Controls.Add(this.tableLayoutPanel1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PrevForm";
            this.Text = "コードビュー";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        public System.Windows.Forms.TextBox reportText;
        private System.Windows.Forms.Button doCopyAndCloseButton;
        private System.Windows.Forms.Button doCancelButton;
    }
}
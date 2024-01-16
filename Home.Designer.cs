namespace ProvaN3_SolicitaSuporte
{
    partial class Home
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
            this.btn_SendImage = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lbl_BodyEmail = new System.Windows.Forms.Label();
            this.txtBox_BodyEmail = new System.Windows.Forms.TextBox();
            this.pnl_Email = new System.Windows.Forms.Panel();
            this.txtBox_Subject = new System.Windows.Forms.TextBox();
            this.txtBox_From = new System.Windows.Forms.TextBox();
            this.txtBox_To = new System.Windows.Forms.TextBox();
            this.lbl_Subject = new System.Windows.Forms.Label();
            this.lbl_From = new System.Windows.Forms.Label();
            this.lbl_To = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.pnl_Email.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_SendImage
            // 
            this.btn_SendImage.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_SendImage.Location = new System.Drawing.Point(2, 328);
            this.btn_SendImage.Name = "btn_SendImage";
            this.btn_SendImage.Size = new System.Drawing.Size(362, 33);
            this.btn_SendImage.TabIndex = 1;
            this.btn_SendImage.Text = "ENVIAR";
            this.btn_SendImage.UseVisualStyleBackColor = true;
            this.btn_SendImage.Click += new System.EventHandler(this.btn_SendImage_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.lbl_BodyEmail);
            this.panel1.Controls.Add(this.txtBox_BodyEmail);
            this.panel1.Location = new System.Drawing.Point(2, 117);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(358, 196);
            this.panel1.TabIndex = 2;
            // 
            // lbl_BodyEmail
            // 
            this.lbl_BodyEmail.AutoSize = true;
            this.lbl_BodyEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_BodyEmail.Location = new System.Drawing.Point(3, 7);
            this.lbl_BodyEmail.Name = "lbl_BodyEmail";
            this.lbl_BodyEmail.Size = new System.Drawing.Size(67, 13);
            this.lbl_BodyEmail.TabIndex = 4;
            this.lbl_BodyEmail.Text = "Mensagem";
            // 
            // txtBox_BodyEmail
            // 
            this.txtBox_BodyEmail.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtBox_BodyEmail.Location = new System.Drawing.Point(3, 28);
            this.txtBox_BodyEmail.Multiline = true;
            this.txtBox_BodyEmail.Name = "txtBox_BodyEmail";
            this.txtBox_BodyEmail.Size = new System.Drawing.Size(350, 126);
            this.txtBox_BodyEmail.TabIndex = 0;
            // 
            // pnl_Email
            // 
            this.pnl_Email.Controls.Add(this.txtBox_Subject);
            this.pnl_Email.Controls.Add(this.txtBox_From);
            this.pnl_Email.Controls.Add(this.txtBox_To);
            this.pnl_Email.Controls.Add(this.lbl_Subject);
            this.pnl_Email.Controls.Add(this.lbl_From);
            this.pnl_Email.Controls.Add(this.lbl_To);
            this.pnl_Email.Location = new System.Drawing.Point(5, 13);
            this.pnl_Email.Name = "pnl_Email";
            this.pnl_Email.Size = new System.Drawing.Size(355, 98);
            this.pnl_Email.TabIndex = 3;
            // 
            // txtBox_Subject
            // 
            this.txtBox_Subject.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtBox_Subject.Location = new System.Drawing.Point(62, 70);
            this.txtBox_Subject.Name = "txtBox_Subject";
            this.txtBox_Subject.Size = new System.Drawing.Size(273, 20);
            this.txtBox_Subject.TabIndex = 5;
            // 
            // txtBox_From
            // 
            this.txtBox_From.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtBox_From.Location = new System.Drawing.Point(62, 38);
            this.txtBox_From.Name = "txtBox_From";
            this.txtBox_From.ReadOnly = true;
            this.txtBox_From.Size = new System.Drawing.Size(273, 20);
            this.txtBox_From.TabIndex = 4;
            this.txtBox_From.Text = "PROVA_N3HIPER@OUTLOOK.COM";
            this.txtBox_From.WordWrap = false;
            // 
            // txtBox_To
            // 
            this.txtBox_To.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtBox_To.Location = new System.Drawing.Point(62, 9);
            this.txtBox_To.Name = "txtBox_To";
            this.txtBox_To.Size = new System.Drawing.Size(273, 20);
            this.txtBox_To.TabIndex = 3;
            // 
            // lbl_Subject
            // 
            this.lbl_Subject.AutoSize = true;
            this.lbl_Subject.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Subject.Location = new System.Drawing.Point(5, 70);
            this.lbl_Subject.Name = "lbl_Subject";
            this.lbl_Subject.Size = new System.Drawing.Size(56, 13);
            this.lbl_Subject.TabIndex = 2;
            this.lbl_Subject.Text = "Assunto:";
            // 
            // lbl_From
            // 
            this.lbl_From.AutoSize = true;
            this.lbl_From.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_From.Location = new System.Drawing.Point(5, 38);
            this.lbl_From.Name = "lbl_From";
            this.lbl_From.Size = new System.Drawing.Size(27, 13);
            this.lbl_From.TabIndex = 1;
            this.lbl_From.Text = "De:";
            // 
            // lbl_To
            // 
            this.lbl_To.AutoSize = true;
            this.lbl_To.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_To.Location = new System.Drawing.Point(5, 9);
            this.lbl_To.Name = "lbl_To";
            this.lbl_To.Size = new System.Drawing.Size(37, 13);
            this.lbl_To.TabIndex = 0;
            this.lbl_To.Text = "Para:";
            // 
            // Home
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(365, 362);
            this.Controls.Add(this.pnl_Email);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btn_SendImage);
            this.Name = "Home";
            this.Text = "Solicita Suporte";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.pnl_Email.ResumeLayout(false);
            this.pnl_Email.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btn_SendImage;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox txtBox_BodyEmail;
        private System.Windows.Forms.Panel pnl_Email;
        private System.Windows.Forms.Label lbl_BodyEmail;
        private System.Windows.Forms.Label lbl_Subject;
        private System.Windows.Forms.Label lbl_From;
        private System.Windows.Forms.Label lbl_To;
        private System.Windows.Forms.TextBox txtBox_Subject;
        private System.Windows.Forms.TextBox txtBox_From;
        private System.Windows.Forms.TextBox txtBox_To;
    }
}


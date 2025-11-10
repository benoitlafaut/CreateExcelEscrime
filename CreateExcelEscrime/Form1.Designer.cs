namespace CreateExcelEscrime
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
            this.btnCreateExcel = new System.Windows.Forms.Button();
            this.btnComptabiliser = new System.Windows.Forms.Button();
            this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
            this.textBoxMontant = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBoxDestinataire = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBoxPériode = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.comboBoxRaison = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnCreateWord = new System.Windows.Forms.Button();
            this.comboBoxDépensesRecettes = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.comboBoxPots = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btnCreateExcel
            // 
            this.btnCreateExcel.Location = new System.Drawing.Point(490, 246);
            this.btnCreateExcel.Name = "btnCreateExcel";
            this.btnCreateExcel.Size = new System.Drawing.Size(154, 76);
            this.btnCreateExcel.TabIndex = 1;
            this.btnCreateExcel.Text = "Create Excel";
            this.btnCreateExcel.UseVisualStyleBackColor = true;
            this.btnCreateExcel.Click += new System.EventHandler(this.btnCreateExcel_Click);
            // 
            // btnComptabiliser
            // 
            this.btnComptabiliser.Location = new System.Drawing.Point(257, 246);
            this.btnComptabiliser.Name = "btnComptabiliser";
            this.btnComptabiliser.Size = new System.Drawing.Size(189, 76);
            this.btnComptabiliser.TabIndex = 2;
            this.btnComptabiliser.Text = "Ajout dans la comptabilité";
            this.btnComptabiliser.UseVisualStyleBackColor = true;
            this.btnComptabiliser.Click += new System.EventHandler(this.btnComptabiliser_Click);
            // 
            // monthCalendar1
            // 
            this.monthCalendar1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.monthCalendar1.Location = new System.Drawing.Point(18, 18);
            this.monthCalendar1.Name = "monthCalendar1";
            this.monthCalendar1.TabIndex = 3;
            this.monthCalendar1.TitleForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.monthCalendar1.TrailingForeColor = System.Drawing.Color.Red;
            this.monthCalendar1.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateChanged);
            // 
            // textBoxMontant
            // 
            this.textBoxMontant.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxMontant.Location = new System.Drawing.Point(347, 61);
            this.textBoxMontant.Name = "textBoxMontant";
            this.textBoxMontant.Size = new System.Drawing.Size(143, 24);
            this.textBoxMontant.TabIndex = 5;
            this.textBoxMontant.TextChanged += new System.EventHandler(this.textBoxMontant_TextChanged);
            // 
            // label1
            // 
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Location = new System.Drawing.Point(257, 60);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 25);
            this.label1.TabIndex = 6;
            this.label1.Text = "Montant :";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // comboBoxDestinataire
            // 
            this.comboBoxDestinataire.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxDestinataire.FormattingEnabled = true;
            this.comboBoxDestinataire.Location = new System.Drawing.Point(347, 21);
            this.comboBoxDestinataire.Name = "comboBoxDestinataire";
            this.comboBoxDestinataire.Size = new System.Drawing.Size(172, 26);
            this.comboBoxDestinataire.TabIndex = 7;
            // 
            // label2
            // 
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.Location = new System.Drawing.Point(257, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 26);
            this.label2.TabIndex = 8;
            this.label2.Text = "Destinataire :";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label3.Location = new System.Drawing.Point(257, 130);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 26);
            this.label3.TabIndex = 10;
            this.label3.Text = "Période :";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // comboBoxPériode
            // 
            this.comboBoxPériode.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxPériode.FormattingEnabled = true;
            this.comboBoxPériode.Location = new System.Drawing.Point(347, 130);
            this.comboBoxPériode.Name = "comboBoxPériode";
            this.comboBoxPériode.Size = new System.Drawing.Size(95, 26);
            this.comboBoxPériode.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label4.Location = new System.Drawing.Point(257, 95);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 26);
            this.label4.TabIndex = 12;
            this.label4.Text = "Raison :";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // comboBoxRaison
            // 
            this.comboBoxRaison.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxRaison.FormattingEnabled = true;
            this.comboBoxRaison.Location = new System.Drawing.Point(347, 95);
            this.comboBoxRaison.Name = "comboBoxRaison";
            this.comboBoxRaison.Size = new System.Drawing.Size(297, 26);
            this.comboBoxRaison.TabIndex = 11;
            // 
            // label5
            // 
            this.label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label5.Location = new System.Drawing.Point(18, 179);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(227, 23);
            this.label5.TabIndex = 13;
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnCreateWord
            // 
            this.btnCreateWord.Location = new System.Drawing.Point(18, 294);
            this.btnCreateWord.Name = "btnCreateWord";
            this.btnCreateWord.Size = new System.Drawing.Size(95, 42);
            this.btnCreateWord.TabIndex = 14;
            this.btnCreateWord.Text = "Create Word";
            this.btnCreateWord.UseVisualStyleBackColor = true;
            this.btnCreateWord.Visible = false;
            this.btnCreateWord.Click += new System.EventHandler(this.btnCreateWord_Click);
            // 
            // comboBoxDépensesRecettes
            // 
            this.comboBoxDépensesRecettes.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxDépensesRecettes.FormattingEnabled = true;
            this.comboBoxDépensesRecettes.Location = new System.Drawing.Point(371, 176);
            this.comboBoxDépensesRecettes.Name = "comboBoxDépensesRecettes";
            this.comboBoxDépensesRecettes.Size = new System.Drawing.Size(170, 26);
            this.comboBoxDépensesRecettes.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label6.Location = new System.Drawing.Point(257, 176);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(108, 26);
            this.label6.TabIndex = 16;
            this.label6.Text = "Dépenses/Recettes";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label7
            // 
            this.label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label7.Location = new System.Drawing.Point(257, 208);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(108, 26);
            this.label7.TabIndex = 17;
            this.label7.Text = "Pot";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // comboBoxPots
            // 
            this.comboBoxPots.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxPots.FormattingEnabled = true;
            this.comboBoxPots.Location = new System.Drawing.Point(371, 208);
            this.comboBoxPots.Name = "comboBoxPots";
            this.comboBoxPots.Size = new System.Drawing.Size(170, 26);
            this.comboBoxPots.TabIndex = 18;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(657, 348);
            this.Controls.Add(this.comboBoxPots);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.comboBoxDépensesRecettes);
            this.Controls.Add(this.btnCreateWord);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.comboBoxRaison);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.comboBoxPériode);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.comboBoxDestinataire);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxMontant);
            this.Controls.Add(this.monthCalendar1);
            this.Controls.Add(this.btnComptabiliser);
            this.Controls.Add(this.btnCreateExcel);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCreateExcel;
        private System.Windows.Forms.Button btnComptabiliser;
        private System.Windows.Forms.MonthCalendar monthCalendar1;
        private System.Windows.Forms.TextBox textBoxMontant;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBoxDestinataire;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBoxPériode;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox comboBoxRaison;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnCreateWord;
        private System.Windows.Forms.ComboBox comboBoxDépensesRecettes;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox comboBoxPots;
    }
}


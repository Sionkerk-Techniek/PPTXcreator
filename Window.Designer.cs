namespace PPTXcreator
{
    partial class Window
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabGeneral = new System.Windows.Forms.TabPage();
            this.tabLiturgie = new System.Windows.Forms.TabPage();
            this.dateTimePickerNu = new System.Windows.Forms.DateTimePicker();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxVoorgangerNuNaam = new System.Windows.Forms.TextBox();
            this.textBoxVoorgangerNuTitel = new System.Windows.Forms.TextBox();
            this.textBoxVoorgangerNextTitel = new System.Windows.Forms.TextBox();
            this.textBoxVoorgangerNextNaam = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxVoorgangerNuPlaats = new System.Windows.Forms.TextBox();
            this.textBoxVoorgangerNextPlaats = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.dateTimePickerNext = new System.Windows.Forms.DateTimePicker();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.tabGeneral.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabGeneral);
            this.tabControl1.Controls.Add(this.tabLiturgie);
            this.tabControl1.Font = new System.Drawing.Font("Arial", 8.830189F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(4, 4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(532, 366);
            this.tabControl1.TabIndex = 0;
            // 
            // tabGeneral
            // 
            this.tabGeneral.Controls.Add(this.groupBox3);
            this.tabGeneral.Controls.Add(this.groupBox2);
            this.tabGeneral.Controls.Add(this.groupBox1);
            this.tabGeneral.Location = new System.Drawing.Point(4, 25);
            this.tabGeneral.Name = "tabGeneral";
            this.tabGeneral.Padding = new System.Windows.Forms.Padding(3);
            this.tabGeneral.Size = new System.Drawing.Size(524, 337);
            this.tabGeneral.TabIndex = 0;
            this.tabGeneral.Text = "Algemeen";
            this.tabGeneral.UseVisualStyleBackColor = true;
            // 
            // tabLiturgie
            // 
            this.tabLiturgie.Location = new System.Drawing.Point(4, 25);
            this.tabLiturgie.Name = "tabLiturgie";
            this.tabLiturgie.Padding = new System.Windows.Forms.Padding(3);
            this.tabLiturgie.Size = new System.Drawing.Size(524, 344);
            this.tabLiturgie.TabIndex = 1;
            this.tabLiturgie.Text = "Liturgie";
            this.tabLiturgie.UseVisualStyleBackColor = true;
            // 
            // dateTimePickerNu
            // 
            this.dateTimePickerNu.CustomFormat = "yyyy-MM-dd HH:mm";
            this.dateTimePickerNu.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickerNu.Location = new System.Drawing.Point(132, 34);
            this.dateTimePickerNu.MinDate = new System.DateTime(2020, 1, 1, 0, 0, 0, 0);
            this.dateTimePickerNu.Name = "dateTimePickerNu";
            this.dateTimePickerNu.Size = new System.Drawing.Size(175, 22);
            this.dateTimePickerNu.TabIndex = 2;
            this.dateTimePickerNu.Value = new System.DateTime(2021, 1, 3, 0, 0, 0, 0);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBoxVoorgangerNextPlaats);
            this.groupBox1.Controls.Add(this.textBoxVoorgangerNuPlaats);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.textBoxVoorgangerNextNaam);
            this.groupBox1.Controls.Add(this.textBoxVoorgangerNextTitel);
            this.groupBox1.Controls.Add(this.textBoxVoorgangerNuTitel);
            this.groupBox1.Controls.Add(this.textBoxVoorgangerNuNaam);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Font = new System.Drawing.Font("Arial", 8.830189F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(6, 15);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(512, 97);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Voorgangers";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 8.830189F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(14, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Deze dienst:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 8.830189F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(14, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 16);
            this.label2.TabIndex = 1;
            this.label2.Text = "Volgende dienst:";
            // 
            // textBoxVoorgangerNuNaam
            // 
            this.textBoxVoorgangerNuNaam.Location = new System.Drawing.Point(173, 32);
            this.textBoxVoorgangerNuNaam.Name = "textBoxVoorgangerNuNaam";
            this.textBoxVoorgangerNuNaam.Size = new System.Drawing.Size(175, 22);
            this.textBoxVoorgangerNuNaam.TabIndex = 4;
            this.textBoxVoorgangerNuNaam.Text = "naam";
            // 
            // textBoxVoorgangerNuTitel
            // 
            this.textBoxVoorgangerNuTitel.AutoCompleteCustomSource.AddRange(new string[] {
            "ds. ",
            "prof. ",
            "prop. ",
            "dr. "});
            this.textBoxVoorgangerNuTitel.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.textBoxVoorgangerNuTitel.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.textBoxVoorgangerNuTitel.CharacterCasing = System.Windows.Forms.CharacterCasing.Lower;
            this.textBoxVoorgangerNuTitel.Location = new System.Drawing.Point(132, 32);
            this.textBoxVoorgangerNuTitel.Name = "textBoxVoorgangerNuTitel";
            this.textBoxVoorgangerNuTitel.Size = new System.Drawing.Size(36, 22);
            this.textBoxVoorgangerNuTitel.TabIndex = 5;
            this.textBoxVoorgangerNuTitel.Text = "titel";
            // 
            // textBoxVoorgangerNextTitel
            // 
            this.textBoxVoorgangerNextTitel.AutoCompleteCustomSource.AddRange(new string[] {
            "ds. ",
            "prof. ",
            "prop. ",
            "dr. "});
            this.textBoxVoorgangerNextTitel.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.textBoxVoorgangerNextTitel.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.textBoxVoorgangerNextTitel.CharacterCasing = System.Windows.Forms.CharacterCasing.Lower;
            this.textBoxVoorgangerNextTitel.Location = new System.Drawing.Point(132, 60);
            this.textBoxVoorgangerNextTitel.Name = "textBoxVoorgangerNextTitel";
            this.textBoxVoorgangerNextTitel.Size = new System.Drawing.Size(36, 22);
            this.textBoxVoorgangerNextTitel.TabIndex = 6;
            this.textBoxVoorgangerNextTitel.Text = "titel";
            // 
            // textBoxVoorgangerNextNaam
            // 
            this.textBoxVoorgangerNextNaam.Location = new System.Drawing.Point(173, 60);
            this.textBoxVoorgangerNextNaam.Name = "textBoxVoorgangerNextNaam";
            this.textBoxVoorgangerNextNaam.Size = new System.Drawing.Size(175, 22);
            this.textBoxVoorgangerNextNaam.TabIndex = 7;
            this.textBoxVoorgangerNextNaam.Text = "naam";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(354, 35);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(23, 16);
            this.label3.TabIndex = 8;
            this.label3.Text = "uit";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(354, 63);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(23, 16);
            this.label4.TabIndex = 9;
            this.label4.Text = "uit";
            // 
            // textBoxVoorgangerNuPlaats
            // 
            this.textBoxVoorgangerNuPlaats.Location = new System.Drawing.Point(383, 32);
            this.textBoxVoorgangerNuPlaats.Name = "textBoxVoorgangerNuPlaats";
            this.textBoxVoorgangerNuPlaats.Size = new System.Drawing.Size(119, 22);
            this.textBoxVoorgangerNuPlaats.TabIndex = 10;
            this.textBoxVoorgangerNuPlaats.Text = "plaats";
            // 
            // textBoxVoorgangerNextPlaats
            // 
            this.textBoxVoorgangerNextPlaats.Location = new System.Drawing.Point(383, 60);
            this.textBoxVoorgangerNextPlaats.Name = "textBoxVoorgangerNextPlaats";
            this.textBoxVoorgangerNextPlaats.Size = new System.Drawing.Size(119, 22);
            this.textBoxVoorgangerNextPlaats.TabIndex = 11;
            this.textBoxVoorgangerNextPlaats.Text = "plaats";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dateTimePickerNext);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.dateTimePickerNu);
            this.groupBox2.Location = new System.Drawing.Point(6, 126);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(512, 94);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Datum en tijd";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial", 8.830189F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(13, 36);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(86, 16);
            this.label5.TabIndex = 1;
            this.label5.Text = "Deze dienst:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Arial", 8.830189F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(13, 64);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(113, 16);
            this.label6.TabIndex = 12;
            this.label6.Text = "Volgende dienst:";
            // 
            // dateTimePickerNext
            // 
            this.dateTimePickerNext.CustomFormat = "yyyy-MM-dd HH:mm";
            this.dateTimePickerNext.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickerNext.Location = new System.Drawing.Point(132, 62);
            this.dateTimePickerNext.MinDate = new System.DateTime(2020, 1, 1, 0, 0, 0, 0);
            this.dateTimePickerNext.Name = "dateTimePickerNext";
            this.dateTimePickerNext.Size = new System.Drawing.Size(175, 22);
            this.dateTimePickerNext.TabIndex = 13;
            this.dateTimePickerNext.Value = new System.DateTime(2021, 1, 3, 0, 0, 0, 0);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.textBox1);
            this.groupBox3.Controls.Add(this.textBox2);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Location = new System.Drawing.Point(6, 234);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(512, 97);
            this.groupBox3.TabIndex = 5;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Organisten";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(132, 62);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(175, 22);
            this.textBox1.TabIndex = 15;
            this.textBox1.Text = "naam";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(132, 34);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(175, 22);
            this.textBox2.TabIndex = 14;
            this.textBox2.Text = "naam";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Arial", 8.830189F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(14, 65);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(113, 16);
            this.label7.TabIndex = 13;
            this.label7.Text = "Volgende dienst:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Arial", 8.830189F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(14, 37);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(86, 16);
            this.label8.TabIndex = 12;
            this.label8.Text = "Deze dienst:";
            // 
            // Window
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.ClientSize = new System.Drawing.Size(539, 374);
            this.Controls.Add(this.tabControl1);
            this.Font = new System.Drawing.Font("Arial", 8.830189F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "Window";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PPTX creator";
            this.tabControl1.ResumeLayout(false);
            this.tabGeneral.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabGeneral;
        private System.Windows.Forms.TabPage tabLiturgie;
        private System.Windows.Forms.DateTimePicker dateTimePickerNu;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox textBoxVoorgangerNuNaam;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxVoorgangerNuTitel;
        private System.Windows.Forms.TextBox textBoxVoorgangerNextTitel;
        private System.Windows.Forms.TextBox textBoxVoorgangerNextNaam;
        private System.Windows.Forms.TextBox textBoxVoorgangerNextPlaats;
        private System.Windows.Forms.TextBox textBoxVoorgangerNuPlaats;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DateTimePicker dateTimePickerNext;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
    }
}


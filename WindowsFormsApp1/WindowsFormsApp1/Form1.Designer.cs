namespace WindowsFormsApp1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnFileWord = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtFileWord = new System.Windows.Forms.TextBox();
            this.btnFileExcel = new System.Windows.Forms.Button();
            this.txtFileExcel = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.textBox9 = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.textBox11 = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.btnConvert = new System.Windows.Forms.Button();
            this.label18 = new System.Windows.Forms.Label();
            this.txtSave = new System.Windows.Forms.TextBox();
            this.btnLuu = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnFileWord
            // 
            this.btnFileWord.Location = new System.Drawing.Point(415, 113);
            this.btnFileWord.Name = "btnFileWord";
            this.btnFileWord.Size = new System.Drawing.Size(75, 23);
            this.btnFileWord.TabIndex = 0;
            this.btnFileWord.Text = "Browse";
            this.btnFileWord.UseVisualStyleBackColor = true;
            this.btnFileWord.Click += new System.EventHandler(this.btnFileWord_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(302, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(134, 25);
            this.label1.TabIndex = 1;
            this.label1.Text = "PHẦN MỀM";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(149, 57);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(464, 25);
            this.label2.TabIndex = 1;
            this.label2.Text = "THÊM DANH SÁCH TỪ EXCEL VÀO WORD";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Tai Le", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(662, 520);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(126, 14);
            this.label4.TabIndex = 1;
            this.label4.Text = "© kimnhuttruong 2020";
            // 
            // txtFileWord
            // 
            this.txtFileWord.Location = new System.Drawing.Point(17, 113);
            this.txtFileWord.Name = "txtFileWord";
            this.txtFileWord.ReadOnly = true;
            this.txtFileWord.Size = new System.Drawing.Size(368, 20);
            this.txtFileWord.TabIndex = 2;
            this.txtFileWord.Text = "đường dẫn file word sẽ hiện tại đây";
            // 
            // btnFileExcel
            // 
            this.btnFileExcel.Location = new System.Drawing.Point(415, 158);
            this.btnFileExcel.Name = "btnFileExcel";
            this.btnFileExcel.Size = new System.Drawing.Size(75, 23);
            this.btnFileExcel.TabIndex = 0;
            this.btnFileExcel.Text = "Browse";
            this.btnFileExcel.UseVisualStyleBackColor = true;
            this.btnFileExcel.Click += new System.EventHandler(this.btnFileExcel_Click);
            // 
            // txtFileExcel
            // 
            this.txtFileExcel.Location = new System.Drawing.Point(17, 158);
            this.txtFileExcel.Name = "txtFileExcel";
            this.txtFileExcel.ReadOnly = true;
            this.txtFileExcel.Size = new System.Drawing.Size(368, 20);
            this.txtFileExcel.TabIndex = 2;
            this.txtFileExcel.Text = "đường dẫn file excel sẽ hiện tại đây";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(17, 243);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(62, 13);
            this.label7.TabIndex = 4;
            this.label7.Text = "Cách dùng:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(112, 243);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(49, 13);
            this.label8.TabIndex = 5;
            this.label8.Text = "File word";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(326, 243);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(49, 13);
            this.label9.TabIndex = 5;
            this.label9.Text = "File word";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(112, 262);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(57, 13);
            this.label10.TabIndex = 5;
            this.label10.Text = "<<name>>";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(318, 262);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(87, 13);
            this.label11.TabIndex = 5;
            this.label11.Text = "Kim Nhựt Trường";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(229, 243);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(25, 13);
            this.label12.TabIndex = 5;
            this.label12.Text = "==>";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(112, 296);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(51, 13);
            this.label13.TabIndex = 5;
            this.label13.Text = "File excel";
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(34, 334);
            this.textBox4.Name = "textBox4";
            this.textBox4.ReadOnly = true;
            this.textBox4.Size = new System.Drawing.Size(113, 20);
            this.textBox4.TabIndex = 6;
            this.textBox4.Text = "name";
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(146, 334);
            this.textBox5.Name = "textBox5";
            this.textBox5.ReadOnly = true;
            this.textBox5.Size = new System.Drawing.Size(52, 20);
            this.textBox5.TabIndex = 6;
            this.textBox5.Text = "age";
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(34, 353);
            this.textBox6.Name = "textBox6";
            this.textBox6.ReadOnly = true;
            this.textBox6.Size = new System.Drawing.Size(113, 20);
            this.textBox6.TabIndex = 6;
            this.textBox6.Text = "Kim Nhựt Trường";
            // 
            // textBox7
            // 
            this.textBox7.Location = new System.Drawing.Point(146, 353);
            this.textBox7.Name = "textBox7";
            this.textBox7.ReadOnly = true;
            this.textBox7.Size = new System.Drawing.Size(52, 20);
            this.textBox7.TabIndex = 6;
            this.textBox7.Text = "18";
            // 
            // textBox8
            // 
            this.textBox8.Location = new System.Drawing.Point(34, 372);
            this.textBox8.Name = "textBox8";
            this.textBox8.ReadOnly = true;
            this.textBox8.Size = new System.Drawing.Size(113, 20);
            this.textBox8.TabIndex = 6;
            this.textBox8.Text = "Nguyễn Văn A";
            // 
            // textBox9
            // 
            this.textBox9.Location = new System.Drawing.Point(146, 372);
            this.textBox9.Name = "textBox9";
            this.textBox9.ReadOnly = true;
            this.textBox9.Size = new System.Drawing.Size(52, 20);
            this.textBox9.TabIndex = 6;
            this.textBox9.Text = "22";
            // 
            // textBox10
            // 
            this.textBox10.Location = new System.Drawing.Point(34, 391);
            this.textBox10.Name = "textBox10";
            this.textBox10.ReadOnly = true;
            this.textBox10.Size = new System.Drawing.Size(113, 20);
            this.textBox10.TabIndex = 6;
            this.textBox10.Text = "Trần Thị B";
            // 
            // textBox11
            // 
            this.textBox11.Location = new System.Drawing.Point(146, 391);
            this.textBox11.Name = "textBox11";
            this.textBox11.ReadOnly = true;
            this.textBox11.Size = new System.Drawing.Size(52, 20);
            this.textBox11.TabIndex = 6;
            this.textBox11.Text = "30";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(241, 360);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(25, 13);
            this.label14.TabIndex = 5;
            this.label14.Text = "==>";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(295, 360);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(115, 13);
            this.label15.TabIndex = 5;
            this.label15.Text = "Các file word có tên,stt";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(295, 379);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(155, 13);
            this.label16.TabIndex = 5;
            this.label16.Text = "Ex: 1_KimNhutTruong_22.docx";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(313, 392);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(126, 13);
            this.label17.TabIndex = 5;
            this.label17.Text = "2_NguyenVanA_18.docx";
            // 
            // btnConvert
            // 
            this.btnConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnConvert.Location = new System.Drawing.Point(665, 233);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(75, 29);
            this.btnConvert.TabIndex = 0;
            this.btnConvert.Text = "convert";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label18.ForeColor = System.Drawing.Color.Red;
            this.label18.Location = new System.Drawing.Point(628, 212);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(143, 18);
            this.label18.TabIndex = 5;
            this.label18.Text = "Nếu đã sẵn sàng?";
            // 
            // txtSave
            // 
            this.txtSave.Location = new System.Drawing.Point(17, 197);
            this.txtSave.Name = "txtSave";
            this.txtSave.ReadOnly = true;
            this.txtSave.Size = new System.Drawing.Size(368, 20);
            this.txtSave.TabIndex = 2;
            this.txtSave.Text = "đường dẫn file lưu sẽ hiện tại đây";
            // 
            // btnLuu
            // 
            this.btnLuu.Location = new System.Drawing.Point(415, 194);
            this.btnLuu.Name = "btnLuu";
            this.btnLuu.Size = new System.Drawing.Size(75, 23);
            this.btnLuu.TabIndex = 0;
            this.btnLuu.Text = "Browse";
            this.btnLuu.UseVisualStyleBackColor = true;
            this.btnLuu.Click += new System.EventHandler(this.btnLuu_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 543);
            this.Controls.Add(this.textBox11);
            this.Controls.Add(this.textBox10);
            this.Controls.Add(this.textBox9);
            this.Controls.Add(this.textBox8);
            this.Controls.Add(this.textBox7);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtSave);
            this.Controls.Add(this.txtFileExcel);
            this.Controls.Add(this.txtFileWord);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnConvert);
            this.Controls.Add(this.btnLuu);
            this.Controls.Add(this.btnFileExcel);
            this.Controls.Add(this.btnFileWord);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Phần mềm thêm danh sách từ excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnFileWord;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtFileWord;
        private System.Windows.Forms.Button btnFileExcel;
        private System.Windows.Forms.TextBox txtFileExcel;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.TextBox textBox7;
        private System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.TextBox textBox9;
        private System.Windows.Forms.TextBox textBox10;
        private System.Windows.Forms.TextBox textBox11;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.TextBox txtSave;
        private System.Windows.Forms.Button btnLuu;
    }
}


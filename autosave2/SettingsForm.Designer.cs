/*
 * Created by SharpDevelop.
 * User: dmitriev_k
 * Date: 03.05.2017
 * Time: 11:30
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace Autosave
{
    partial class SettingsForm
    {
        /// <summary>
        /// Designer variable used to keep track of non-visual components.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Disposes resources used by the form.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// This method is required for Windows Forms designer support.
        /// Do not change the method contents inside the source code editor. The Forms designer might
        /// not be able to load this method if it was changed manually.
        /// </summary>
        private void InitializeComponent()
        {
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.button2 = new System.Windows.Forms.Button();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // checkBox1
            // 
            this.checkBox1.Location = new System.Drawing.Point(13, 12);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(195, 24);
            this.checkBox1.TabIndex = 0;
            this.checkBox1.Text = "Автосохранение включено";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(248, 77);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(66, 24);
            this.comboBox1.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(11, 80);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(231, 23);
            this.label1.TabIndex = 2;
            this.label1.Text = "Интервал между автосохранениями";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(320, 80);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 23);
            this.label2.TabIndex = 3;
            this.label2.Text = "мин";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(-15, 148);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(231, 23);
            this.label3.TabIndex = 2;
            this.label3.Text = "Имя папки автосохранения";
            this.label3.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(196, 145);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(230, 22);
            this.textBox1.TabIndex = 4;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(10, 222);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(231, 23);
            this.label4.TabIndex = 2;
            this.label4.Text = "Положение папки автосохранения";
            this.label4.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(15, 249);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(783, 22);
            this.textBox2.TabIndex = 4;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(252, 216);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(546, 23);
            this.button1.TabIndex = 5;
            this.button1.Text = "Задать расположение папки автосохранения";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1Click);
            // 
            // checkBox2
            // 
            this.checkBox2.Location = new System.Drawing.Point(16, 186);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(814, 24);
            this.checkBox2.TabIndex = 0;
            this.checkBox2.Text = "Создавать папку автосохранения в папке документа (настройка: \"Положение папки авт" +
    "осохранения\" игнорируется)";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(601, 290);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(197, 52);
            this.button2.TabIndex = 5;
            this.button2.Text = "Применить";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Button2Click);
            // 
            // checkBox3
            // 
            this.checkBox3.Location = new System.Drawing.Point(14, 42);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(485, 24);
            this.checkBox3.TabIndex = 6;
            this.checkBox3.Text = "При открытии документа для чтения автосохранять";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(320, 110);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(66, 24);
            this.comboBox2.TabIndex = 1;
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(392, 113);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(51, 23);
            this.label5.TabIndex = 3;
            this.label5.Text = "дней";
            this.label5.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(4, 113);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(313, 23);
            this.label6.TabIndex = 2;
            this.label6.Text = "Время хранения автосохраненных документов";
            this.label6.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // checkBox4
            // 
            this.checkBox4.Location = new System.Drawing.Point(214, 12);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(485, 24);
            this.checkBox4.TabIndex = 6;
            this.checkBox4.Text = "Сброс таймера при сохранении";
            this.checkBox4.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            this.label7.Font = new System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label7.ForeColor = System.Drawing.Color.Green;
            this.label7.Location = new System.Drawing.Point(659, 12);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(139, 42);
            this.label7.TabIndex = 7;
            this.label7.Text = "00:00";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(807, 356);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.checkBox4);
            this.Controls.Add(this.checkBox3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.checkBox2);
            this.Controls.Add(this.checkBox1);
            this.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Настройки автосохранения";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.CheckBox checkBox4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        public System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.CheckBox checkBox3;
        private System.Windows.Forms.Button button2;
        public System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.Button button1;
        public System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label4;
        public System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.ComboBox comboBox1;
        public System.Windows.Forms.CheckBox checkBox1;
    }
}

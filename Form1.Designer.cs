using System;
using System.Text;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DKMail
{
    partial class Form1
    {
        private string[] recipients;
        public Form1()
        {
            string[] config = null;
            if(File.Exists("mailer.config")) 
                config = File.ReadAllLines("mailer.config", System.Text.Encoding.UTF8);
            else if(File.Exists(Application.UserAppDataPath + "\\mailer.config")) 
                config = File.ReadAllLines(Application.UserAppDataPath + "\\mailer.config", System.Text.Encoding.UTF8);
            else
            {
                MessageBox.Show("Файл конфигурации mailer.config\nдолжен быть в текущей папке или в папке\n" + Application.UserAppDataPath
                    , "Ошибка конфигурации");
                Application.Exit();
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            InitializeComponent();
            recipients = new string[config.Length+1];
            this.comboBox1.Items.Insert(0, "Выберите категорию отчёта");
            this.comboBox1.SelectedIndex = 0;
            
            for(var i=0;i<config.Length; i++)
            {
                string[] tokens = config[i].Split(';');
                comboBox1.Items.Add(tokens[0]);
                recipients[i] = tokens[1];
            }

            this.listBox1.DragDrop += new System.Windows.Forms.DragEventHandler(this._DragDrop);
            this.listBox1.DragEnter += new System.Windows.Forms.DragEventHandler(this._DragEnter);
            this.textBox1.DragDrop += new System.Windows.Forms.DragEventHandler(this._DragDrop);
            this.textBox1.DragEnter += new System.Windows.Forms.DragEventHandler(this._DragEnter);
            this.comboBox1.DragDrop += new System.Windows.Forms.DragEventHandler(this._DragDrop);
            this.comboBox1.DragEnter += new System.Windows.Forms.DragEventHandler(this._DragEnter);

            this.DragEnter += _DragEnter;
            this.DragDrop += _DragDrop;

            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.button2.Click += new System.EventHandler(this.button2_Click);
            textBox1.ForeColor = Color.Gray;
            textBox1.Text = "Введите текст сообщения...";
            this.textBox1.Enter += new System.EventHandler(textBox1_Enter);
            this.textBox1.Leave += new System.EventHandler(textBox1_Leave);
            this.ActiveControl = button1;
        }


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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.AllowDrop = true;
            this.textBox1.Location = new System.Drawing.Point(0, 149);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(284, 164);
            this.textBox1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.AllowDrop = true;
            this.button1.Location = new System.Drawing.Point(0, 310);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(284, 49);
            this.button1.TabIndex = 1;
            this.button1.Text = "ПОСЛАТЬ ПОЧТУ";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // listBox1
            // 
            this.listBox1.AllowDrop = true;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(0, 26);
            this.listBox1.Margin = new System.Windows.Forms.Padding(0);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(252, 121);
            this.listBox1.TabIndex = 2;
            // 
            // comboBox1
            // 
            this.comboBox1.AllowDrop = true;
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(-1, -1);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(284, 21);
            this.comboBox1.TabIndex = 3;
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.button2.Location = new System.Drawing.Point(253, 22);
            this.button2.Margin = new System.Windows.Forms.Padding(0);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(31, 124);
            this.button2.TabIndex = 4;
            this.button2.Text = "У д а л и т ь";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 359);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.Text = "Отправка отчёта";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListBox listBox1;
        private ComboBox comboBox1;
        private Button button2;

        private void _DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            int i;
            for (i = 0; i < s.Length; i++)
                listBox1.Items.Add(s[i]);

        }
        private void _DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

      
        private void button2_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex >= 0)
                listBox1.Items.Remove(listBox1.SelectedItem);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            var ol = new Outlook.Application();
            Outlook.MailItem mail = ol.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.Subject = comboBox1.Items[comboBox1.SelectedIndex].ToString();
            Outlook.AddressEntry currentUser = ol.Session.CurrentUser.AddressEntry;
            if (currentUser.Type == "EX")
            {
                mail.Body = textBox1.Text;
                mail.Recipients.Add(recipients[comboBox1.SelectedIndex-1]);
                mail.Recipients.ResolveAll();
                for (int i = 0; i < listBox1.Items.Count; i++)
                    mail.Attachments.Add(listBox1.Items[i],
                    Outlook.OlAttachmentType.olByValue, Type.Missing,
                    Type.Missing);
                mail.Send();
                this.Close();
            }
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            if (textBox1.Text == "Введите текст сообщения...")
            {
                textBox1.Text = "";
                textBox1.ForeColor = Color.Black;
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.Text = "Введите текст сообщения...";
                textBox1.ForeColor = Color.Gray;
            }
        }

    }
}


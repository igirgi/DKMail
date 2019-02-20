using System;
using System.Collections;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DKMail
{
    partial class Form1
    {
        public Form1()
        {
            string[] config = null;            
            if (File.Exists("DKMail.config")) 
                config = File.ReadAllLines("DKMail.config", System.Text.Encoding.UTF8);
            else if (File.Exists(Path.Combine(Directory.GetParent(Application.UserAppDataPath).FullName, "DKMail.config")))
                config = File.ReadAllLines(Path.Combine(Directory.GetParent(Application.UserAppDataPath).FullName, "DKMail.config"), System.Text.Encoding.UTF8);
            else if (File.Exists(Path.Combine(Directory.GetParent(Application.CommonAppDataPath).FullName, "DKMail.config")))
                config = File.ReadAllLines(Path.Combine(Directory.GetParent(Application.CommonAppDataPath).FullName, "DKMail.config"), System.Text.Encoding.UTF8);
            else
            {
                MessageBox.Show("Файл конфигурации DKMail.config\nдолжен быть в текущей папке или в папке\n" + Application.UserAppDataPath
                    , "Ошибка конфигурации");
                Application.Exit();
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            InitializeComponent();
            groupedComboBox1.ValueMember = "Mail";
            groupedComboBox1.DisplayMember = "Display";
            groupedComboBox1.GroupMember = "Group";
            
            ArrayList ds = new ArrayList();
            ds.Add(new gitem(String.Empty, "Выберите категорию обращения", String.Empty));
                                                
            for(var i=0;i<config.Length; i++)
            {
                string[] tokens = config[i].Split(';');
                ds.Add(new gitem(tokens[0], tokens[1], tokens[2]));
            }
            
            groupedComboBox1.DataSource = ds;
            
            this.listBox1.DragDrop += new System.Windows.Forms.DragEventHandler(this._DragDrop);
            this.listBox1.DragEnter += new System.Windows.Forms.DragEventHandler(this._DragEnter);
            this.textBox1.DragDrop += new System.Windows.Forms.DragEventHandler(this._DragDrop);
            this.textBox1.DragEnter += new System.Windows.Forms.DragEventHandler(this._DragEnter);
            this.groupedComboBox1.DragDrop += new System.Windows.Forms.DragEventHandler(this._DragDrop);
            this.groupedComboBox1.DragEnter += new System.Windows.Forms.DragEventHandler(this._DragEnter);
            this.button1.DragDrop += new System.Windows.Forms.DragEventHandler(this._DragDrop);
            this.button1.DragEnter += new System.Windows.Forms.DragEventHandler(this._DragEnter);
            this.button3.DragDrop += new System.Windows.Forms.DragEventHandler(this._DragDrop);
            this.button3.DragEnter += new System.Windows.Forms.DragEventHandler(this._DragEnter);

            this.DragEnter += _DragEnter;
            this.DragDrop += _DragDrop;

            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.button2.Click += new System.EventHandler(this.button2_Click);
            this.button3.Click += new System.EventHandler(this.button3_Click);

            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_Select);

            textBox1.ForeColor = Color.Gray;
            textBox1.Text = "Введите текст обращения...";
            this.textBox1.Enter += new System.EventHandler(textBox1_Enter);
            this.textBox1.Leave += new System.EventHandler(textBox1_Leave);
            this.ActiveControl = button3;
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
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupedComboBox1 = new GroupedComboBox();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.AllowDrop = true;
            this.textBox1.Location = new System.Drawing.Point(0, 164);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(416, 149);
            this.textBox1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.AllowDrop = true;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(0, 310);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(416, 49);
            this.button1.TabIndex = 1;
            this.button1.Text = "Отправить обращение";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // listBox1
            // 
            this.listBox1.AllowDrop = true;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(0, 45);
            this.listBox1.Margin = new System.Windows.Forms.Padding(0);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(331, 95);
            this.listBox1.TabIndex = 2;
            // 
            // button2
            // 
            this.button2.Enabled = false;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.button2.Location = new System.Drawing.Point(334, 90);
            this.button2.Margin = new System.Windows.Forms.Padding(0);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(82, 31);
            this.button2.TabIndex = 4;
            this.button2.Text = "Удалить";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button3.Location = new System.Drawing.Point(334, 45);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(82, 31);
            this.button3.TabIndex = 5;
            this.button3.Text = "Добавить";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(0, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(401, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Файлы - прилложения.  Нажмите кномку \"Добавить\" или натащите мышкой.";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(0, 148);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(97, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Текст обращения";
            // 
            // groupedComboBox1
            // 
            this.groupedComboBox1.DataSource = null;
            this.groupedComboBox1.FormattingEnabled = true;
            this.groupedComboBox1.Location = new System.Drawing.Point(0, 5);
            this.groupedComboBox1.Name = "groupedComboBox1";
            this.groupedComboBox1.Size = new System.Drawing.Size(416, 21);
            this.groupedComboBox1.TabIndex = 8;
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(413, 359);
            this.Controls.Add(this.groupedComboBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.Text = "    Система заявок";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListBox listBox1;
        private Button button2;
        private Button button3;
        private Label label1;
        private Label label2;
        private GroupedComboBox groupedComboBox1;


        private void _DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);            
            for (int i = 0; i < s.Length; i++)
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

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Multiselect = true;
            openFileDialog1.Filter = "Любой файл|*.*";
            openFileDialog1.Title = "Выберите файлы приложения";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                for (int i = 0; i < openFileDialog1.FileNames.Length; i++)
                    listBox1.Items.Add(openFileDialog1.FileNames[i]);
            } 
        }      
        private void button2_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex >= 0)
                listBox1.Items.Remove(listBox1.SelectedItem);
            button2.Enabled = (listBox1.SelectedIndex >= 0);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (groupedComboBox1.SelectedIndex <= 0)
            {
                MessageBox.Show("Выберите категорию обращения\nиз выпадающего списка сверху", "Ошибка");
                return;
            }
            
            var ol = new Outlook.Application();
            try
            {
                Outlook.MailItem mail = ol.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                mail.Subject = ((gitem) groupedComboBox1.SelectedItem).Display;
                Outlook.AddressEntry currentUser = ol.Session.CurrentUser.AddressEntry;
                if (currentUser.Type == "EX")
                {
                    mail.Body = textBox1.Text.Equals("Введите текст обращения...") ? " " : textBox1.Text;
                    mail.Recipients.Add(((gitem)groupedComboBox1.SelectedItem).Mail);
                    mail.Recipients.ResolveAll();
                    for (int i = 0; i < listBox1.Items.Count; i++)
                        mail.Attachments.Add(listBox1.Items[i],
                        Outlook.OlAttachmentType.olByValue, Type.Missing,
                        Type.Missing);
                    mail.Send();
                    button1.BackColor = Color.Yellow;
                    button1.ForeColor = Color.Red;
                    button1.Text = "ОБРАЩЕНИЕ ОТПРАВЛЕНО";
                    System.Threading.Thread.Sleep(500);
                    this.Close();
                }
            }catch (COMException ce)
            {
                MessageBox.Show("Чтобы отослать сообщение, у вас должен быть установлен Аутлук и вы должны дать разрешение на его использование", "Ошибка Аутлука");
            }
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            if (textBox1.Text == "Введите текст обращения...")
            {
                textBox1.Text = "";
                textBox1.ForeColor = Color.Black;
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.Text = "Введите текст обращения...";
                textBox1.ForeColor = Color.Gray;
            }
        }
        private void listBox1_Select(object sender, EventArgs e)
        {
            button2.Enabled = (listBox1.SelectedIndex >= 0);
        }
    }
}


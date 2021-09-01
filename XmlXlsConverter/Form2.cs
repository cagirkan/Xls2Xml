using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace XmlXlsConverter
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            if(!File.Exists(AppDomain.CurrentDomain.BaseDirectory + "settings.xml"))
            {
                createSettingsXml();
            }
            XmlDocument doc = new XmlDocument();
            doc.Load(AppDomain.CurrentDomain.BaseDirectory + "settings.xml");
            string compName = doc.SelectSingleNode("/Settings/COMPANY_SHORT_NAME[1]").InnerText;
            string dbName = doc.SelectSingleNode("/Settings/DB_NAME[1]").InnerText;
            string dbUser = doc.SelectSingleNode("/Settings/DB_USERNAME[1]").InnerText;
            string dbPass = doc.SelectSingleNode("/Settings/DB_PASSWORD[1]").InnerText;
            string compNo = doc.SelectSingleNode("/Settings/COMPANY_NO[1]").InnerText;
            string invPeriod = doc.SelectSingleNode("/Settings/INVOICE_PERIOD[1]").InnerText;
            string filePath = doc.SelectSingleNode("/Settings/FILE_PATH[1]").InnerText;

            textBox1.Text = compName;
            textBox2.Text = dbName;
            textBox3.Text = dbUser;
            textBox4.Text = dbPass;
            textBox5.Text = compNo;
            textBox6.Text = invPeriod;
            textBox7.Text = filePath;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text == "" && textBox5.Text == "" && textBox6.Text == "")
            {
                MessageBox.Show("Lütfen Boş Alan Bırakmayın!");

            }
            else
            {
                var sts = new XmlWriterSettings()
                {
                    Indent = true,
                    Encoding = Encoding.GetEncoding("ISO-8859-9"),
                    ConformanceLevel = ConformanceLevel.Document,
                    IndentChars = ("  "),
                };
                XmlWriter writer = XmlWriter.Create(AppDomain.CurrentDomain.BaseDirectory + "settings.xml", sts);
                writer.WriteStartElement("Settings");
                writer.WriteElementString("FIRST", "NO");
                writer.WriteElementString("COMPANY_SHORT_NAME", textBox1.Text);
                writer.WriteElementString("DB_NAME", textBox2.Text);
                writer.WriteElementString("DB_USERNAME", textBox3.Text);
                writer.WriteElementString("DB_PASSWORD", textBox4.Text);
                writer.WriteElementString("COMPANY_NO", Convert.ToInt32(textBox5.Text).ToString("000"));
                writer.WriteElementString("INVOICE_PERIOD", Convert.ToInt32(textBox6.Text).ToString("00"));
                writer.WriteElementString("FILE_PATH", textBox7.Text);
                writer.WriteEndElement();
                writer.Flush();
                writer.Close();
                MessageBox.Show("Ayarlar Kaydedildi", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string connetionString;
            SqlConnection cnn;
            connetionString = @"Data Source=" + textBox2.Text + ";Initial Catalog=LOGO;User ID=" + textBox3.Text + ";Password=" + textBox4.Text;
            try
            {
                cnn = new SqlConnection(connetionString);
                cnn.Open();
                MessageBox.Show("Veritabanı Bağlantısı Başarılı!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cnn.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Veritabanı Bağlantısı Başarısız. Lütfen Alanları Kontrol Edin", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Custom Description";

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                textBox7.Text = fbd.SelectedPath;
            }
            
        }

        private void createSettingsXml()
        {
            var sts = new XmlWriterSettings()
            {
                Indent = true,
                Encoding = Encoding.GetEncoding("ISO-8859-9"),
                ConformanceLevel = ConformanceLevel.Document,
                IndentChars = ("  "),
            };
            XmlWriter writer = XmlWriter.Create(AppDomain.CurrentDomain.BaseDirectory + "settings.xml", sts);
            writer.WriteStartElement("Settings");
            writer.WriteElementString("FIRST", "YES");
            writer.WriteElementString("COMPANY_SHORT_NAME", "");
            writer.WriteElementString("DB_NAME", "");
            writer.WriteElementString("DB_USERNAME", "");
            writer.WriteElementString("DB_PASSWORD", "");
            writer.WriteElementString("COMPANY_NO", "");
            writer.WriteElementString("INVOICE_PERIOD", "");
            writer.WriteElementString("FILE_PATH", "");
            writer.WriteEndElement();
            writer.Flush();
            writer.Close();
        }
    }
}

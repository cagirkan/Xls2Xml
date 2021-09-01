using System;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace XmlXlsConverter
{
    public partial class Form2 : Form
    {
        static string secretKey = "7548megbilisim587469as1f7dsa4785";
        static string licenseSecretPhase = "m123%M321.";
        public Form2()
        {
            InitializeComponent();
            if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory + "settings.xml"))
            {
                createSettingsXml();
            }
            XmlDocument doc = new XmlDocument();
            doc.Load(AppDomain.CurrentDomain.BaseDirectory + "settings.xml");
            string dbName = doc.SelectSingleNode("/Settings/DB_NAME[1]").InnerText;
            string dbn = doc.SelectSingleNode("/Settings/DBN[1]").InnerText;
            string dbUser = doc.SelectSingleNode("/Settings/DB_USERNAME[1]").InnerText;
            string dbPass = doc.SelectSingleNode("/Settings/DB_PASSWORD[1]").InnerText;
            string compNo = doc.SelectSingleNode("/Settings/COMPANY_NO[1]").InnerText;
            string invPeriod = doc.SelectSingleNode("/Settings/INVOICE_PERIOD[1]").InnerText;
            string filePath = doc.SelectSingleNode("/Settings/FILE_PATH[1]").InnerText;
            string sheetName = doc.SelectSingleNode("/Settings/SHEET_NAME[1]").InnerText;
            bool cariControl = bool.Parse(doc.SelectSingleNode("/Settings/CREATE_ARP[1]").InnerText);
            bool seriControl = bool.Parse(doc.SelectSingleNode("/Settings/CREATE_SLT[1]").InnerText);
            bool hasQuantity = bool.Parse(doc.SelectSingleNode("/Settings/QUANTITY[1]").InnerText);
            string arpPrefix = doc.SelectSingleNode("/Settings/ARP_PREFIX[1]").InnerText;
            string arpCode = doc.SelectSingleNode("/Settings/ARP_CODE[1]").InnerText;
            string whSubStart = doc.SelectSingleNode("/Settings/WH_SUBSTRING_START[1]").InnerText;
            string whSubLen = doc.SelectSingleNode("/Settings/WH_SUBSTRING_LEN[1]").InnerText;
            string whExample = doc.SelectSingleNode("/Settings/WH_EXAMPLE[1]").InnerText;

            textBox1.Text = sheetName;
            textBox2.Text = dbName;
            textBox8.Text = dbn;
            textBox3.Text = dbUser;
            textBox4.Text = dbPass;
            textBox5.Text = compNo;
            textBox6.Text = invPeriod;
            textBox7.Text = filePath;
            checkBox1.Checked = cariControl;
            checkBox2.Checked = seriControl;
            checkBox3.Checked = hasQuantity;
            textBox9.Text = arpPrefix;
            textBox10.Text = arpCode;
            textBox13.Text = whExample;
            numericUpDown1.Value = Convert.ToDecimal(whSubStart);
            numericUpDown2.Value = Convert.ToDecimal(whSubLen);

            label13.Text = "Cari Kod: " + textBox9.Text + "." + textBox10.Text;
            string temp = textBox13.Text;
            string temp1 = temp.Substring(Convert.ToInt32(numericUpDown1.Value), Convert.ToInt32(numericUpDown2.Value));
            label18.Text = "-" + temp1 + "-";

            if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "meg.xml"))
            {
                XmlDocument doc2 = new XmlDocument();
                doc2.Load(AppDomain.CurrentDomain.BaseDirectory + "meg.xml");
                string meg_id = doc2.SelectSingleNode("/MEG/MEG_ID[1]").InnerText;
                string lic_key = AesOperation.DecryptString(secretKey, doc2.SelectSingleNode("/MEG/LICENSE_KEY[1]").InnerText);
                string mac_adr = doc2.SelectSingleNode("/MEG/MAC[1]").InnerText;

                SKGL.Validate validateLicense = new SKGL.Validate();
                validateLicense.secretPhase = licenseSecretPhase;
                validateLicense.Key = lic_key;
                if (validateLicense.IsValid)
                {
                    label10.Text = "Lisanslı Program. Lisans bitiş tarihi: " + validateLicense.ExpireDate + " Kalan gün: " + validateLicense.DaysLeft;
                    label10.ForeColor = System.Drawing.Color.Green;
                    lisansText.Text = "Lisans Doğrulandı";
                    lisansText.Enabled = false;
                    button4.Enabled = false;
                    if(validateLicense.DaysLeft < 10)
                    {
                        label10.ForeColor = System.Drawing.Color.DarkOrange;
                    }
                    if(validateLicense.DaysLeft < 3)
                    {
                        MessageBox.Show("Programın lisansının bitmesine " + validateLicense.DaysLeft + " gün kalmış. Uzatmak isterseniz lütfen Meg Bilişim İle iletişime geçin!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        label10.ForeColor = System.Drawing.Color.Red;
                    }
                }
                else
                {
                    label10.Text = "Program lisanslı değil.";
                    label10.ForeColor = System.Drawing.Color.Red;
                    MessageBox.Show("Programınızın Lisansı Yapılmamış Gözüküyor. Lütfen Meg Bilişim ile iletişime geçiniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                label10.Text = "Program lisanslı değil.";
                label10.ForeColor = System.Drawing.Color.Red;
                MessageBox.Show("Programınızın Lisansı Yapılmamış Gözüküyor. Lütfen Meg Bilişim ile iletişime geçiniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region ButtonEvents
        private void CheckBox1_CheckedChanged(Object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                MessageBox.Show("Cari hesap yoksa oluşturulacak. Varsa direkt işlenecek.", "Dikkat", MessageBoxButtons.OK, MessageBoxIcon.Information); ;
            }
            else
            {
                MessageBox.Show("Cari hesap olup olmadığına bakılmaksızın, yeni cari hesap oluşturulacak.", "Dikkat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CheckBox2_CheckedChanged(Object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                MessageBox.Show("Seri Lot takibi ve kaydı yapılacak. Exceldeki Seri/Lot alanını Boş bırakmayınız.", "Dikkat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Seri/Lot takibi ve kaydı yapılmayacak. Exceldeki Seri/Lot alanı boş bırakılabilir.", "Dikkat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CheckBox3_CheckedChanged(Object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                MessageBox.Show("Miktar işlenmesi yapılacak. Excel'de miktar alanını doldurmayı unutmayın.", "Dikkat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Excelde miktar alanı olmayacak. Miktar 1 olarak girilecek", "Dikkat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Textbox9_TextChanged(Object sender, EventArgs e)
        {
            label13.Text = "Cari Kod: " + textBox9.Text + "." + textBox10.Text;
        }

        private void Textbox10_TextChanged(Object sender, EventArgs e)
        {
            label13.Text = "Cari Kod: " + textBox9.Text + "." + textBox10.Text;
        }

        private void Textbox11_TextChanged(Object sender, EventArgs e)
        {
            string temp = textBox13.Text;
            string temp1 = temp.Substring(Convert.ToInt32(numericUpDown1.Value), Convert.ToInt32(numericUpDown2.Value));
            label18.Text = "-" + temp1 + "-";
        }

        private void Textbox12_TextChanged(Object sender, EventArgs e)
        {
            string temp = textBox13.Text;
            string temp1 = temp.Substring(Convert.ToInt32(numericUpDown1.Value), Convert.ToInt32(numericUpDown2.Value));
            label18.Text = "-" + temp1 + "-";
        }
        #endregion

        #region ButtonClicks
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text == "" && textBox5.Text == "" && textBox6.Text == "" && textBox7.Text == "" && textBox8.Text == "" && textBox9.Text == "" && textBox10.Text == "")
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
                writer.WriteElementString("DB_NAME", textBox2.Text);
                writer.WriteElementString("DBN", textBox8.Text);
                writer.WriteElementString("DB_USERNAME", textBox3.Text);
                writer.WriteElementString("DB_PASSWORD", textBox4.Text);
                writer.WriteElementString("COMPANY_NO", Convert.ToInt32(textBox5.Text).ToString("000"));
                writer.WriteElementString("INVOICE_PERIOD", Convert.ToInt32(textBox6.Text).ToString("00"));
                writer.WriteElementString("FILE_PATH", textBox7.Text);
                writer.WriteElementString("SHEET_NAME", textBox1.Text);
                writer.WriteElementString("CREATE_ARP", checkBox1.Checked.ToString());
                writer.WriteElementString("CREATE_SLT", checkBox2.Checked.ToString());
                writer.WriteElementString("QUANTITY", checkBox3.Checked.ToString());
                writer.WriteElementString("ARP_PREFIX", textBox9.Text);
                writer.WriteElementString("ARP_CODE", textBox10.Text);
                writer.WriteElementString("WH_SUBSTRING_START", numericUpDown1.Text);
                writer.WriteElementString("WH_SUBSTRING_LEN", numericUpDown2.Text);
                writer.WriteElementString("WH_EXAMPLE", textBox13.Text);
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
            connetionString = @"Data Source=" + textBox2.Text + ";Initial Catalog=" + textBox8.Text + ";User ID=" + textBox3.Text + ";Password=" + textBox4.Text;
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
        #endregion

        #region License
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
            writer.WriteElementString("DB_NAME", "");
            writer.WriteElementString("DBN", "");
            writer.WriteElementString("DB_USERNAME", "");
            writer.WriteElementString("DB_PASSWORD", "");
            writer.WriteElementString("COMPANY_NO", "");
            writer.WriteElementString("INVOICE_PERIOD", "");
            writer.WriteElementString("FILE_PATH", "");
            writer.WriteElementString("SHEET_NAME", "");
            writer.WriteElementString("CREATE_ARP", "False");
            writer.WriteElementString("CREATE_SLT", "False");
            writer.WriteElementString("QUANTITY", "False");
            writer.WriteElementString("ARP_PREFIX", "xxx.xx");
            writer.WriteElementString("ARP_CODE", "0001");
            writer.WriteElementString("WH_SUBSTRING_START", "6");
            writer.WriteElementString("WH_SUBSTRING_LEN", "5");
            writer.WriteElementString("WH_EXAMPLE", "00123.00051 - Şube Adı");

            writer.WriteEndElement();
            writer.Flush();
            writer.Close();
        }

        private void createLicenseSettings(string licenseKey)
        {

            //VERITABANINA LICENSE KEY MEG_ID VE MAC ARDESI YAZDIRMA ISLEMI
            String firstMacAddress = NetworkInterface
                .GetAllNetworkInterfaces()
                .Where(nic => nic.OperationalStatus == OperationalStatus.Up && nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                .Select(nic => nic.GetPhysicalAddress().ToString())
                .FirstOrDefault();

            string megId = generateMegID();

            var sts = new XmlWriterSettings()
            {
                Indent = true,
                Encoding = Encoding.GetEncoding("ISO-8859-9"),
                ConformanceLevel = ConformanceLevel.Document,
                IndentChars = ("  "),
            };
            //MEG Settings
            string encryptedLicenseKey = AesOperation.EncryptString(secretKey, licenseKey);
            XmlWriter writer2 = XmlWriter.Create(AppDomain.CurrentDomain.BaseDirectory + "meg.xml", sts);
            writer2.WriteStartElement("MEG");
            writer2.WriteElementString("PRG_NAME", "XLS TO XML CONVERTER");
            writer2.WriteElementString("LICENSE_TYPE", "365");
            writer2.WriteElementString("MEG_ID", megId);
            writer2.WriteElementString("LICENSE_KEY", encryptedLicenseKey);
            //VERITABANINDAN KONTROL SAĞLANMALI
            writer2.WriteElementString("MAC", firstMacAddress);
            writer2.WriteEndElement();
            writer2.Flush();
            writer2.Close();

            //insertLicenseToDb(licenseKey, megId, firstMacAddress);
        }

        private void insertLicenseToDb(string licenseKey, string megID, string macAddress)
        {
            string connetionString, sql;
            SqlConnection cnn;
            connetionString = @"Data Source=" + textBox2.Text + ";Initial Catalog=" + textBox8.Text + ";User ID=" + textBox3.Text + ";Password=" + textBox4.Text;
            sql = "INSERT INTO License (LicenseKey, MEGID, MacAddress, UserCount)VALUES ('" + licenseKey + "', '" + megID + "', '" + macAddress + "', 1)";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            if (cmd.ExecuteNonQuery() <= 0)
            {
                MessageBox.Show("INSERT HATA");
            }
            cnn.Close();
        }

        private string generateMegID()
        {
            Random rnd = new Random();
            int id = rnd.Next(1000000000, int.MaxValue);
            return "MEG-XML-" + id.ToString("000-000-000");
        }

        private bool isLicensed(string megID, string licenseKey)
        {
            String macAddr = NetworkInterface
                .GetAllNetworkInterfaces()
                .Where(nic => nic.OperationalStatus == OperationalStatus.Up && nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                .Select(nic => nic.GetPhysicalAddress().ToString())
                .FirstOrDefault();
            string connetionString, sql, dbLicenseKey = "", dbMegID = "", dbMacAddress = "";
            int userCount = 0;
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + textBox2.Text + ";Initial Catalog=DbXls2Xml;User ID=" + textBox3.Text + ";Password=" + textBox4.Text;
            sql = "SELECT LicenseKey, MEGID, MacAddress, UserCount FROM License WHERE LicenseKey = '" + licenseKey + "'";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            if (dataReader.Read())
            {
                dbLicenseKey = dataReader.GetValue(0).ToString();
                dbMegID = dataReader.GetValue(1).ToString();
                dbMacAddress = dataReader.GetValue(2).ToString();
                userCount = int.Parse(dataReader.GetValue(3).ToString());
            }
            cnn.Close();

            if (megID == dbMegID && licenseKey == dbLicenseKey && macAddr == dbMacAddress && userCount <= 1)
            {
                return true;
            }
            return false;
        }
        
            private void button4_Click(object sender, EventArgs e)
        {
            SKGL.Validate validateLicense = new SKGL.Validate();
            validateLicense.secretPhase = licenseSecretPhase;
            validateLicense.Key = lisansText.Text;
            if (validateLicense.IsValid)
            {
                createLicenseSettings(validateLicense.Key);
                MessageBox.Show("Lisanslama işlemi Başarılı.", "Tebrikler");
            }
            else
            {
                MessageBox.Show("Lisanslama işlemi Başarısız. Lütfen Doğru Anahtar Girin.", "Başarısız", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    }
}

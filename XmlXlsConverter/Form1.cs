using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;


namespace XmlXlsConverter
{
    public partial class Form1 : Form
    {
        static string xlsFilePath, xlsFileName, xlsFileExtension;
        static string secretKey = "7548megbilisim587469as1f7dsa4785";
        static string licenseSecretPhase = "m123%M321.";
        string dbName, dbUser, dbPass, compNo, invPeriod, filePath, sheetName, dbn, arpPrefix, arpSeri, whSubStart, whSubLen;
        bool dbError = false, cariControl, seriControl, hasQuantity, licensed = true;
        List<string> errorList = new List<string>();

        public Form1()
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            Thread thread = new Thread(new ThreadStart(StartForm));
            thread.Start();
            Control.CheckForIllegalCrossThreadCalls = false;
            Thread.Sleep(5000);
            InitializeComponent();
            thread.Abort();
            if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory + "settings.xml"))
            {
                MessageBox.Show("Uygulama Ayarlarınız Henüz Yapılmamış Gözüküyor. Lütfen Ayarlar Butonuna Tıklayıp Ayarlarınızı Girin", "Merhaba", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Form2 f2 = new Form2();
                f2.ShowDialog();
            }
            if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory + "meg.xml") || !validator())
            {
                MessageBox.Show("Programınızın Lisansı Yapılmamış Gözüküyor. Lütfen Meg Bilişim ile iletişime geçiniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error);
                licensed = false;
            }
            getSettings();
            textBox3.Text = sheetName;
        }

        public void StartForm()
        {
            Application.Run(new SplashScreen());
        }

        private void getSettings()
        {
            string tempS;
            XmlDocument doc = new XmlDocument();
            doc.Load(AppDomain.CurrentDomain.BaseDirectory + "settings.xml");
            dbName = doc.SelectSingleNode("/Settings/DB_NAME[1]").InnerText;
            dbn = doc.SelectSingleNode("/Settings/DBN[1]").InnerText;
            dbUser = doc.SelectSingleNode("/Settings/DB_USERNAME[1]").InnerText;
            dbPass = doc.SelectSingleNode("/Settings/DB_PASSWORD[1]").InnerText;
            compNo = doc.SelectSingleNode("/Settings/COMPANY_NO[1]").InnerText;
            invPeriod = doc.SelectSingleNode("/Settings/INVOICE_PERIOD[1]").InnerText;
            tempS = doc.SelectSingleNode("/Settings/FILE_PATH[1]").InnerText;
            sheetName = doc.SelectSingleNode("/Settings/SHEET_NAME[1]").InnerText;
            cariControl = bool.Parse(doc.SelectSingleNode("/Settings/CREATE_ARP[1]").InnerText);
            seriControl = bool.Parse(doc.SelectSingleNode("/Settings/CREATE_SLT[1]").InnerText);
            arpPrefix = doc.SelectSingleNode("/Settings/ARP_PREFIX[1]").InnerText;
            arpSeri = doc.SelectSingleNode("/Settings/ARP_CODE[1]").InnerText;
            hasQuantity = bool.Parse(doc.SelectSingleNode("/Settings/QUANTITY[1]").InnerText);
            whSubStart = doc.SelectSingleNode("/Settings/WH_SUBSTRING_START[1]").InnerText;
            whSubLen = doc.SelectSingleNode("/Settings/WH_SUBSTRING_LEN[1]").InnerText;
            filePath = tempS.Replace("/", "\\");
            filePath += "\\";
        }
        private bool validator()
        {
            XmlDocument doc2 = new XmlDocument();
            doc2.Load(AppDomain.CurrentDomain.BaseDirectory + "meg.xml");
            string lic_key = AesOperation.DecryptString(secretKey, doc2.SelectSingleNode("/MEG/LICENSE_KEY[1]").InnerText);

            SKGL.Validate validateLicense = new SKGL.Validate();
            validateLicense.secretPhase = licenseSecretPhase;
            validateLicense.Key = lic_key;
            if (validateLicense.IsValid && !validateLicense.IsExpired)
            {
                return true;
            }
            return false;
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.megyazilim.com.tr");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = ofd.FileName;
                xlsFilePath = ofd.FileName;
                int pos = xlsFilePath.LastIndexOf("\\");
                int pos2 = xlsFilePath.LastIndexOf(".");
                int len = xlsFilePath.Length;
                xlsFileName = xlsFilePath.Substring(pos, pos2 - pos);
                int posExt = len - pos2;
                xlsFileExtension = xlsFilePath.Substring(pos2 + 1);

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            getSettings();
            if (validateLicense())
            {
                if (textBox3.Text == "" || xlsFilePath == "")
                {
                    MessageBox.Show("Lütfen dosya seçin ve sayfa adını yazın!");
                }
                else
                {
                    string pageName = textBox3.Text;
                    try
                    {
                        OleDbConnection MyConnection;
                        DataSet ds;
                        OleDbDataAdapter MyCommand;
                        if (xlsFileExtension == "xls")
                            MyConnection = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + xlsFilePath + "';Extended Properties=Excel 8.0;");
                        else
                            MyConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + xlsFilePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"");
                        MyCommand = new OleDbDataAdapter("select * from [" + pageName + "$] WHERE NOT ([Fatura Numarası] = '')", MyConnection);
                        MyCommand.TableMappings.Add("Table", "Fields");
                        ds = new DataSet();
                        MyCommand.Fill(ds);
                        MyConnection.Close();
                        dataGridView1.DataSource = ds.Tables[0];
                        button4.Enabled = false;
                        dbError = false;
                        BuildXml(ds);
                        button4.Enabled = true;
                        string errorString = "XML dosyası oluşturuldu ancal aşağıdaki hata(lar) nedeni ile içeri aktarım yapılamaz!.\n\n Lütfen hataları kontrol edip tekrar deneyin \n\n";
                        foreach (string err in errorList)
                        {
                            errorString += err + "\n";
                        }
                        if (errorList.Count != 0)
                            MessageBox.Show(errorString, "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        MessageBox.Show("XML dosyası " + xlsFileName + ".xml ismi ile seçili dosya olunda oluşturuldu.");
                        errorList.Clear();
                        errorString = "";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
            else
            {
                MessageBox.Show("Programınızın Lisansı Yapılmamış Gözüküyor. Lütfen Meg Bilişim ile iletişime geçiniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Dictionary<string, string> getSerialDetails(string serialNo)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            string connetionString, sql, slRef, serialCode = "";
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=" + dbn + ";User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT LOGICALREF, STATE, CODE from LG_" + compNo + "_" + invPeriod + "_SERILOTN where CODE = '" + serialNo + "'";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            if (dataReader.Read())
            {
                dict["STATE"] = dataReader.GetValue(1).ToString();
                slRef = dataReader.GetValue(0).ToString();
                serialCode = dataReader.GetValue(2).ToString();
            }
            else
            {
                dict["STATE"] = "";
                slRef = "";
                errorList.Add("Seri No (" + serialNo + ") Mevcut değil!");
            }
            cnn.Close();
            if (slRef != "")
                dict["SOURCE_SLT_REFERENCE"] = getSltRef(slRef);
            else
                dict["SOURCE_SLT_REFERENCE"] = "";
            return dict;
        }

        private string getSltRef(string slRef)
        {
            string connetionString, sql, sltRef;
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=" + dbn + ";User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT LOGICALREF FROM LG_" + compNo + "_" + invPeriod + "_SLTRANS WHERE SLREF = " + slRef;
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            if (dataReader.Read())
            {
                sltRef = dataReader.GetValue(0).ToString();
            }
            else
            {
                sltRef = "";
                errorList.Add("Seri No Mevcut değil. Lütfen kontrol edip tekrar deneyin.");
            }
            cnn.Close();
            return sltRef;
        }

        private string createArpCode()
        {
            string connetionString, sql, arpCode;
            int code, prefixLen = arpPrefix.Length + 2;
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=" + dbn + ";User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT MAX(CONVERT(INT,SUBSTRING(CODE, " + prefixLen + ", 100))) from LG_" + compNo + "_CLCARD WHERE CODE LIKE'" + arpPrefix + ".%'";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            if (dataReader.Read())
            {
                string data = dataReader.GetValue(0).ToString();
                if (data == "")
                    return arpPrefix + "." + arpSeri;
                code = Convert.ToInt32(dataReader.GetValue(0).ToString());
                code++;
                arpCode = arpPrefix + "." + code.ToString("000");
            }
            else
            {
                arpCode = arpPrefix + "." + arpSeri;
            }
            cnn.Close();
            return arpCode;
        }

        private Dictionary<string, string> getArpCode(string taxnr)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            string connetionString, sql, arpCode = "", eInvoice = "";
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=" + dbn + ";User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT CODE, ACCEPTEINV FROM LG_" + compNo + "_CLCARD WHERE TAXNR = '" + taxnr + "'";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            if (dataReader.Read())
            {
                arpCode = dataReader.GetValue(0).ToString();
                eInvoice = dataReader.GetValue(1).ToString();
            }
            cnn.Close();
            dict["arpCode"] = arpCode;
            if (eInvoice == "")
                eInvoice = "0";
            dict["eInvoice"] = eInvoice;
            return dict;
        }

        private Dictionary<string, string> getMasterCodeFromDef(string stokAdi)
        {
            string connetionString, sql;
            Dictionary<string, string> fields = new Dictionary<string, string>();
            Dictionary<string, string> unitInfo = new Dictionary<string, string>();
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=" + dbn + ";User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT LOGICALREF, CODE FROM LG_" + compNo + "_ITEMS WHERE NAME = '" + stokAdi + "'";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            if (dataReader.Read())
            {
                fields["logicalRef"] = dataReader.GetValue(0).ToString();
                fields["masterCode"] = dataReader.GetValue(1).ToString();
            }
            else
            {
                errorList.Add(stokAdi + " isminde malzeme bulunamadı.");
                return fields;
            }
            unitInfo = getUnitInfo(stokAdi);
            fields["unitCode"] = unitInfo["code"];
            fields["globalCode"] = unitInfo["globalCode"];
            fields["conv1"] = unitInfo["conv1"];
            fields["conv2"] = unitInfo["conv2"];
            cnn.Close();
            return fields;
        }

        private Dictionary<string, string> getMasterCode(string seriNo)
        {
            string connetionString, sql;
            Dictionary<string, string> unitInfo;
            Dictionary<string, string> fields = new Dictionary<string, string>();
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=" + dbn + ";User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT CODE, NAME FROM LG_" + compNo + "_ITEMS WHERE LOGICALREF = (SELECT TOP 1 ITEMREF FROM LG_" + compNo + "_" + invPeriod + "_SERILOTN WHERE CODE = '" + seriNo + "')";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            if (dataReader.Read())
            {
                fields["masterCode"] = dataReader.GetValue(0).ToString();
                fields["masterName"] = dataReader.GetValue(1).ToString();
            }
            else
            {
                return fields;
            }
            unitInfo = getUnitInfo(fields["masterName"]);
            fields["unitCode"] = unitInfo["code"];
            fields["globalCode"] = unitInfo["globalCode"];
            fields["conv1"] = unitInfo["conv1"];
            fields["conv2"] = unitInfo["conv2"];
            cnn.Close();
            return fields;
        }

        private Dictionary<string, string> getUnitInfo(string masterCode)
        {
            string connetionString, sql;
            Dictionary<string, string> dict = new Dictionary<string, string>();
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=" + dbn + ";User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT CODE, GLOBALCODE, CONVFACT1, CONVFACT2 FROM LG_" + compNo + "_UNITSETL WHERE UNITSETREF = (SELECT TOP 1 UNITSETREF FROM LG_" + compNo + "_ITEMS WHERE NAME = '" + masterCode + "')";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            if (dataReader.Read())
            {
                dict["code"] = dataReader.GetValue(0).ToString();
                dict["globalCode"] = dataReader.GetValue(1).ToString();
                dict["conv1"] = dataReader.GetValue(2).ToString();
                dict["conv2"] = dataReader.GetValue(3).ToString();
            }
            cnn.Close();
            return dict;
        }

        private string getAmbarKodu(string name)
        {
            string connetionString, sql, ambarKodu = "0";
            if (name == "00005")
            {
                return "9";
            }
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=" + dbn + ";User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT NR FROM L_CAPIWHOUSE WHERE NAME = '" + Convert.ToInt32(name) + "' AND FIRMNR = " + Convert.ToInt32(compNo);
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            if (dataReader.Read())
            {
                ambarKodu = dataReader.GetValue(0).ToString();
            }
            cnn.Close();
            return ambarKodu;
        }

        private bool seridenAmbarKontrol(string seriNo, string ambarNo)
        {
            string connetionString, sql, ambarKodu = "";
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=" + dbn + ";User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT TOP 1 INVENNO FROM LG_" + compNo + "_" + invPeriod + "_SLTRANS WHERE SLREF = (SELECT TOP 1 LOGICALREF FROM LG_" + compNo + "_" + invPeriod + "_SERILOTN WHERE CODE = '" + seriNo + "' ) ORDER BY DATE_ DESC, LOGICALREF DESC";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            if (dataReader.Read())
            {
                ambarKodu = dataReader.GetValue(0).ToString();
                if(ambarKodu != null && ambarKodu == ambarNo)
                {
                    return true;
                }
            }
            cnn.Close();
            if (ambarKodu == "")
            {
                //errorList.Add(seriNo + " seri numarası kaydı hiçbir ambarda bulunamadı.");
                return true;
            }
            else
                errorList.Add(seriNo + " seri numaralı ürün " + ambarNo + " kodlu ambarda değil." + ambarKodu + " no'lu ambarda kayıtlıdır!");
            return false;
        }

        private void insertCari(string name, string tc, string arpCode, string eInvoice)
        {
            string connetionString, sql;
            SqlConnection cnn;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=" + dbn + ";User ID=" + dbUser + ";Password=" + dbPass;

            sql = "INSERT INTO LG_" + compNo + "_CLCARD(ACTIVE, CARDTYPE, CODE, DEFINITION_, TAXNR, PURCHBRWS, SALESBRWS, IMPBRWS, EXPBRWS, FINBRWS, ACCEPTEINV) VALUES(0, 3, '" + arpCode + "','" + name + "','" + tc + "', 1, 1, 1, 1, 1, " + eInvoice + ")";

            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            if (!dbError && cmd.ExecuteNonQuery() <= 0)
            {
                dbError = true;
                errorList.Add("Cari oluşturulamadı!, Lütfen SQL ayarlarınızı kontrol edin.");
            }
            cnn.Close();
        }

        private void insertSeri(string itemRef, string seriNo)
        {
            string connetionString, sql;
            SqlConnection cnn;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=" + dbn + ";User ID=" + dbUser + ";Password=" + dbPass;
            sql = "INSERT INTO LG_" + compNo + "_" + invPeriod + "_SERILOTN (ITEMREF, SLTYPE, CODE, NAME, STATE, SITEID, RECSTATUS, ORGLOGICREF, WFSTATUS, ORGLOGOID, CAPIBLOCK_CREATEDBY, CAPIBLOCK_CREADEDDATE, CAPIBLOCK_CREATEDHOUR, CAPIBLOCK_CREATEDMIN, CAPIBLOCK_CREATEDSEC, CAPIBLOCK_MODIFIEDBY, CAPIBLOCK_MODIFIEDHOUR,CAPIBLOCK_MODIFIEDMIN, CAPIBLOCK_MODIFIEDSEC,VARIANTREF, GROUPLOTNO)" +
                " VALUES(" + itemRef + ",2," + seriNo + ",'', 0,0, 1, 0, 0, '', 1, '2021-08-24', 11,40,24,0,0,0,0,0,'')";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            if (!dbError && cmd.ExecuteNonQuery() <= 0)
            {
                dbError = true;
                errorList.Add("Seri/Lot Oluşturulamadı!, Lütfen SQL ayarlarınızı kontrol edin.");
            }
            cnn.Close();
        }

        private bool serialControl(string seriNo)
        {
            string connetionString, sql;
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=" + dbn + ";User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT CODE FROM LG_" + compNo + "_" + invPeriod + "_SERILOTN WHERE CODE = '" + seriNo + "'";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            if (dataReader.Read())
            {
                if (seriNo == dataReader.GetValue(0).ToString())
                {
                    return false;
                }
            }
            cnn.Close();
            return true;
        }

        private string yaziyaCevir(decimal tutar)
        {
            string sTutar = tutar.ToString("F2").Replace('.', ','); // Replace('.',',') ondalık ayracının . olma durumu için
            string lira = sTutar.Substring(0, sTutar.IndexOf(',')); //tutarın tam kısmı
            string kurus = sTutar.Substring(sTutar.IndexOf(',') + 1, 2);
            string yazi = "";
            string[] birler = { "", "Bir", "İki", "Üç", "Dört", "Beş", "Altı", "Yedi", "Sekiz", "Dokuz" };
            string[] onlar = { "", "On", "Yirmi", "Otuz", "Kırk", "Elli", "Altmış", "Yetmiş", "Seksen", "Doksan" };
            string[] binler = { "Katrilyon", "Trilyon", "Milyar", "Milyon", "Bin", "" }; //KATRİLYON'un önüne ekleme yapılarak artırabilir.
            int grupSayisi = binler.Length; //sayıdaki 3'lü grup sayısı. katrilyon içi 6. (1.234,00 daki grup sayısı 2'dir.)
            lira = lira.PadLeft(grupSayisi * 3, '0'); //sayının soluna '0' eklenerek sayı 'grup sayısı x 3' basakmaklı yapılıyor.
            string grupDegeri;
            for (int i = 0; i < grupSayisi * 3; i += 3) //sayı 3'erli gruplar halinde ele alınıyor.
            {
                grupDegeri = "";
                if (lira.Substring(i, 1) != "0")
                    grupDegeri += birler[Convert.ToInt32(lira.Substring(i, 1))] + "Yüz"; //yüzler
                if (grupDegeri == "BirYüz") //biryüz düzeltiliyor.
                    grupDegeri = "Yüz";
                grupDegeri += onlar[Convert.ToInt32(lira.Substring(i + 1, 1))]; //onlar
                grupDegeri += birler[Convert.ToInt32(lira.Substring(i + 2, 1))]; //birler
                if (grupDegeri != "") //binler
                    grupDegeri += binler[i / 3];
                if (grupDegeri == "BirBin") //birbin düzeltiliyor.
                    grupDegeri = "Bin";
                yazi += grupDegeri;
            }
            if (yazi != "")
                yazi += " TL";
            int yaziUzunlugu = yazi.Length;
            if (kurus.Substring(0, 1) != "0") //kuruş onlar
                yazi += onlar[Convert.ToInt32(kurus.Substring(0, 1))];
            if (kurus.Substring(1, 1) != "0") //kuruş birler
                yazi += birler[Convert.ToInt32(kurus.Substring(1, 1))];
            if (yazi.Length > yaziUzunlugu)
                yazi += " kuruş";
            return yazi;
        }

        private bool validateLicense()
        {
            if (licensed)
            {
                XmlDocument doc2 = new XmlDocument();
                doc2.Load(AppDomain.CurrentDomain.BaseDirectory + "meg.xml");
                string lic_key = AesOperation.DecryptString(secretKey, doc2.SelectSingleNode("/MEG/LICENSE_KEY[1]").InnerText);

                SKGL.Validate validateLicense = new SKGL.Validate();
                validateLicense.secretPhase = licenseSecretPhase;
                validateLicense.Key = lic_key;
                if (validateLicense.IsValid)
                {
                    licensed = true;
                    return true;
                }
            }

            return false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();
            button4.Enabled = true;
        }

        private void BuildXml(DataSet ds)
        {
            bool satis = true;
            if (satinalmaRadio.Checked)
                satis = false;

            var sts = new XmlWriterSettings()
            {
                Indent = true,
                Encoding = Encoding.GetEncoding("ISO-8859-9"),
                ConformanceLevel = ConformanceLevel.Document,
                IndentChars = ("  "),
            };

            XmlWriter writer = XmlWriter.Create(filePath + xlsFileName + ".xml", sts);
            if (satis)
                writer.WriteStartElement("SALES_INVOICES");
            else
                writer.WriteStartElement("PURCHASE_INVOICES");

            foreach (DataRow row in ds.Tables[0].Rows)
            {
                string faturaNo = row["Fatura Numarası"].ToString(),
                    cariAdi = row["Firma/Kişi Adı"].ToString(),
                    cariTc = row["VKN/TCKN"].ToString(),
                    toplamTutar = row["Ödenecek Tutar"].ToString().Replace(",", "."),
                    not = row["Not"].ToString(),
                    seriNo = row["Seri/IMEI No"].ToString(),
                    sube = row["Şube"].ToString(),
                    ambarAdi = sube.Substring(int.Parse(whSubStart), int.Parse(whSubLen)),
                    stokAdi = row["Stok Adı"].ToString(),
                    ozelKod = "2",
                    toplamTutarYazi = yaziyaCevir(Math.Round(Convert.ToDecimal(toplamTutar), 2)),
                    tarih = row["Fatura Tarihi"].ToString().Replace(".", "/"),
                    tarihFormat = "dd/MM/yyyy",
                    miktar = "1";

                if (hasQuantity)
                    miktar = row["Miktar"].ToString();
                DateTime faturaTarih = DateTime.Parse(tarih);

                string formatliTarih = faturaTarih.ToString(tarihFormat);
                string cariKodu, eInvoice = "0";

                if (cariControl)
                {
                    cariKodu = getArpCode(cariTc)["arpCode"];
                    eInvoice = getArpCode(cariTc)["eInvoice"];
                    if (cariKodu == "")
                    {
                        DialogResult dr = new DialogResult();
                        dr = MessageBox.Show(cariTc + " numarasına sahip cari bulunamadı. Oluşturmak ister misiniz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (dr == DialogResult.Yes)
                        {
                            cariKodu = createArpCode();
                            insertCari(cariAdi, cariTc, cariKodu, eInvoice);
                            MessageBox.Show("Yeni cari " + cariKodu + " koduyla oluşturuldu.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show(cariTc + " nosuna cari oluşturmadınız. Bilgileri kontrol edip tekrar deneyin", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                {
                    if (satis)
                        cariKodu = createArpCode();
                    else
                    {
                        cariKodu = getArpCode(cariTc)["arpCode"];
                        if (cariKodu == "")
                            errorList.Add(cariTc + " numarasına sahip cari bulunamadı.");
                    }
                }

                string ambarKodu = getAmbarKodu(ambarAdi);
                if (satis && !seridenAmbarKontrol(seriNo, ambarKodu))
                    break;
                Decimal kdvOraniD = 18;
                Decimal kdvsizTutarD = Math.Round(Convert.ToDecimal(toplamTutar) / (1 + (kdvOraniD / 100)), 2);
                Decimal kdvD = Math.Round(Convert.ToDecimal(toplamTutar) - kdvsizTutarD, 2);
                string kdvOrani = kdvOraniD.ToString().Replace(",", ".");
                string kdvsizTutar = kdvsizTutarD.ToString().Replace(",", ".");
                string kdv = kdvD.ToString().Replace(",", ".");
                Dictionary<string, string> masterFields;
                string masterCode = "", masterDef = "", itemRef = "", unitCode = "", globalCode = "", conv1 = "", conv2 = "";

                if (seriControl && satis)
                {
                    masterFields = getMasterCode(seriNo);
                    if (masterFields.ContainsKey("masterCode"))
                    {
                        masterCode = masterFields["masterCode"];
                        masterDef = masterFields["masterName"];
                        unitCode = masterFields["unitCode"];
                        globalCode = masterFields["globalCode"];
                        conv1 = masterFields["conv1"];
                        conv2 = masterFields["conv2"];
                    }
                }
                else
                {
                    masterFields = getMasterCodeFromDef(stokAdi);
                    if (masterFields.ContainsKey("masterCode"))
                    {
                        itemRef = masterFields["logicalRef"];
                        masterCode = masterFields["masterCode"];
                        masterDef = stokAdi;
                        unitCode = masterFields["unitCode"];
                        globalCode = masterFields["globalCode"];
                        conv1 = masterFields["conv1"];
                        conv2 = masterFields["conv2"];
                    }
                }

                Dictionary<string, string> serialDetails = new Dictionary<string, string>();

                if (seriControl)
                {
                    if (satis)
                    {
                        serialDetails = getSerialDetails(seriNo);
                        if (serialDetails["STATE"] == "2")
                            errorList.Add("Seri No (" + seriNo + ") Daha Önceden İşlenmiş!");
                        else
                            if (!cariControl)
                            insertCari(cariAdi, cariTc, cariKodu, "0");

                    }
                    else
                    {
                        if (serialControl(seriNo))
                            insertSeri(itemRef, seriNo);
                        else
                            errorList.Add(seriNo + " seri nosu daha önce işlenmiş.");
                    }
                }
                else
                {
                    if (!cariControl)
                        insertCari(cariAdi, cariTc, cariKodu, "0");
                }

                writer.WriteStartElement("INVOICE");
                writer.WriteAttributeString("DBOP", "INS");
                writer.WriteElementString("INTERNAL_REFERENCE", "1");
                if (satis)
                    writer.WriteElementString("TYPE", "7");
                else
                    writer.WriteElementString("TYPE", "1");
                writer.WriteElementString("NUMBER", faturaNo);
                writer.WriteElementString("DATE", formatliTarih);
                writer.WriteElementString("TIME", "254936320");
                if (satis)
                {
                    writer.WriteElementString("DOC_NUMBER", faturaNo);
                    writer.WriteElementString("AUXIL_CODE", ozelKod);
                }
                writer.WriteElementString("ARP_CODE", cariKodu);
                writer.WriteElementString("SOURCE_WH", ambarKodu);
                writer.WriteElementString("SOURCE_COST_GRP", ambarKodu);
                writer.WriteElementString("POST_FLAGS", "247");
                writer.WriteElementString("VAT_RATE", "18");
                writer.WriteElementString("TOTAL_DISCOUNTED", kdvsizTutar);
                //writer.WriteElementString("TOTAL_VAT", kdv);
                //writer.WriteElementString("TOTAL_GROSS", kdvsizTutar);
                //writer.WriteElementString("TOTAL_NET", toplamTutar);
                writer.WriteElementString("NOTES1", not);
                writer.WriteElementString("TC_NET", toplamTutar);
                writer.WriteElementString("RC_XRATE", "1");
                writer.WriteElementString("RC_NET", toplamTutar);
                writer.WriteElementString("VAT_INCLUDED_GRS", "1");
                writer.WriteElementString("CREATED_BY", "1");
                writer.WriteElementString("DATE_CREATED", formatliTarih);
                writer.WriteElementString("HOUR_CREATED", faturaTarih.ToString("hh"));
                writer.WriteElementString("MIN_CREATED", faturaTarih.ToString("mm"));
                writer.WriteElementString("SEC_CREATED", faturaTarih.ToString("ss"));
                writer.WriteElementString("CURRSEL_TOTALS", "1");
                writer.WriteElementString("DATA_REFERENCE", "0");
                writer.WriteStartElement("DISPATCHES");
                writer.WriteStartElement("DISPATCH");
                writer.WriteElementString("INTERNAL_REFERENCE", "1");
                if (satis)
                    writer.WriteElementString("TYPE", "7");
                else
                    writer.WriteElementString("TYPE", "1");
                writer.WriteElementString("NUMBER", faturaNo);
                writer.WriteElementString("DATE", formatliTarih);
                writer.WriteElementString("TIME", "254936320");
                writer.WriteElementString("INVOICE_NUMBER", faturaNo);
                writer.WriteElementString("ARP_CODE", cariKodu);
                writer.WriteElementString("SOURCE_WH", ambarKodu);
                writer.WriteElementString("SOURCE_COST_GRP", ambarKodu);
                writer.WriteElementString("INVOICED", "1");
                writer.WriteElementString("TOTAL_DISCOUNTED", kdvsizTutar);
                //writer.WriteElementString("TOTAL_VAT", kdv);
                //writer.WriteElementString("TOTAL_GROSS", kdvsizTutar);
                //writer.WriteElementString("TOTAL_NET", toplamTutar);
                writer.WriteElementString("RC_RATE", "1");
                writer.WriteElementString("RC_NET", toplamTutar);
                writer.WriteElementString("CREATED_BY", "1");
                writer.WriteElementString("DATE_CREATED", formatliTarih);
                writer.WriteElementString("HOUR_CREATED", faturaTarih.ToString("hh"));
                writer.WriteElementString("MIN_CREATED", faturaTarih.ToString("mm"));
                writer.WriteElementString("SEC_CREATED", faturaTarih.ToString("ss"));
                writer.WriteElementString("CURRSEL_TOTALS", "1");
                writer.WriteElementString("DATA_REFERENCE", "0");
                writer.WriteElementString("ORIG_NUMBER", "0000000000000001");
                writer.WriteStartElement("ORGLOGOID");
                writer.WriteRaw("");
                writer.WriteEndElement();
                writer.WriteElementString("DEDUCTIONPART1", "2");
                writer.WriteElementString("DEDUCTIONPART2", "3");
                writer.WriteElementString("AFFECT_RISK", "0");
                writer.WriteElementString("DISP_STATUS", "1");
                writer.WriteElementString("SHIP_DATE", formatliTarih);
                writer.WriteElementString("SHIP_TIME", "254936356");
                writer.WriteElementString("DOC_DATE", formatliTarih);
                writer.WriteElementString("DOC_TIME", "254936320");
                if (eInvoice == "1")
                    writer.WriteElementString("EINVOICE", "1");
                writer.WriteEndElement(); //DISPATCH
                writer.WriteEndElement(); //DISPATTCHES
                writer.Flush();
                writer.WriteStartElement("TRANSACTIONS");
                writer.WriteStartElement("TRANSACTION");
                writer.WriteElementString("INTERNAL_REFERENCE", "1");
                writer.WriteElementString("TYPE", "0");
                writer.WriteElementString("MASTER_CODE", masterCode); //SERI
                writer.WriteElementString("SOURCEINDEX", ambarKodu);
                writer.WriteElementString("SOURCECOSTGRP", ambarKodu);
                writer.WriteElementString("QUANTITY", miktar);
                writer.WriteElementString("PRICE", toplamTutar);
                writer.WriteElementString("TOTAL", toplamTutar);
                writer.WriteElementString("RC_XRATE", "1");
                writer.WriteElementString("UNIT_CODE", unitCode);
                writer.WriteElementString("UNIT_CONV1", conv1);
                writer.WriteElementString("UNIT_CONV2", conv2);
                writer.WriteElementString("VAT_INCLUDED", "1");
                writer.WriteElementString("VAT_RATE", kdvOrani);
                writer.WriteElementString("VAT_AMOUNT", kdv);//kdv
                //writer.WriteElementString("VAT_BASE", kdvsizTutar);
                writer.WriteElementString("BILLED", "1");
                //writer.WriteElementString("TOTAL_NET", kdvsizTutar);
                writer.WriteElementString("DATA_REFERENCE", "0");
                writer.WriteElementString("DISPATCH_NUMBER", "0000000000000001");
                if (seriControl)
                {
                    writer.WriteStartElement("SL_DETAILS");
                    writer.WriteStartElement("SERIAL_LOT_TRN");
                    writer.WriteElementString("INTERNAL_REFERENCE", "9");
                    if (satis)
                    {
                        writer.WriteElementString("SOURCE_MT_REFERENCE", "1");
                        writer.WriteElementString("SOURCE_SLT_REFERENCE", serialDetails["SOURCE_SLT_REFERENCE"]);
                        writer.WriteElementString("SOURCE_QUANTITY", "1");
                        writer.WriteElementString("IOCODE", "4");
                    }
                    else
                    {
                        writer.WriteElementString("SOURCE_MT_REFERENCE", "0");
                        writer.WriteElementString("SOURCE_SLT_REFERENCE", "0");
                        writer.WriteElementString("SOURCE_QUANTITY", "0");
                        writer.WriteElementString("IOCODE", "1");
                    }
                    writer.WriteElementString("SOURCE_WH", ambarKodu);
                    writer.WriteElementString("SL_TYPE", "2");
                    writer.WriteElementString("SL_CODE", seriNo);
                    writer.WriteElementString("MU_QUANTITY", "1");
                    writer.WriteElementString("UNIT_CODE", unitCode);
                    writer.WriteElementString("QUANTITY", "1");
                    if (!satis)
                    {
                        writer.WriteElementString("REM_QUANTITY", "1");
                        writer.WriteElementString("LU_REM_QUANTITY", "1");
                    }
                    writer.WriteElementString("UNIT_CONV1", conv1);
                    writer.WriteElementString("UNIT_CONV2", conv2);
                    writer.WriteElementString("DATE_EXPIRED", faturaTarih.AddMonths(-1).ToString("dd/MM/yyyy"));
                    if (!satis)
                    {
                        writer.WriteElementString("OUT_COST", toplamTutar);
                        writer.WriteElementString("TC_OUT_COST", toplamTutar);
                    }
                    writer.WriteElementString("DATA_REFERENCE", "0");
                    writer.WriteElementString("ORGLOGOID", "");
                    writer.WriteElementString("ORGLINKREF", "0");
                    writer.WriteEndElement();//SERIAL_LOT_TRN 
                    writer.WriteEndElement();//SL_DETAILS
                }
                writer.WriteStartElement("DETAILS");
                writer.WriteRaw("");
                writer.WriteEndElement();//DETAILS
                writer.WriteElementString("DIST_ORD_REFERENCE", "0");
                writer.WriteStartElement("CAMPAIGN_INFOS");
                writer.WriteStartElement("CAMPAIGN_INFO");
                writer.WriteRaw("");
                writer.WriteEndElement();//CAMPAIGN_INFO
                writer.WriteEndElement();//CAMPAIGN_INFOS
                writer.WriteElementString("MULTI_ADD_TAX", "0");
                writer.WriteElementString("EDT_CURR", "160");
                writer.WriteElementString("EDT_PRICE", toplamTutar);
                writer.WriteStartElement("ORGLOGOID");
                writer.WriteRaw("");
                writer.WriteEndElement();
                writer.WriteStartElement("GENIUSFLDSLIST");
                writer.WriteRaw("");
                writer.WriteEndElement();
                writer.WriteStartElement("DEFNFLDSLIST");
                writer.WriteRaw("");
                writer.WriteEndElement();
                writer.WriteElementString("MONTH", faturaTarih.ToString("MM"));
                writer.WriteElementString("YEAR", faturaTarih.ToString("yyyy"));
                writer.WriteStartElement("PREACCLINES");
                writer.WriteRaw("");
                writer.WriteEndElement();
                writer.WriteElementString("UNIT_GLOBAL_CODE", globalCode);
                writer.WriteElementString("EDTCURR_GLOBAL_CODE", "TL");
                writer.WriteElementString("MASTER_DEF", masterDef);
                writer.WriteElementString("FOREIGN_TRADE_TYPE", "0");
                writer.WriteElementString("DISTRIBUTION_TYPE_WHS", "0");
                writer.WriteElementString("DISTRIBUTION_TYPE_FNO", "0");
                if (!satis)
                    writer.WriteElementString("FUTURE_MONTH_BEGDATE", formatliTarih);
                writer.WriteEndElement();//TRANSACTION


                writer.WriteEndElement();//TRANSACTIONS
                writer.WriteStartElement("PAYMENT_LIST");
                writer.WriteStartElement("PAYMENT");
                writer.WriteElementString("INTERNAL_REFERENCE", "0");
                writer.WriteElementString("DATE", formatliTarih);
                writer.WriteElementString("MODULENR", "4");
                if (!satis)
                    writer.WriteElementString("SIGN", "1");
                writer.WriteElementString("TRCODE", "7");
                writer.WriteElementString("TOTAL", toplamTutar);
                writer.WriteElementString("PROCDATE", formatliTarih);
                writer.WriteElementString("REPORTRATE", "1");
                writer.WriteElementString("DATA_REFERENCE", "0");
                writer.WriteElementString("DISCOUNT_DUEDATE", formatliTarih);
                writer.WriteElementString("PAY_NO", "1");
                writer.WriteStartElement("DISCTRLIST");
                writer.WriteRaw("");
                writer.WriteEndElement();
                writer.WriteElementString("DISCTRDELLIST", "0");
                writer.WriteEndElement();//PAYMENT
                writer.WriteEndElement();//PAYMENT_LIST
                writer.WriteStartElement("ORGLOGOID");
                writer.WriteRaw("");
                writer.WriteEndElement();
                writer.WriteStartElement("DEFNFLDSLIST");
                writer.WriteRaw("");
                writer.WriteEndElement();
                writer.WriteElementString("DEDUCTIONPART1", "2");
                writer.WriteElementString("DEDUCTIONPART2", "3");
                writer.WriteElementString("DATA_LINK_REFERENCE", "1");
                writer.WriteStartElement("INTEL_LIST");
                writer.WriteStartElement("INTEL");
                writer.WriteElementString("LOGICALREF", "0");
                writer.WriteEndElement();//INTEL
                writer.WriteEndElement();//INTEL_LIST
                writer.WriteElementString("AFFECT_RISK", "0");
                writer.WriteStartElement("PREACCLINES");
                writer.WriteRaw("");
                writer.WriteEndElement();
                writer.WriteElementString("DOC_DATE", formatliTarih);
                if (eInvoice == "1")
                    writer.WriteElementString("EINVOICE", "1");
                if (!satis)
                    writer.WriteElementString("ESTATUS", formatliTarih);
                writer.WriteElementString("EDURATION_TYPE", "0");
                writer.WriteElementString("EDTCURR_GLOBAL_CODE", "TL");
                writer.WriteElementString("EINVOICE_TURETPRICESTR", "Sıfır TL");
                writer.WriteElementString("TOTAL_NET_STR", toplamTutarYazi);
                writer.WriteElementString("EXIMVAT", "0");
                writer.WriteElementString("EARCHIVEDETR_INTPAYMENTTYPE", "0");
                writer.WriteStartElement("OKCINFO_LIST");
                writer.WriteStartElement("OKCINFO");
                writer.WriteElementString("INTERNAL_REFERENCE", "0");
                writer.WriteEndElement();//OKCINFO
                writer.WriteEndElement();//OKCINFO_LIST
                writer.WriteStartElement("LABEL_LIST");
                writer.WriteRaw("");
                writer.WriteEndElement();
                writer.WriteEndElement();//INVOICE
            }
            writer.WriteEndElement();
            writer.Flush();
            writer.Close();
        }
    }
}


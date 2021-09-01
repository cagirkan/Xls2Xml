using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace XmlXlsConverter
{
    public partial class AddClCard : Form
    {
        string dbName, dbUser, dbPass, compNo, countryCode;
        Dictionary<string, int> countryDict = new Dictionary<string, int>();
        Dictionary<string, int> cityDict = new Dictionary<string, int>();
        Dictionary<string, int> townDict = new Dictionary<string, int>();

        public AddClCard()
        {
            InitializeComponent();
            getSettings();
            Form1 form1 = new Form1();
            //textBox1.Text = form1.;
            countryDict = GetCountriesFromDBTable();
            listBox1.Items.Add("TÜRKİYE");
            listBox1.SelectedIndex = 0;
            foreach (KeyValuePair<string, int> item in countryDict)
            {
                if (item.Key == "TÜRKİYE")
                    continue;
                listBox1.Items.Add(item.Key);
            }
        }

        private void getSettings()
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(AppDomain.CurrentDomain.BaseDirectory + "settings.xml");
            dbName = doc.SelectSingleNode("/Settings/DB_NAME[1]").InnerText;
            dbUser = doc.SelectSingleNode("/Settings/DB_USERNAME[1]").InnerText;
            dbPass = doc.SelectSingleNode("/Settings/DB_PASSWORD[1]").InnerText;
            compNo = doc.SelectSingleNode("/Settings/COMPANY_NO[1]").InnerText;

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            cityDict = GetCitiesFromDB(countryDict[listBox1.SelectedItem.ToString()]);
            foreach (KeyValuePair<string, int> item in cityDict)
            {
                listBox2.Items.Add(item.Key);
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            townDict = GetTownsFromDB(countryDict[listBox1.SelectedItem.ToString()], cityDict[listBox2.SelectedItem.ToString()]); 
            foreach (KeyValuePair<string, int> item in townDict)
            {
                listBox3.Items.Add(item.Key);
            }
        }

        private Dictionary<string, int> GetCountriesFromDBTable()
        {
            Dictionary<string, int> countryDict = new Dictionary<string, int>();
            string connetionString, sql;
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=LOGO;User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT NAME, COUNTRYNR FROM L_COUNTRY";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            while (dataReader.Read())
            {
                countryDict[dataReader.GetValue(0).ToString()] = Convert.ToInt32(dataReader.GetValue(1));
            }
            cnn.Close();
            return countryDict;
        }

        private Dictionary<string, int> GetCitiesFromDB(int countryNo)
        {
            Dictionary<string, int> dict = new Dictionary<string, int>();
            string connetionString, sql;
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=LOGO;User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT NAME, CODE FROM L_CITY WHERE COUNTRY = " + countryNo.ToString();
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            while (dataReader.Read())
            {
                dict[dataReader.GetValue(0).ToString()] = Convert.ToInt32(dataReader.GetValue(1));
            }
            cnn.Close();
            return dict;
        }

        private Dictionary<string, int> GetTownsFromDB(int countryNo, int cityNo)
        {
            Dictionary<string, int> dict = new Dictionary<string, int>();
            string connetionString, sql;
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=LOGO;User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT NAME, CODE FROM L_TOWN WHERE CNTRNR = " + countryNo.ToString() + " AND CTYREF = " + cityNo.ToString();
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            while (dataReader.Read())
            {
                dict[dataReader.GetValue(0).ToString()] = Convert.ToInt32(dataReader.GetValue(1));
            }
            cnn.Close();
            return dict;
        }

        private void getCountryCode()
        {
            string connetionString, sql;
            SqlConnection cnn;
            SqlDataReader dataReader;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=LOGO;User ID=" + dbUser + ";Password=" + dbPass;
            sql = "SELECT CODE FROM L_COUNTRY WHERE NAME ='" + listBox1.SelectedItem.ToString() + "'";
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            dataReader = cmd.ExecuteReader();
            if (dataReader.Read())
            {
                countryCode = dataReader.GetValue(0).ToString();
            }
            cnn.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            getCountryCode();
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("Lütfen Boş Alanları Doldurunuz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if ((listBox2.Items.Count == 0 || listBox3.Items.Count == 0) || (listBox1.SelectedIndex == -1 || listBox2.SelectedIndex == -1 || listBox3.SelectedIndex == -1))
            {
                MessageBox.Show("Lütfen Ülke, İl ve ilçe seçimlerinizi yapınız!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string connetionString, sql;
            SqlConnection cnn;
            connetionString = @"Data Source=" + dbName + ";Initial Catalog=LOGO;User ID=" + dbUser + ";Password=" + dbPass;
            if (checkBox1.Checked)
            {
                sql = "INSERT INTO LG_" + compNo + "_CLCARD(ACTIVE, CARDTYPE, CODE, DEFINITION_, ACCEPTEINV, CITY, COUNTRY, TOWN, CITYCODE, COUNTRYCODE, TOWNCODE) VALUES(0, 3, '" 
                    + textBox1.Text + "','" + textBox2.Text + "', 1, '" + listBox2.SelectedItem.ToString() +"', '" + listBox1.SelectedItem.ToString() + "', '" + listBox3.SelectedItem.ToString() + "', " + cityDict[listBox2.SelectedItem.ToString()] + ", '" +
                   countryCode + "', " + townDict[listBox3.SelectedItem.ToString()] + ")";
            }
            else
            {
                sql = "INSERT INTO LG_" + compNo + "_CLCARD(ACTIVE, CARDTYPE, CODE, DEFINITION_, ACCEPTEINV, CITY, COUNTRY, TOWN, CITYCODE, COUNTRYCODE, TOWNCODE) VALUES(0, 3, '"
                    + textBox1.Text + "','" + textBox2.Text + "', 0, '" + listBox2.SelectedItem.ToString() + "', '" + listBox1.SelectedItem.ToString() + "', '" + listBox3.SelectedItem.ToString() + "', " + cityDict[listBox2.SelectedItem.ToString()] + ", '" +
                   countryCode + "', " + townDict[listBox3.SelectedItem.ToString()] + ")";
            }
            cnn = new SqlConnection(connetionString);
            cnn.Open();
            SqlCommand cmd = new SqlCommand(sql, cnn);
            if (cmd.ExecuteNonQuery() > 0)
            {
                MessageBox.Show(textBox1.Text + " Kodlu Cari başarıyla oluşturuldu!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            else
                MessageBox.Show(textBox1.Text + " Kodlu Cari oluşturulamadı!, Lütfen SQL ayarlarınızı kontrol edin.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            cnn.Close();
        }
    }
}

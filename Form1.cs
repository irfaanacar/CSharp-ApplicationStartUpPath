using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ApplicationStartupPath
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OleDbConnection baglan=new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+Application.StartupPath+"\\Kisiler.mdb"); //Application Startup Path bulunduğu klasördeki dosyaları tarar .
        private void verilerigoster()
        {
            listView1.Items.Clear();
            baglan.Open();
            OleDbCommand komut = new OleDbCommand();
            komut.Connection = baglan;
            komut.CommandText=("Select * from Kisibilgi");
            OleDbDataReader oku=komut.ExecuteReader();
            while(oku.Read())
            {
                ListViewItem ekle=new ListViewItem();
                ekle.Text = oku["Id"].ToString();
                ekle.SubItems.Add(oku["Ad"].ToString());
                ekle.SubItems.Add(oku["Soyad"].ToString());
                ekle.SubItems.Add(oku["Şehir"].ToString());
                ekle.SubItems.Add(oku["Meslek"].ToString());
                ekle.SubItems.Add(oku["Yaş"].ToString());
                listView1.Items.Add(ekle);
            }
            baglan.Close();
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            verilerigoster();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            baglan.Open();
            OleDbCommand komut = new OleDbCommand("insert into Kisibilgi (Ad,Soyad,Şehir,Meslek,Yaş) Values('" + textBox1.Text.ToString() + "','" + textBox2.Text.ToString() + "','" + textBox3.Text.ToString() + "','" + textBox4.Text.ToString() + "','" + textBox5.Text.ToString() + "')", baglan);
            komut.ExecuteNonQuery();
            baglan.Close();
            verilerigoster();
        }
    }
}

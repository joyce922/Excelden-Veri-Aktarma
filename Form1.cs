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

namespace WindowsFormsApplication18
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog opfd = new OpenFileDialog();
            if(opfd.ShowDialog ()== DialogResult .OK)
                textselect.Text = opfd.FileName;
        
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string stringconn = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + textselect.Text + "; Extended Properties='Excel 12.0;HDR=NO'";
            OleDbConnection conn = new OleDbConnection(stringconn);
            if (textselect.Text != "")
            {
                OleDbDataAdapter da = new OleDbDataAdapter("Select * from [" + textchoice.Text + "$]", conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            else
                MessageBox.Show("ERROR");
        }
    }
}

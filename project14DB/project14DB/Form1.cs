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

namespace project14DB
{
    public partial class Form1 : Form
    {
        //hammonds string
        //String conn_string = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = D:\\Customer.accdb; Persist Security Info = false";
        //my string
        String conn_string = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = F:\\CSharpDb.accdb; Persist Security Info = false";
        
        String q = "";
        String errormsg = "";
        OleDbConnection conn = null;


        public Form1()
        {
            InitializeComponent();
        }

        private void runQuery()
        {
            q = queryBox1.Text;
            try
            {
                OleDbCommand cmd = new OleDbCommand(q, conn);
                OleDbDataAdapter a = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();
                a.SelectCommand = cmd;
                a.Fill(dt);
                results.DataSource = dt;
                results.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void runQueryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            runQuery();
            this.Cursor = Cursors.Default;

        }

        private void connectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                conn = new OleDbConnection(conn_string);
                conn.Open();
                MessageBox.Show("Connection Open");
                connectToolStripMenuItem.Enabled = false;
                disconnectToolStripMenuItem.Enabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void disconnectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                
                conn.Close();
                MessageBox.Show("Connection Closed");
                connectToolStripMenuItem.Enabled = true;
                disconnectToolStripMenuItem.Enabled = false;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void Form1_FormClosed(object sender, EventArgs e)
        {
            disconnectToolStripMenuItem.PerformClick();
        }

    }


}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;

namespace AutomateMapping
{
    public partial class Login : Form
    {
        private OracleConnection ConnectionProd;
        private bool flagHispeed = false;

        public Login()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {
            btnLogin.Enabled = true;
            flagHispeed = true;
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (String.IsNullOrEmpty(txtUser.Text))
            {
                MessageBox.Show("Please input Username.");
            }
            else if(String.IsNullOrEmpty(txtPassword.Text))
            {
                MessageBox.Show("Please input Password.");
            }
            else
            {
                string user = txtUser.Text;
                string password = txtPassword.Text;

                try
                {
                    ConnectionProd = new OracleConnection();

                    string connString = "Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.19.193.20)(PORT = 1560))" +
                        "(CONNECT_DATA = (SID = TEST03)));User Id=" + user + "; Password=" + password + "; Min Pool Size=10; Max Pool Size =20";

                    //string connString = @"Data Source= (DESCRIPTION =(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 150.4.2.2)(PORT = 1521)) )" +
                    //   "(CONNECT_DATA =(SERVICE_NAME = TAPRD)));User ID=" + user + ";Password=" + password + ";";

                    ConnectionProd.ConnectionString = connString;
                    ConnectionProd.Open();

                    if (ConnectionProd.State == ConnectionState.Open)
                    {
                        btnLogin.Enabled = false;

                        if (flagHispeed == true)
                        {
                            InputHispeed inputHispeed = new InputHispeed(ConnectionProd, user);
                            this.Hide();

                            inputHispeed.Show();
                        }                
                    }
                    else
                    {
                        btnLogin.Enabled = true;
                        DialogResult result = MessageBox.Show("Please try again!!" + "\r\n" + "Cannot connect to database.",
                       "Warning", MessageBoxButtons.OKCancel);
                        if (result == DialogResult.Cancel)
                        {
                            Application.Exit();
                        }
                    }

                }
                catch (Exception ex)
                {
                    DialogResult result = MessageBox.Show("Please try again!! " + "\r\n" + "Connection database failed" + "\r\n" + ex.Message,
                        "Confirmation", MessageBoxButtons.OKCancel);

                    if (result == DialogResult.Cancel)
                    {
                        ConnectionProd.Close();
                        ConnectionProd.Dispose();

                        Application.Exit();
                    }
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                }
            }
        }
    }
}

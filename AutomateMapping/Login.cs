using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OracleClient;
using System.Drawing;

namespace AutomateMapping
{
    public partial class Login : Form
    {
        private OracleConnection ConnectionProd;
        private bool flagHispeed = false;
        int mov, movX, movY;

        public Login()
        {
            InitializeComponent();
        }

        #region "Drop Shadow"
        private const int CS_DropShadow = 0x00020000;

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= CS_DropShadow;
                return cp;
            }
        }
        #endregion

        #region "Event Handler"
        private void Login_Load(object sender, EventArgs e)
        {
            btnLogin.Enabled = true;
            flagHispeed = true;
            txtUser.Select();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            execute();
        }
     
        private void labelClose_Click(object sender, EventArgs e)
        {
            if (ConnectionProd != null)
            {
                if (ConnectionProd.State == ConnectionState.Open)
                {
                    ConnectionProd.Close();
                    ConnectionProd.Dispose();
                }
            }

            Application.Exit();
        }

        private void Login_SizeChanged(object sender, EventArgs e)
        {
            labelClose.Location = new Point(panel5.Width - 30, labelClose.Location.Y);
            pictureBox3.Location = new Point(txtUser.Location.X - 32, txtUser.Location.Y - 4);
            pictureBox4.Location = new Point(txtPassword.Location.X - 32, txtPassword.Location.Y - 4);
            int half = (Size.Width - panel1.Width) / 2;
            label1.Location = new Point((panel1.Width + half) - (label1.Size.Width / 2), pictureBox1.Location.Y + 77);
            label.Location = new Point(label.Location.X, panel2.Location.Y - 39);
        }

        private void txtPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                execute();
            }
        }

        private void panel5_MouseUp(object sender, MouseEventArgs e)
        {
            mov = 0;
        }

        private void panel5_MouseMove(object sender, MouseEventArgs e)
        {
            if (mov == 1)
            {
                this.SetDesktopLocation(MousePosition.X - movX, MousePosition.Y - movY);
            }
        }

        private void Login_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ConnectionProd != null)
            {
                if (ConnectionProd.State == ConnectionState.Open)
                {
                    ConnectionProd.Close();
                    ConnectionProd.Dispose();
                }
            }

            GC.Collect();
        }

        private void panel5_MouseDown(object sender, MouseEventArgs e)
        {
            mov = 1;
            movX = e.X;
            movY = e.Y;
        }
        #endregion

        #region "private method"
        private void execute()
        {
            Cursor.Current = Cursors.WaitCursor;

            if (String.IsNullOrEmpty(txtUser.Text))
            {
                MessageBox.Show("Please input Username.");
            }
            else if (String.IsNullOrEmpty(txtPassword.Text))
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

                    //string connString = @"Data Source= (DESCRIPTION =(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.19.217.162)(PORT = 1559))) " +
                    //                "(CONNECT_DATA =(SERVICE_NAME = CVMDEV)));User Id=" + user + "; Password=" + password + ";";

                    string connString = @"Data Source= (DESCRIPTION =(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 150.4.2.2)(PORT = 1521)) )" +
                       "(CONNECT_DATA =(SERVICE_NAME = TAPRD)));User ID=" + user + ";Password=" + password + ";";

                    ConnectionProd.ConnectionString = connString;
                    ConnectionProd.Open();

                    if (ConnectionProd.State == ConnectionState.Open)
                    {
                        btnLogin.Enabled = false;

                        if (flagHispeed == true)
                        {
                            InputHispeed inputHispeed = new InputHispeed(ConnectionProd, user, "");
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
        #endregion

    }
}

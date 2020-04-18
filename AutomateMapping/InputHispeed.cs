using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace AutomateMapping
{
    public partial class InputHispeed : Form
    {
        private string implementer;
        private string filename;
        private string fileDesc;
        private string folder;
        private OracleConnection ConnectionProd;
        //For move form
        int mov, movX, movY;
        public InputHispeed(OracleConnection con, string user)
        {
            InitializeComponent();
            implementer = user;
            ConnectionProd = con;
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

        private void InputHispeed_Load(object sender, EventArgs e)
        {
            txtImp.Text = implementer;
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            txtInput.Clear();

            filename = "";
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filename = openFileDialog1.FileName;
                txtInput.Text = filename;
                txtInput.SelectionStart = txtInput.Text.Length;
                txtInput.SelectionLength = 0;
            }

            Cursor.Current = Cursors.Default;
        }

        private void btnOpenDesc_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            txtDescFile.Clear();

            fileDesc = "";
            openFileDialog2.Filter = "Excel Files|*.xls;*.xlsx";

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                fileDesc = openFileDialog2.FileName;
                txtDescFile.Text = fileDesc;
                txtDescFile.SelectionStart = txtDescFile.Text.Length;
                txtDescFile.SelectionLength = 0;
            }

            Cursor.Current = Cursors.Default;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            txtOutput.Clear();

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                folder = folderBrowserDialog1.SelectedPath;

                txtOutput.Text = folder;
                txtOutput.SelectionStart = txtOutput.Text.Length;
                txtOutput.SelectionLength = 0;
            }

            Cursor.Current = Cursors.Default;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if(String.IsNullOrEmpty(txtDescFile.Text) ||
                String.Equals(txtDescFile.Text, "X:/xxxx/xxxx/xxxx/file.xlsx"))
            {
                fileDesc = null;
            }

            if (String.IsNullOrEmpty(txtUr.Text))
            {
                MessageBox.Show("Please input Ur.NO#");
            }
            else if (String.IsNullOrEmpty(txtImp.Text))
            {
                MessageBox.Show("Please input Implementer");
            }
            else if (String.IsNullOrEmpty(txtInput.Text) ||
                String.Equals(txtInput.Text, "X:/xxxx/xxxx/xxxx/file.xlsx"))
            {
                MessageBox.Show("Please input requirement file");
            }
            else if (String.IsNullOrEmpty(txtOutput.Text) || 
                String.Equals(txtOutput.Text,"X:/xxxx/xxxx/xxxx/file.xlsx"))
            {
                MessageBox.Show("Please select output path");
            }
            else
            {
                MainHispeed main = new MainHispeed(ConnectionProd, filename, fileDesc, implementer, txtUr.Text, txtOutput.Text);

                this.Hide();
                main.Show();
            }

            Cursor.Current = Cursors.Default;
        }

        private void panel5_MouseDown(object sender, MouseEventArgs e)
        {
            mov = 1;
            movX = e.X;
            movY = e.Y;
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            this.Close();
            Login login = new Login();
            login.Show();
        }

        private void labelClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void panel5_MouseMove(object sender, MouseEventArgs e)
        {
            if (mov == 1)
            {
                this.SetDesktopLocation(MousePosition.X - movX, MousePosition.Y - movY);
            }
        }

        private void panel5_MouseUp(object sender, MouseEventArgs e)
        {
            mov = 0;
        }
    }
}

using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data.OracleClient;
using System.Data;

namespace AutomateMapping
{
    public partial class InputHispeed : Form
    {
        #region "Private Field"
        /// <summary>
        /// implementer
        /// </summary>
        private string implementer;
        /// <summary>
        /// Ur no
        /// </summary>
        private string urNo;
        /// <summary>
        /// Requirement file (xls)
        /// </summary>
        private string filename;
        /// <summary>
        /// File New MKT Code for get PName
        /// </summary>
        private string fileDesc;
        /// <summary>
        /// Output Path
        /// </summary>
        private string folder;
        /// <summary>
        /// Connection of Production
        /// </summary>
        private OracleConnection ConnectionProd;
        /// <summary>
        /// Variable for move form
        /// </summary>
        private int mov, movX, movY, w;
        #endregion

        #region "init"
        public InputHispeed(OracleConnection con, string user, string ur)
        {
            InitializeComponent();
            implementer = user;
            urNo = ur;
            ConnectionProd = con;
        }
        #endregion

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

        #region  "Event Handler"
        private void InputHispeed_Load(object sender, EventArgs e)
        {
            ToolTip toolTip = new ToolTip();
            toolTip.ShowAlways = true;
            toolTip.SetToolTip(btnLogout, "Log out");

            txtImp.Text = implementer;
            txtUr.Text = urNo;

            txtUr.Select();

            w = this.Size.Width;
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
                txtInput.ScrollToCaret();
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
                txtDescFile.ScrollToCaret();
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
                txtOutput.ScrollToCaret();
            }

            Cursor.Current = Cursors.Default;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (String.IsNullOrEmpty(txtDescFile.Text) ||
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
                String.Equals(txtOutput.Text, "X:/xxxx/xxxx/xxxx/file.xlsx"))
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

        private void InputHispeed_SizeChanged(object sender, EventArgs e)
        {
            Rectangle rec = Screen.PrimaryScreen.WorkingArea;

            if (rec.Height < 900)
            {
                if (w == 0)
                {
                    this.Size = new Size(this.Size.Width, btnNext.Location.Y + 98);
                }
                else
                {
                    this.Size = new Size(w, btnNext.Location.Y + 98);
                }

            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void InputHispeed_FormClosing(object sender, FormClosingEventArgs e)
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

        private void btnLogout_Click(object sender, EventArgs e)
        {
            this.Close();
            Login login = new Login();
            login.Show();
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

        #endregion"
    }
}

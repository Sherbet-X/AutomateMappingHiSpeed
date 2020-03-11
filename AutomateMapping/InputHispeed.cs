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
    public partial class InputHispeed : Form
    {
        private string implementer;
        private string filename;
        private string fileDesc;
        private string folder;
        private OracleConnection ConnectionProd;
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

            filename = "";
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filename = openFileDialog1.FileName;
                lableInput.Text = filename;
            }

            Cursor.Current = Cursors.Default;
        }

        private void btnOpenDesc_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            fileDesc = "";
            openFileDialog2.Filter = "Excel Files|*.xls;*.xlsx";

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                fileDesc = openFileDialog2.FileName;
                labelDescFile.Text = fileDesc;
            }

            Cursor.Current = Cursors.Default;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                folder = folderBrowserDialog1.SelectedPath;

                labelOutput.Text = folder;
            }

            Cursor.Current = Cursors.Default;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (String.IsNullOrEmpty(txtUr.Text))
            {
                MessageBox.Show("Please input Ur.NO#");
            }
            else if (String.IsNullOrEmpty(txtImp.Text))
            {
                MessageBox.Show("Please input Implementer");
            }
            else if (String.IsNullOrEmpty(lableInput.Text))
            {
                MessageBox.Show("Please input requirement file");
            }
            else if (String.IsNullOrEmpty(labelOutput.Text))
            {
                MessageBox.Show("Please select output path");
            }
            else
            {
                MainHispeed main = new MainHispeed(ConnectionProd, filename);

                this.Hide();
                main.Show();
            }

            Cursor.Current = Cursors.Default;
        }
    }
}

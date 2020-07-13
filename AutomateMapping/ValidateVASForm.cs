using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OracleClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutomateMapping
{
    public partial class ValidateVASForm : Form
    {
        /// <summary>
        /// Connection of Production
        /// </summary>
        private OracleConnection ConnectionProd;
        /// <summary>
        /// Connection of CVMDEV (Database for validate data)
        /// </summary>
        private OracleConnection ConnectionTemp;
        /// <summary>
        /// Validation class
        /// </summary>
        Validation validation;
        /// <summary>
        /// List of sheetName from file requirement
        /// </summary>
        List<string> sheets = new List<string>();

        string filename, outputPath, implementer, urNo;

        int mov, movX, movY;

        bool flagVASProd, flagVASRule, flagNotAllow;

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void ValidateVASForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ConnectionProd != null)
            {
                if (ConnectionProd.State == ConnectionState.Open)
                {
                    ConnectionProd.Close();
                    ConnectionProd.Dispose();
                }
            }

            if (ConnectionTemp != null)
            {
                if (ConnectionTemp.State == ConnectionState.Open)
                {
                    ConnectionTemp.Close();
                    ConnectionTemp.Dispose();
                }
            }

            GC.Collect();
        }

        private void btnMaximize_Click(object sender, EventArgs e)
        {
            if (this.WindowState != FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
            }
        }

        private void ValidateVASForm_SizeChanged(object sender, EventArgs e)
        {
            //location panel vasCriteria&MKTNotAllow
            panel3.Location = new Point((this.Width / 2) - 49, 13);
            panel4.Location = new Point(this.Width - (panel4.Width + 26), 13);

            //size line
            pictureBox4.Location = new Point(panel2.Location.X + panel2.Width + 4, 31);
            pictureBox4.Width = (panel3.Location.X - pictureBox4.Location.X) - 5;

            pictureBox5.Location = new Point(panel3.Location.X + panel3.Width + 4, 31);
            pictureBox5.Width = (panel4.Location.X - pictureBox5.Location.X) - 5;

        }

        public ValidateVASForm(OracleConnection con, string file, string user, string ur, string fileOut)
        {
            InitializeComponent();

            ConnectionProd = con;
            filename = file;
            outputPath = fileOut;
            implementer = user;
            urNo = ur;
        }

        private void ValidateVASForm_Load(object sender, EventArgs e)
        {
            userControlCriteria1.Hide();
            userControlNotAllow1.Hide();
            userControlVASProduct1.Hide();

            flagVASProd = false;
            flagVASRule = false;
            flagNotAllow = false;

            btnBack.Visible = false;
            
            try
            {
                ConnectionTemp = new OracleConnection();

                string connStringTmp = @"Data Source=(DESCRIPTION =(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.19.217.162)(PORT = 1559))) " +
                                    "(CONNECT_DATA =(SERVICE_NAME = CVMDEV)));User Id= EPCSUPUSR; Password=EPCSUPUSR_55;";

                ConnectionTemp.ConnectionString = connStringTmp;
                ConnectionTemp.Open();

                validation = new Validation(ConnectionProd, ConnectionTemp);

                //Get all sheet name from excel file
                sheets = validation.ToExcelsSheetList(filename);

                if (sheets.Contains("New VAS code (VCare&CCBS)"))
                {
                    //validate new vas code
                    userControlVASProduct1.Show();
                    userControlVASProduct1.BringToFront();
                    flagVASProd = true;

                    //validate vas product
                    ValidateVASProd();
                }
                else if(sheets.Contains("VAS New Sale(SMART UI)"))
                {
                    //validate new sale
                    //picturebox.Image = project.Properties.Resources.imgfromresource
                    pictureBox4.BackColor = Color.Silver;
                    lblProd.ForeColor = Color.Silver;
                }
                else if(sheets.Contains("main offer Not allow"))
                {
                    //validate offer not allow
                }
                else
                {
                    //show message error don't have sheet
                }
            }
            catch(Exception ex)
            {

            }
            
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            btnBack.Visible = true;
            if(flagVASProd == true && flagVASRule == false && flagNotAllow == false)
            {
                //step2
                pictureBox2.Image = AutomateMapping.Properties.Resources.numeric_2_circle_36;
                pictureBox1.Image = AutomateMapping.Properties.Resources.icons8_1st_36;
                pictureBox4.BackColor = Color.Silver;
                lblProd.ForeColor = Color.Silver;

                userControlVASProduct1.Hide();
                userControlNotAllow1.Hide();

                userControlCriteria1.Show();
                userControlCriteria1.BringToFront();

                if (sheets.Contains("VAS New Sale(SMART UI)"))
                {
                    flagVASRule = true;

                    //validate vas rule
                }
                else if(sheets.Contains("main offer Not allow"))
                {
                    //step3
                    flagNotAllow = true;
                    pictureBox3.Image = AutomateMapping.Properties.Resources.numeric_3_circle_36;
                    pictureBox2.Image = AutomateMapping.Properties.Resources.icons8_circled_2_c_36;
                    pictureBox5.BackColor = Color.Silver;
                    lblCri.ForeColor = Color.Silver;

                    //validate mkt not allow
                }
            }
            else if(flagVASProd == true && flagVASRule == true && flagNotAllow == false)
            {
                //step3
                flagNotAllow = true;
                pictureBox3.Image = AutomateMapping.Properties.Resources.numeric_3_circle_36;
                pictureBox2.Image = AutomateMapping.Properties.Resources.icons8_circled_2_c_36;
                pictureBox5.BackColor = Color.Silver;
                lblCri.ForeColor = Color.Silver;

                userControlVASProduct1.Hide();
                userControlCriteria1.Hide();

                userControlNotAllow1.Show();
                userControlNotAllow1.BringToFront();
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            if(flagVASProd == true && flagVASRule == true && flagNotAllow == true)
            {
                //back to step2
                flagNotAllow = false;

                //step2
                pictureBox3.Image = AutomateMapping.Properties.Resources.numeric_3_circle_outline_36;
                pictureBox2.Image = AutomateMapping.Properties.Resources.numeric_2_circle_36;

                pictureBox5.BackColor = Color.WhiteSmoke;
                lblMKT.ForeColor = Color.Black;
                lblCri.ForeColor = Color.Black;

                userControlVASProduct1.Hide();
                userControlNotAllow1.Hide();

                userControlCriteria1.Show();
                userControlCriteria1.BringToFront();
            }
            else if(flagVASProd == true && flagVASRule == true && flagNotAllow == false)
            {
                //back to step1
                flagVASRule = false;
                btnBack.Visible = false;

                //step1
                pictureBox2.Image = AutomateMapping.Properties.Resources.numeric_2_circle_outline_36;
                pictureBox1.Image = AutomateMapping.Properties.Resources.numeric_1_circle_36;

                pictureBox4.BackColor = Color.WhiteSmoke;
                lblProd.ForeColor = Color.Black;

                userControlCriteria1.Hide();
                userControlNotAllow1.Hide();

                userControlVASProduct1.Show();
                userControlVASProduct1.BringToFront();

            }
        }

        private void panel5_MouseMove(object sender, MouseEventArgs e)
        {
            if (mov == 1)
            {
                this.SetDesktopLocation(MousePosition.X - movX, MousePosition.Y - movY);
            }
        }

        private void panel5_MouseDown(object sender, MouseEventArgs e)
        {
            mov = 1;
            movX = e.X;
            movY = e.Y;
        }

        private void panel5_MouseUp(object sender, MouseEventArgs e)
        {
            mov = 0;
        }

        private void ValidateVASProd()
        {

        }
    }
}

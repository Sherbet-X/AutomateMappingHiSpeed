using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OracleClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutomateMapping
{
    public partial class MainVASUpSpeed : Form
    {
        #region "Private Field"
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
        /// <summary>
        /// List Prodtype from DB Master
        /// </summary>
        DataTable tableProdType = new DataTable();
        /// <summary>
        /// List channel from DB Master
        /// </summary>
        List<string[]> lstChannel = new List<string[]>();
        /// <summary>
        /// List province from DB Master
        /// </summary>
        List<string[]> lstProvince = new List<string[]>();
        /// <summary>
        /// List vas type from DB master
        /// </summary>
        List<string[]> lstVasType = new List<string[]>();
        /// <summary>
        /// List vas group from DB master
        /// </summary>
        List<string[]> lstVasGroup = new List<string[]>();
        /// <summary>
        /// List vas channel(product) from DB master
        /// </summary>
        List<string[]> lstVasChannel = new List<string[]>();

        string filename, outputPath, implementer, urNo, validateLog;
        /// <summary>
        /// Variable for keep id, Use to script export criteria file
        /// </summary>
        string lstID, lstCode, lstOffer, lstCodeforOffer, lstUpdateID, existingID, existingCode, existingOffer;
        /// <summary>
        /// For move form
        /// </summary>
        int mov, movX, movY;
        /// <summary>
        /// There is a process about new vas_code(vas_product)
        /// </summary>
        bool flagVasCode = false;
        /// <summary>
        /// There is a process about new vas sale for SmartUI
        /// </summary>
        bool flagVasSale = false;
        /// <summary>
        /// There is a process about Main offer not allow
        /// </summary>
        bool flagNotAllow = false;
        /// <summary>
        /// There is a process about Update date vas(SmartUI)
        /// </summary>
        bool flagUpdate = false;
        /// <summary>
        /// Use to focus row in datagridview
        /// </summary>
        List<int> indexListbox = new List<int>();
        /// <summary>
        /// DataGridView
        /// </summary>
        DataGridView dataGridCode, dataGridSale, dataGridMKT, dataGridUpdate;

        ExportScript exportScript;
        #endregion

        #region "init"
        public MainVASUpSpeed(OracleConnection con, string file, string user, string ur, string fileOut)
        {
            InitializeComponent();

            ConnectionProd = con;
            filename = file;
            outputPath = fileOut;
            implementer = user;
            urNo = ur;
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

        #region "Event Handler"
        private void MainVASUpSpeed_Load(object sender, EventArgs e)
        {
            Application.UseWaitCursor = true;
            Cursor.Current = Cursors.WaitCursor;

            userControlCriteria1.Hide();
            userControlNotAllow1.Hide();
            userControlVASProduct1.Hide();
            userControlUpdate1.Hide();

            flagVasCode = false;
            flagVasSale = false;
            flagNotAllow = false;
            flagUpdate = false;

            btnLog.Visible = false;
            btnBack.Visible = false;
            btnExe.Visible = false;
            btnExe.Enabled = true;

            exportScript = new ExportScript();

            double widthRatio = Screen.PrimaryScreen.Bounds.Width;
            double heightRatio = Screen.PrimaryScreen.Bounds.Height;

            //Different resolutions cause different screen display and widescreen cannot start maximize
            //Set default screen when starting first time
            if (widthRatio >= 1366 && heightRatio >= 768)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
                this.Size = new Size((int)(widthRatio + 250), (int)(heightRatio + 180));
            }

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

                if (sheets.Any(s => s.Equals("New VAS code (VCare&CCBS)", StringComparison.OrdinalIgnoreCase)))
                {
                    toolStripStatusLabel1.Text = "Loading data from sheet[New VAS Code (VCare&CCBS)";

                    flagVasCode = true;

                    userControlVASProduct1.Show();
                    userControlVASProduct1.BringToFront();

                    DgvSettings dgvSettings = new DgvSettings();
                    List<string> lstHeader = new List<string>();

                    //Set header dataGridView
                    lstHeader.Add("Code");
                    lstHeader.Add("Desc");
                    lstHeader.Add("Type");
                    lstHeader.Add("Rule");
                    lstHeader.Add("Price");
                    lstHeader.Add("Channel");
                    lstHeader.Add("Group");
                    lstHeader.Add("StartDate");
                    lstHeader.Add("Desc on bill (TH)");
                    lstHeader.Add("Desc on bill (En)");

                    int i = sheets.FindIndex(x => x.Equals("New VAS code (VCare&CCBS)", StringComparison.OrdinalIgnoreCase));
                    dgvSettings.SetDgv(userControlVASProduct1.GetDataGridView, filename, sheets[i]+"$B3:K", lstHeader);

                    //validate vas product
                    backgroundWorker1.RunWorkerAsync("VasProd");
                }
                else if (sheets.Any(s => s.Equals("VAS New Sale(SMART UI)", StringComparison.OrdinalIgnoreCase)))
                {
                    toolStripStatusLabel1.Text = "Loading data from sheet[VAS New Sale(SMART UI)]";

                    flagVasSale = true;

                    pictureBox1.Image = AutomateMapping.Properties.Resources.icons8_1st_36;//1 grey
                    pictureBox2.Image = AutomateMapping.Properties.Resources.numeric_2_circle_36;//2 black
                    pictureBox5.BackColor = Color.Silver;
                    lblProd.ForeColor = Color.Silver;

                    userControlVASProduct1.Hide();
                    userControlNotAllow1.Hide();

                    userControlCriteria1.Show();
                    userControlCriteria1.BringToFront();

                    DgvSettings dgvSettings = new DgvSettings();
                    List<string> lstHeader = new List<string>();

                    //Set header dataGridView
                    lstHeader.Add("Code");
                    lstHeader.Add("Desc");
                    lstHeader.Add("Speed");
                    lstHeader.Add("Price");
                    lstHeader.Add("Channel");
                    lstHeader.Add("Allow MKT");
                    lstHeader.Add("OrderType");
                    lstHeader.Add("Product");
                    lstHeader.Add("Province");
                    lstHeader.Add("Adv Month");
                    lstHeader.Add("Download From");
                    lstHeader.Add("Download To");
                    lstHeader.Add("Upload From");
                    lstHeader.Add("Upload To");
                    lstHeader.Add("Speed not allow");
                    lstHeader.Add("Price From");
                    lstHeader.Add("Price To");
                    lstHeader.Add("StartDate");
                    lstHeader.Add("EndDate");

                    int i = sheets.FindIndex(x => x.Equals("VAS New Sale(SMART UI)", StringComparison.OrdinalIgnoreCase));
                    dgvSettings.SetDgv(userControlCriteria1.GetDataGridView, filename, sheets[i]+"$B4:T", lstHeader);

                    //validate vas sale
                    backgroundWorker1.RunWorkerAsync("VasSale");
                }
                else if (sheets.Any(s => s.Equals("main offer not allow", StringComparison.OrdinalIgnoreCase)))
                {
                    toolStripStatusLabel1.Text = "Loading data from sheet[Main Offer Not Allow]";

                    flagNotAllow = true;

                    pictureBox1.Image = AutomateMapping.Properties.Resources.icons8_1st_36;//1 grey
                    pictureBox2.Image = AutomateMapping.Properties.Resources.icons8_circled_2_c_36;//2 grey
                    pictureBox3.Image = AutomateMapping.Properties.Resources.numeric_3_circle_36;//3 black

                    pictureBox5.BackColor = Color.Silver;
                    pictureBox6.BackColor = Color.Silver;
                    lblProd.ForeColor = Color.Silver;
                    lblCri.ForeColor = Color.Silver;

                    userControlVASProduct1.Hide();
                    userControlCriteria1.Hide();

                    userControlNotAllow1.Show();
                    userControlNotAllow1.BringToFront();

                    DgvSettings dgvSettings = new DgvSettings();
                    List<string> lstHeader = new List<string>();

                    lstHeader.Add("Code");
                    lstHeader.Add("Not Allow Main Offer");
                    lstHeader.Add("Active Flag");

                    int i = sheets.FindIndex(x => x.Equals("main offer Not allow", StringComparison.OrdinalIgnoreCase));
                    dgvSettings.SetDgv(userControlNotAllow1.GetDataGridView, filename, sheets[i]+"$B2:D", lstHeader);

                    //validate offer not allow
                    backgroundWorker1.RunWorkerAsync("MKTNotAllow");
                }
                else if (sheets.Any(s => s.Equals("Update Date VAS (SMART UI)", StringComparison.OrdinalIgnoreCase)))
                {
                    toolStripStatusLabel1.Text = "Loading data from sheet[Update Date VAS (SMART UI)]";

                    flagUpdate = true;

                    pictureBox1.Image = AutomateMapping.Properties.Resources.icons8_1st_36;//1 grey
                    pictureBox2.Image = AutomateMapping.Properties.Resources.icons8_circled_2_c_36;//2 grey
                    pictureBox3.Image = AutomateMapping.Properties.Resources.icons8_circled_3_c_36;//3 grey
                    pictureBox4.Image = AutomateMapping.Properties.Resources.icons8_circled_4_36bk;//4 black

                    pictureBox5.BackColor = Color.Silver;
                    pictureBox6.BackColor = Color.Silver;
                    pictureBox7.BackColor = Color.Silver;

                    lblCri.ForeColor = Color.Silver;
                    lblProd.ForeColor = Color.Silver;
                    lblMKT.ForeColor = Color.Silver;

                    userControlVASProduct1.Hide();
                    userControlCriteria1.Hide();
                    userControlNotAllow1.Hide();

                    userControlUpdate1.Show();
                    userControlUpdate1.BringToFront();

                    DgvSettings dgvSettings = new DgvSettings();
                    List<string> lstHeader = new List<string>();

                    lstHeader.Add("VAS_ID");
                    lstHeader.Add("StartDate");
                    lstHeader.Add("EndDate");
                    lstHeader.Add("Code");
                    lstHeader.Add("Type");
                    lstHeader.Add("Status");
                    lstHeader.Add("Rule");
                    lstHeader.Add("Channel");

                    int i = sheets.FindIndex(x => x.Equals("Update Date VAS (SMART UI)", StringComparison.OrdinalIgnoreCase));
                    dgvSettings.SetDgv(userControlUpdate1.GetDataGridView, filename, sheets[i]+"$A3:I", lstHeader);

                    //validate update date
                    backgroundWorker1.RunWorkerAsync("Update");
                }
                else
                {
                    //show message error don't have sheet
                    MessageBox.Show("The relevant sheet was not found." + "\r\n" + "The program will close now.", "Automate Mapping Tool");
                    Application.Exit();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Couldn't load the data.Please try again later." + "\r\n" + "Detail : " + ex.Message, "Automate Mapping Tool"
                    , MessageBoxButtons.OK, MessageBoxIcon.Error);

                Application.Exit();
            }   
            finally
            {
                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Do you want to close this application?", "Automate Mapping Tool"
                , MessageBoxButtons.OKCancel,MessageBoxIcon.Question);
            if(dialogResult == DialogResult.OK)
            {
                Application.Exit();
            }
           
        }

        private void MainVASUpSpeed_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (backgroundWorker1.IsBusy)
            {
                backgroundWorker1.CancelAsync();
            }

            if (backgroundWorker2.IsBusy)
            {
                backgroundWorker2.CancelAsync();
            }

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

        private void btnMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnMaximize_Click(object sender, EventArgs e)
        {
            DataGridView dataGrid = new DataGridView();

            if (dataGridCode != null && dataGridCode.Visible)
            {
                dataGrid = dataGridCode;
            }
            else if (dataGridSale != null && dataGridSale.Visible)
            {
                dataGrid = dataGridSale;
            }
            else if (dataGridMKT != null && dataGridMKT.Visible)
            {
                dataGrid = dataGridMKT;
            }
            else if (dataGridUpdate != null && dataGridUpdate.Visible)
            {
                dataGrid = dataGridUpdate;
            }

            if (this.WindowState != FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Maximized;

                dataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;

                if (dataGrid != dataGridMKT || dataGrid != dataGridUpdate)
                {
                    dataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                }
            }
        }

        private void MainVASUpSpeed_SizeChanged(object sender, EventArgs e)
        {
            int box = this.Width / 4;
            panel3.Location = new Point((box - panel3.Width) / 2, 20);
            panel4.Location = new Point(box + (box - panel4.Width) / 2, 20);
            panel5.Location = new Point(box * 2 + ((box - panel5.Width) / 2), 20);
            panel6.Location = new Point(box * 3 + ((box - panel6.Width) / 2), 20);

            //size line
            pictureBox5.Location = new Point(panel3.Location.X + panel3.Width + 4, 41);
            pictureBox5.Width = panel4.Location.X - pictureBox5.Location.X - 5;

            pictureBox6.Location = new Point(panel4.Location.X + panel4.Width + 4, 41);
            pictureBox6.Width = panel5.Location.X - pictureBox6.Location.X - 5;

            pictureBox7.Location = new Point(panel5.Location.X + panel5.Width + 4, 41);
            pictureBox7.Width = panel6.Location.X - pictureBox7.Location.X - 5;

            int w = this.Size.Width;

            if (lblProd.Width < 100)
            {
                lblProd.Size = new Size(lblProd.Width + 40, lblProd.Height + 9);
                lblProd.Font = new Font("Roboto", 7);

                lblCri.Size = new Size(lblCri.Width + 40, lblCri.Height + 9);
                lblCri.Font = new Font("Roboto", 7);

                lblMKT.Size = new Size(lblMKT.Width + 42, lblMKT.Height + 9);
                lblMKT.Font = new Font("Roboto", 7);

                lblUpdate.Size = new Size(lblUpdate.Width + 42, lblUpdate.Height + 9);
                lblUpdate.Font = new Font("Roboto", 7);
            }

            int section = (statusStrip1.Location.Y - 300) / 5;

            userControlVASProduct1.Size = new Size(w, (section * 2) + (section / 2));
            userControlCriteria1.Size = new Size(w, (section * 2) + (section / 2));
            userControlNotAllow1.Size = new Size(w, (section * 2) + (section / 2));
            userControlUpdate1.Size = new Size(w, (section * 2) + (section / 2));

            listBox1.Location = new Point(0, userControlVASProduct1.Location.Y + userControlVASProduct1.Height + (section / 2));
            listBox1.Size = new Size(w, section + (section / 2));

            btnValidate.Location = new Point(w - btnValidate.Width, listBox1.Location.Y - btnValidate.Height);

            if (btnNext.Size.Height < 62)
            {
                btnNext.Size = new Size(164, 62);
                btnBack.Size = new Size(164, 62);
                btnExe.Size = new Size(164, 62);
                btnLog.Size = new Size(164, 62);

                btnNext.Font = new Font("Roboto", 11);
                btnBack.Font = new Font("Roboto", 11);
                btnExe.Font = new Font("Roboto", 11);
                btnLog.Font = new Font("Roboto", 11);
            }

            int section2 = (statusStrip1.Location.Y - (listBox1.Location.Y + listBox1.Height)) / 4;
            btnNext.Location = new Point(w - btnNext.Width - 40, listBox1.Location.Y + listBox1.Height + section2);
            btnExe.Location = new Point(w - btnExe.Width - 40, listBox1.Location.Y + listBox1.Height + section2);
            btnBack.Location = new Point(btnNext.Location.X - 30 - btnBack.Width, btnNext.Location.Y);
            btnLog.Location = new Point(w - btnLog.Width - 40, listBox1.Location.Y + listBox1.Height + section2);
        }

        private void listBox1_Click(object sender, EventArgs e)
        {
            DataGridView dataGrid = new DataGridView();
            if(dataGridCode != null && dataGridCode.Visible)
            {
                dataGrid = dataGridCode;
            }
            else if(dataGridSale != null && dataGridSale.Visible)
            {
                dataGrid = dataGridSale;
            }
            else if(dataGridMKT != null && dataGridMKT.Visible)
            {
                dataGrid = dataGridMKT;
            }
            else if(dataGridUpdate != null && dataGridUpdate.Visible)
            {
                dataGrid = dataGridUpdate;
            }

            dataGrid.ClearSelection();
            if (listBox1.SelectedItem != null)
            {
                int selected = listBox1.SelectedIndex;
                dataGrid.Rows[indexListbox[selected]].Selected = true;
                dataGrid.FirstDisplayedScrollingRowIndex = indexListbox[selected];
                dataGrid.Focus();
            }
        }

        private void btnValidate_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy)
            {
                backgroundWorker1.CancelAsync();
            }

            if (userControlVASProduct1.Visible)
            {
                backgroundWorker1.RunWorkerAsync("VasProd");
            }
            else if(userControlCriteria1.Visible)
            {
                backgroundWorker1.RunWorkerAsync("VasSale");
            }
            else if(userControlNotAllow1.Visible)
            {
                backgroundWorker1.RunWorkerAsync("MKTNotAllow");
            }
            else if(userControlUpdate1.Visible)
            {
                backgroundWorker1.RunWorkerAsync("Update");
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy)
            {
                backgroundWorker1.CancelAsync();
            }

            if (listBox1.Items.Count <= 0)
            {
                Application.UseWaitCursor = true;
                Cursor.Current = Cursors.WaitCursor;

                if (userControlVASProduct1.Visible)
                {
                    //step2                   
                    if (sheets.Any(s => s.Equals("VAS New Sale(SMART UI)", StringComparison.OrdinalIgnoreCase)))
                    {
                        toolStripStatusLabel1.Text = "Loading data from sheet[VAS New Sale(SMART UI)]";

                        flagVasSale = true;

                        pictureBox1.Image = AutomateMapping.Properties.Resources.icons8_1st_36;//1 grey
                        pictureBox2.Image = AutomateMapping.Properties.Resources.numeric_2_circle_36;//2 black

                        pictureBox5.BackColor = Color.Silver;//line1 silver
                        lblProd.ForeColor = Color.Silver;

                        userControlVASProduct1.Hide();
                        userControlNotAllow1.Hide();

                        userControlCriteria1.Show();
                        userControlCriteria1.BringToFront();

                        DgvSettings dgvSettings = new DgvSettings();
                        List<string> lstHeader = new List<string>();

                        lstHeader.Add("Code");
                        lstHeader.Add("Desc");
                        lstHeader.Add("Speed");
                        lstHeader.Add("Price");
                        lstHeader.Add("Channel");
                        lstHeader.Add("Allow MKT");
                        lstHeader.Add("OrderType");
                        lstHeader.Add("Product");
                        lstHeader.Add("Province");
                        lstHeader.Add("Adv Month");
                        lstHeader.Add("Download From");
                        lstHeader.Add("Download To");
                        lstHeader.Add("Upload From");
                        lstHeader.Add("Upload To");
                        lstHeader.Add("Speed not allow");
                        lstHeader.Add("Price From");
                        lstHeader.Add("Price To");
                        lstHeader.Add("StartDate");
                        lstHeader.Add("EndDate");

                        int i = sheets.FindIndex(x => x.Equals("VAS New Sale(SMART UI)", StringComparison.OrdinalIgnoreCase));
                        dgvSettings.SetDgv(userControlCriteria1.GetDataGridView, filename, sheets[i]+"$B4:T", lstHeader);

                        //validate vas new sale
                        backgroundWorker1.RunWorkerAsync("VasSale");
                    }
                    else if (sheets.Any(s => s.Equals("main offer not allow", StringComparison.OrdinalIgnoreCase)))
                    {
                        toolStripStatusLabel1.Text = "Loading data from sheet[Main Offer Not Allow]";

                        flagNotAllow = true;

                        //step3
                        pictureBox1.Image = AutomateMapping.Properties.Resources.icons8_1st_36;//1 grey
                        pictureBox2.Image = AutomateMapping.Properties.Resources.icons8_circled_2_c_36;//2 grey
                        pictureBox3.Image = AutomateMapping.Properties.Resources.numeric_3_circle_36;//3 black

                        pictureBox5.BackColor = Color.Silver;
                        pictureBox6.BackColor = Color.Silver;
                        lblProd.ForeColor = Color.Silver;
                        lblCri.ForeColor = Color.Silver;

                        userControlVASProduct1.Hide();
                        userControlCriteria1.Hide();

                        userControlNotAllow1.Show();
                        userControlNotAllow1.BringToFront();

                        DgvSettings dgvSettings = new DgvSettings();
                        List<string> lstHeader = new List<string>();

                        lstHeader.Add("Code");
                        lstHeader.Add("Not Allow Main Offer");
                        lstHeader.Add("Active Flag");

                        int i = sheets.FindIndex(x => x.Equals("main offer Not allow", StringComparison.OrdinalIgnoreCase));
                        dgvSettings.SetDgv(userControlNotAllow1.GetDataGridView, filename, sheets[i]+"$B2:D", lstHeader);

                        //validate mkt not allow
                        backgroundWorker1.RunWorkerAsync("MKTNotAllow");

                    }
                }
                else if (userControlCriteria1.Visible)
                {
                    if (sheets.Any(s => s.Equals("main offer not allow", StringComparison.OrdinalIgnoreCase)))
                    {
                        //step3
                        toolStripStatusLabel1.Text = "Loading data from sheet[Main Offer Not Allow]";

                        flagNotAllow = true;

                        pictureBox1.Image = AutomateMapping.Properties.Resources.icons8_1st_36;//1 grey
                        pictureBox2.Image = AutomateMapping.Properties.Resources.icons8_circled_2_c_36;//2 grey
                        pictureBox3.Image = AutomateMapping.Properties.Resources.numeric_3_circle_36;//3 black

                        pictureBox5.BackColor = Color.Silver;
                        pictureBox6.BackColor = Color.Silver;
                        lblProd.ForeColor = Color.Silver;
                        lblCri.ForeColor = Color.Silver;

                        userControlVASProduct1.Hide();
                        userControlCriteria1.Hide();

                        userControlNotAllow1.Show();
                        userControlNotAllow1.BringToFront();

                        DgvSettings dgvSettings = new DgvSettings();
                        List<string> lstHeader = new List<string>();

                        lstHeader.Add("Code");
                        lstHeader.Add("Not Allow Main Offer");
                        lstHeader.Add("Active Flag");

                        int i = sheets.FindIndex(x => x.Equals("main offer Not allow", StringComparison.OrdinalIgnoreCase));
                        dgvSettings.SetDgv(userControlNotAllow1.GetDataGridView, filename, sheets[i]+"$B2:D", lstHeader);

                        //validate mkt not allow
                        backgroundWorker1.RunWorkerAsync("MKTNotAllow");
                    }
                }
            }
            else
            {
                MessageBox.Show("There is incorrect data." + "\r\n" + "Please correct it.", "Automate Mapping Tool"
                    , MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            if (userControlNotAllow1.Visible)
            {
                //back to step vas new sale   
                if (sheets.Any(s => s.Equals("VAS New Sale(SMART UI)", StringComparison.OrdinalIgnoreCase)))
                {
                    flagNotAllow = false;

                    btnNext.Visible = true;
                    btnExe.Visible = false;
                    if (sheets.Any(s => s.Equals("New VAS code (VCare&CCBS)", StringComparison.OrdinalIgnoreCase)))
                    {
                        btnBack.Visible = true;
                    }
                    else
                    {
                        btnBack.Visible = false;
                    }

                    //step2
                    pictureBox3.Image = AutomateMapping.Properties.Resources.numeric_3_circle_outline_36;//3 white
                    pictureBox2.Image = AutomateMapping.Properties.Resources.numeric_2_circle_36;//2 black

                    pictureBox6.BackColor = Color.WhiteSmoke;
                    lblMKT.ForeColor = Color.Black;
                    lblCri.ForeColor = Color.Black;

                    userControlVASProduct1.Hide();
                    userControlNotAllow1.Hide();

                    userControlCriteria1.Show();
                    userControlCriteria1.BringToFront();
                }
                else if (sheets.Any(s => s.Equals("New VAS code (VCare&CCBS)", StringComparison.OrdinalIgnoreCase)))
                {
                    //back to step1
                    flagVasSale = false;

                    btnNext.Visible = true;
                    btnExe.Visible = false;
                    btnBack.Visible = false;

                    //step1
                    pictureBox3.Image = AutomateMapping.Properties.Resources.numeric_3_circle_outline_36;//3 white
                    pictureBox2.Image = AutomateMapping.Properties.Resources.numeric_2_circle_outline_36;//2 white
                    pictureBox1.Image = AutomateMapping.Properties.Resources.numeric_1_circle_36; //1 black

                    pictureBox5.BackColor = Color.WhiteSmoke;
                    pictureBox6.BackColor = Color.WhiteSmoke;

                    lblProd.ForeColor = Color.Black;
                    lblCri.ForeColor = Color.Black;
                    lblMKT.ForeColor = Color.Black;

                    userControlCriteria1.Hide();
                    userControlNotAllow1.Hide();

                    userControlVASProduct1.Show();
                    userControlVASProduct1.BringToFront();
                }
            }
            else if (userControlCriteria1.Visible)
            {
                if (sheets.Any(s => s.Equals("New VAS code (VCare&CCBS)", StringComparison.OrdinalIgnoreCase)))
                {
                    //back to step1
                    flagVasSale = false;

                    btnBack.Visible = false;
                    btnNext.Visible = true;
                    btnExe.Visible = false;

                    //step1
                    pictureBox2.Image = AutomateMapping.Properties.Resources.numeric_2_circle_outline_36;//2 white
                    pictureBox1.Image = AutomateMapping.Properties.Resources.numeric_1_circle_36;//1 black

                    pictureBox5.BackColor = Color.WhiteSmoke;
                    pictureBox6.BackColor = Color.WhiteSmoke;

                    lblCri.ForeColor = Color.Black;
                    lblProd.ForeColor = Color.Black;

                    userControlCriteria1.Hide();
                    userControlNotAllow1.Hide();

                    userControlVASProduct1.Show();
                    userControlVASProduct1.BringToFront();
                }
            }
        }

        private void btnExe_Click(object sender, EventArgs e)
        {
            Application.UseWaitCursor = true;
            Cursor.Current = Cursors.WaitCursor;

            btnExe.Enabled = false;

            backgroundWorker1.CancelAsync();

            backgroundWorker2.RunWorkerAsync();
        }

        private void btnLog_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(validateLog))
            {
                MessageBox.Show("The verification process is complete. No errors occurred during process.", "Automate Mapping Tool");
            }
            else
            {
                string strFilePath = outputPath + "\\LOG_VALIDATE_" + urNo.ToUpper() + "_" + DateTime.Now.ToString("ddMMyyyy") + ".txt";
                using (StreamWriter writer = new StreamWriter(strFilePath, true))
                {
                    writer.Write(validateLog);
                }

                MessageBox.Show("Log file has been written successfully." + "\r\n" + "Program will be closing", "Automate Mapping Tool");

                Application.Exit();
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

        private void panel5_MouseMove(object sender, MouseEventArgs e)
        {
            if (mov == 1)
            {
                this.SetDesktopLocation(MousePosition.X - movX, MousePosition.Y - movY);
            }
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            this.Close();
            Login login = new Login();
            login.Show();
        }
        #endregion

        #region "BackgroundWorker"
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string process = e.Argument.ToString();

            if (process == "VasProd")
            {
                ValidateVasCode();

                btnBack.Visible = false;

                if (sheets.Contains("VAS New Sale(SMART UI)", StringComparer.OrdinalIgnoreCase) || 
                    sheets.Contains("main offer Not allow", StringComparer.OrdinalIgnoreCase))
                {
                    btnNext.Visible = true;
                }
                else
                {
                    btnNext.Visible = false;
                    btnExe.Visible = true;
                }
            }
            else if (process == "VasSale")
            {
                ValidateVasSale();

                if (sheets.Contains("New VAS code (VCare&CCBS)", StringComparer.OrdinalIgnoreCase))
                {
                    btnBack.Visible = true;
                }
                else
                {
                    btnBack.Visible = false;
                }

                if (sheets.Contains("main offer Not allow", StringComparer.OrdinalIgnoreCase))
                {
                    btnNext.Visible = true;
                }
                else
                {
                    btnNext.Visible = false;
                    btnExe.Visible = true;
                }
            }
            else if (process == "MKTNotAllow")
            {
                ValidateMKTNotAllow();

                btnNext.Visible = false;
                btnExe.Visible = true;

                if (sheets.Contains("New VAS code (VCare&CCBS)", StringComparer.OrdinalIgnoreCase) 
                    || sheets.Contains("VAS New Sale(SMART UI)", StringComparer.OrdinalIgnoreCase))
                {
                    btnBack.Visible = true;
                }
            }
            else if (process == "Update")
            {
                btnBack.Visible = false;

                ValidateUpdateVASSmartUI();

                if (sheets.Contains("New VAS code (VCare&CCBS)", StringComparer.OrdinalIgnoreCase) == false 
                    && sheets.Contains("VAS New Sale(SMART UI)", StringComparer.OrdinalIgnoreCase) == false
                    && sheets.Contains("main offer Not allow", StringComparer.OrdinalIgnoreCase) == false)
                {
                    btnNext.Visible = false;
                    btnExe.Visible = true;
                }
            }

        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            toolStripProgressBar1.Value = e.ProgressPercentage;

            //if (e.ProgressPercentage == 0 || e.ProgressPercentage == 100)
            //{
            //    Application.UseWaitCursor = false;
            //    Cursor.Current = Cursors.Default;
            //    return;
            //}
            //else
            //{
            //    this.Cursor = Cursors.WaitCursor;
            //    return;
            //}
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            toolStripProgressBar1.Value = 0;

            if (listBox1.Items.Count > 0)
            {
                btnLog.Visible = true;

                btnNext.Visible = false;
                btnExe.Visible = false;
                btnBack.Visible = false;
            }
            else
            {
                btnLog.Visible = false;
                btnExe.Visible = true;
            }

            Application.UseWaitCursor = false;
            Cursor.Current = Cursors.Default;
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            if (flagVasCode)
            {
                InsertNewVASCode();
                flagVasCode = false;
            }

            if (flagVasSale)
            {
                InsertVASNewSale();
                flagVasSale = false;
            }

            if (flagNotAllow)
            {
                InsertMKTNotAllow();
                flagNotAllow = false;
            }

            if (flagUpdate)
            {
                UpdateDateVASSmartUI();
                flagUpdate = false;

                ExportCriteria();
            }
            else
            {
                if (sheets.Contains("Update Date VAS (SMART UI)", StringComparer.OrdinalIgnoreCase))
                {
                    DialogResult dialogResult = MessageBox.Show("Do you want go to the process update date.", "Automate Mapping Tool"
                        , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dialogResult == DialogResult.Yes)
                    {
                        if (backgroundWorker1.IsBusy == true)
                        {
                            backgroundWorker1.CancelAsync();
                        }

                        if (backgroundWorker2.IsBusy == true)
                        {
                            backgroundWorker2.CancelAsync();
                        }

                        Application.UseWaitCursor = true;
                        Cursor.Current = Cursors.WaitCursor;

                        toolStripStatusLabel1.Text = "Loading data for update date...";

                        flagUpdate = true;

                        btnBack.Visible = false;

                        pictureBox1.Image = AutomateMapping.Properties.Resources.icons8_1st_36;//1 grey
                        pictureBox2.Image = AutomateMapping.Properties.Resources.icons8_circled_2_c_36;//2 grey
                        pictureBox3.Image = AutomateMapping.Properties.Resources.icons8_circled_3_c_36;//3 grey
                        pictureBox4.Image = AutomateMapping.Properties.Resources.icons8_circled_4_36bk;//4 black

                        pictureBox5.BackColor = Color.Silver;
                        pictureBox6.BackColor = Color.Silver;
                        pictureBox7.BackColor = Color.Silver;

                        lblCri.ForeColor = Color.Silver;
                        lblProd.ForeColor = Color.Silver;
                        lblMKT.ForeColor = Color.Silver;

                        userControlVASProduct1.Hide();
                        userControlCriteria1.Hide();
                        userControlNotAllow1.Hide();

                        DgvSettings dgvSettings = new DgvSettings();
                        List<string> lstHeader = new List<string>();

                        lstHeader.Add("VAS_ID");
                        lstHeader.Add("StartDate");
                        lstHeader.Add("EndDate");
                        lstHeader.Add("Code");
                        lstHeader.Add("Type");
                        lstHeader.Add("Status");
                        lstHeader.Add("Rule");
                        lstHeader.Add("Channel");

                        int i = sheets.FindIndex(x => x.Equals("Update Date VAS (SMART UI)", StringComparison.OrdinalIgnoreCase));
                        dgvSettings.SetDgv(userControlUpdate1.GetDataGridView, filename, sheets[i] + "$A3:I", lstHeader);

                        userControlUpdate1.Show();
                        userControlUpdate1.BringToFront();

                        backgroundWorker1.RunWorkerAsync("Update");
                    }
                    else
                    {
                        ExportCriteria();
                    }
                }
                else
                {
                    ExportCriteria();
                }
            }

            if (String.IsNullOrEmpty(existingID) == false)
            {
                lstID = null;
                lstCode = null;
                lstOffer = null;
                lstCodeforOffer = null;

                lstUpdateID = existingID;
            }

            if (String.IsNullOrEmpty(existingCode) == false)
            {
                lstID = null;
                lstCode = null;
                lstOffer = null;
                lstCodeforOffer = null;
            }

        }

        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            toolStripProgressBar1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Application.UseWaitCursor = false;
            Cursor.Current = Cursors.Default;
        }
        #endregion

        private void ValidateVasCode()
        {
            try
            {
                Application.UseWaitCursor = true;
                Cursor.Current = Cursors.WaitCursor;

                toolStripStatusLabel1.Text = "Checking New VASCode...";

                dataGridCode = userControlVASProduct1.GetDataGridView;

                backgroundWorker1.ReportProgress(10);

                InitialValue(dataGridCode);

                backgroundWorker1.ReportProgress(15);

                //Get vas type from DB
                if (lstVasType.Count <= 0)
                {
                    lstVasType = validation.GetVasType;
                }

                //Get vas group from DB
                if (lstVasGroup.Count <= 0)
                {
                    lstVasGroup = validation.GetVasGroup;
                }

                //Get vas channel from DB
                if (lstVasChannel.Count <= 0)
                {
                    lstVasChannel = validation.GetVasChannel;
                }

                backgroundWorker1.ReportProgress(30);

                if (lstVasType.Count <= 0 || lstVasGroup.Count <= 0 || lstVasChannel.Count <= 0)
                {
                    MessageBox.Show("An error occurred while retrieving data from the database.Please try again!!", "Automate Mapping Tool");
                    backgroundWorker1.ReportProgress(0);
                }
                else
                {
                    for (int i = 0; i < dataGridCode.RowCount; i++)
                    {
                        string code = dataGridCode.Rows[i].Cells[0].Value.ToString().ToUpper().Trim();
                        string desc = dataGridCode.Rows[i].Cells[1].Value.ToString().Trim();
                        string type = dataGridCode.Rows[i].Cells[2].Value.ToString().ToUpper().Trim();
                        string rule = dataGridCode.Rows[i].Cells[3].Value.ToString().Trim();
                        string price = dataGridCode.Rows[i].Cells[4].Value.ToString().Trim();
                        string channel = dataGridCode.Rows[i].Cells[5].Value.ToString().Trim();
                        string group = dataGridCode.Rows[i].Cells[6].Value.ToString().ToUpper().Trim();
                        string start = dataGridCode.Rows[i].Cells[7].Value.ToString().Trim();

                        string txt = "SELECT * FROM VAS_PRODUCT WHERE VAS_CODE = '" + code + "' AND VAS_CHANNEL = '" + channel +
                            "' AND VAS_TYPE = '" + type + "' AND PARENT_VAS_CODE = '" + group + "'";

                        OracleCommand command = new OracleCommand(txt, ConnectionProd);
                        OracleDataReader reader = command.ExecuteReader();
                        if (reader.HasRows == false)
                        {
                            if (String.IsNullOrEmpty(code))
                            {
                                //write log
                                string msg = "VASCode is null or empty";
                                listBox1.Items.Add(msg);
                                indexListbox.Add(i);
                                hilightRow("prod", "code", i, dataGridCode);

                                validateLog += "[row : " + i + 4 + "]     " + msg + "\r\n";
                            }
                            else
                            {
                                if (code.Length != 15)
                                {
                                    //write log
                                    string msg = "VASCode fotmat is not supported";
                                    listBox1.Items.Add(msg);
                                    indexListbox.Add(i);
                                    hilightRow("prod", "code", i, dataGridCode);

                                    validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                        + msg + "\r\n";
                                }
                            }

                            if (String.IsNullOrEmpty(desc))
                            {
                                //write log
                                string msg = "VAS Description is null or empty";
                                listBox1.Items.Add(msg);
                                indexListbox.Add(i);
                                hilightRow("prod", "desc", i, dataGridCode);

                                validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                        + msg + "\r\n";
                            }

                            if (String.IsNullOrEmpty(type))
                            {
                                //write log
                                string msg = "VAS_Type is null or empty";
                                listBox1.Items.Add(msg);
                                indexListbox.Add(i);
                                hilightRow("prod", "type", i, dataGridCode);

                                validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                        + msg + "\r\n";
                            }
                            else
                            {
                                string msgType = validation.CheckType(lstVasType, type);
                                if (msgType != "Success")
                                {
                                    listBox1.Items.Add(msgType);
                                    indexListbox.Add(i);
                                    hilightRow("prod", "type", i, dataGridCode);

                                    validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                            + msgType + "\r\n";
                                }
                            }

                            if (String.IsNullOrEmpty(channel))
                            {
                                //write log
                                string msg = "VAS_Channel is null or empty";
                                listBox1.Items.Add(msg);
                                indexListbox.Add(i);
                                hilightRow("prod", "channel", i, dataGridCode);

                                validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                        + msg + "\r\n";
                            }
                            else
                            {
                                string msgCh = validation.CheckVasChannel(lstVasChannel, channel);
                                if (msgCh != "Success")
                                {
                                    listBox1.Items.Add(msgCh);
                                    indexListbox.Add(i);
                                    hilightRow("prod", "channel", i, dataGridCode);

                                    validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                            + msgCh + "\r\n";
                                }
                            }

                            if (String.IsNullOrEmpty(rule))
                            {
                                //write log
                                string msg = "VAS_RULE is null or empty";
                                listBox1.Items.Add(msg);
                                indexListbox.Add(i);
                                hilightRow("prod", "rule", i, dataGridCode);

                                validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                        + msg + "\r\n";
                            }

                            if (String.IsNullOrEmpty(price))
                            {
                                //write log
                                string msg = "Price is null or empty";
                                listBox1.Items.Add(msg);
                                indexListbox.Add(i);
                                hilightRow("prod", "price", i, dataGridCode);
                            }
                            else
                            {
                                if (double.TryParse(price, out _) == false)
                                {
                                    //write log not numeric
                                    string msg = "Price is not a numeric";
                                    listBox1.Items.Add(msg);
                                    indexListbox.Add(i);
                                    hilightRow("prod", "price", i, dataGridCode);

                                    validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                        + msg + "\r\n";
                                }
                                else
                                {
                                    if (Convert.ToInt32(price) < 0)
                                    {
                                        string msg = "Price is a negative number";
                                        listBox1.Items.Add(msg);
                                        indexListbox.Add(i);
                                        hilightRow("prod", "price", i, dataGridCode);

                                        validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                            + msg + "\r\n";
                                    }
                                }
                            }

                            if (String.IsNullOrEmpty(group))
                            {
                                string msg = "VAS_GROUP is null or empty";
                                listBox1.Items.Add(msg);
                                indexListbox.Add(i);
                                hilightRow("prod", "group", i, dataGridCode);

                                validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                        + msg + "\r\n";
                            }
                            else
                            {
                                string msgGroup = validation.CheckGroup(lstVasGroup, group);
                                if (msgGroup != "Success")
                                {
                                    listBox1.Items.Add(msgGroup);
                                    indexListbox.Add(i);
                                    hilightRow("prod", "group", i, dataGridCode);

                                    validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                            + msgGroup + "\r\n";
                                }
                            }

                            start = validation.ChangeFormatDate(start);
                            if(String.IsNullOrEmpty(start))
                            {
                                string msg = "StartDate is null or empty";
                                listBox1.Items.Add(msg);
                                indexListbox.Add(i);
                                hilightRow("prod", "start", i, dataGridCode);

                                validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                        + msg + "\r\n";
                            }
                            else if (start == "Invalid")
                            {
                                string msg = "StartDate fotmat is not supported";
                                listBox1.Items.Add(msg);
                                indexListbox.Add(i);
                                hilightRow("prod", "start", i, dataGridCode);

                                validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                        + msg + "\r\n";
                            }
                            else
                            {
                                if (Convert.ToDateTime(start) < DateTime.Now)
                                {
                                    //write log
                                    string msg = "StartDate < Sysdate";
                                    listBox1.Items.Add(msg);
                                    indexListbox.Add(i);
                                    hilightRow("prod", "start", i, dataGridCode);

                                    validateLog += "[(row:" + i + 4 + ") VASCode:" + code + ", Channel:" + channel + "]     "
                                            + msg + "\r\n";
                                }
                            }
                        }

                        for (int j = i + 1; j < dataGridCode.RowCount; j++)
                        {
                            string nextCode = dataGridCode.Rows[j].Cells[0].Value.ToString().ToUpper().Trim();
                            string nextDesc = dataGridCode.Rows[j].Cells[1].Value.ToString().Trim();
                            string nextType = dataGridCode.Rows[j].Cells[2].Value.ToString().ToUpper().Trim();
                            string nextRule = dataGridCode.Rows[j].Cells[3].Value.ToString().Trim();
                            string nextPrice = dataGridCode.Rows[j].Cells[4].Value.ToString().Trim();
                            string nextChannel = dataGridCode.Rows[j].Cells[5].Value.ToString().Trim();
                            string nextGroup = dataGridCode.Rows[i].Cells[6].Value.ToString().Trim();
                            string nextStart = dataGridCode.Rows[j].Cells[7].Value.ToString().Trim();

                            if (code == nextCode && desc == nextDesc && type == nextType && rule == nextRule && price == nextPrice
                                && channel == nextChannel && start == nextStart && group == nextGroup)
                            {
                                //write log data dup
                                listBox1.Items.Add("Duplicate record: " + i + " and record: " + j);
                                indexListbox.Add(i);
                                dataGridCode.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                                dataGridCode.Rows[j].DefaultCellStyle.BackColor = Color.Yellow;

                                validateLog += "Duplicate record: " + i + 3 + " and record: " + j + 3 + "\r\n";
                            }
                        }

                        backgroundWorker1.ReportProgress(30 + ((i + 1) * 70 / dataGridCode.RowCount));

                    }
                }

                toolStripStatusLabel1.Text = "Validation Completed!!";
            }
            catch (Exception e)
            {
                backgroundWorker1.CancelAsync();
                toolStripStatusLabel1.Text = "Failed to validate new vas code";
                MessageBox.Show("There was a problem during the validation new vas code process.Please try again later." + "\r\n" +
                    "Detail : " + e.Message, "Automate Mapping Tool", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;
            }

        }
   
        private void ValidateVasSale()
        {
            try
            {
                Application.UseWaitCursor = true;
                Cursor.Current = Cursors.WaitCursor;

                toolStripStatusLabel1.Text = "Checking New VAS Sale for SmartUI...";

                dataGridSale = userControlCriteria1.GetDataGridView;

                backgroundWorker1.ReportProgress(10);

                InitialValue(dataGridSale);

                backgroundWorker1.ReportProgress(15);

                //Get prodtype from DB
                if (tableProdType.Rows.Count <= 0)
                {
                    tableProdType = validation.GetProdType();
                }

                //Get Channel from DB
                if (lstChannel.Count <= 0 || lstChannel is null)
                {
                    lstChannel = validation.GetChannelFromDB;
                }

                //Get province from DB
                if (lstProvince.Count <= 0 || lstProvince is null)
                {
                    lstProvince = validation.GetProvFromDB;
                }

                backgroundWorker1.ReportProgress(30);

                if (tableProdType.Rows.Count <= 0 || lstChannel.Count <= 0 || lstProvince.Count <= 0)
                {
                    MessageBox.Show("An error occurred while retrieving data from the database.Please try again!!", "Automate Mapping Tool");
                    backgroundWorker1.ReportProgress(0);
                }
                else
                {
                    for (int i = 0; i < dataGridSale.RowCount; i++)
                    {
                        string code = dataGridSale.Rows[i].Cells[0].Value.ToString().Trim();
                        string speed = dataGridSale.Rows[i].Cells[2].Value.ToString().Trim();
                        string price = dataGridSale.Rows[i].Cells[3].Value.ToString().Trim();
                        string channel = dataGridSale.Rows[i].Cells[4].Value.ToString();
                        string offer = dataGridSale.Rows[i].Cells[5].Value.ToString().Trim();
                        string order = dataGridSale.Rows[i].Cells[6].Value.ToString();
                        string product = dataGridSale.Rows[i].Cells[7].Value.ToString();
                        string province = dataGridSale.Rows[i].Cells[8].Value.ToString().Trim();
                        string advMonth = dataGridSale.Rows[i].Cells[9].Value.ToString();
                        string downloadF = dataGridSale.Rows[i].Cells[10].Value.ToString();
                        string downloadT = dataGridSale.Rows[i].Cells[11].Value.ToString();
                        string uploadF = dataGridSale.Rows[i].Cells[12].Value.ToString();
                        string uploadT = dataGridSale.Rows[i].Cells[13].Value.ToString();
                        string priceF = dataGridSale.Rows[i].Cells[15].Value.ToString();
                        string priceT = dataGridSale.Rows[i].Cells[16].Value.ToString();
                        string start = dataGridSale.Rows[i].Cells[17].Value.ToString();
                        string end = dataGridSale.Rows[i].Cells[18].Value.ToString();

                        //Check VasCode
                        if (String.IsNullOrEmpty(code))
                        {
                            //write log
                            string msg = "VASCode is null or empty";
                            listBox1.Items.Add(msg);
                            indexListbox.Add(i);
                            hilightRow("rule", "code", i, dataGridSale);

                            validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msg + "\r\n";
                        }
                        else
                        {
                            if (code.Length != 15)
                            {
                                //write log
                                string msg = "VASCode fotmat is not supported";
                                listBox1.Items.Add(msg);
                                indexListbox.Add(i);
                                hilightRow("rule", "code", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msg + "\r\n";
                            }
                        }

                        //Check channel
                        channel = Regex.Replace(channel, "ALL", "ALL", RegexOptions.IgnoreCase);
                        dataGridSale.Rows[i].Cells[4].Value = channel;

                        string msgChannel = validation.CheckChannel(lstChannel, channel, end);
                        if (msgChannel != "Success")
                        {
                            listBox1.Items.Add(msgChannel);
                            indexListbox.Add(i);
                            hilightRow("rule", "channel", i, dataGridSale);

                            validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msgChannel + "\r\n";
                        }

                        //check allow main offer
                        if (String.IsNullOrEmpty(offer))
                        {
                            string msg = "Main Offer is null or empty.";
                            listBox1.Items.Add(msg);
                            indexListbox.Add(i);
                            hilightRow("rule", "offer", i, dataGridSale);

                            validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                         + msg + "\r\n";
                        }
                        else
                        {
                            string msgAllowOffer = validation.CheckAllowOffer(offer);
                            if (msgAllowOffer != "Success")
                            {
                                listBox1.Items.Add(msgAllowOffer);
                                indexListbox.Add(i);
                                hilightRow("rule", "offer", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                             + msgAllowOffer + "\r\n";
                            }
                        }

                        //check order type
                        if (String.IsNullOrEmpty(order))
                        {
                            string msg = "Order Type is null or empty.";
                            listBox1.Items.Add(msg);
                            indexListbox.Add(i);
                            hilightRow("rule", "order", i, dataGridSale);

                            validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msg + "\r\n";
                        }
                        else
                        {
                            order = Regex.Replace(order, "NEW", "New", RegexOptions.IgnoreCase);
                            order = Regex.Replace(order, "CHANGE", "Change", RegexOptions.IgnoreCase);
                            dataGridSale.Rows[i].Cells[6].Value = order;

                            string msgOrder = validation.CheckOrderType(order);
                            if (msgOrder != "Success")
                            {
                                listBox1.Items.Add(msgOrder);
                                indexListbox.Add(i);
                                hilightRow("rule", "order", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                            + msgOrder + "\r\n";
                            }
                        }
                       
                        //check product
                        bool hasProd = false;
                        foreach (DataRow row in tableProdType.Rows)
                        {
                            string mediaDB = row[0].ToString();

                            if (product == mediaDB)
                            {
                                hasProd = true;
                                break;
                            }
                        }

                        if (hasProd == false)
                        {
                            //write log
                            string msg = "Not found Product: " + product + " on Master Data.";
                            listBox1.Items.Add(msg);
                            indexListbox.Add(i);
                            hilightRow("rule", "product", i, dataGridSale);

                            validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msg + "\r\n";
                        }

                        //Check province
                        if(String.IsNullOrEmpty(province))
                        {
                            string msg = "Province is null or empty.";
                            listBox1.Items.Add(msg);
                            indexListbox.Add(i);
                            hilightRow("rule", "province", i, dataGridSale);

                            validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msg + "\r\n";
                        }
                        else
                        {
                            string msgProvince = validation.CheckProvince(province);
                            if (msgProvince != "Success")
                            {
                                listBox1.Items.Add(msgProvince);
                                indexListbox.Add(i);
                                hilightRow("rule", "province", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                            + msgProvince + "\r\n";
                            }
                        }
                        
                        //check allow advance month
                        if(String .IsNullOrEmpty(advMonth))
                        {
                            string msg = "Allow advance month is null or empty.";
                            listBox1.Items.Add(msg);
                            indexListbox.Add(i);
                            hilightRow("rule", "month", i, dataGridSale);

                            validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msg + "\r\n";
                        }
                        else
                        {
                            string msgAdvMonth = validation.CheckAllowAdvMonth(advMonth);
                            if (msgAdvMonth != "Success")
                            {
                                listBox1.Items.Add(msgAdvMonth);
                                indexListbox.Add(i);
                                hilightRow("rule", "month", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                            + msgAdvMonth + "\r\n";
                            }
                        }

                        //Check DownloadSpeed
                        if(String.IsNullOrEmpty(downloadF))
                        {
                            string msg = "Download(From) is null or empty.";
                            listBox1.Items.Add(msg);
                            indexListbox.Add(i);
                            hilightRow("rule", "downF", i, dataGridSale);

                            validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msg + "\r\n";
                        }
                        else if(String.IsNullOrEmpty(downloadT))
                        {
                            string msg = "Download(To) is null or empty.";
                            listBox1.Items.Add(msg);
                            indexListbox.Add(i);
                            hilightRow("rule", "downT", i, dataGridSale);

                            validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msg + "\r\n";
                        }
                        else
                        {
                            string[] msgDownload = validation.CheckSpeedVAS(downloadF, downloadT);
                            if (msgDownload[0] != "Success")
                            {
                                listBox1.Items.Add(msgDownload[0]);
                                indexListbox.Add(i);
                                hilightRow("rule", "downF", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                            + msgDownload[0] + "\r\n";
                            }
                            if (msgDownload[1] != "Success")
                            {
                                listBox1.Items.Add(msgDownload[1]);
                                indexListbox.Add(i);
                                hilightRow("rule", "downT", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                            + msgDownload[1] + "\r\n";
                            }
                            if (msgDownload[2] != "Success")
                            {
                                listBox1.Items.Add(msgDownload[2]);
                                indexListbox.Add(i);
                                hilightRow("rule", "downF", i, dataGridSale);
                                hilightRow("rule", "downT", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                            + msgDownload[2] + "\r\n";
                            }
                        }

                        //Check UploadSpeed
                        if (String.IsNullOrEmpty(uploadF))
                        {
                            string msg = "Upload(From) is null or empty.";
                            listBox1.Items.Add(msg);
                            indexListbox.Add(i);
                            hilightRow("rule", "upF", i, dataGridSale);

                            validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msg + "\r\n";
                        }
                        else if(String.IsNullOrEmpty(uploadT))
                        {
                            string msg = "Upload(To) is null or empty.";
                            listBox1.Items.Add(msg);
                            indexListbox.Add(i);
                            hilightRow("rule", "upT", i, dataGridSale);

                            validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msg + "\r\n";
                        }
                        else
                        {
                            string[] msgUpload = validation.CheckSpeedVAS(uploadF, uploadT);
                            if (msgUpload[0] != "Success")
                            {
                                listBox1.Items.Add(msgUpload[0]);
                                indexListbox.Add(i);
                                hilightRow("rule", "upF", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                            + msgUpload[0] + "\r\n";
                            }
                            if (msgUpload[1] != "Success")
                            {
                                listBox1.Items.Add(msgUpload[1]);
                                indexListbox.Add(i);
                                hilightRow("rule", "upT", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                            + msgUpload[1] + "\r\n";
                            }
                            if (msgUpload[2] != "Success")
                            {
                                listBox1.Items.Add(msgUpload[2]);
                                indexListbox.Add(i);
                                hilightRow("rule", "upF", i, dataGridSale);
                                hilightRow("rule", "upT", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                            + msgUpload[2] + "\r\n";
                            }
                        }
                        
                        //Check price
                        if(String.IsNullOrEmpty(priceF))
                        {
                            string msg = "Price(From) is null or empty.";
                            listBox1.Items.Add(msg);
                            indexListbox.Add(i);
                            hilightRow("rule", "priceF", i, dataGridSale);

                            validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msg + "\r\n";
                        }
                        else if(String.IsNullOrEmpty(priceT))
                        {
                            string msg = "Price(To) is null or empty.";
                            listBox1.Items.Add(msg);
                            indexListbox.Add(i);
                            hilightRow("rule", "priceT", i, dataGridSale);

                            validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msg + "\r\n";
                        }
                        else
                        {
                            string[] msgPrice = validation.CheckPrice(priceF, priceT);
                            if (msgPrice[0] != "Success")
                            {
                                listBox1.Items.Add(msgPrice[0]);
                                indexListbox.Add(i);
                                hilightRow("rule", "priceF", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                            + msgPrice[0] + "\r\n";
                            }
                            if (msgPrice[1] != "Success")
                            {
                                listBox1.Items.Add(msgPrice[1]);
                                indexListbox.Add(i);
                                hilightRow("rule", "priceT", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                            + msgPrice[1] + "\r\n";
                            }
                            if (msgPrice[2] != "Success")
                            {
                                listBox1.Items.Add(msgPrice[2]);
                                indexListbox.Add(i);
                                hilightRow("rule", "priceF", i, dataGridSale);
                                hilightRow("rule", "priceT", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                            + msgPrice[2] + "\r\n";
                            }
                        }

                        //Check Date
                        string msgDate = validation.CheckDate(start, end);
                        if (msgDate != "Success")
                        {
                            if (msgDate == "Start Date fotmat is not supported")
                            {
                                listBox1.Items.Add(msgDate);
                                indexListbox.Add(i);
                                hilightRow("rule", "start", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msgDate + "\r\n";
                            }
                            else if (msgDate == "End Date fotmat is not supported")
                            {
                                listBox1.Items.Add(msgDate);
                                indexListbox.Add(i);
                                hilightRow("rule", "end", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msgDate + "\r\n";
                            }
                            else
                            {
                                listBox1.Items.Add(msgDate);
                                indexListbox.Add(i);
                                hilightRow("rule", "start", i, dataGridSale);
                                hilightRow("rule", "end", i, dataGridSale);

                                validateLog += "[(row:" + i + 5 + ") VASCode:" + code + ", Speed:" + speed + ", Price:" + price + "]     "
                                        + msgDate + "\r\n";
                            }
                        }

                        backgroundWorker1.ReportProgress(30 + ((i + 1) * 70 / dataGridSale.RowCount));
                    }

                }

                toolStripStatusLabel1.Text = "Validation Completed!!";
            }
            catch (Exception e)
            {
                backgroundWorker1.CancelAsync();
                toolStripStatusLabel1.Text = "Failed to validate new vas sale for SmartUI";
                MessageBox.Show("There was a problem during the validation new vas sale process.Please try again later." + "\r\n" +
                    "Detail : " + e.Message, "Automate Mapping Tool", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;
            }
        }

        private void ValidateMKTNotAllow()
        {
            try
            {
                Application.UseWaitCursor = true;
                Cursor.Current = Cursors.WaitCursor;

                toolStripStatusLabel1.Text = "Checking MKT Not Allow...";

                dataGridMKT = userControlNotAllow1.GetDataGridView;

                backgroundWorker1.ReportProgress(15);

                InitialValue(dataGridMKT);

                backgroundWorker1.ReportProgress(30);

                for (int i = 0; i < dataGridMKT.RowCount; i++)
                {
                    string code = dataGridMKT.Rows[i].Cells[0].Value.ToString().Trim();
                    string mkt = dataGridMKT.Rows[i].Cells[1].Value.ToString().Trim();
                    string flag = dataGridMKT.Rows[i].Cells[2].Value.ToString().Trim();

                    if (String.IsNullOrEmpty(code))
                    {
                        //write log
                        string msg = "VASCode is null or empty.";
                        listBox1.Items.Add(msg);
                        indexListbox.Add(i);
                        hilightRow("notAllow", "code", i, dataGridMKT);

                        validateLog += "[(row:" + i + 3 + ") VASCode:" + code + ", MKT:" + mkt + "]     "
                                    + msg + "\r\n";
                    }

                    if (code.Length != 15)
                    {
                        //write log
                        string msg = "VASCode fotmat is not supported";
                        listBox1.Items.Add(msg);
                        indexListbox.Add(i);
                        hilightRow("notAllow", "code", i, dataGridMKT);

                        validateLog += "[(row:" + i + 3 + ") VASCode:" + code + ", MKT:" + mkt + "]     "
                                    + msg + "\r\n";
                    }

                    if (mkt.Contains("-") == false)
                    {
                        //write log
                        string msg = "MKTCode fotmat is not supported";
                        listBox1.Items.Add(msg);
                        indexListbox.Add(i);
                        hilightRow("notAllow", "mkt", i, dataGridMKT);

                        validateLog += "[(row:" + i + 3 + ") VASCode:" + code + ", MKT:" + mkt + "]     "
                                    + msg + "\r\n";
                    }

                    if (!flag.Equals("Y") && !flag.Equals("N"))
                    {
                        //write log
                        string msg = "Invalid value of Active_Flag";
                        listBox1.Items.Add(msg);
                        indexListbox.Add(i);
                        hilightRow("notAllow", "flag", i, dataGridMKT);

                        validateLog += "[(row:" + i + 3 + ") VASCode:" + code + ", MKT:" + mkt + "]     "
                                    + msg + "\r\n";
                    }

                    backgroundWorker1.ReportProgress(30 + ((i + 1) * 70 / dataGridMKT.RowCount));
                }

                toolStripStatusLabel1.Text = "Validation Completed!!";
            }
            catch (Exception e)
            {
                backgroundWorker1.CancelAsync();
                toolStripStatusLabel1.Text = "Failed to validate main offer not allow";
                MessageBox.Show("There was a problem during the validation main offer not allow process.Please try again later." + "\r\n" +
                    "Detail : " + e.Message, "Automate Mapping Tool", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;
            }
        }

        private void ValidateUpdateVASSmartUI()
        {
            try
            {
                Application.UseWaitCursor = true;
                Cursor.Current = Cursors.WaitCursor;

                toolStripStatusLabel1.Text = "Checking data for update date...";

                dataGridUpdate = userControlUpdate1.GetDataGridView;

                backgroundWorker1.ReportProgress(15);

                InitialValue(dataGridUpdate);

                backgroundWorker1.ReportProgress(30);

                for (int i = 0; i < dataGridUpdate.RowCount; i++)
                {
                    string id = dataGridUpdate.Rows[i].Cells[0].Value.ToString().Trim();
                    string start = dataGridUpdate.Rows[i].Cells[1].Value.ToString().Trim();
                    string end = dataGridUpdate.Rows[i].Cells[2].Value.ToString().Trim();

                    if (String.IsNullOrEmpty(id))
                    {
                        //write log id is null
                        string msg = "VAS_ID is null or empty.";
                        listBox1.Items.Add(msg);
                        indexListbox.Add(i);
                        hilightRow("update", "id", i, dataGridUpdate);

                        validateLog += "[row:" + i + 4 + "]     " + msg + "\r\n";
                    }
                    else
                    {
                        if (String.IsNullOrEmpty(start))
                        {
                            if (Convert.ToDateTime(end) < DateTime.Now)
                            {
                                string msg = "EndDate is invalid";
                                listBox1.Items.Add(msg);
                                indexListbox.Add(i);
                                hilightRow("update", "end", i, dataGridUpdate);

                                validateLog += "[(row:" + i + 4 + ") VAS_ID: " + id + "]     " + msg + "\r\n";
                            }
                        }
                        else
                        {
                            //check format startdate
                            start = validation.ChangeFormatDate(start);

                            if (start == "Invalid")
                            {
                                string msg = "Start Date fotmat is not supported";
                                listBox1.Items.Add(msg);
                                indexListbox.Add(i);
                                hilightRow("update", "start", i, dataGridUpdate);

                                validateLog += "[(row:" + i + 4 + ") VAS_ID: " + id + "]     " + msg + "\r\n";
                            }
                            else
                            {
                                if (String.IsNullOrEmpty(end) == false)
                                {
                                    //check format enddate
                                    end = validation.ChangeFormatDate(end);

                                    if (end == "Invalid")
                                    {
                                        string msg = "End Date fotmat is not supported";
                                        listBox1.Items.Add(msg);
                                        indexListbox.Add(i);
                                        hilightRow("update", "end", i, dataGridUpdate);

                                        validateLog += "[(row:" + i + 4 + ") VAS_ID: " + id + "]     " + msg + "\r\n";
                                    }
                                    else
                                    {
                                        if (Convert.ToDateTime(end) < DateTime.Now)
                                        {
                                            string msg = "EndDate is invalid";
                                            listBox1.Items.Add(msg);
                                            indexListbox.Add(i);
                                            hilightRow("update", "end", i, dataGridUpdate);

                                            validateLog += "[(row:" + i + 4 + ") VAS_ID: " + id + "]     " + msg + "\r\n";
                                        }
                                    }
                                }
                            }
                        }
                    }

                    backgroundWorker1.ReportProgress(30 + ((i + 1) * 70 / dataGridUpdate.RowCount));
                }

                if (listBox1.Items.Count <= 0)
                {
                    btnExe.Enabled = true;
                }

                toolStripStatusLabel1.Text = "Validation Completed!!";
            }
            catch (Exception e)
            {
                backgroundWorker1.CancelAsync();
                toolStripStatusLabel1.Text = "Failed to validate update date";
                MessageBox.Show("There was a problem during the validation update date for SmartUI process.Please try again later." + "\r\n" +
                    "Detail : " + e.Message, "Automate Mapping Tool", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;
            }
        }

        private void InsertNewVASCode()
        {
            string log = "", sql = "";

            try
            {
                Application.UseWaitCursor = true;
                Cursor.Current = Cursors.WaitCursor;

                toolStripStatusLabel1.Text = "Inserting New VASCode...";

                //Create an OracleCommand object using the connection object
                OracleCommand command = ConnectionProd.CreateCommand();
                OracleTransaction transaction = null;

                for (int i = 0; i < dataGridCode.RowCount; i++)
                {
                    using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                    {
                        command.Transaction = transaction;

                        string code = dataGridCode.Rows[i].Cells[0].Value.ToString().Trim();
                        string desc = dataGridCode.Rows[i].Cells[1].Value.ToString().Trim();
                        string type = dataGridCode.Rows[i].Cells[2].Value.ToString().Trim();
                        string rule = dataGridCode.Rows[i].Cells[3].Value.ToString().Trim();
                        string price = dataGridCode.Rows[i].Cells[4].Value.ToString().Trim();
                        string channel = dataGridCode.Rows[i].Cells[5].Value.ToString().Trim();
                        string group = dataGridCode.Rows[i].Cells[6].Value.ToString().Trim();
                        string start = dataGridCode.Rows[i].Cells[7].Value.ToString().Trim();

                        start = validation.ChangeFormatDate(start);
                        if (String.IsNullOrEmpty(start))
                        {
                            start = "";
                        }

                        string txt = "SELECT * FROM VAS_PRODUCT WHERE VAS_CODE = '" + code + "' AND VAS_CHANNEL = '" + channel + "'";

                        command.CommandText = txt;
                        OracleDataReader reader = command.ExecuteReader();
                        reader.Read();

                        if (reader.HasRows)
                        {
                            string vasRule = reader["VAS_RULE"].ToString();
                            string vasPrice = reader["VAS_PRICE"].ToString();
                            string vasStatus = reader["VAS_STATUS"].ToString();
                            DateTime date = Convert.ToDateTime(reader["VAS_START_DATE"]);

                            if (rule == vasRule)
                            {
                                if (price == vasPrice)
                                {
                                    if (vasStatus == "Active")
                                    {
                                        if (date != Convert.ToDateTime(start))
                                        {
                                            //write log
                                            log += "[VAS_CODE : " + code + " ,Channel : " + channel + "] VAS_START_DATE(file) is not equal to VAS_START_DATE(DB)" + "\r\n";
                                        }
                                    }
                                    else
                                    {
                                        try
                                        {
                                            //update status
                                            string cmdTxt = "UPDATE VAS_PRODUCT SET VAS_STATUS = 'Active', VAS_START_DATE = " +
                                                "TO_DATE('" + start + "','dd/mm/yyyy'), VAS_END_DATE = null WHERE VAS_CODE = '" + code +
                                                "' AND VAS_TYPE = '" + type + "' AND VAS_RULE = '" + rule + "' AND VAS_CHANNEL = '" + channel + "'";
                                            sql += cmdTxt + ";" + "\r\n";

                                            command.CommandText = cmdTxt;
                                            command.ExecuteNonQuery();
                                            transaction.Commit();

                                            lstCode += "'" + code + "',";
                                        }
                                        catch (Exception e)
                                        {
                                            //write log cannot update
                                            transaction.Rollback();

                                            log += "Failed to update data VAS_Code: " + code + " Channel: " + channel + " Rule: " + rule + " into database" + "\r\n" +
                                                    "Detail :" + e.Message + "\r\n" + "\r\n";
                                        }
                                    }
                                }
                                else
                                {
                                    //write log
                                    log += "[VAS_CODE : " + code + " ,Channel : " + channel + "] VAS_PRICE(file) is not equal to VAS_PRICE(DB)" + "\r\n";
                                }
                            }
                            else
                            {
                                //write log
                                log += "[VAS_CODE : " + code + " ,Channel : " + channel + "] VAS_RULE(file) is not equal to VAS_RULE(DB)" + "\r\n";
                            }
                            reader.Close();
                        }
                        else
                        {
                            try
                            {
                                //insert
                                string cmdTxt = "INSERT INTO VAS_PRODUCT VALUES ('" + code + "','" + desc + "','" + type + "','Active','" + rule + "','" +
                                    price + "',null,'" + channel + "',null,'" + group + "',TO_DATE('" + start + "','dd/mm/yyyy')" +
                                    ",TO_DATE('','dd/mm/yyyy'),sysdate,'" + implementer + "',null,null)";
                                sql += cmdTxt + ";" + "\r\n";

                                command.CommandText = cmdTxt;
                                command.ExecuteNonQuery();

                                transaction.Commit();

                                lstCode += "'" + code + "',";
                            }
                            catch (Exception ex)
                            {
                                //write log cannot insert
                                transaction.Rollback();

                                log += "Failed to insert vas_code:  " + code + " Channel: " + channel + " Rule: " + rule + " into database" + "\r\n" +
                                    "Detail :" + ex.Message + "\r\n" + "\r\n";
                            }
                        }

                        backgroundWorker2.ReportProgress((i + 1) * 80 / dataGridCode.RowCount);
                    }
                }

                toolStripStatusLabel1.Text = "Already inserted new vas code";
            }
            catch (Exception e)
            {
                toolStripProgressBar1.Value = 0;
                toolStripStatusLabel1.Text = "Failed to insert new vas code";
                log += "There was a problem during the insert process." + "\r\n" + "Detail : " + e.Message + "\r\n" + "\r\n";
            }
            finally
            {
                if (String.IsNullOrEmpty(log) == false)
                {
                    //write log
                    string logPath = outputPath + "\\Log_New_VASCode" + urNo.ToUpper() + ".txt";
                    using (StreamWriter writer = new StreamWriter(logPath, true))
                    {
                        writer.Write(log);
                    }
                }

                if (String.IsNullOrEmpty(sql) == false)
                {
                    //write log
                    string path = outputPath + "\\Script_New_VASCode" + urNo.ToUpper() + ".txt";
                    using (StreamWriter writer = new StreamWriter(path, true))
                    {
                        writer.Write(sql);
                    }
                }

                backgroundWorker2.ReportProgress(100);

                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;
            }
        }

        private void InsertVASNewSale()
        {
            string log = "", sql = "", existing = "", id = "";

            try
            {
                Application.UseWaitCursor = true;
                Cursor.Current = Cursors.WaitCursor;

                toolStripStatusLabel1.Text = "Inserting VAS Sale for SmartUI...";

                OracleCommand cmd = null;
                OracleTransaction transaction = null;

                ReserveID reserveID = new ReserveID();
                reserveID.Reserve(ConnectionProd, ConnectionTemp, "VAS", implementer, urNo);

                //Get MinID
                cmd = new OracleCommand("SELECT MAX(DC_ID) FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_ID LIKE 'VAS%'", ConnectionProd);
                OracleDataReader reader = cmd.ExecuteReader();
                reader.Read();
                int minID = Convert.ToInt32(Convert.ToString(reader[0]).Substring(3)) + 1;
                reader.Close();

                cmd = ConnectionProd.CreateCommand();

                backgroundWorker2.ReportProgress(10);

                for (int i = 0; i < dataGridSale.RowCount; i++)
                {
                    string code = dataGridSale.Rows[i].Cells[0].Value.ToString().Trim();
                    string speed = dataGridSale.Rows[i].Cells[2].Value.ToString().Trim();
                    string channel = dataGridSale.Rows[i].Cells[4].Value.ToString();
                    string offer = dataGridSale.Rows[i].Cells[5].Value.ToString().Trim();
                    string order = dataGridSale.Rows[i].Cells[6].Value.ToString();
                    string product = dataGridSale.Rows[i].Cells[7].Value.ToString().Trim();
                    string province = dataGridSale.Rows[i].Cells[8].Value.ToString().Trim();
                    string month = dataGridSale.Rows[i].Cells[9].Value.ToString().Trim();
                    string downloadF = dataGridSale.Rows[i].Cells[10].Value.ToString().Trim();
                    string downloadT = dataGridSale.Rows[i].Cells[11].Value.ToString().Trim();
                    string uploadF = dataGridSale.Rows[i].Cells[12].Value.ToString().Trim();
                    string uploadT = dataGridSale.Rows[i].Cells[13].Value.ToString().Trim();
                    string priceF = dataGridSale.Rows[i].Cells[15].Value.ToString().Trim();
                    string priceT = dataGridSale.Rows[i].Cells[16].Value.ToString().Trim();
                    string start = dataGridSale.Rows[i].Cells[17].Value.ToString().Trim();
                    string end = dataGridSale.Rows[i].Cells[18].Value.ToString().Trim();

                    cmd.CommandText = "SELECT * FROM VAS_PRODUCT WHERE VAS_CODE = '" + code + "'";
                    reader = cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        //main offer
                        string[] lstOffer;
                        if (offer.Contains(','))
                        {
                            lstOffer = offer.Split(',');
                        }
                        else
                        {
                            lstOffer = new string[1];
                            lstOffer[0] = offer;
                        }

                        //OrderType
                        string[] lstOrder;
                        if (order.Contains(","))
                        {
                            lstOrder = order.Split(',');
                        }
                        else
                        {
                            lstOrder = new string[1];
                            lstOrder[0] = order;
                        }

                        //Province
                        string[] lstProv;
                        if (province.Contains(","))
                        {
                            lstProv = province.Split(',');
                        }
                        else
                        {
                            lstProv = new string[1];
                            lstProv[0] = province;
                        }

                        //Channel
                        string[] lstChannel;
                        if (channel.Contains(","))
                        {
                            lstChannel = channel.Split(',');
                        }
                        else
                        {
                            lstChannel = new string[1];
                            lstChannel[0] = channel;
                        }

                        //ConvertSpeed2K
                        if (downloadF != "ALL")
                        {
                            downloadF = Convert.ToString(validation.ConvertUOM2K(downloadF, Regex.Replace(downloadF, "[0-9]", "")));
                        }

                        if (downloadT != "ALL")
                        {
                            downloadT = Convert.ToString(validation.ConvertUOM2K(downloadT, Regex.Replace(downloadT, "[0-9]", "")));
                        }

                        if (uploadF != "ALL")
                        {
                            uploadF = Convert.ToString(validation.ConvertUOM2K(uploadF, Regex.Replace(uploadF, "[0-9]", "")));
                        }

                        if (uploadT != "ALL")
                        {
                            uploadT = Convert.ToString(validation.ConvertUOM2K(uploadT, Regex.Replace(uploadT, "[0-9]", "")));
                        }

                        //Change format Date
                        start = validation.ChangeFormatDate(start);
                        end = validation.ChangeFormatDate(end);

                        for (int j = 0; j < lstOffer.Length; j++)
                        {
                            for (int k = 0; k < lstOrder.Length; k++)
                            {
                                for (int l = 0; l < lstProv.Length; l++)
                                {
                                    for (int m = 0; m < lstChannel.Length; m++)
                                    {
                                        cmd.CommandText = "SELECT * FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_ID IN " +
                                            "(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_GROUPID = '" + code + "' " +
                                            "AND DC_ID IN(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + lstOffer[j].Trim() + "' AND DC_TYPE = 'PROMOTION_CODE') " +
                                            "AND DC_ID IN(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + lstOrder[k].Trim() + "' AND DC_TYPE = 'ORDER_TYPE') " +
                                            "AND DC_ID IN(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + product + "' AND DC_TYPE = 'PRODUCT') " +
                                            "AND DC_ID IN(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + lstProv[l].Trim() + "' AND DC_TYPE = 'PROVINCE') " +
                                            "AND DC_ID IN(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + lstChannel[m].Trim() + "' AND DC_TYPE = 'SALE_CHANNEL') " +
                                            "AND DC_ID IN(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + month + "' AND DC_TYPE = 'ALLOW_ADVANCE_MONTH') " +
                                            "AND DC_ID IN(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + downloadF + "' AND DC_TYPE = 'DL_FROM') " +
                                            "AND DC_ID IN(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + downloadT + "' AND DC_TYPE = 'DL_TO') " +
                                            "AND DC_ID IN(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + uploadF + "' AND DC_TYPE = 'UL_FROM') " +
                                            "AND DC_ID IN(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + uploadT + "' AND DC_TYPE = 'UL_TO') " +
                                            "AND DC_ID IN(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + priceF + "' AND DC_TYPE = 'PRICE_FROM') " +
                                            "AND DC_ID IN(SELECT DC_ID FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_VALUE = '" + priceT + "' AND DC_TYPE = 'PRICE_TO')) ";

                                        reader = cmd.ExecuteReader();
                                        if (reader.HasRows)
                                        {
                                            reader.Read();
                                            string dcID = reader["DC_ID"].ToString();
                                            string sDate = reader["DC_START_DT"].ToString();
                                            string eDate = reader["DC_END_DT"].ToString();

                                            /*VAS_ID,START_DATE,END_DATE,VAS_CODE,VAS_NAME,VAS_PRICE,VAS_STATUS,VAS_RULE,VAS_CHANNEL,SALE_CHANNEL,PROMOTION_CODE
,ORDER_TYPE,ALLOW_ADVANCE_MONTH,DOWNLOAD_FROM,DOWNLOAD_TO,UPLOAD_FROM,UPLOAD_TO
,PRICE_FROM,PRICE_TO,PROVINCE,PRODUCT,PARENT_VAS_CODE,VAS_TYPE*/
                                            //write existing data
                                            existing += dcID +sDate+eDate+ code + ", Channel: " + lstChannel[m] + ", Main Offer: " + lstOffer[j] + ", Order: " + lstOrder[k] +
                                                ", Province: " + lstProv[l] + ", Speed: " + speed + ", StartDate: " + sDate + ", EndDate: " + eDate + "\r\n";

                                            existingID += "'" + dcID + "',";
                                        }
                                        else
                                        {
                                            using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                                            {
                                                cmd.Transaction = transaction;
                                                try
                                                {
                                                    id = "VAS" + string.Format("{0:0000000}", minID);

                                                    //insert new vas for smartUI
                                                    for (int n = 1; n <= 12; n++)
                                                    {
                                                        object[] obj = new object[11];

                                                        obj[0] = "VAS" + string.Format("{0:0000000}", minID);

                                                        obj[3] = code;

                                                        obj[4] = start;

                                                        obj[5] = end;

                                                        obj[6] = "Y";

                                                        obj[7] = DateTime.Now.ToString("dd/MM/yyyy");
                                                        obj[8] = implementer;
                                                        obj[9] = "";
                                                        obj[10] = "";

                                                        switch (n)
                                                        {
                                                            case 1:
                                                                obj[1] = "PROMOTION_CODE";
                                                                obj[2] = lstOffer[j].Trim();
                                                                break;
                                                            case 2:
                                                                obj[1] = "ORDER_TYPE";
                                                                obj[2] = lstOrder[k].Trim();
                                                                break;
                                                            case 3:
                                                                obj[1] = "PRODUCT";
                                                                obj[2] = product;
                                                                break;
                                                            case 4:
                                                                obj[1] = "PROVINCE";
                                                                obj[2] = lstProv[l].Trim();
                                                                break;
                                                            case 5:
                                                                obj[1] = "SALE_CHANNEL";
                                                                obj[2] = lstChannel[m].Trim();
                                                                break;
                                                            case 6:
                                                                obj[1] = "ALLOW_ADVANCE_MONTH";
                                                                obj[2] = month;
                                                                break;
                                                            case 7:
                                                                obj[1] = "DL_FROM";
                                                                obj[2] = downloadF;
                                                                break;
                                                            case 8:
                                                                obj[1] = "DL_TO";
                                                                obj[2] = downloadT;
                                                                break;
                                                            case 9:
                                                                obj[1] = "UL_FROM";
                                                                obj[2] = uploadF;
                                                                break;
                                                            case 10:
                                                                obj[1] = "UL_TO";
                                                                obj[2] = uploadT;
                                                                break;
                                                            case 11:
                                                                obj[1] = "PRICE_FROM";
                                                                obj[2] = priceF;
                                                                break;
                                                            case 12:
                                                                obj[1] = "PRICE_TO";
                                                                obj[2] = priceT;
                                                                break;
                                                        }

                                                        cmd.CommandText = "INSERT INTO DISCOUNT_CRITERIA_MAPPING VALUES ('" + id + "','" + obj[1] + "','" + obj[2] + "','" + code +
                                                            "',to_date('" + start + "','dd/mm/yyyy')," + "to_date('" + end + "', 'dd/mm/yyyy'),'Y',sysdate,'" +
                                                            implementer + "',null,null)";
                                                        cmd.ExecuteNonQuery();

                                                        sql += cmd.CommandText + ";" + "\r\n";

                                                        lstID += "'" + id + "',";
                                                    }

                                                    transaction.Commit();
                                                }
                                                catch (Exception ex)
                                                {
                                                    transaction.Rollback();

                                                    //write log
                                                    log += "Failed to insert ID: " + id + ", VASCode:" + code + ", Channel:" + lstChannel[m] + ", Main offer:" + lstOffer[j] +
                                                        ", Order: " + lstOrder[k] + ", Province:" + lstProv[l] + " into database" + "\r\n" + "Detail of system : " +
                                                        ex.Message + "\r\n" + "\r\n";
                                                }
                                            }
                                        }
                                        minID += 1;
                                    }

                                }
                            }
                        }
                    }
                    else
                    {
                        //write log not found vascode in vas_product
                        log += "Not found VASCode[" + code + "] in database table[VAS_PRODUCT]" + "\r\n";
                    }

                    backgroundWorker2.ReportProgress(10 + ((i + 1) * 70 / dataGridSale.RowCount));
                }

                //Update ReserveID
                reserveID.UpdateReserveID(ConnectionTemp, ConnectionProd, "VAS", urNo);

                backgroundWorker2.ReportProgress(90);

                toolStripStatusLabel1.Text = "Already inserted new vas sale for SmartUI";
            }
            catch (Exception e)
            {
                toolStripProgressBar1.Value = 0;
                toolStripStatusLabel1.Text = "Failed to insert new vas sale for SmartUI";
                log += "There was a problem during the insert process." + "\r\n" + "Detail : " + e.Message + "\r\n" + "\r\n";
            }
            finally
            {
                if (String.IsNullOrEmpty(log) == false)
                {
                    //write log
                    string logPath = outputPath + "\\Log_New_VAS_SmartUI" + urNo.ToUpper() + ".txt";
                    using (StreamWriter writer = new StreamWriter(logPath, true))
                    {
                        writer.Write(log);
                    }
                }

                if (String.IsNullOrEmpty(existing) == false)
                {
                    string tmp = "VAS_ID, START_DATE, END_DATE, VAS_CODE, VAS_NAME, VAS_PRICE, VAS_STATUS, " +
                        "VAS_RULE, VAS_CHANNEL, SALE_CHANNEL, PROMOTION_CODE, ORDER_TYPE, ALLOW_ADVANCE_MONTH, " +
                        "DOWNLOAD_FROM, DOWNLOAD_TO, UPLOAD_FROM, UPLOAD_TO, PRICE_FROM, PRICE_TO, PROVINCE, " +
                        "PRODUCT, PARENT_VAS_CODE, VAS_TYPE" + "\r\n" + "\r\n" + existing;

                    string logPath = outputPath + "\\ExistingData" + urNo.ToUpper() + ".txt";
                    using (StreamWriter writer = new StreamWriter(logPath, true))
                    {
                        writer.Write(tmp);
                    }
                }

                if (String.IsNullOrEmpty(sql) == false)
                {
                    string path = outputPath + "\\Script_New_VAS_SmartUI" + urNo.ToUpper() + ".txt";
                    using (StreamWriter writer = new StreamWriter(path, true))
                    {
                        writer.Write(sql);
                    }
                }

                backgroundWorker2.ReportProgress(100);

                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;
            }
        }

        private void InsertMKTNotAllow()
        {
            string log = "", sql = "", existing = "";

            try
            {
                Application.UseWaitCursor = true;
                Cursor.Current = Cursors.WaitCursor;

                toolStripStatusLabel1.Text = "Inserting MKT Not Allow...";

                OracleCommand cmd = null;
                OracleDataReader reader = null;
                OracleTransaction transaction = null;
                cmd = ConnectionProd.CreateCommand();

                for (int i = 0; i < dataGridMKT.RowCount; i++)
                {
                    using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                    {
                        cmd.Transaction = transaction;

                        string code = dataGridMKT.Rows[i].Cells[0].Value.ToString().Trim();
                        string offer = dataGridMKT.Rows[i].Cells[1].Value.ToString().Trim();
                        string active = dataGridMKT.Rows[i].Cells[2].Value.ToString().Trim();

                        cmd.CommandText = "SELECT * FROM VAS_PRODUCT WHERE VAS_CODE = '" + code + "'";
                        reader = cmd.ExecuteReader();
                        if (reader.HasRows)
                        {
                            try
                            {
                                cmd.CommandText = "SELECT * FROM TMP_CONFLICTS_FEATURE WHERE FEATURE_CODE = '" + code +
                                    "' AND CONFLICT_CODE = '" + offer + "'";
                                reader = cmd.ExecuteReader();
                                reader.Read();
                                if (reader.HasRows)
                                {
                                    //check active flag
                                    string flag = reader["ACTIVE"].ToString();
                                    if (flag == active)
                                    {
                                        //write log existing data
                                        existing += "VASCode: " + code + ", Main offer: " + offer + ", Flag:" + flag + "\r\n";
                                    }
                                    else
                                    {
                                        //update flag
                                        cmd.CommandText = "UPDATE TMP_CONFLICTS_FEATURE SET ACTIVE = '" + active +
                                            "' WHERE FEATURE_CODE = '" + code + "' AND CONFLICT_CODE = '" + offer + "'";
                                        cmd.ExecuteNonQuery();

                                        sql += cmd.CommandText + ";" + "\r\n";

                                        lstOffer += "'" + offer + "',";
                                        lstCodeforOffer += "'" + code + "',";
                                    }
                                    reader.Close();
                                }
                                else
                                {
                                    //insert data
                                    cmd.CommandText = "INSERT INTO TMP_CONFLICTS_FEATURE VALUES ('" + code + "','" + offer + "','Y')";
                                    cmd.ExecuteNonQuery();

                                    sql += cmd.CommandText + ";" + "\r\n";

                                    lstOffer += "'" + offer + "',";
                                    lstCodeforOffer += "'" + code + "',";
                                }

                                transaction.Commit();
                            }
                            catch (Exception e)
                            {
                                transaction.Rollback();
                                //write log
                                log += "Failed to insert or update VASCode[" + code + "], Main offer[" + offer + "] into database" +
                                    "\r\n" + "Detail of system: " + e.Message + "\r\n" + "\r\n";
                            }
                        }
                        else
                        {
                            //write log not found vas product 
                            log += "Not found VASCode[" + code + "] in database table[VAS_PRODUCT]" + "\r\n";
                        }

                        //backgroundWorker2.ReportProgress((i + 1) * 80 / dataGridMKT.RowCount);
                        toolStripProgressBar1.Value = (i + 1) * 80 / dataGridMKT.RowCount;
                    }
                }

                toolStripStatusLabel1.Text = "Already inserted MKT Code Not Allow";
            }
            catch(Exception e)
            {
                toolStripProgressBar1.Value = 0;
                toolStripStatusLabel1.Text = "Failed to insert MKT Not Allow";
                log += "There was a problem during the insert process." + "\r\n" + "Detail : " + e.Message + "\r\n" + "\r\n";
            }
            finally
            {
                //write log
                if (String.IsNullOrEmpty(log) == false)
                {
                    //write log
                    string logPath = outputPath + "\\Log_AllowMainOffer" + urNo.ToUpper() + ".txt";
                    using (StreamWriter writer = new StreamWriter(logPath, true))
                    {
                        writer.Write(log);
                    }
                }

                if (String.IsNullOrEmpty(existing) == false)
                {
                    string logPath = outputPath + "\\Existing_MainOffer" + urNo.ToUpper() + ".txt";
                    using (StreamWriter writer = new StreamWriter(logPath, true))
                    {
                        writer.Write(existing);
                    }
                }

                if (String.IsNullOrEmpty(sql) == false)
                {
                    string path = outputPath + "\\Script_AllowMainOffer" + urNo.ToUpper() + ".txt";
                    using (StreamWriter writer = new StreamWriter(path, true))
                    {
                        writer.Write(sql);
                    }
                }

                backgroundWorker2.ReportProgress(100);
                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;
            }
        }

        private void UpdateDateVASSmartUI()
        {
            string cmdTxt, log = "", sql = "", id = "";

            try
            {
                Application.UseWaitCursor = true;
                Cursor.Current = Cursors.WaitCursor;

                toolStripStatusLabel1.Text = "Updating...";

                OracleDataReader reader = null;
                OracleTransaction transaction = null;
                OracleCommand cmd = ConnectionProd.CreateCommand();

                for (int i = 0; i < dataGridUpdate.RowCount; i++)
                {
                    using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                    {
                        cmd.Transaction = transaction;
                        try
                        {
                            id = dataGridUpdate.Rows[i].Cells[0].Value.ToString().Trim();
                            string start = dataGridUpdate.Rows[i].Cells[1].Value.ToString().Trim();
                            string end = dataGridUpdate.Rows[i].Cells[2].Value.ToString().Trim();

                            start = validation.ChangeFormatDate(start);
                            end = validation.ChangeFormatDate(end);

                            cmd.CommandText = "SELECT * FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_ID = '" + id + "'";
                            reader = cmd.ExecuteReader();
                            reader.Read();
                            if (reader.HasRows)
                            {
                                DateTime startDate = Convert.ToDateTime(start);
                                DateTime endDate = Convert.ToDateTime(end);
                                DateTime startDB = new DateTime();
                                DateTime endDB = new DateTime();

                                if (reader["DC_START_DT"] != DBNull.Value)
                                {
                                    startDB = Convert.ToDateTime(reader["DC_START_DT"]);
                                }

                                if (reader["DC_END_DT"] != DBNull.Value)
                                {
                                    endDB = Convert.ToDateTime(reader["DC_END_DT"]);
                                }

                                if (startDate > DateTime.Now)
                                {
                                    if (startDB > DateTime.Now)
                                    {
                                        //update db.dc_start_dt = file.start
                                        cmdTxt = "UPDATE DISCOUNT_CRITERIA_MAPPING SET DC_START_DT = TO_DATE('" + start + "', " +
                                            "'dd/MM/yyyy') WHERE DC_ID = '" + id + "'";
                                        cmd.CommandText = cmdTxt;
                                        cmd.ExecuteNonQuery();

                                        sql += cmdTxt + ";" + "\r\n";

                                        lstUpdateID += "'" + id + "',";
                                    }
                                    else
                                    {
                                        if ((endDB == null || endDB < DateTime.Now) && endDate == DateTime.Now)
                                        {
                                            //update db.dc_end_dt = datetime
                                            cmdTxt = "UPDATE DISCOUNT_CRITERIA_MAPPING SET DC_END_DT = sysdate WHERE DC_ID = '" + id + "'";
                                            cmd.CommandText = cmdTxt;
                                            cmd.ExecuteNonQuery();

                                            sql += cmdTxt + ";" + "\r\n";

                                            //update db.dc_start_dt = file.startDate
                                            cmdTxt = "UPDATE DISCOUNT_CRITERIA_MAPPING SET DC_START_DT = TO_DATE('" + start + "', " +
                                            "'dd/MM/yyyy') WHERE DC_ID = '" + id + "'";
                                            cmd.CommandText = cmdTxt;
                                            cmd.ExecuteNonQuery();

                                            sql += cmdTxt + ";" + "\r\n";

                                            lstUpdateID += "'" + id + "',";
                                        }
                                        else
                                        {
                                            if ((endDB == null || endDB < DateTime.Now) && (String.IsNullOrEmpty(end) || endDate > DateTime.Now))
                                            {
                                                //update db.dc_end_dt = file.enddate
                                                cmdTxt = "UPDATE DISCOUNT_CRITERIA_MAPPING SET DC_END_DT = TO_DATE('" + end + "', " +
                                                "'dd/MM/yyyy') WHERE DC_ID = '" + id + "'";
                                                cmd.CommandText = cmdTxt;
                                                cmd.ExecuteNonQuery();

                                                sql += cmdTxt + ";" + "\r\n";

                                                //update db.dc_start_dt = file.startDate
                                                cmdTxt = "UPDATE DISCOUNT_CRITERIA_MAPPING SET DC_START_DT = TO_DATE('" + start + "', " +
                                                "'dd/MM/yyyy') WHERE DC_ID = '" + id + "'";
                                                cmd.CommandText = cmdTxt;
                                                cmd.ExecuteNonQuery();

                                                sql += cmdTxt + ";" + "\r\n";

                                                lstUpdateID += "'" + id + "',";
                                            }
                                            else
                                            {
                                                log += "[VAS_ID: " + id + "] End Date is invalid" + "\r\n";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (endDate == DateTime.Now)
                                    {
                                        //update db.dc_end_dt = datetime
                                        cmdTxt = "UPDATE DISCOUNT_CRITERIA_MAPPING SET DC_END_DT = sysdate WHERE DC_ID = '" + id + "'";
                                        cmd.CommandText = cmdTxt;
                                        cmd.ExecuteNonQuery();

                                        sql += cmdTxt + ";" + "\r\n";

                                        lstUpdateID += "'" + id + "',";
                                    }
                                    else
                                    {
                                        if (endDate > DateTime.Now || String.IsNullOrEmpty(end))
                                        {
                                            //update db.dc_end_dt = file.enddate
                                            cmdTxt = "UPDATE DISCOUNT_CRITERIA_MAPPING SET DC_END_DT = TO_DATE('" + end + "', " +
                                                "'dd/MM/yyyy') WHERE DC_ID = '" + id + "'";
                                            cmd.CommandText = cmdTxt;
                                            cmd.ExecuteNonQuery();

                                            sql += cmdTxt + ";" + "\r\n";

                                            lstUpdateID += "'" + id + "',";
                                        }
                                        else
                                        {
                                            log += "[VAS_ID: " + id + "] End Date is invalid" + "\r\n";
                                        }
                                    }
                                }

                                transaction.Commit();
                            }
                            else
                            {
                                log += "Not found VAS_ID: " + id + " on database." + "\r\n";
                            }
                        }
                        catch (Exception e)
                        {
                            transaction.Rollback();

                            log += "Failed to update VAS_ID: " + id + " on database" + "\r\n" + "Detail of system : " +
                                e.Message + "\r\n";
                        }

                        backgroundWorker2.ReportProgress((i + 1) * 80 / dataGridUpdate.RowCount);
                    }
                }

                toolStripStatusLabel1.Text = "Already updated date";
            }
            catch (Exception e)
            {
                toolStripProgressBar1.Value = 0;
                toolStripStatusLabel1.Text = "Failed to update date vas criteria";
                log += "There was a problem during the update process." + "\r\n" + "Detail : " + e.Message + "\r\n" + "\r\n";
            }
            finally
            {
                //write log
                if (String.IsNullOrEmpty(log) == false)
                {
                    //write log
                    string logPath = outputPath + "\\Log_UpdateDateSmartUI" + urNo.ToUpper() + ".txt";
                    using (StreamWriter writer = new StreamWriter(logPath, true))
                    {
                        writer.Write(log);
                    }
                }

                if (String.IsNullOrEmpty(sql) == false)
                {
                    string path = outputPath + "\\Script_UpdateDateSmartUI" + urNo.ToUpper() + ".txt";
                    using (StreamWriter writer = new StreamWriter(path, true))
                    {
                        writer.Write(sql);
                    }
                }

                backgroundWorker2.ReportProgress(100);
                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;
            }
        }

        private void ExportCriteria()
        {
            Excel.Application xlApp = null;
            Excel.Workbook workbook = null;
            Excel.Workbook workbook2 = null;

            if(String.IsNullOrEmpty(lstID) && String.IsNullOrEmpty(lstUpdateID) && String.IsNullOrEmpty(lstCode)
                && String.IsNullOrEmpty(lstOffer) && String.IsNullOrEmpty(lstCodeforOffer) && String.IsNullOrEmpty(existingID))
            {
                toolStripStatusLabel1.Text = "Success";
                MessageBox.Show("There are no files to export.", "Success", MessageBoxButtons.OK);
            }
            else
            {
                try
                {
                    Application.UseWaitCursor = true;
                    Cursor.Current = Cursors.WaitCursor;

                    toolStripStatusLabel1.Text = "Exporting file...";

                    if (String.IsNullOrEmpty(lstID) == false)
                    {
                        lstID = lstID.Substring(0, lstID.Length - 1);
                    }

                    if (String.IsNullOrEmpty(lstUpdateID) == false)
                    {
                        lstUpdateID = lstUpdateID.Substring(0, lstUpdateID.Length - 1);
                    }

                    if (String.IsNullOrEmpty(lstCode) == false)
                    {
                        lstCode = lstCode.Substring(0, lstCode.Length - 1);
                    }

                    if (String.IsNullOrEmpty(lstOffer) == false)
                    {
                        lstOffer = lstOffer.Substring(0, lstOffer.Length - 1);
                    }

                    if (String.IsNullOrEmpty(lstCodeforOffer) == false)
                    {
                        lstCodeforOffer = lstCodeforOffer.Substring(0, lstCodeforOffer.Length - 1);
                    }

                    if (String.IsNullOrEmpty(existingID) == false)
                    {
                        existingID = existingID.Substring(0, existingID.Length - 1);
                    }

                    toolStripProgressBar1.Value = 10;

                    xlApp = new Excel.Application();
                    xlApp.Visible = false;
                    xlApp.DisplayAlerts = false;
                    workbook = xlApp.Workbooks.Add();
                    workbook2 = xlApp.Workbooks.Add();

                    //Script
                    string sqlVasProd = "SELECT * FROM VAS_PRODUCT WHERE VAS_CODE IN (" + lstCode + ")";

                    string sqlEliRule = "SELECT DC_ID VAS_ID, VAS_CODE, VAS_NAME, VAS_PRICE, VAS_STATUS, VAS_RULE" +
                                 ",to_char(trunc(DC_START_DT),'dd/mm/yyyy') START_DATE, to_char(trunc(DC_END_DT),'dd/mm/yyyy') END_DATE" +
                                 ",VAS_CHANNEL,SALE_CHANNEL,PROMOTION_CODE,ORDER_TYPE,ALLOW_ADVANCE_MONTH,DOWNLOAD_FROM,DOWNLOAD_TO" +
                                 ",UPLOAD_FROM,UPLOAD_TO,PRICE_FROM,PRICE_TO,PROVINCE,PRODUCT,PARENT_VAS_CODE,VAS_TYPE " +
                                 "FROM ( " +
                                    " SELECT * FROM " +
                                         "(SELECT DC1.DC_START_DT, DC1.DC_END_DT, DC1.DC_ID, DC1.DC_GROUPID " +
                                            ", NVL(PRODUCT, 'ALL') PRODUCT, NVL(PROMOTION_CODE, 'ALL') PROMOTION_CODE" +
                                            ", NVL(ORDER_TYPE, 'ALL') ORDER_TYPE, NVL(PROVINCE, 'ALL') PROVINCE" +
                                            ", NVL(SALE_CHANNEL, 'ALL') SALE_CHANNEL, NVL(ALLOW_ADVANCE_MONTH, 'ALL')ALLOW_ADVANCE_MONTH " +
                                            ", NVL(DOWNLOAD_FROM, 'ALL') DOWNLOAD_FROM, NVL(DOWNLOAD_TO, 'ALL') DOWNLOAD_TO" +
                                            ", NVL(UPLOAD_FROM, 'ALL') UPLOAD_FROM, NVL(UPLOAD_TO, 'ALL') UPLOAD_TO" +
                                            ", NVL(PRICE_FROM, 'ALL') PRICE_FROM, NVL(PRICE_TO, 'ALL') PRICE_TO " +
                                            "FROM " +
                                               "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRODUCT', DC_VALUE, 'ALL') PRODUCT " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PRODUCT'  AND DC_ACTIVE_FLAG = 'Y') DC1," +
                                               "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PROMOTION_CODE', DC_VALUE, 'ALL') PROMOTION_CODE " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PROMOTION_CODE' AND DC_ACTIVE_FLAG = 'Y') DC2," +
                                               "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'ORDER_TYPE', DC_VALUE, 'ALL') ORDER_TYPE " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'ORDER_TYPE' AND DC_ACTIVE_FLAG = 'Y') DC3," +
                                               "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PROVINCE', DC_VALUE, 'ALL') PROVINCE " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PROVINCE' AND DC_ACTIVE_FLAG = 'Y') DC4," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'SALE_CHANNEL', DC_VALUE, 'ALL') SALE_CHANNEL " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'SALE_CHANNEL' AND DC_ACTIVE_FLAG = 'Y') DC5," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'ALLOW_ADVANCE_MONTH', DC_VALUE, 'ALL') ALLOW_ADVANCE_MONTH " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'ALLOW_ADVANCE_MONTH' AND DC_ACTIVE_FLAG = 'Y' ) DC6," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M' end DOWNLOAD_FROM " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'DL_FROM' AND DC_ACTIVE_FLAG = 'Y') DC7," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M' end DOWNLOAD_TO " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'DL_TO' AND DC_ACTIVE_FLAG = 'Y') DC8," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M' end UPLOAD_FROM " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'UL_FROM' AND DC_ACTIVE_FLAG = 'Y') DC9," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M'  end UPLOAD_TO " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'UL_TO' AND DC_ACTIVE_FLAG = 'Y')  DC10," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRICE_FROM', DC_VALUE, 'ALL') PRICE_FROM " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PRICE_FROM' AND DC_ACTIVE_FLAG = 'Y') DC11," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRICE_TO', DC_VALUE, 'ALL') PRICE_TO " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PRICE_TO' AND DC_ACTIVE_FLAG = 'Y') DC12 " +
                                            "WHERE DC1.DC_ID = DC2.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC3.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC4.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC5.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC6.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC7.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC8.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC9.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC10.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC11.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC12.DC_ID(+) " +
                                           ")" +
                                        ") DISCOUNT, VAS_PRODUCT " +
                                    "WHERE DISCOUNT.DC_GROUPID = VAS_PRODUCT.VAS_CODE AND VAS_CHANNEL = 'VAS_SMARTUI' AND SALE_CHANNEL<> 'SS' " +
                                    "AND DC_ID IN (" + lstID + ") " +
                                    "union " +
                                    "SELECT DC_ID VAS_ID, VAS_CODE, VAS_NAME, VAS_PRICE, VAS_STATUS, VAS_RULE " +
                                    ",to_char(trunc(DC_START_DT), 'dd/mm/yyyy') START_DATE, to_char(trunc(DC_END_DT), 'dd/mm/yyyy') END_DATE " +
                                    ",VAS_CHANNEL,SALE_CHANNEL,PROMOTION_CODE,ORDER_TYPE,ALLOW_ADVANCE_MONTH,DOWNLOAD_FROM,DOWNLOAD_TO, " +
                                    "UPLOAD_FROM,UPLOAD_TO,PRICE_FROM,PRICE_TO,PROVINCE,PRODUCT,PARENT_VAS_CODE,VAS_TYPE " +
                                    "FROM( " +
                                        "SELECT * FROM " +
                                            "(SELECT DC1.DC_START_DT, DC1.DC_END_DT, DC1.DC_ID, DC1.DC_GROUPID " +
                                                ", NVL(PRODUCT, 'ALL') PRODUCT, NVL(PROMOTION_CODE, 'ALL') PROMOTION_CODE " +
                                                ", NVL(ORDER_TYPE, 'ALL') ORDER_TYPE, NVL(PROVINCE, 'ALL') PROVINCE " +
                                                ", NVL(SALE_CHANNEL, 'ALL') SALE_CHANNEL, NVL(ALLOW_ADVANCE_MONTH, 'ALL') ALLOW_ADVANCE_MONTH " +
                                                ", NVL(DOWNLOAD_FROM, 'ALL') DOWNLOAD_FROM, NVL(DOWNLOAD_TO, 'ALL') DOWNLOAD_TO " +
                                                ", NVL(UPLOAD_FROM, 'ALL') UPLOAD_FROM, NVL(UPLOAD_TO, 'ALL') UPLOAD_TO " +
                                                ", NVL(PRICE_FROM, 'ALL') PRICE_FROM, NVL(PRICE_TO, 'ALL') PRICE_TO " +
                                                "FROM " +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRODUCT', DC_VALUE, 'ALL') PRODUCT " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PRODUCT' AND DC_ACTIVE_FLAG = 'Y') DC1," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PROMOTION_CODE', DC_VALUE, 'ALL') PROMOTION_CODE " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PROMOTION_CODE' AND DC_ACTIVE_FLAG = 'Y') DC2," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'ORDER_TYPE', DC_VALUE, 'ALL') ORDER_TYPE " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'ORDER_TYPE' AND DC_ACTIVE_FLAG = 'Y') DC3," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PROVINCE', DC_VALUE, 'ALL') PROVINCE " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PROVINCE' AND DC_ACTIVE_FLAG = 'Y' ) DC4," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'SALE_CHANNEL', DC_VALUE, 'ALL') SALE_CHANNEL " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'SALE_CHANNEL' AND DC_ACTIVE_FLAG = 'Y') DC5," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'ALLOW_ADVANCE_MONTH', DC_VALUE, 'ALL') ALLOW_ADVANCE_MONTH " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'ALLOW_ADVANCE_MONTH' AND DC_ACTIVE_FLAG = 'Y' ) DC6," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M' end DOWNLOAD_FROM " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'DL_FROM' AND DC_ACTIVE_FLAG = 'Y') DC7," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M' end DOWNLOAD_TO " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'DL_TO' AND DC_ACTIVE_FLAG = 'Y') DC8," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M' end UPLOAD_FROM " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'UL_FROM' AND DC_ACTIVE_FLAG = 'Y') DC9," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M'  end UPLOAD_TO " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'UL_TO' AND DC_ACTIVE_FLAG = 'Y') DC10," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRICE_FROM', DC_VALUE, 'ALL') PRICE_FROM " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PRICE_FROM' AND DC_ACTIVE_FLAG = 'Y' ) DC11," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRICE_TO', DC_VALUE, 'ALL') PRICE_TO " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PRICE_TO' AND DC_ACTIVE_FLAG = 'Y') DC12 " +
                                                "WHERE DC1.DC_ID = DC2.DC_ID(+) " +
                                                    "AND DC1.DC_ID = DC3.DC_ID(+) " +
                                                    "AND DC1.DC_ID = DC4.DC_ID(+) " +
                                                    "AND DC1.DC_ID = DC5.DC_ID(+) " +
                                                    "AND DC1.DC_ID = DC6.DC_ID(+)" +
                                                    "AND DC1.DC_ID = DC7.DC_ID(+)" +
                                                    "AND DC1.DC_ID = DC8.DC_ID(+)" +
                                                    "AND DC1.DC_ID = DC9.DC_ID(+)" +
                                                    "AND DC1.DC_ID = DC10.DC_ID(+)" +
                                                    "AND DC1.DC_ID = DC11.DC_ID(+)" +
                                                    "AND DC1.DC_ID = DC12.DC_ID(+)" +
                                             ")) DISCOUNT, VAS_PRODUCT " +
                                    "WHERE DISCOUNT.DC_GROUPID = VAS_PRODUCT.VAS_CODE AND VAS_CHANNEL<> 'VAS_SMARTUI' AND SALE_CHANNEL = 'SS' " +
                                    "AND DC_ID IN (" + lstID + ")";

                    string sqlUpdate = "SELECT DC_ID VAS_ID, VAS_CODE, VAS_NAME, VAS_PRICE, VAS_STATUS, VAS_RULE" +
                                 ",to_char(trunc(DC_START_DT),'dd/mm/yyyy') START_DATE, to_char(trunc(DC_END_DT),'dd/mm/yyyy') END_DATE" +
                                 ",VAS_CHANNEL,SALE_CHANNEL,PROMOTION_CODE,ORDER_TYPE,ALLOW_ADVANCE_MONTH,DOWNLOAD_FROM,DOWNLOAD_TO" +
                                 ",UPLOAD_FROM,UPLOAD_TO,PRICE_FROM,PRICE_TO,PROVINCE,PRODUCT,PARENT_VAS_CODE,VAS_TYPE " +
                                 "FROM ( " +
                                    " SELECT * FROM " +
                                         "(SELECT DC1.DC_START_DT, DC1.DC_END_DT, DC1.DC_ID, DC1.DC_GROUPID " +
                                            ", NVL(PRODUCT, 'ALL') PRODUCT, NVL(PROMOTION_CODE, 'ALL') PROMOTION_CODE" +
                                            ", NVL(ORDER_TYPE, 'ALL') ORDER_TYPE, NVL(PROVINCE, 'ALL') PROVINCE" +
                                            ", NVL(SALE_CHANNEL, 'ALL') SALE_CHANNEL, NVL(ALLOW_ADVANCE_MONTH, 'ALL')ALLOW_ADVANCE_MONTH " +
                                            ", NVL(DOWNLOAD_FROM, 'ALL') DOWNLOAD_FROM, NVL(DOWNLOAD_TO, 'ALL') DOWNLOAD_TO" +
                                            ", NVL(UPLOAD_FROM, 'ALL') UPLOAD_FROM, NVL(UPLOAD_TO, 'ALL') UPLOAD_TO" +
                                            ", NVL(PRICE_FROM, 'ALL') PRICE_FROM, NVL(PRICE_TO, 'ALL') PRICE_TO " +
                                            "FROM " +
                                               "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRODUCT', DC_VALUE, 'ALL') PRODUCT " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PRODUCT'  AND DC_ACTIVE_FLAG = 'Y') DC1," +
                                               "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PROMOTION_CODE', DC_VALUE, 'ALL') PROMOTION_CODE " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PROMOTION_CODE' AND DC_ACTIVE_FLAG = 'Y') DC2," +
                                               "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'ORDER_TYPE', DC_VALUE, 'ALL') ORDER_TYPE " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'ORDER_TYPE' AND DC_ACTIVE_FLAG = 'Y') DC3," +
                                               "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PROVINCE', DC_VALUE, 'ALL') PROVINCE " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PROVINCE' AND DC_ACTIVE_FLAG = 'Y') DC4," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'SALE_CHANNEL', DC_VALUE, 'ALL') SALE_CHANNEL " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'SALE_CHANNEL' AND DC_ACTIVE_FLAG = 'Y') DC5," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'ALLOW_ADVANCE_MONTH', DC_VALUE, 'ALL') ALLOW_ADVANCE_MONTH " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'ALLOW_ADVANCE_MONTH' AND DC_ACTIVE_FLAG = 'Y' ) DC6," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M' end DOWNLOAD_FROM " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'DL_FROM' AND DC_ACTIVE_FLAG = 'Y') DC7," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M' end DOWNLOAD_TO " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'DL_TO' AND DC_ACTIVE_FLAG = 'Y') DC8," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M' end UPLOAD_FROM " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'UL_FROM' AND DC_ACTIVE_FLAG = 'Y') DC9," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M'  end UPLOAD_TO " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'UL_TO' AND DC_ACTIVE_FLAG = 'Y')  DC10," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRICE_FROM', DC_VALUE, 'ALL') PRICE_FROM " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PRICE_FROM' AND DC_ACTIVE_FLAG = 'Y') DC11," +
                                                "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRICE_TO', DC_VALUE, 'ALL') PRICE_TO " +
                                                    "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PRICE_TO' AND DC_ACTIVE_FLAG = 'Y') DC12 " +
                                            "WHERE DC1.DC_ID = DC2.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC3.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC4.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC5.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC6.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC7.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC8.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC9.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC10.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC11.DC_ID(+) " +
                                            "AND DC1.DC_ID = DC12.DC_ID(+) " +
                                           ")" +
                                        ") DISCOUNT, VAS_PRODUCT " +
                                    "WHERE DISCOUNT.DC_GROUPID = VAS_PRODUCT.VAS_CODE AND VAS_CHANNEL = 'VAS_SMARTUI' AND SALE_CHANNEL<> 'SS' " +
                                    "AND DC_ID IN (" + lstUpdateID + ") " +
                                    "union " +
                                    "SELECT DC_ID VAS_ID, VAS_CODE, VAS_NAME, VAS_PRICE, VAS_STATUS, VAS_RULE " +
                                    ",to_char(trunc(DC_START_DT), 'dd/mm/yyyy') START_DATE, to_char(trunc(DC_END_DT), 'dd/mm/yyyy') END_DATE " +
                                    ",VAS_CHANNEL,SALE_CHANNEL,PROMOTION_CODE,ORDER_TYPE,ALLOW_ADVANCE_MONTH,DOWNLOAD_FROM,DOWNLOAD_TO, " +
                                    "UPLOAD_FROM,UPLOAD_TO,PRICE_FROM,PRICE_TO,PROVINCE,PRODUCT,PARENT_VAS_CODE,VAS_TYPE " +
                                    "FROM( " +
                                        "SELECT * FROM " +
                                            "(SELECT DC1.DC_START_DT, DC1.DC_END_DT, DC1.DC_ID, DC1.DC_GROUPID " +
                                                ", NVL(PRODUCT, 'ALL') PRODUCT, NVL(PROMOTION_CODE, 'ALL') PROMOTION_CODE " +
                                                ", NVL(ORDER_TYPE, 'ALL') ORDER_TYPE, NVL(PROVINCE, 'ALL') PROVINCE " +
                                                ", NVL(SALE_CHANNEL, 'ALL') SALE_CHANNEL, NVL(ALLOW_ADVANCE_MONTH, 'ALL') ALLOW_ADVANCE_MONTH " +
                                                ", NVL(DOWNLOAD_FROM, 'ALL') DOWNLOAD_FROM, NVL(DOWNLOAD_TO, 'ALL') DOWNLOAD_TO " +
                                                ", NVL(UPLOAD_FROM, 'ALL') UPLOAD_FROM, NVL(UPLOAD_TO, 'ALL') UPLOAD_TO " +
                                                ", NVL(PRICE_FROM, 'ALL') PRICE_FROM, NVL(PRICE_TO, 'ALL') PRICE_TO " +
                                                "FROM " +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRODUCT', DC_VALUE, 'ALL') PRODUCT " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PRODUCT' AND DC_ACTIVE_FLAG = 'Y') DC1," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PROMOTION_CODE', DC_VALUE, 'ALL') PROMOTION_CODE " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PROMOTION_CODE' AND DC_ACTIVE_FLAG = 'Y') DC2," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'ORDER_TYPE', DC_VALUE, 'ALL') ORDER_TYPE " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'ORDER_TYPE' AND DC_ACTIVE_FLAG = 'Y') DC3," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PROVINCE', DC_VALUE, 'ALL') PROVINCE " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PROVINCE' AND DC_ACTIVE_FLAG = 'Y' ) DC4," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'SALE_CHANNEL', DC_VALUE, 'ALL') SALE_CHANNEL " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'SALE_CHANNEL' AND DC_ACTIVE_FLAG = 'Y') DC5," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'ALLOW_ADVANCE_MONTH', DC_VALUE, 'ALL') ALLOW_ADVANCE_MONTH " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'ALLOW_ADVANCE_MONTH' AND DC_ACTIVE_FLAG = 'Y' ) DC6," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M' end DOWNLOAD_FROM " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'DL_FROM' AND DC_ACTIVE_FLAG = 'Y') DC7," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M' end DOWNLOAD_TO " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'DL_TO' AND DC_ACTIVE_FLAG = 'Y') DC8," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M' end UPLOAD_FROM " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'UL_FROM' AND DC_ACTIVE_FLAG = 'Y') DC9," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, dc_value, CASE when dc_value = 'ALL' THEN  'ALL' ELSE to_char(to_number(dc_value)/ 1024)|| 'M'  end UPLOAD_TO " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'UL_TO' AND DC_ACTIVE_FLAG = 'Y') DC10," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRICE_FROM', DC_VALUE, 'ALL') PRICE_FROM " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PRICE_FROM' AND DC_ACTIVE_FLAG = 'Y' ) DC11," +
                                                    "(SELECT DC_START_DT, DC_END_DT, DC_ID, DC_GROUPID, DECODE(DC_TYPE, 'PRICE_TO', DC_VALUE, 'ALL') PRICE_TO " +
                                                        "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE = 'PRICE_TO' AND DC_ACTIVE_FLAG = 'Y') DC12 " +
                                                "WHERE DC1.DC_ID = DC2.DC_ID(+) " +
                                                    "AND DC1.DC_ID = DC3.DC_ID(+) " +
                                                    "AND DC1.DC_ID = DC4.DC_ID(+) " +
                                                    "AND DC1.DC_ID = DC5.DC_ID(+) " +
                                                    "AND DC1.DC_ID = DC6.DC_ID(+)" +
                                                    "AND DC1.DC_ID = DC7.DC_ID(+)" +
                                                    "AND DC1.DC_ID = DC8.DC_ID(+)" +
                                                    "AND DC1.DC_ID = DC9.DC_ID(+)" +
                                                    "AND DC1.DC_ID = DC10.DC_ID(+)" +
                                                    "AND DC1.DC_ID = DC11.DC_ID(+)" +
                                                    "AND DC1.DC_ID = DC12.DC_ID(+)" +
                                             ")) DISCOUNT, VAS_PRODUCT " +
                                    "WHERE DISCOUNT.DC_GROUPID = VAS_PRODUCT.VAS_CODE AND VAS_CHANNEL<> 'VAS_SMARTUI' AND SALE_CHANNEL = 'SS' " +
                                    "AND DC_ID IN (" + lstUpdateID + ")";

                    string sqlPriceandBill = "SELECT distinct SIEBEL.TP_DAT,SIEBEL.MKT_CODE MAIN_OFFER , SIEBEL.NAME MAIN_OFFER_NAME, SIEBEL.SPEED,SIEBEL.BIL_FREQ,FE.PRICE_HISPEED PRICE_NULL_IS_599 " +
                                             "from " +
                                                "(select DISTINCT prod_x.attrib_04 mkt_code, prod.name, prod_x.attrib_34 speed, adsl.* " +
                                                "from( " +
                                                        "select a.row_id, a.ref_number_3 SRV_CD, b.internet_package SRV_CD_ROWID, b.adsl_promotion MKT_ROWID, PP.* " +
                                                        "from SIEBEL.S_PROD_INT A,SIEBEL.CX_ADSL_INTERNET_PACKAGES_X B, " +
                                                                "(select c.row_id PP_ROW_ID, c.sef_cd, c.tp_dat, c.bil_freq " +
                                                                "from siebel.cx_price_plan_map_x c " +
                                                                "where TP_DAT IN('TRUFTTX') and rc_oc_uc_flg = 'R' and active_flag = 'Y' " +
                                                                "and company_code = 'TI' and ccbs_flag = 'Y' and bn_flag = 'Y' and csm_param_1 is not null) PP " +
                                                        "where a.ref_number_3 = pp.SEF_CD and a.row_id = b.internet_package)ADSL,SIEBEL.S_PROD_INT PROD,SIEBEL.S_PROD_INT_X PROD_X " +
                                                "where adsl.MKT_ROWID = prod.row_id and prod.row_id = prod_x.row_id) SIEBEL, " +
                                                   "(select distinct a.p_id, a.status,  a.p_code Propo, a.p_code || '-' || suffix PROMOTION ,a.p_name,b.price PRICE_HISPEED, a.order_type ,A.PRODTYPE " +
                                                   "from hispeed_promotion a, hispeed_speed_promotion b where a.p_id = b.p_id and a.status in ('Active', 'Pending') " +
                                                   "and UPPER(a.order_type) in ('NEW', 'CHANGE') and a.prodtype in ('HISPEED_FBTH_NF', 'FIBER_TO_HOME')) FE " +
                                             "where SIEBEL.MKT_CODE = FE.PROMOTION " +
                                             "UNION " +
                                             "SELECT * FROM " +
                                                "(SELECT distinct SIEBEL.TP_DAT, SIEBEL.MKT_CODE MAIN_OFFER, SIEBEL.NAME MAIN_OFFER_NAME, SIEBEL.SPEED, SIEBEL.BIL_FREQ, TO_NUMBER (NULL) PRICE_NULL_IS_599 " +
                                                "from " +
                                                    "(select DISTINCT prod_x.attrib_04 mkt_code, prod.name, prod_x.attrib_34 speed, adsl.* " +
                                                    "from " +
                                                        "(select a.row_id, a.ref_number_3 SRV_CD, b.internet_package SRV_CD_ROWID, b.adsl_promotion MKT_ROWID, PP.* " +
                                                        "from SIEBEL.S_PROD_INT A,SIEBEL.CX_ADSL_INTERNET_PACKAGES_X B,(select c.row_id PP_ROW_ID, c.sef_cd, c.tp_dat, c.bil_freq from siebel.cx_price_plan_map_x c " +
                                                    "where TP_DAT IN ('TRUFTTX') and rc_oc_uc_flg = 'R' and active_flag = 'Y' and company_code = 'TI' and ccbs_flag = 'Y' and bn_flag = 'Y' and csm_param_1 is not null) PP " +
                                                "where a.ref_number_3 = pp.SEF_CD   and a.row_id = b.internet_package)ADSL,SIEBEL.S_PROD_INT PROD,SIEBEL.S_PROD_INT_X PROD_X " +
                                             "where adsl.MKT_ROWID = prod.row_id and prod.row_id = prod_x.row_id) SIEBEL " +
                                             "MINUS " +
                                             "SELECT distinct  SIEBEL.TP_DAT,SIEBEL.MKT_CODE MAIN_OFFER, SIEBEL.NAME MAIN_OFFER_NAME, SIEBEL.SPEED,SIEBEL.BIL_FREQ,TO_NUMBER(NULL) PRICE_NULL_IS_599 " +
                                             "from " +
                                                "(select  DISTINCT prod_x.attrib_04 mkt_code, prod.name, prod_x.attrib_34 speed, adsl.* " +
                                                "from(select a.row_id, a.ref_number_3 SRV_CD, b.internet_package SRV_CD_ROWID, b.adsl_promotion MKT_ROWID, PP.* " +
                                                      "from SIEBEL.S_PROD_INT A,SIEBEL.CX_ADSL_INTERNET_PACKAGES_X B,(select c.row_id PP_ROW_ID, c.sef_cd, c.tp_dat, c.bil_freq from siebel.cx_price_plan_map_x c " +
                                                "WHERE TP_DAT IN('TRUFTTX') and rc_oc_uc_flg = 'R' and active_flag = 'Y' and company_code = 'TI' and ccbs_flag = 'Y' and bn_flag = 'Y' and csm_param_1 is not null) PP " +
                                             "where a.ref_number_3 = pp.SEF_CD and a.row_id = b.internet_package)ADSL,SIEBEL.S_PROD_INT PROD,SIEBEL.S_PROD_INT_X PROD_X " +
                                             "where adsl.MKT_ROWID = prod.row_id and prod.row_id = prod_x.row_id) SIEBEL, " +
                                             "(select distinct a.p_id, a.status,  a.p_code Propo, a.p_code || '-' || suffix PROMOTION ,a.p_name,b.price PRICE_HISPEED, a.order_type ,A.PRODTYPE " +
                                             "from hispeed_promotion a, hispeed_speed_promotion b where a.p_id = b.p_id and UPPER(a.order_type) in ('NEW', 'CHANGE') and a.prodtype in ('HISPEED_FBTH_NF', 'FIBER_TO_HOME')) FE " +
                                             "WHERE SIEBEL.MKT_CODE = FE.PROMOTION)";

                    string sqlMKTNotAllow = "SELECT * FROM TMP_CONFLICTS_FEATURE WHERE FEATURE_CODE IN(" + lstCodeforOffer + ") AND CONFLICT_CODE IN(" + lstOffer + ") ORDER BY FEATURE_CODE,CONFLICT_CODE,ACTIVE";

                    toolStripProgressBar1.Value = 20;

                    if (String.IsNullOrEmpty(lstID) == false)
                    {
                        Excel.Worksheet sheet1 = workbook.ActiveSheet as Excel.Worksheet;
                        sheet1.Name = "VAS_Eligibility";
                        sheet1.get_Range("A1", "W1").Interior.Color = Excel.XlRgbColor.rgbLightSeaGreen;
                        sheet1.get_Range("A1", "W1").Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                        WriteData(sheet1, sqlEliRule);

                        /* Excel.Worksheet sheet2 = workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
                         sheet2.Name = "MainOffer(Price&BillFreq)";
                         sheet2.get_Range("A1", "F1").Interior.Color = Excel.XlRgbColor.rgbAqua;
                         sheet2.get_Range("A1", "F1").Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                         WriteData(sheet2, sqlPriceandBill);*/
                    }
                    else
                    {
                        if (String.IsNullOrEmpty(lstCode) == false)
                        {
                            Excel.Worksheet sheet1 = workbook.ActiveSheet as Excel.Worksheet;
                            sheet1.Name = "VAS_Product";
                            sheet1.get_Range("A1", "P1").Interior.Color = Excel.XlRgbColor.rgbLightSeaGreen;
                            sheet1.get_Range("A1", "P1").Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                            WriteData(sheet1, sqlVasProd);
                        }
                    }

                    if (String.IsNullOrEmpty(lstCodeforOffer) == false)
                    {
                        Excel.Worksheet sheet;
                        if (workbook.Sheets.Count >= 1)
                        {
                            sheet = workbook.ActiveSheet as Excel.Worksheet;
                        }
                        else
                        {
                            sheet = workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
                        }
                        
                        sheet.Name = "MainOffer_NotAllow";
                        sheet.get_Range("A1", "C1").Interior.Color = Excel.XlRgbColor.rgbLightSeaGreen;
                        sheet.get_Range("A1", "C1").Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                        WriteData(sheet, sqlMKTNotAllow);

                        flagNotAllow = false;
                    }

                    if (String.IsNullOrEmpty(lstUpdateID) == false)
                    {
                        Excel.Worksheet sheet;
                        if (workbook.Sheets.Count >= 1)
                        {
                            sheet = workbook.ActiveSheet as Excel.Worksheet;
                        }
                        else
                        {
                            sheet = workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
                        }

                        sheet.Name = "VAS_Eligibility(Update)";
                        sheet.get_Range("A1", "W1").Interior.Color = Excel.XlRgbColor.rgbAqua;
                        sheet.get_Range("A1", "W1").Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                        WriteData(sheet, sqlUpdate);
                    }

                    workbook.Sheets[1].Activate();
                    string path = outputPath + "\\VAS_Criteria" + urNo.ToUpper() + ".xlsx";
                    workbook.SaveAs(path, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    toolStripStatusLabel1.Text = "Already exported file";
                    toolStripProgressBar1.Value = 100;
                    MessageBox.Show("Criteria file has been exported successfully.", "Success", MessageBoxButtons.OK);
                }
                catch (Exception ex)
                {
                    toolStripProgressBar1.Value = 0;
                    toolStripStatusLabel1.Text = "Failed to export data";
                    MessageBox.Show("An error occurred while exporting file criteria." + "\r\n" + "Detail of System : " + ex.Message
                        , "Failed to export", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    workbook.Close();
                    xlApp.Quit();

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    Cursor.Current = Cursors.Default;
                    Application.UseWaitCursor = false;
                }
            }
        }

        private void WriteData(Excel.Worksheet sheet, string query)
        {
            toolStripStatusLabel1.Text = "Writing Sheet[" + sheet.Name + "]...";

            OracleDataAdapter adapter = new OracleDataAdapter(query, ConnectionProd);
            DataTable dt = new DataTable();
            adapter.Fill(dt);

            toolStripProgressBar1.Value = 10;

            //Set column heading
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
            }

            toolStripProgressBar1.Value = 20;

            //Write data
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    DateTime date;
                    if (DateTime.TryParse(dt.Rows[i][j].ToString(), out date))
                    {
                        sheet.Cells[i + 2, j + 1] = string.Format("{0:dd/MMM/yyyy}", date);
                    }
                    else
                    {
                        sheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                    }
                }

                toolStripProgressBar1.Value = (i + 1) * 80 / dt.Rows.Count;
            }
            adapter.Dispose();
        }

        private void hilightRow(string type, string key, int indexRow, DataGridView gridView)
        {
            Dictionary<string, int> indexProd = new Dictionary<string, int>
            {{"code",0 },{"desc",1},{"type",2},{"rule",3},{"price",4},{"channel",5 },{"group",6},{"start",7}};

            Dictionary<string, int> indexRules = new Dictionary<string, int>
            {{"code",0 },{"channel",4},{"offer",5},{"order",6},{"product",7},{"province",8},{"month",9},
            {"downF",10},{"downT",11},{"upF",12},{"upT",13},{"priceF",15},{"priceT",16},{"start",17},{"end",18}};

            Dictionary<string, int> indexMKT = new Dictionary<string, int>
            {{"code",0 },{"mkt",1},{"flag",2}};

            Dictionary<string, int> indexUpdate = new Dictionary<string, int>
            {{"id",0 },{"start",1},{"end",2}};

            if (type == "prod")
            {
                int indexCol = indexProd[key];
                gridView.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }
            else if(type == "rule")
            {
                int indexCol = indexRules[key];
                gridView.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }
            else if(type == "notAllow")
            {
                int indexCol = indexMKT[key];
                gridView.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }           
            else
            {
                int indexCol = indexUpdate[key];
                gridView.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }
        }

        private void InitialValue(DataGridView dataGrid)
        {
            //Clear selection
            dataGrid.ClearSelection();

            for (int i = 0; i < dataGrid.RowCount; i++)
            {
                for (int j = 0; j < dataGrid.ColumnCount; j++)
                {
                    dataGrid.Rows[i].Cells[j].Style.BackColor = Color.Empty;
                }

                dataGrid.Rows[i].DefaultCellStyle.BackColor = Color.Empty;
            }

            //Clear list index
            indexListbox.Clear();
            //Clear listbox
            listBox1.Items.Clear();

            validateLog = "";
        }
    }

    class ExportScript
    {
        private string _elirule;
        private string _offerNotAllow;
        private string _priceAndBill;
        private string _vasProduct;

        public string VasEligibility(string id)
        {
            _elirule = "SELECT DC_ID VAS_ID,to_char(trunc(DC_START_DT),'dd/mm/yyyy') START_DATE,to_char(trunc(DC_END_DT),'dd/mm/yyyy') END_DATE" +
",VAS_CODE,VAS_NAME,VAS_PRICE,VAS_STATUS,VAS_RULE,VAS_CHANNEL,SALE_CHANNEL,PROMOTION_CODE,ORDER_TYPE,ALLOW_ADVANCE_MONTH" +
",DOWNLOAD_FROM,DOWNLOAD_TO,UPLOAD_FROM,UPLOAD_TO,PRICE_FROM,PRICE_TO,PROVINCE,PRODUCT,PARENT_VAS_CODE,VAS_TYPE " +
"FROM ( " +
    "SELECT * FROM " +
            "( SELECT DC1.DC_START_DT , DC1.DC_END_DT ,DC1.DC_ID,DC1.DC_GROUPID" +
                ",NVL (PRODUCT,'ALL') PRODUCT,NVL (PROMOTION_CODE,'ALL')PROMOTION_CODE ,NVL (ORDER_TYPE,'ALL')ORDER_TYPE" +
                ",NVL (PROVINCE,'ALL')PROVINCE ,NVL (SALE_CHANNEL,'ALL')SALE_CHANNEL ,NVL (ALLOW_ADVANCE_MONTH,'ALL')ALLOW_ADVANCE_MONTH" +
                ",NVL (DOWNLOAD_FROM,'ALL')DOWNLOAD_FROM ,NVL (DOWNLOAD_TO,'ALL')DOWNLOAD_TO ,NVL (UPLOAD_FROM,'ALL')UPLOAD_FROM ,NVL (UPLOAD_TO,'ALL')UPLOAD_TO" +
                ",NVL (PRICE_FROM,'ALL')PRICE_FROM ,NVL (PRICE_TO,'ALL')PRICE_TO " +
                    "FROM " +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'PRODUCT',DC_VALUE,'ALL') PRODUCT " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='PRODUCT' AND DC_ACTIVE_FLAG='Y') DC1," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'PROMOTION_CODE',DC_VALUE,'ALL') PROMOTION_CODE " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='PROMOTION_CODE' AND DC_ACTIVE_FLAG='Y' ) DC2," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'ORDER_TYPE',DC_VALUE,'ALL') ORDER_TYPE " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='ORDER_TYPE' AND DC_ACTIVE_FLAG='Y') DC3," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'PROVINCE',DC_VALUE,'ALL') PROVINCE " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='PROVINCE' AND DC_ACTIVE_FLAG='Y') DC4," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'SALE_CHANNEL',DC_VALUE,'ALL') SALE_CHANNEL " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='SALE_CHANNEL' AND DC_ACTIVE_FLAG='Y') DC5," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'ALLOW_ADVANCE_MONTH',DC_VALUE,'ALL') ALLOW_ADVANCE_MONTH " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='ALLOW_ADVANCE_MONTH' AND DC_ACTIVE_FLAG='Y') DC6," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,dc_value,CASE when dc_value = 'ALL' THEN 'ALL' ELSE to_char(to_number(dc_value)/1024)||'M' end DOWNLOAD_FROM " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='DL_FROM' AND DC_ACTIVE_FLAG='Y') DC7," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,dc_value,CASE when dc_value = 'ALL' THEN 'ALL' ELSE to_char(to_number(dc_value)/1024)||'M' end DOWNLOAD_TO " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='DL_TO' AND DC_ACTIVE_FLAG='Y') DC8," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,dc_value,CASE when dc_value = 'ALL' THEN 'ALL' ELSE to_char(to_number(dc_value)/1024)||'M' end UPLOAD_FROM " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='UL_FROM' AND DC_ACTIVE_FLAG='Y') DC9," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,dc_value,CASE when dc_value = 'ALL' THEN 'ALL' ELSE to_char(to_number(dc_value)/1024)||'M' end UPLOAD_TO " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='UL_TO' AND DC_ACTIVE_FLAG='Y') DC10," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'PRICE_FROM',DC_VALUE,'ALL') PRICE_FROM " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='PRICE_FROM' AND DC_ACTIVE_FLAG='Y') DC11," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'PRICE_TO',DC_VALUE,'ALL') PRICE_TO " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='PRICE_TO' AND DC_ACTIVE_FLAG='Y') DC12 " +
                    "WHERE DC1.DC_ID = DC2.DC_ID (+) " +
                    "AND DC1.DC_ID = DC3.DC_ID (+) " +
                    "AND DC1.DC_ID = DC4.DC_ID (+) " +
                    "AND DC1.DC_ID = DC5.DC_ID (+) " +
                    "AND DC1.DC_ID = DC6.DC_ID (+) " +
                    "AND DC1.DC_ID = DC7.DC_ID (+) " +
                    "AND DC1.DC_ID = DC8.DC_ID (+) " +
                    "AND DC1.DC_ID = DC9.DC_ID (+) " +
                    "AND DC1.DC_ID = DC10.DC_ID (+) " +
                    "AND DC1.DC_ID = DC11.DC_ID (+) " +
                    "AND DC1.DC_ID = DC12.DC_ID (+) " +
            ") " +
    ") DISCOUNT, VAS_PRODUCT " +
"WHERE DISCOUNT.DC_GROUPID = VAS_PRODUCT.VAS_CODE " +
"AND DC_ID > 'VAS1000000' " +
"AND VAS_CHANNEL = 'VAS_SMARTUI' " +
"AND SALE_CHANNEL <> 'SS' " +
"AND DC_ID IN (" + id + ") " +
"union " +
"SELECT DC_ID VAS_ID,to_char(trunc(DC_START_DT),'dd/mm/yyyy') START_DATE,to_char(trunc(DC_END_DT),'dd/mm/yyyy') END_DATE" +
",VAS_CODE,VAS_NAME,VAS_PRICE,VAS_STATUS,VAS_RULE,VAS_CHANNEL,SALE_CHANNEL,PROMOTION_CODE,ORDER_TYPE,ALLOW_ADVANCE_MONTH" +
",DOWNLOAD_FROM,DOWNLOAD_TO,UPLOAD_FROM,UPLOAD_TO,PRICE_FROM,PRICE_TO,PROVINCE,PRODUCT,PARENT_VAS_CODE,VAS_TYPE " +
"FROM ( " +
     "SELECT * FROM " +
            "( SELECT DC1.DC_START_DT , DC1.DC_END_DT ,DC1.DC_ID,DC1.DC_GROUPID" +
                ",NVL (PRODUCT,'ALL') PRODUCT,NVL (PROMOTION_CODE,'ALL')PROMOTION_CODE ,NVL (ORDER_TYPE,'ALL')ORDER_TYPE " +
                ",NVL (PROVINCE,'ALL')PROVINCE ,NVL (SALE_CHANNEL,'ALL')SALE_CHANNEL ,NVL (ALLOW_ADVANCE_MONTH,'ALL')ALLOW_ADVANCE_MONTH " +
                ",NVL (DOWNLOAD_FROM,'ALL')DOWNLOAD_FROM ,NVL (DOWNLOAD_TO,'ALL')DOWNLOAD_TO ,NVL (UPLOAD_FROM,'ALL')UPLOAD_FROM ,NVL (UPLOAD_TO,'ALL')UPLOAD_TO " +
                ",NVL (PRICE_FROM,'ALL')PRICE_FROM ,NVL (PRICE_TO,'ALL')PRICE_TO " +
                    "FROM " +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'PRODUCT',DC_VALUE,'ALL') PRODUCT " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='PRODUCT'  AND DC_ACTIVE_FLAG='Y') DC1," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'PROMOTION_CODE',DC_VALUE,'ALL') PROMOTION_CODE " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='PROMOTION_CODE' AND DC_ACTIVE_FLAG='Y') DC2," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'ORDER_TYPE',DC_VALUE,'ALL') ORDER_TYPE " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='ORDER_TYPE' AND DC_ACTIVE_FLAG='Y') DC3," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'PROVINCE',DC_VALUE,'ALL') PROVINCE " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='PROVINCE' AND DC_ACTIVE_FLAG='Y') DC4," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'SALE_CHANNEL',DC_VALUE,'ALL') SALE_CHANNEL " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='SALE_CHANNEL' AND DC_ACTIVE_FLAG='Y') DC5," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'ALLOW_ADVANCE_MONTH',DC_VALUE,'ALL') ALLOW_ADVANCE_MONTH " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='ALLOW_ADVANCE_MONTH' AND DC_ACTIVE_FLAG='Y') DC6," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,dc_value,CASE when dc_value = 'ALL' THEN 'ALL' ELSE to_char(to_number(dc_value)/1024)||'M' end DOWNLOAD_FROM " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='DL_FROM' AND DC_ACTIVE_FLAG='Y') DC7," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,dc_value,CASE when dc_value = 'ALL' THEN 'ALL' ELSE to_char(to_number(dc_value)/1024)||'M' end DOWNLOAD_TO " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='DL_TO' AND DC_ACTIVE_FLAG='Y') DC8," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,dc_value,CASE when dc_value = 'ALL' THEN 'ALL' ELSE to_char(to_number(dc_value)/1024)||'M' end UPLOAD_FROM " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='UL_FROM' AND DC_ACTIVE_FLAG='Y') DC9," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,dc_value,CASE when dc_value = 'ALL' THEN 'ALL' ELSE to_char(to_number(dc_value)/1024)||'M' end UPLOAD_TO " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='UL_TO' AND DC_ACTIVE_FLAG='Y') DC10," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'PRICE_FROM',DC_VALUE,'ALL') PRICE_FROM " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='PRICE_FROM' AND DC_ACTIVE_FLAG='Y') DC11," +
                        "(SELECT DC_START_DT,DC_END_DT,DC_ID,DC_GROUPID,DECODE (DC_TYPE,'PRICE_TO',DC_VALUE,'ALL') PRICE_TO " +
                            "FROM DISCOUNT_CRITERIA_MAPPING WHERE DC_TYPE='PRICE_TO' AND DC_ACTIVE_FLAG='Y') DC12 " +
                    "WHERE DC1.DC_ID = DC2.DC_ID (+) " +
                    "AND DC1.DC_ID = DC3.DC_ID (+) " +
                    "AND DC1.DC_ID = DC4.DC_ID (+) " +
                    "AND DC1.DC_ID = DC5.DC_ID (+) " +
                    "AND DC1.DC_ID = DC6.DC_ID (+) " +
                    "AND DC1.DC_ID = DC7.DC_ID (+) " +
                    "AND DC1.DC_ID = DC8.DC_ID (+) " +
                    "AND DC1.DC_ID = DC9.DC_ID (+) " +
                    "AND DC1.DC_ID = DC10.DC_ID (+) " +
                    "AND DC1.DC_ID = DC11.DC_ID (+) " +
                    "AND DC1.DC_ID = DC12.DC_ID (+) " +
            ") " +
      ") DISCOUNT, VAS_PRODUCT " +
"WHERE DISCOUNT.DC_GROUPID = VAS_PRODUCT.VAS_CODE " +
"AND DC_ID > 'VAS1000000' " +
"AND VAS_CHANNEL <> 'VAS_SMARTUI' " +
"AND SALE_CHANNEL IN ( 'SS','ALL') " +
"AND DC_ID IN (" + id + ")";

            return _elirule;
        }
    }
}

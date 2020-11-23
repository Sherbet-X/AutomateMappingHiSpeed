using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.OracleClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutomateMapping
{
    public partial class MainHispeed : Form
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

        private string filename , fileDesc, implementer, urNo, outputPath, validateLog, logHispeed, 
            logCampaign, func, expHisp, expCamp, id, tolPack, tvsPack;
        /// <summary>
        /// Keep PName value
        /// Key : MKT, Value : Description
        /// </summary>
        Dictionary<string, string> lstPname = new Dictionary<string, string>();
        /// <summary>
        /// List channel from DB Master
        /// </summary>
        List<string[]> lstChannel = new List<string[]>();
        /// <summary>
        /// List SubProfile from DB Master
        /// </summary>
        List<string[]> lstSubProfile = new List<string[]>();
        /// <summary>
        /// List Speed from DB Master
        /// </summary>
        Dictionary<int, string[]> lstSpeedMast = new Dictionary<int, string[]>();
        /// <summary>
        /// List ExtraProfile from DB Master
        /// </summary>
        List<string[]> lstExtraProfile = new List<string[]>();
        /// <summary>
        /// List Contract from DB Master
        /// </summary>
        DataTable tableContract = new DataTable();
        /// <summary>
        /// List Prodtype from DB Master
        /// </summary>
        DataTable tableProdType = new DataTable();
        /// <summary>
        /// Keep index of listbox(log viewer)
        /// </summary>
        List<int> indexListbox = new List<int>();
        /// <summary>
        /// List of sheetName from file requirement
        /// </summary>
        List<string> sheets = new List<string>();
        /// <summary>
        /// Variable for move form
        /// </summary>
        int mov, movX, movY;
        /// <summary>
        /// Validation class
        /// </summary>
        Validation validation;
        #endregion

        #region "init"
        public MainHispeed(OracleConnection con, string file, string fDesc, string user, string ur, string fileOut)
        {
            InitializeComponent();
            ConnectionProd = con;
            filename = file;
            fileDesc = fDesc;
            implementer = user;
            urNo = ur;
            outputPath = fileOut;
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
        private void MainHispeed_Load(object sender, EventArgs e)
        {
            dataGridView1.Hide();
            Application.UseWaitCursor = true;
            Cursor.Current = Cursors.WaitCursor;

            toolStripStatusLabel1.ForeColor = Color.White;
            btnLog.Visible = true;
            btnExe.Enabled = true;

            double widthRatio = Screen.PrimaryScreen.Bounds.Width;
            double heightRatio = Screen.PrimaryScreen.Bounds.Height;

            //Different resolutions cause different screen display and widescreen cannot start maximize
            //Set default screen when starting first time
            if (widthRatio >= 1366 && heightRatio >= 768)
            {
                this.WindowState = FormWindowState.Maximized;
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
                this.Size = new Size((int)(widthRatio + 74), (int)(heightRatio + 16));
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }

            toolStripStatusLabel1.Text = "Loding Excel File...";

            #region "Set ToolTip"
            toolTip1.ShowAlways = true;
            toolTip1.SetToolTip(btnValidate, "Click to validate again");

            ToolTip toolTip2 = new ToolTip { ShowAlways = true };
            toolTip2.SetToolTip(btnHome, "Back to home page");

            ToolTip toolTip3 = new ToolTip { ShowAlways = true };
            toolTip3.SetToolTip(btnLogout, "Log out");
            #endregion

            try
            {
                ConnectionTemp = new OracleConnection();

                string connStringTmp = @"Data Source=(DESCRIPTION =(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.19.217.162)(PORT = 1559))) " +
                                    "(CONNECT_DATA =(SERVICE_NAME = CVMDEV)));User Id= EPCSUPUSR; Password=EPCSUPUSR_55;";

                ConnectionTemp.ConnectionString = connStringTmp;
                ConnectionTemp.Open();

                //Get all sheet name from excel file
                sheets = ToExcelsSheetList(filename);

                DgvSettings dgvSettings = new DgvSettings();
                List<string> lstHeader = new List<string>();

                if (sheets.Contains("HiSpeed Promotion"))
                {
                    //Set header view hispeed promotion
                    lstHeader.Add("Media");
                    lstHeader.Add("MKTCode");
                    lstHeader.Add("Speed");
                    lstHeader.Add("Sub Profile");
                    lstHeader.Add("Extra Profile");
                    lstHeader.Add("Price");
                    lstHeader.Add("Order Type");
                    lstHeader.Add("Channel");
                    lstHeader.Add("Modem Type");
                    lstHeader.Add("Docsis Type");
                    lstHeader.Add("Effective");
                    lstHeader.Add("Expire");
                    lstHeader.Add("Entry Code");
                    lstHeader.Add("Install Code");

                    dgvSettings.SetDgv(dataGridView1, filename, "HiSpeed Promotion$B3:O", lstHeader);
                    dataGridView1.Show();
                    func = "Hispeed";
                }
                else if(sheets.Contains("Campaign Mapping"))
                {
                    lstHeader.Clear();

                    lstHeader.Add("Type");
                    lstHeader.Add("Campaign Name");
                    lstHeader.Add("TOL Package");
                    lstHeader.Add("TOL Discount");
                    lstHeader.Add("TVS Package");
                    lstHeader.Add("TVS Discount");

                    dgvSettings.SetDgv(dataGridView1, filename, "Campaign Mapping$B2:G", lstHeader);
                    dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    dataGridView1.Show();
                    func = "Campaign";
                }
                else
                {
                    MessageBox.Show("Sheet name : 'HiSpeed Promotion' and 'Campaign Mapping' Not Found!!");
                    Application.Exit();
                }

                if (dataGridView1.AutoSizeColumnsMode != DataGridViewAutoSizeColumnsMode.Fill)
                {
                    dataGridView1.BackgroundColor = Color.White;
                }
                else
                {
                    dataGridView1.BackgroundColor = Color.FromArgb(79, 172, 254);
                }

                backgroundWorker1.RunWorkerAsync(func);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace, "Loding Data Failed!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                toolStripStatusLabel1.Text = "";
                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;
            }
        }

        private void MainHispeed_SizeChanged(object sender, EventArgs e)
        {
            int w = this.Size.Width;
            int h = this.Size.Height;

            btnClose.Location = new Point(w - 50, 0);
            btnMaximize.Location = new Point(btnClose.Location.X - 50, 0);
            btnMinimize.Location = new Point(btnMaximize.Location.X - 50, 0);
            btnLogout.Location = new Point(btnMinimize.Location.X - 50, 0);

            dataGridView1.Size = new Size(w, (int)(((h / 2) * 2) / 3));
            listBox1.Location = new Point(0, dataGridView1.Height + 161);
            listBox1.Size = new Size(w, (int)(((h / 2) * 2) / 3) - 45);

            btnValidate.Location = new Point(w - 39, listBox1.Location.Y - 31);
            btnHome.Location = new Point(w - 81, listBox1.Location.Y - 31);

            btnExe.Location = new Point(w - 152, listBox1.Location.Y + listBox1.Height + 30);
            btnLog.Location = new Point(btnExe.Location.X - 210, listBox1.Location.Y + listBox1.Height + 30);

            labelLogViewer.Location = new Point(12, listBox1.Location.Y - 45);
            labelHead.Location = new Point(58, 10);

            if (dataGridView1.AutoSizeColumnsMode != DataGridViewAutoSizeColumnsMode.Fill)
            {
                dataGridView1.BackgroundColor = Color.White;
            }
            else
            {
                dataGridView1.BackgroundColor = Color.FromArgb(79, 172, 254);
            }

            labelFunction.Location = new Point(40, dataGridView1.Location.Y - 43);
        }

        private void labelClose_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy)
            {
                backgroundWorker1.CancelAsync();
            }
            if (backgroundWorker2.IsBusy)
            {
                backgroundWorker2.CancelAsync();
            }
            if (backgroundWorker3.IsBusy)
            {
                backgroundWorker3.CancelAsync();
            }

            Environment.Exit(0);
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            this.Close();
            Login login = new Login();
            login.Show();
        }

        private void btnHome_Click(object sender, EventArgs e)
        {
            InputHispeed inputHispeed = new InputHispeed(ConnectionProd, implementer, urNo);
            this.Close();
            inputHispeed.Show();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Do you want to close this application?", "Automate Mapping Tool"
                , MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.OK)
            {
                Application.Exit();
            }
        }

        private void btnMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnMaximize_Click(object sender, EventArgs e)
        {
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            double widthRatio = Screen.PrimaryScreen.Bounds.Width;
            double heightRatio = Screen.PrimaryScreen.Bounds.Height;
            int w = Size.Width;

            if (this.WindowState != FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;

                if (widthRatio > 1366 && heightRatio > 768)
                {
                    this.Size = new Size(w, btnExe.Location.Y + 75);
                }
                else
                {
                    btnExe.Location = new Point(btnExe.Location.X, statusStrip1.Location.Y - 46);
                    btnLog.Location = new Point(btnLog.Location.X, statusStrip1.Location.Y - 46);
                }
            }
        }

        private void btnValidate_Click(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();
            dataGridView1.Update();
            btnExe.Enabled = true;
            btnLog.Visible = true;

            if (backgroundWorker1.IsBusy)
            {
                backgroundWorker1.CancelAsync();
            }

            backgroundWorker1.RunWorkerAsync(func);

            dataGridView1.Refresh();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string process = e.Argument.ToString();

            if (process == "Hispeed")
            {
                ValidateHiSpeed();
            }

            if (process == "Campaign")
            {
                ValidateCampaign();
            }
        }
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            toolStripProgressBar1.Value = e.ProgressPercentage;
        }
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            toolStripStatusLabel1.Text = "Validation Completed!!";
            toolStripProgressBar1.Value = 0;
            Cursor.Current = Cursors.Default;
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            string process = e.Argument.ToString();
            if (process == "Hispeed")
            {
                MappingHiSpeed();
            }
            else if (process == "Campaign")
            {
                MappingCampaign();
            }
        }
        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            toolStripProgressBar1.Value = e.ProgressPercentage;
        }
        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            toolStripStatusLabel1.Text = "";
        }

        private void backgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            if (String.IsNullOrEmpty(id) == false)
            {
                ExportHiSpeed();
            }
            else if (String.IsNullOrEmpty(logHispeed) == false)
            {
                string logPath = outputPath + "\\Log_HiSpeedPromotion_" + urNo.ToUpper() + ".txt";
                using (StreamWriter writer = new StreamWriter(logPath, true))
                {
                    writer.Write(logHispeed);
                }
            }

            if (String.IsNullOrEmpty(tolPack) == false)
            {
                if (backgroundWorker3.IsBusy)
                {
                    backgroundWorker3.CancelAsync();
                }

                ExportCampaign();
            }
        }

        private void backgroundWorker3_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            toolStripProgressBar1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            toolStripStatusLabel1.Text = "Finish!!";
            toolStripProgressBar1.Value = 0;

            if (String.IsNullOrEmpty(logHispeed) == false || String.IsNullOrEmpty(logCampaign) == false)
            {
                MessageBox.Show("Please read log file.", "Successfully", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                MessageBox.Show("The data has been exported successfully", "Successfully", MessageBoxButtons.OK);
            }

            Cursor.Current = Cursors.Default;
        }

        private void listBox1_Click(object sender, EventArgs e)
        {
            dataGridView1.ClearSelection();
            if (listBox1.SelectedItem != null)
            {
                int selected = listBox1.SelectedIndex;
                dataGridView1.Rows[indexListbox[selected]].Selected = true;
                dataGridView1.FirstDisplayedScrollingRowIndex = indexListbox[selected];
                dataGridView1.Focus();
            }
        }

        private void btnExe_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(validateLog))
            {
                btnLog.Visible = false;
                btnExe.Enabled = false;

                if (func == "Hispeed" && (String.IsNullOrEmpty(id)))
                {
                    backgroundWorker2.RunWorkerAsync("Hispeed");
                }
                else if (func == "Campaign")
                {
                    if (backgroundWorker2.IsBusy == true)
                    {
                        backgroundWorker2.CancelAsync();
                    }
                    
                    backgroundWorker2.RunWorkerAsync("Campaign");
                }
            }
            else
            {
                MessageBox.Show("Please review this file carefully before clicking button execute!!");
            }

        }

        private void btnLog_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(validateLog))
            {
                MessageBox.Show("The verification process is complete. No errors occurred during process.");
            }
            else
            {
                string strFilePath = outputPath + "\\LOG_VALIDATE_" + urNo.ToUpper() + "_" + DateTime.Now.ToString("ddMMyyyy") + ".txt";
                using (StreamWriter writer = new StreamWriter(strFilePath, true))
                {
                    writer.Write(validateLog);
                }

                MessageBox.Show("Log file has been written successfully." + "\r\n" + "Program will be closing");

                Application.Exit();
            }
        }

        private void MainHispeed_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (backgroundWorker1.IsBusy)
            {
                backgroundWorker1.CancelAsync();
            }
            if (backgroundWorker2.IsBusy)
            {
                backgroundWorker2.CancelAsync();
            }
            if (backgroundWorker3.IsBusy)
            {
                backgroundWorker3.CancelAsync();
            }

            if (ConnectionProd != null)
            {
                if (ConnectionProd.State == ConnectionState.Open)
                {
                    ConnectionProd.Close();
                    ConnectionProd.Dispose();
                }
            }

            if(ConnectionTemp != null)
            {
                if(ConnectionTemp.State == ConnectionState.Open)
                {
                    ConnectionTemp.Close();
                    ConnectionTemp.Dispose();
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
        #endregion

        #region "Private Method"
        private void ValidateHiSpeed()
        {
            try
            {
                Application.UseWaitCursor = true;
                Cursor.Current = Cursors.WaitCursor;

                labelFunction.Text = "Hi-Speed Promotion";
                toolStripStatusLabel1.Text = "Checking HiSpeed...";

                InitialValue();
                validation = new Validation(ConnectionProd, ConnectionTemp);
                backgroundWorker1.ReportProgress(3);

                //Get Description
                if (String.IsNullOrEmpty(fileDesc) == false)
                {
                    lstPname = validation.GetDescription(fileDesc);
                }
                //Get Channel from DB
                if (lstChannel.Count <= 0 || lstChannel is null)
                {
                    lstChannel = validation.GetChannelFromDB;
                }
                //Get SubProfile from DB
                if (lstSubProfile.Count <= 0)
                {
                    lstSubProfile = validation.GetSubProfile;
                }
                //Get Extra profile from DB
                if (lstExtraProfile.Count <= 0)
                {
                    lstExtraProfile = validation.GetExtraProfile;
                }
                //Get speed from DB
                if (lstSpeedMast.Count <= 0)
                {
                    lstSpeedMast = validation.GetSpeedFromDB;
                }
                //Get contract from DB
                if (tableContract.Rows.Count <= 0)
                {
                    tableContract = validation.GetContract();
                }
                //Get prodtype from DB
                if (tableProdType.Rows.Count <= 0)
                {
                    tableProdType = validation.GetProdType();
                }

                backgroundWorker1.ReportProgress(20);

                if (lstChannel.Count <= 0 || lstSubProfile.Count <= 0 || lstExtraProfile.Count <= 0 ||
                    lstSpeedMast.Count <= 0 || tableContract.Rows.Count <= 0 || tableProdType.Rows.Count <= 0)
                {
                    MessageBox.Show("An error occurred while retrieving data from the database.Please try again!!");
                    backgroundWorker1.ReportProgress(0);
                }
                else
                {
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        string mkt = dataGridView1.Rows[i].Cells[1].Value.ToString().Trim();
                        string speed = dataGridView1.Rows[i].Cells[2].Value.ToString().Trim();
                        string subProfile = dataGridView1.Rows[i].Cells[3].Value.ToString().Trim();
                        string extra = dataGridView1.Rows[i].Cells[4].Value.ToString().ToUpper().Trim();
                        string order = dataGridView1.Rows[i].Cells[6].Value.ToString().Trim();
                        string channel = dataGridView1.Rows[i].Cells[7].Value.ToString().Trim();
                        string start = dataGridView1.Rows[i].Cells[10].Value.ToString().Trim();
                        string end = dataGridView1.Rows[i].Cells[11].Value.ToString().Trim();
                        string entry = dataGridView1.Rows[i].Cells[12].Value.ToString().Trim();
                        string install = dataGridView1.Rows[i].Cells[13].Value.ToString().Trim();

                        #region "Speed"
                        string[] msgSpeed = validation.CheckSpeed(lstSpeedMast, mkt, speed);
                        if (msgSpeed[0] != "Success")
                        {
                            listBox1.Items.Add(msgSpeed[0]);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "mkt", i);

                            validateLog += "[MKT:" + mkt + ", Speed:" + speed + "]     " + msgSpeed[0] + "\r\n";
                        }

                        if (msgSpeed[1] != "Success")
                        {
                            listBox1.Items.Add(msgSpeed[1]);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "speed", i);

                            validateLog += "[MKT:" + mkt + ", Speed:" + speed + "]     " + msgSpeed[1] + "\r\n";
                        }

                        if (msgSpeed[2] != "Success" && msgSpeed[2] != null)
                        {
                            listBox1.Items.Add(msgSpeed[2]);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "mkt", i);
                            hilightRow("Hispeed", "speed", i);

                            validateLog += "[MKT:" + mkt + ", Speed:" + speed + "]     " + msgSpeed[2] + "\r\n";
                        }
                        #endregion

                        #region"Extra Profile"
                        string msgExtra = validation.CheckExtra(lstExtraProfile, extra);
                        if (msgExtra != "Success")
                        {
                            listBox1.Items.Add(msgExtra);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "extra", i);

                            validateLog += "[MKT:" + mkt + ", Speed:" + speed + "]     " + msgExtra + "\r\n";
                        }
                        #endregion

                        #region "SubProfile"
                        string msgSub = validation.CheckSubProfile(lstSubProfile, subProfile);
                        if (msgSub != "Success")
                        {
                            listBox1.Items.Add(msgSub);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "subProfile", i);

                            validateLog += "[MKT:" + mkt + ", Speed:" + speed + "]     " + msgSub + "\r\n";
                        }
                        #endregion

                        #region "OrderType"
                        order = Regex.Replace(order, "NEW", "New", RegexOptions.IgnoreCase);
                        order = Regex.Replace(order, "CHANGE", "Change", RegexOptions.IgnoreCase);
                        dataGridView1.Rows[i].Cells[6].Value = order;

                        string msgOrder = validation.CheckOrderType(order);
                        if (msgOrder != "Success")
                        {
                            listBox1.Items.Add(msgOrder);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "order", i);

                            validateLog += "[MKT:" + mkt + ", Speed:" + speed + "]     " + msgOrder + "\r\n";
                        }
                        #endregion

                        #region "Channel"
                        channel = Regex.Replace(channel, "ALL", "DEFAULT", RegexOptions.IgnoreCase);
                        dataGridView1.Rows[i].Cells[7].Value = channel;

                        string msgChannel = validation.CheckChannel(lstChannel, channel, end);
                        if (msgChannel != "Success")
                        {
                            listBox1.Items.Add(msgChannel);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "channel", i);

                            validateLog += "[MKT:" + mkt + ", Speed:" + speed + "]     " + msgChannel + "\r\n";
                        }
                        #endregion

                        #region "Date"
                        string msgDate = validation.CheckDate(start, end);
                        if (msgDate != "Success")
                        {
                            if (msgDate == "StartDate fotmat is not supported")
                            {
                                listBox1.Items.Add(msgDate);
                                indexListbox.Add(i);
                                hilightRow("Hispeed", "start", i);

                                validateLog += "[MKT:" + mkt + ", Speed:" + speed + "]     " + msgDate + "\r\n";
                            }
                            else if (msgDate == "EndDate fotmat is not supported")
                            {
                                listBox1.Items.Add(msgDate);
                                indexListbox.Add(i);
                                hilightRow("Hispeed", "end", i);

                                validateLog += "[MKT:" + mkt + ", Speed:" + speed + "]     " + msgDate + "\r\n";
                            }
                            else
                            {
                                listBox1.Items.Add(msgDate);
                                indexListbox.Add(i);
                                hilightRow("Hispeed", "start", i);
                                hilightRow("Hispeed", "end", i);

                                validateLog += "[MKT:" + mkt + ", Speed:" + speed + "]     " + msgDate + "\r\n";
                            }
                        }
                        #endregion

                        #region "Contract"
                        string msgContract = validation.CheckContract(tableContract, entry, install);
                        if (msgContract != "Success")
                        {
                            listBox1.Items.Add(msgContract);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "entry", i);
                            hilightRow("Hispeed", "install", i);

                            validateLog += "[MKT:" + mkt + ", Speed:" + speed + "]     " + msgContract + "\r\n";
                        }
                        #endregion

                        backgroundWorker1.ReportProgress(20 + ((i + 1) * 80 / dataGridView1.RowCount));
                    }
                }

                if (String.IsNullOrEmpty(validateLog))
                {
                    btnLog.Visible = false;
                }


            }
            catch
            {}
            finally 
            {
                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;

                this.Refresh();
            }
        }
                

        private void ValidateCampaign()
        {
            Application.UseWaitCursor = true;
            Cursor.Current = Cursors.WaitCursor;

            labelFunction.Text = "Campaign Mapping";
            toolStripStatusLabel1.Text = "Checking Campaign...";

            InitialValue();

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                string requestType = dataGridView1.Rows[i].Cells[0].Value.ToString().Trim();
                string campaignName = dataGridView1.Rows[i].Cells[1].Value.ToString().Trim();
                string tolPackage = dataGridView1.Rows[i].Cells[2].Value.ToString().Trim();
                string tolDiscount = dataGridView1.Rows[i].Cells[3].Value.ToString().Trim();
                string tvsPackage = dataGridView1.Rows[i].Cells[4].Value.ToString().Trim();
                string tvsDiscount = dataGridView1.Rows[i].Cells[5].Value.ToString().Trim();

                if (String.IsNullOrEmpty(campaignName) == false &&
                    String.IsNullOrEmpty(tolPackage) == false)
                {
                    if (requestType == "Insert" || requestType.Contains("Update"))
                    {
                        if (String.IsNullOrEmpty(tvsPackage))
                        {
                            //write log
                            listBox1.Items.Add("Not found TVS_Package of '" + tolPackage + "'");
                            indexListbox.Add(i);
                            hilightRow("Campaign", "tvsPack", i);

                            validateLog += "Not found TVS_Package of '" + tolPackage + "'" + "\r\n";
                        }
                        else
                        {
                            if(tolPackage.EndsWith("G"))
                            {
                                string[] val = tolPackage.Split('-');
                                val[1] = Regex.Replace(val[1], "[^0-9]", "");

                                tolPackage = val[0] + "-" + Convert.ToInt32(val[1]) * 1000;
                            }

                            //CHECK SUB PROFILE AND MKT
                            string txt = "SELECT P.BUNDLE_CAMPAIGN,P.P_CODE || '-' || S.SPEED_ID AS MKTCODE " +
                                    "FROM HISPEED_PROMOTION P, HISPEED_SPEED_PROMOTION S WHERE P.P_ID = S.P_ID " +
                                    "AND P_CODE || '-' || SPEED_ID = '" + tolPackage + "' AND BUNDLE_CAMPAIGN = '" + campaignName + "'";

                            OracleCommand command = new OracleCommand(txt, ConnectionProd);
                            OracleDataReader reader = command.ExecuteReader();
                            if (reader.HasRows)
                            {
                                for (int j = i + 1; j < dataGridView1.RowCount; j++)
                                {
                                    string requestTypeN = dataGridView1.Rows[j].Cells[0].Value.ToString().Trim();
                                    string campaignNameN = dataGridView1.Rows[j].Cells[1].Value.ToString().Trim();
                                    string tolPackageN = dataGridView1.Rows[j].Cells[2].Value.ToString().Trim();
                                    string tolDiscountN = dataGridView1.Rows[j].Cells[3].Value.ToString().Trim();
                                    string tvsPackageN = dataGridView1.Rows[j].Cells[4].Value.ToString().Trim();
                                    string tvsDiscountN = dataGridView1.Rows[j].Cells[5].Value.ToString().Trim();

                                    if (requestType == requestTypeN && campaignName == campaignNameN &&
                                        tolPackage == tolPackageN && tolDiscount == tolDiscountN && tvsPackage == tvsPackageN
                                        && tvsDiscount == tvsDiscountN)
                                    {
                                        listBox1.Items.Add("Duplicate record: " + i + " and record: " + j);
                                        indexListbox.Add(i);
                                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                                        dataGridView1.Rows[j].DefaultCellStyle.BackColor = Color.Yellow;

                                        validateLog += "Duplicate record: " + i + " and record: " + j + "\r\n";
                                    }
                                }
                            }
                            else
                            {
                                listBox1.Items.Add("TOL_PACKAGE: " + tolPackage + " and CAMPAIGN: " + campaignName +
                                    " not found on table HISPEED_PROMOTION");
                                indexListbox.Add(i);
                                dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Red;

                                validateLog += "TOL_PACKAGE: " + tolPackage + " and CAMPAIGN: " + campaignName +
                                    " not found on table HISPEED_PROMOTION" + "\r\n";
                            }
                        }
                    }
                    else
                    {
                        //write log
                        listBox1.Items.Add("Request type is wrong!");
                        indexListbox.Add(i);
                        hilightRow("Campaign", "type", i);

                        validateLog += "Request type of package[" + tolPackage + "] is wrong!" + "\r\n";
                    }
                }

                backgroundWorker1.ReportProgress((i + 1) * 100 / dataGridView1.RowCount);
            }

            if (String.IsNullOrEmpty(validateLog))
            {
                btnLog.Visible = false;
            }

            Application.UseWaitCursor = false;
            Cursor.Current = Cursors.Default;

            this.Refresh();
        }

        private void MappingHiSpeed()
        {
            try
            {
                Application.UseWaitCursor = true;
                Cursor.Current = Cursors.WaitCursor;

                toolStripStatusLabel1.Text = "Inserting Hi-Speed...";

                ReserveID reserveID = new ReserveID();
                reserveID.Reserve(ConnectionProd, ConnectionTemp, "Hispeed", implementer, urNo);

                backgroundWorker2.ReportProgress(5);

                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    string media = dataGridView1.Rows[i].Cells[0].Value.ToString().Trim();
                    string mkt = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    string speed = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    string sub = dataGridView1.Rows[i].Cells[3].Value.ToString().Trim();
                    string extra = dataGridView1.Rows[i].Cells[4].Value.ToString().Trim();
                    double price = Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString().Trim());
                    string order = dataGridView1.Rows[i].Cells[6].Value.ToString().Trim();
                    string channel = dataGridView1.Rows[i].Cells[7].Value.ToString();
                    string modemType = dataGridView1.Rows[i].Cells[8].Value.ToString().Trim();
                    string docsis = dataGridView1.Rows[i].Cells[9].Value.ToString().Trim();
                    string start = dataGridView1.Rows[i].Cells[10].Value.ToString().Trim();
                    string end = dataGridView1.Rows[i].Cells[11].Value.ToString().Trim();
                    string entry = dataGridView1.Rows[i].Cells[12].Value.ToString().Trim();
                    string install = dataGridView1.Rows[i].Cells[13].Value.ToString().Trim();

                    string[] lstMkt = mkt.Split('-');
                    mkt = lstMkt[0].Trim();

                    int suffix;
                    if (lstMkt[1].Contains("G"))
                    {
                        lstMkt[1] = Regex.Replace(lstMkt[1], "[^0-9]", "");
                        suffix = Convert.ToInt32(lstMkt[1]) * 1000;
                    }
                    else
                    {
                        if (lstMkt[1] == "00" || lstMkt[1] == "01")
                        {
                            suffix = Convert.ToInt32(Regex.Replace((speed.Split('/'))[0], "[^0-9]", ""));
                        }
                        else
                        {
                            suffix = Convert.ToInt32(lstMkt[1]);
                        }
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

                    //SubProfile = STL
                    if (sub.StartsWith("STL"))
                    {
                        sub = "N";
                    }

                    //Get P_Name
                    string p_name = GetPName(dataGridView1.Rows[i].Cells[1].Value.ToString().Trim());
                    if (p_name == dataGridView1.Rows[i].Cells[1].Value.ToString().Trim())
                    {
                        listBox1.Items.Add("Not found description of MKT [P_NAME] in file or database!!");
                        indexListbox.Add(i);
                        dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.Yellow;
                    }

                    //Convert upload speed
                    string[] sp = speed.Split('/');
                    int uploadK = validation.ConvertUOM2K(sp[1], Regex.Replace(sp[1], "[0-9]", ""));

                    //Change format date
                    start = validation.ChangeFormatDate(start);
                    end = validation.ChangeFormatDate(end);

                    for (int j = 0; j < lstOrder.Length; j++)
                    {
                        List<string> lstData = new List<string>();
                        string tmpExtra = "";
                        /* string txt = "SELECT P.P_ID, P.P_CODE, P.ORDER_TYPE,C.SALE_CHANNEL ,P.START_DATE,P.END_DATE,P.STATUS, " +
                             "S.SPEED_ID DOWNLOAD, S.UPLOAD_SPEED / 1024 UPLOAD,S.PRICE FROM HISPEED_PROMOTION P, " +
                             "HISPEED_SPEED_PROMOTION S,HISPEED_CHANNEL_PROMOTION C WHERE P.P_ID = S.P_ID AND P.P_ID = C.P_ID " +
                             "AND P_CODE = '" + mkt + "' AND SPEED_ID = '" + suffix + "' AND ORDER_TYPE = '" + lstOrder[j] + "'";*/

                        if (String.IsNullOrEmpty(extra))
                        {
                            tmpExtra = "IS NULL";
                        }
                        else
                        {
                            tmpExtra = "= '" + extra + "'";
                        }
                        string txt = "SELECT P.P_ID, P.P_CODE, P.ORDER_TYPE,C.SALE_CHANNEL ,P.START_DATE,P.END_DATE, " +
                                        "P.STATUS, S.SPEED_ID DOWNLOAD, S.UPLOAD_SPEED/1024 UPLOAD,S.PRICE,P.EXTRA_MESSAGE, " +
                                        "P.PRODTYPE,P.BUNDLE_CAMPAIGN " +
                                     "FROM HISPEED_PROMOTION P, HISPEED_SPEED_PROMOTION S,HISPEED_CHANNEL_PROMOTION C " +
                                     "WHERE P.P_ID = S.P_ID  AND P.P_ID = C.P_ID AND P_CODE = '" + mkt + "' " +
                                     "AND SPEED_ID = '" + suffix + "' AND ORDER_TYPE = '" + lstOrder[j] + "' " +
                                     "AND P.EXTRA_MESSAGE " + tmpExtra + " AND P.BUNDLE_CAMPAIGN = '" + sub + "'";

                        OracleCommand cmd = new OracleCommand(txt, ConnectionProd);
                        OracleDataReader reader = cmd.ExecuteReader();

                        if (reader.HasRows)
                        {
                            //Existing
                            reader.Read();
                            int id = Convert.ToInt32(reader["P_ID"]);

                            ExistingData(id, suffix, uploadK, channel, price, start, end, dataGridView1.Rows[i].Cells[1].Value.ToString());
                        }
                        else
                        {
                            //New promotion
                            lstData.Add(media);
                            lstData.Add(mkt);
                            lstData.Add(uploadK.ToString());
                            lstData.Add(sub);
                            lstData.Add(extra);
                            lstData.Add(price.ToString());
                            lstData.Add(lstOrder[j]);
                            lstData.Add(channel);
                            lstData.Add(modemType);
                            lstData.Add(docsis);
                            lstData.Add(start);
                            lstData.Add(end);
                            lstData.Add(entry);
                            lstData.Add(install);
                            lstData.Add(p_name);
                            lstData.Add(suffix.ToString());

                            NewHiSpeedData(lstData);
                        }
                    }

                    backgroundWorker2.ReportProgress(5 + ((i + 1) * 90 / dataGridView1.RowCount));
                }

                //Update ReserveID
                reserveID.UpdateReserveID(ConnectionTemp, ConnectionProd, "Hispeed", urNo);
                backgroundWorker2.ReportProgress(100);

                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;

                if (sheets.Contains("Campaign Mapping"))
                {
                    DialogResult dialogResult = MessageBox.Show("The process mapping Hi-Speed promotion has been completed." + "\r\n" +
                        "Do you want go to the process mapping campaign?", "Complete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dialogResult == DialogResult.Yes)
                    {
                        Cursor.Current = Cursors.WaitCursor;
                        this.UseWaitCursor = true;

                        btnExe.Enabled = true;
                        btnLog.Visible = true;

                        func = "Campaign";
                        //validate campaign
                        DgvSettings dgvSettings = new DgvSettings();
                        List<string> lstHeader = new List<string>();

                        lstHeader.Clear();

                        lstHeader.Add("Type");
                        lstHeader.Add("Campaign Name");
                        lstHeader.Add("TOL Package");
                        lstHeader.Add("TOL Discount");
                        lstHeader.Add("TVS Package");
                        lstHeader.Add("TVS Discount");

                        InitialValue();

                        dataGridView1.DataSource = null;

                        dgvSettings.SetDgv(dataGridView1, filename, "Campaign Mapping$B2:G", lstHeader);
                        dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                        //validate campaign
                        if (backgroundWorker1.IsBusy)
                        {
                            backgroundWorker1.CancelAsync();
                        }
                        backgroundWorker1.RunWorkerAsync(func);
                    }
                    else
                    {
                        if (backgroundWorker2.IsBusy)
                        {
                            backgroundWorker2.CancelAsync();
                        }
                        if (backgroundWorker1.IsBusy)
                        {
                            backgroundWorker1.CancelAsync();
                        }

                        //export hispeed
                        backgroundWorker3.RunWorkerAsync();
                    }
                }
                else
                {
                    //export hispeed
                    backgroundWorker3.RunWorkerAsync();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while processing HiSpeed Function. Please try again!!" + "\r\n" +
                    "System Detail : " + ex.Message + "\r\n" + ex.StackTrace, "Automate Mapping Tool"
                    , MessageBoxButtons.OK, MessageBoxIcon.Error);

                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;
            }
        }

        private void NewHiSpeedData(List<string> data)
        {
            //Get data
            string media = data[0];
            string mkt = data[1];
            string uploadK = data[2];
            string sub = data[3];
            string extra = data[4];
            string price = data[5];
            string order = data[6];
            string channel = data[7];
            string modemType = data[8];
            string docsis = data[9];
            string start = data[10];
            string end = data[11];
            string entryF = data[12];
            string installF = data[13];
            string pName = data[14];
            string suffix = data[15];
            string onetimeCode, modem;

            if (String.IsNullOrEmpty(channel))
            {
                MessageBox.Show("Found channel of new MKT :" + mkt + " is empty.");
                Application.Exit();
            }
            else
            {
                try
                {
                    //Get MAX P_ID
                    OracleCommand cmd;
                    cmd = new OracleCommand("SELECT MAX(P_ID)+1 FROM HISPEED_PROMOTION WHERE P_ID LIKE '20%'", ConnectionProd);
                    OracleDataReader reader = cmd.ExecuteReader();
                    reader.Read();
                    int minID = Convert.ToInt32(reader[0]);

                    //Get prodtype
                    string prodType = "";
                    foreach (DataRow row in tableProdType.Rows)
                    {
                        string mediaDB = row[0].ToString();
                        string orderDB = row[1].ToString();

                        if (media == mediaDB && order == orderDB)
                        {
                            prodType = row[2].ToString();
                            break;
                        }
                    }

                    //Get Contract
                    string term = "", entry = "", install = "";
                    foreach (DataRow row in tableContract.Rows)
                    {
                        string entDB = row[0].ToString();
                        string insDB = row[2].ToString();

                        if (entryF == entDB && installF == insDB)
                        {
                            entry = row[1].ToString();
                            install = row[3].ToString();
                            term = row[5].ToString();
                            break;
                        }
                    }

                    //OnetimeCode
                    if(media == "FIBER_TO_HOME")
                    {
                        onetimeCode = "ENFP";
                    }
                    else
                    {
                        onetimeCode = "";
                    }

                    //modem
                    if(order == "New")
                    {
                        modem = "BM";
                    }
                    else
                    {
                        modem = "BM,NM";
                    }

                    if (prodType == "")
                    {
                        logHispeed += "Failed to insert ID: " + minID + " MKT: " + mkt + " Order: " + order + " Speed: "
                                + suffix + " into database" + "\r\n" +
                                "Detail : Incorrect media" + "\r\n" + "\r\n";
                    }
                    else if(Convert.ToDateTime(start).Date >= DateTime.Now.Date == false)
                    {
                        logHispeed += "Cannot to insert ID: " + minID + " MKT: " + mkt + " Order: " + order + " Speed: "
                                + suffix + " into database" + "\r\n" +
                                "Detail : StartDate < Sysdate" + "\r\n" + "\r\n";
                    }
                    else
                    {
                        //insert new data
                        cmd = ConnectionProd.CreateCommand();
                        OracleTransaction transaction = null;

                        if (start == DateTime.Now.ToString("dd/MM/yyyy"))
                        {
                            start = DateTime.Now.ToString();
                        }
                        
                        using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                        {
                            try
                            {
                                cmd.Transaction = transaction;
                                //Insert into hispeed_promotion
                                cmd.CommandText = "INSERT INTO HISPEED_PROMOTION VALUES (" + minID + ", '" + mkt + "', '" + mkt + "', '" + pName +
                                    "', '" + pName + "', '" + order + "', 'Active','" + extra + "','"+onetimeCode+"',0,0,'Y','Y','',0,'N','0','Y','Y','N','" + prodType +
                                    "', sysdate, sysdate, '" + term + "',0,'TI', TO_DATE('" + start + "','dd/mm/yyyy'), " +
                                    "TO_DATE('" + end + "','dd/mm/yyyy'), 'M', '" + mkt + "','N','N','Y', '" + entry + "', '" +
                                    install + "','" + modem + "','N','" + sub + "','')";
                                expHisp += cmd.CommandText + ";" + "\r\n";
                                cmd.ExecuteNonQuery();

                                //Insert into hispeed_speed_promotion
                                cmd.CommandText = "INSERT INTO HISPEED_SPEED_PROMOTION  VALUES (" + suffix + ", " + minID + ", " +
                                    price + ", null, 'Y', '" + suffix + "', '" + modemType + "', " + "'" + uploadK + "', '" + docsis + "')";
                                expHisp += cmd.CommandText + ";" + "\r\n";
                                cmd.ExecuteNonQuery();

                                string[] arrChannel;
                                if (channel.Contains(","))
                                {
                                    arrChannel = channel.Split(',');
                                }
                                else
                                {
                                    arrChannel = new string[1];
                                    arrChannel[0] = channel;
                                }

                                //Insert into hispeed_channel_promotion
                                for (int i = 0; i < arrChannel.Length; i++)
                                {
                                    cmd.CommandText = "INSERT INTO HISPEED_CHANNEL_PROMOTION VALUES(" + minID + ", '" + arrChannel[i].Trim() +
                                        "', TO_DATE('" + start + "','dd/MM/yyyy'), TO_DATE('" + end + "','dd/MM/yyyy'), 'S')";
                                    expHisp += cmd.CommandText + ";" + "\r\n";
                                    cmd.ExecuteNonQuery();
                                }
                                expHisp += "\r\n";

                                transaction.Commit();

                                id += "," + minID;
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();
                                logHispeed += "Failed to insert ID: " + minID + " MKT: " + mkt + " Order: " + order + " Speed: "
                                    + suffix + " into database" + "\r\n" +
                                    "Detail of system :" + ex.Message + "\r\n" + ex.StackTrace + "\r\n" + "\r\n";
                            }
                        }
                    }
                }
                catch (Exception)
                {}
            }
        }

        private void ExistingData(int id, int suffix, int upload, string channel, double price, string start, 
            string end, string mkt)
        {
            int suffixDB = -1, uploadDB = -1;
            double priceDB = 999999;
            string active = "";

            OracleCommand cmd = new OracleCommand("SELECT * FROM HISPEED_SPEED_PROMOTION WHERE P_ID = " + id, ConnectionProd);
            OracleDataReader reader = cmd.ExecuteReader();
            reader.Read();
            if (reader.HasRows)
            {
                suffixDB = Convert.ToInt32(reader["SUFFIX"]);
                uploadDB = Convert.ToInt32(reader["UPLOAD_SPEED"]);
                active = reader["ACTIVE_PRICE"].ToString();
                priceDB = Convert.ToDouble(reader["PRICE"]);
            }

            cmd = ConnectionProd.CreateCommand();
            OracleTransaction transaction;
            using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
            {
                cmd.Transaction = transaction;
                try
                {
                    if (suffix == suffixDB && upload == uploadDB)
                    {
                        if (active == "Y")
                        {
                            if (price == priceDB)
                            {
                                if (String.IsNullOrEmpty(channel))
                                {
                                    //update h.c end date = enddate file where p.id and hc.enddate is null or hc.enddate >= sysdate
                                    cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET END_DATE = TO_DATE('" + end + "', 'dd/MM/yyyy')" +
                                                " WHERE P_ID = " + id + " AND(END_DATE >= trunc(sysdate) OR END_DATE IS NULL)";
                                    expHisp += cmd.CommandText + ";" + "\r\n";
                                    cmd.ExecuteNonQuery();                                    
                                }
                                else
                                {
                                    Dictionary<string, string[]> lstChannelDB = new Dictionary<string, string[]>();

                                    cmd.CommandText = "SELECT* FROM HISPEED_CHANNEL_PROMOTION WHERE P_ID = " + id;
                                    reader = cmd.ExecuteReader();

                                    while (reader.Read())
                                    {
                                        string[] date = new string[2];
                                        date[0] = reader["START_DATE"].ToString();
                                        date[1] = reader["END_DATE"].ToString();
                                        lstChannelDB.Add(reader["SALE_CHANNEL"].ToString(), date);
                                    }

                                    //channel from file
                                    string[] lstCh;
                                    if (channel.Contains(','))
                                    {
                                        lstCh = channel.Split(',');
                                    }
                                    else
                                    {
                                        lstCh = new string[1];
                                        lstCh[0] = channel;
                                    }

                                    for (int i = 0; i < lstCh.Length; i++)
                                    {
                                        string ch = lstCh[i].Trim();
                                        DateTime startDB = new DateTime();
                                        DateTime endDB = new DateTime();
                                        DateTime startF = new DateTime();
                                        DateTime endF = new DateTime();

                                        //existing channel
                                        if (lstChannelDB.Keys.Contains(ch))
                                        {
                                            string[] date = lstChannelDB[ch];
                                            if (String.IsNullOrEmpty(date[0]) == false)
                                            {
                                                startDB = Convert.ToDateTime(date[0]);
                                            }

                                            if (String.IsNullOrEmpty(date[1]) == false)
                                            {
                                                endDB = Convert.ToDateTime(date[1]);
                                            }

                                            if (String.IsNullOrEmpty(start) == false)
                                            {
                                                startF = Convert.ToDateTime(start);
                                            }

                                            if (String.IsNullOrEmpty(end) == false)
                                            {
                                                endF = Convert.ToDateTime(end);
                                            }

                                            if (endDB.Date < DateTime.Now.Date)
                                            {
                                                if (startF.Date == DateTime.Now.Date)
                                                {
                                                    //update startdate == datetime.now
                                                    cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET START_DATE = sysdate " +
                                                        "WHERE P_ID = " + id + " AND SALE_CHANNEL = '" + ch + "'";
                                                    expHisp += cmd.CommandText + ";" + "\r\n";
                                                    cmd.ExecuteNonQuery();
                                                }
                                                else
                                                {
                                                    if (startF.Date > DateTime.Now.Date)
                                                    {
                                                        //update start date by date on file
                                                        cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET START_DATE = TO_DATE('" +
                                                            start + "', 'dd/MM/yyyy') WHERE P_ID = " + id + " AND SALE_CHANNEL = '" + ch + "'";
                                                        expHisp += cmd.CommandText + ";" + "\r\n";
                                                        cmd.ExecuteNonQuery();
                                                    }
                                                    else
                                                    {
                                                        if (endF.Date == DateTime.Now.Date)
                                                        {
                                                            //update enddate = datetime sysdate
                                                            cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET END_DATE = sysdate " +
                                                                "WHERE P_ID = " + id + " AND SALE_CHANNEL = '" + ch + "'";
                                                            expHisp += cmd.CommandText + ";" + "\r\n";
                                                            cmd.ExecuteNonQuery();
                                                        }
                                                        else
                                                        {
                                                            //update enddate = end on file
                                                            cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET END_DATE = TO_DATE('" +
                                                            end + "', 'dd/MM/yyyy') WHERE P_ID = " + id + " AND SALE_CHANNEL = '" + ch + "'";
                                                            expHisp += cmd.CommandText + ";" + "\r\n";
                                                            cmd.ExecuteNonQuery();
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (startDB.Date > DateTime.Now.Date)
                                                {
                                                    //update start = date sysdate
                                                    cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET START_DATE = sysdate " +
                                                                "WHERE P_ID = " + id + " AND SALE_CHANNEL = '" + ch + "'";
                                                    expHisp += cmd.CommandText + ";" + "\r\n";
                                                    cmd.ExecuteNonQuery();
                                                }
                                                else
                                                {
                                                    if (endF.Date == DateTime.Now.Date)
                                                    {
                                                        //update enddate = datetime sysdate
                                                        cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET END_DATE = sysdate " +
                                                                "WHERE P_ID = " + id + " AND SALE_CHANNEL = '" + ch + "'";
                                                        expHisp += cmd.CommandText + ";" + "\r\n";
                                                        cmd.ExecuteNonQuery();
                                                    }
                                                    else
                                                    {
                                                        //update enddate = end on file
                                                        cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET END_DATE = TO_DATE('" +
                                                            end + "', 'dd/MM/yyyy') WHERE P_ID = " + id + " AND SALE_CHANNEL = '" + ch + "'";
                                                        expHisp += cmd.CommandText + ";" + "\r\n";
                                                        cmd.ExecuteNonQuery();
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            //insert new chnnel
                                            cmd.CommandText = "INSERT INTO HISPEED_CHANNEL_PROMOTION VALUES(" + id + ", '" +
                                                ch + "', TO_DATE('" + start + "','dd/MM/yyyy'), TO_DATE('" + end + "','dd/MM/yyyy'), 'S')";
                                            expHisp += cmd.CommandText + ";" + "\r\n";
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                }
                            }
                            else
                            {
                                DialogResult dialogResult = MessageBox.Show("Do you want to update price [" + price + "] on databse?",
                                    "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                if (dialogResult == DialogResult.Yes)
                                {
                                    //update new price to DB 
                                    cmd.CommandText = "UPDATE HISPEED_SPEED_PROMOTION SET PRICE = " + price + " WHERE P_ID = " + id;
                                    expHisp += cmd.CommandText + ";" + "\r\n";
                                    cmd.ExecuteNonQuery();
                                }
                                else
                                {
                                    logHispeed += "MKT: " + mkt + ", price[" + price + "] on file is not matching price[" +
                                        priceDB + "] on DB" + "\r\n" + "\r\n";
                                }
                            }
                        }
                        else
                        {
                            if (price == priceDB)
                            {
                                //update active price = y
                                cmd.CommandText = "UPDATE HISPEED_SPEED_PROMOTION SET ACTIVE_PRICE = 'Y' WHERE P_ID = " + id;
                                expHisp += cmd.CommandText + ";" + "\r\n";
                                cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                DialogResult dialogResult = MessageBox.Show("Do you want to update price [" + price + "] on databse?",
                                    "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                if (dialogResult == DialogResult.Yes)
                                {
                                    //update new price to DB and active = y
                                    cmd.CommandText = "UPDATE HISPEED_SPEED_PROMOTION SET ACTIVE_PRICE = 'Y', PRICE = " + price +
                                        " WHERE P_ID = " + id;
                                    expHisp += cmd.CommandText + ";" + "\r\n";
                                    cmd.ExecuteNonQuery();
                                }
                                else
                                {
                                    logHispeed += "MKT: " + mkt + ", price[" + price + "] on file is not matching price[" +
                                        priceDB + "] on DB" + "\r\n" + "\r\n";
                                }
                            }
                        }
                    }
                    else
                    {
                        logHispeed += "Download or Upload Speed of " + mkt + " not matching on database." + "\r\n" + "\r\n";
                    }

                    transaction.Commit();

                    this.id += "," + id;
                }
                catch (Exception ex)
                {
                    transaction.Rollback();

                    logHispeed += "Failed to update data ID: " + id + " MKT: " + mkt + " Speed: " + suffix + " into database" + "\r\n" +
                                "Detail of system :" + ex.Message + "\r\n" + ex.StackTrace + "\r\n" + "\r\n";
                }
            }
        }

        private void MappingCampaign()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                this.UseWaitCursor = true;

                toolStripStatusLabel1.Text = "Inserting Campaign...";

                OracleTransaction transaction = null;
                OracleCommand cmd = null;

                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    string requestType = dataGridView1.Rows[i].Cells[0].Value.ToString().Trim();
                    string campaignName = dataGridView1.Rows[i].Cells[1].Value.ToString().Trim();
                    string tolPackage = dataGridView1.Rows[i].Cells[2].Value.ToString().Trim();
                    string tolDiscount = dataGridView1.Rows[i].Cells[3].Value.ToString().Trim();
                    string tvsPackage = dataGridView1.Rows[i].Cells[4].Value.ToString().Trim();
                    string tvsDiscount = dataGridView1.Rows[i].Cells[5].Value.ToString().Trim();

                    if (String.IsNullOrEmpty(campaignName) == false &&
                        String.IsNullOrEmpty(tolPackage) == false)
                    {
                        using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                        {
                            cmd = ConnectionProd.CreateCommand();
                            cmd.Transaction = transaction;
                            string status = "", valTOLDisc, valTVSDisc;

                            if (String.IsNullOrEmpty(tolDiscount))
                            {
                                valTOLDisc = "IS NULL";
                            }
                            else
                            {
                                valTOLDisc = "= '" + tolDiscount + "'";
                            }

                            if (String.IsNullOrEmpty(tvsDiscount))
                            {
                                valTVSDisc = "IS NULL";
                            }
                            else
                            {
                                valTVSDisc = "= '" + tvsDiscount + "'";
                            }

                            try
                            {
                                string txt = "SELECT * FROM CAMPAIGN_MAPPING WHERE TOL_PACKAGE = '" + tolPackage + "' AND TOL_DISCOUNT " + valTOLDisc +
                                         " AND TVS_PACKAGE = '" + tvsPackage + "' AND TVS_DISCOUNT " + valTVSDisc + " AND STATUS IN('A', 'I')";

                                //OracleCommand command = new OracleCommand(txt, ConnectionProd);
                                cmd.CommandText = txt;
                                OracleDataReader reader = cmd.ExecuteReader();

                                if (requestType == "Insert")
                                {
                                    if (reader.HasRows)
                                    {
                                        reader.Read();
                                        status = reader["STATUS"].ToString();
                                        if (status == "A")
                                        {
                                            //Already exists data in the database
                                            logCampaign += "Already exists data TOL_PACKAGE: '" + tolPackage + "' Campaign_Name: '" + campaignName +
                                                "' TOL_DISCOUNT: '" + tolDiscount + "' TVS_PACKAGE: '" + tvsPackage + "' TVS_DISCOUNT: '" +
                                                tvsDiscount + "'" + "\r\n";
                                        }
                                        else
                                        {
                                            cmd.CommandText = "UPDATE CAMPAIGN_MAPPING SET STATUS = 'A' WHERE TOL_PACKAGE = '" + tolPackage +
                                                "' AND TOL_DISCOUNT " + valTOLDisc + " AND TVS_PACKAGE = '" + tvsPackage +
                                                "' AND TVS_DISCOUNT " + valTVSDisc;
                                            expCamp += cmd.CommandText + ";" + "\r\n";
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    else
                                    {
                                        //insert new campaign
                                        cmd.CommandText = "INSERT INTO CAMPAIGN_MAPPING(CAMPAIGN_NAME, TOL_PACKAGE, TOL_DISCOUNT, " +
                                            "TVS_PACKAGE, TVS_DISCOUNT, STATUS) VALUES('" + campaignName + "', '" + tolPackage + "', '" + tolDiscount +
                                            "', '" + tvsPackage + "', '" + tvsDiscount + "', 'A')";
                                        expCamp += cmd.CommandText + ";" + "\r\n";
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                                else
                                {
                                    if (reader.HasRows)
                                    {
                                        reader.Read();
                                        status = reader["STATUS"].ToString();

                                        if (status == "A")
                                        {
                                            cmd.CommandText = "UPDATE CAMPAIGN_MAPPING SET STATUS = 'I' WHERE TOL_PACKAGE = '" + tolPackage +
                                                "' AND TOL_DISCOUNT " + valTOLDisc + " AND TVS_PACKAGE = '" + tvsPackage +
                                                "' AND TVS_DISCOUNT " + valTVSDisc;
                                            expCamp += cmd.CommandText + ";" + "\r\n";
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    else
                                    {
                                        logCampaign += "Not found data TOL_PACKAGE: '" + tolPackage + "' Campaign_Name: '" + campaignName +
                                                "' TOL_DISCOUNT: '" + tolDiscount + "' TVS_PACKAGE: '" + tvsPackage + "' TVS_DISCOUNT: '" +
                                                tvsDiscount + "' in database" + "\r\n";
                                    }
                                }

                                transaction.Commit();

                                tolPack += "," + "'" + tolPackage + "'";
                                tvsPack += "," + "'" + tvsPackage + "'";
                            }
                            catch (Exception)
                            {
                                transaction.Rollback();

                                logCampaign += "Failed to insert or update data TOL_PACKAGE: '" + tolPackage + "' Campaign_Name: '" + campaignName +
                                                "' TOL_DISCOUNT: '" + tolDiscount + "' TVS_PACKAGE: '" + tvsPackage + "' TVS_DISCOUNT: '" +
                                                tvsDiscount + "' in database" + "\r\n";
                            }
                        }
                    }
                }

                if (String.IsNullOrEmpty(logCampaign) == false && String.IsNullOrEmpty(tolPack) && String.IsNullOrEmpty(tvsPack))
                {
                    string logPath = outputPath + "\\Log_Campaign_" + urNo.ToUpper() + ".txt";
                    using (StreamWriter writer = new StreamWriter(logPath, true))
                    {
                        writer.Write(logCampaign);
                    }
                }

                backgroundWorker3.RunWorkerAsync();

                this.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;
            }
            catch(Exception ex)
            {
                MessageBox.Show("An error occurred while processing Campaign Mapping Function. Please try again!!" + "\r\n" +
                    "Detail of system : " + ex.Message + "\r\n" + ex.StackTrace, "Automate Mapping Tool"
                    , MessageBoxButtons.OK, MessageBoxIcon.Error);

                Application.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;

            }
        }

        private void ExportHiSpeed()
        {
            Application.UseWaitCursor = true;
            Cursor.Current = Cursors.WaitCursor;

            toolStripStatusLabel1.Text = "Exporting Hi-Speed Promotion...";

            Excel.Application xlApp;
            Excel.Workbook workbook;
            Excel.Worksheet sheet;

            DataTable dt = new DataTable();
            OracleDataAdapter adapter;

            backgroundWorker3.ReportProgress(3);
            //Set excel
            xlApp = new Excel.Application();
            workbook = xlApp.Workbooks.Add(Type.Missing);
            sheet = workbook.ActiveSheet as Excel.Worksheet;
            sheet.Name = "Report";
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;

            if (String.IsNullOrEmpty(id) == false)
            {
                id = id.Substring(1);
            }

            string sql = "SELECT P.START_DATE AS PROMOTION_START_DATE" +
                        ",P.END_DATE AS PROMOTION_END_DATE" +
                        ",P.PRODTYPE" +
                        ",P.P_ID" +
                        ",P_NAME" +
                        ",P_CODE || '-' || SUFFIX AS PROMOTION" +
                        ",ORDER_TYPE" +
                        ",P.STATUS" +
                        ",SALE_CHANNEL" +
                        ",C.START_DATE AS SALE_CHANNEL_START" +
                        ",C.END_DATE AS SALE_CHANNEL_END" +
                        ",BUNDLE_CAMPAIGN" +
                        ",PRICE" +
                        ",NOTIFY_MSG AS TERM" +
                        ",TDS_MODEM_P_CODE AS ENTRY" +
                        ",TDS_ROUTER_P_CODE AS INSTALL" +
                        ",ACTIVE_PRICE" +
                        ",S.SPEED_ID AS DOWNLOAD" +
                        ",S.UPLOAD_SPEED / 1024 AS UPLOAD " +
                        "FROM HISPEED_PROMOTION P, HISPEED_CHANNEL_PROMOTION C, HISPEED_SPEED_PROMOTION S " +
                        "WHERE P.P_ID = S.P_ID AND P.P_ID = C.P_ID " +
                        "AND P.P_ID IN(" + id + ")";

            adapter = new OracleDataAdapter(sql, ConnectionProd);
            adapter.Fill(dt);

            backgroundWorker3.ReportProgress(8);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;

                backgroundWorker3.ReportProgress(8 + ((i + 1) * 12 / dt.Columns.Count));
            }

            sheet.get_Range("A1", "W1").Interior.Color = Excel.XlRgbColor.rgbAquamarine;
            sheet.get_Range("A1", "W1").Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;

            //Write data
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (j == 0 || j == 1 || j == 9 || j == 10)
                    {
                        string date = dt.Rows[i][j].ToString();

                        if (DateTime.TryParse(date, out DateTime dDate))
                        {
                            date = string.Format("{0:dd/MMM/yyyy}", dDate);
                            sheet.Cells[i + 2, j + 1] = date;
                        }
                        else
                        {
                            sheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                        }
                    }
                    else
                    {
                        sheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                    }
                }

                backgroundWorker3.ReportProgress(20 + ((i + 1) * 70 / dt.Rows.Count));
            }

            string exportFile = outputPath + "\\" + urNo.ToUpper() + "_Hispeed_Criteria.xlsx";
            workbook.SaveAs(exportFile, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            workbook.Close();
            xlApp.Quit();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            //export script
            if (String.IsNullOrEmpty(expHisp) == false)
            {
                string sqlPath = outputPath + "\\Script_HiSpeedPromotion_" + urNo.ToUpper() + ".txt";
                using (StreamWriter writer = new StreamWriter(sqlPath, true))
                {
                    writer.Write(expHisp);
                }
            }

            //export log
            if(String.IsNullOrEmpty(logHispeed) == false)
            {
                string logPath = outputPath + "\\Log_HiSpeedPromotion_" + urNo.ToUpper() + ".txt";
                using (StreamWriter writer = new StreamWriter(logPath, true))
                {
                    writer.Write(logHispeed);
                }
            }

            backgroundWorker3.ReportProgress(100);

            Cursor.Current = Cursors.Default;
            Application.UseWaitCursor = false;
        }

        private void ExportCampaign()
        {
            Application.UseWaitCursor = true;
            Cursor.Current = Cursors.WaitCursor;

            toolStripStatusLabel1.Text = "Exporting Campaign...";

            Excel.Application xlApp;
            Excel.Workbook workbook;
            Excel.Worksheet sheet;

            DataTable dt = new DataTable();
            OracleDataAdapter adapter;

            backgroundWorker3.ReportProgress(3);

            xlApp = new Excel.Application();
            workbook = xlApp.Workbooks.Add(Type.Missing);
            sheet = workbook.ActiveSheet as Excel.Worksheet;
            sheet.Name = "Report";
            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;

            tolPack = tolPack.Substring(1);
            tvsPack = tvsPack.Substring(1);

            //export file criteria
            string sql = "SELECT TOL_PACKAGE,TVS_PACKAGE,TOL_DISCOUNT, TVS_DISCOUNT,CAMPAIGN_NAME,STATUS " +
                    "FROM CAMPAIGN_MAPPING WHERE TMV_PACKAGE IS NULL " +
                    "AND TOL_PACKAGE IN(" + tolPack + ")" +
                    "AND TVS_PACKAGE IN(" + tvsPack + ")" +
                    "ORDER BY TOL_PACKAGE";

            adapter = new OracleDataAdapter(sql, ConnectionProd);
            adapter.Fill(dt);

            backgroundWorker3.ReportProgress(8);

            //Set column heading
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;

                backgroundWorker3.ReportProgress(8 + ((i + 1) * 12 / dt.Columns.Count));
            }

            sheet.get_Range("A1", "W1").Interior.Color = Excel.XlRgbColor.rgbAquamarine;
            sheet.get_Range("A1", "W1").Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;

            //Write data
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                }

                backgroundWorker3.ReportProgress(20 + ((i + 1) * 70 / dt.Rows.Count));
            }

            string exportFile = outputPath + "\\" + urNo.ToUpper() + "_CampaignMapping.xlsx";
            workbook.SaveAs(exportFile, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            workbook.Close();
            xlApp.Quit();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            //export script
            if (String.IsNullOrEmpty(expCamp) == false)
            {
                string strFilePath = outputPath + "\\Script_Campaign_" + urNo.ToUpper() + ".txt";
                using (StreamWriter writer = new StreamWriter(strFilePath, true))
                {
                    writer.Write(expCamp);
                }
            }

            //export log
            if (String.IsNullOrEmpty(logCampaign) == false)
            {
                string logPath = outputPath + "\\Log_Campaign_" + urNo.ToUpper() + ".txt";
                using (StreamWriter writer = new StreamWriter(logPath, true))
                {
                    writer.Write(logCampaign);
                }
            }

            backgroundWorker3.ReportProgress(100);

            Cursor.Current = Cursors.Default;
            Application.UseWaitCursor = false;
        }

        /// <summary>
        /// Get PName (Description of package) from file excel
        /// </summary>
        private string GetPName(string mkt)
        {
            string pName;
            if (lstPname.Count > 0)
            {
                if (lstPname.ContainsKey(mkt))
                {
                    pName = mkt + " - " + lstPname[mkt];
                }
                else
                {
                    pName = mkt;
                }
            }
            else
            {
                //string txt = "SELECT X.ATTRIB_04 MKT, S.NAME FROM SIEBEL.S_PROD_INT S , SIEBEL.S_PROD_INT_X  X WHERE S.ROW_ID " +
                //        " = X.ROW_ID AND X.ATTRIB_04 = '" + mkt + "'";

                string txt = "SELECT X.ATTRIB_04 MKT, S.NAME FROM S_PROD_INT S , S_PROD_INT_X  X WHERE S.ROW_ID " +
                        " = X.ROW_ID AND X.ATTRIB_04 = '" + mkt + "'";

                OracleCommand command = new OracleCommand(txt, ConnectionProd);
                OracleDataReader reader = command.ExecuteReader();
                reader.Read();
                if (reader.HasRows)
                {
                    pName = reader["NAME"].ToString();
                    reader.Close();
                }
                else
                {
                    pName = mkt;
                }
            }

            return pName;
        }

        /// <summary>
        /// Get SheetName from file
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <returns></returns>
        private static List<string> ToExcelsSheetList(string excelFilePath)
        {
            List<string> sheets = new List<string>();
            using (OleDbConnection connection =
                    new OleDbConnection((excelFilePath.TrimEnd().ToLower().EndsWith("x"))
                    ? "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + excelFilePath + "';" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
                    : "provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + excelFilePath + "';Extended Properties=Excel 8.0;"))
            {
                connection.Open();
                DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow drSheet in dt.Rows)
                    if (drSheet["TABLE_NAME"].ToString().Contains("$"))
                    {
                        string s = drSheet["TABLE_NAME"].ToString();
                        sheets.Add(s.StartsWith("'") ? s.Substring(1, s.Length - 3) : s.Substring(0, s.Length - 1));
                    }
                connection.Close();
            }

            return sheets;
        }

        private void InitialValue()
        {
            //Clear selection
            dataGridView1.ClearSelection();

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.Empty;
                }

                dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Empty;
            }

            //Clear list index
            indexListbox.Clear();
            //Clear listbox
            listBox1.Items.Clear();

            validateLog = "";
        }

        private void hilightRow(string type, string key, int indexRow)
        {
            Dictionary<string, int> indexDisc = new Dictionary<string, int>
            { {"month",1}, {"channel",2 },{"mkt",3},{"order",4},{"speed",6},{"province",7},{"start",8},{"end",9} };

            Dictionary<string, int> indexVas = new Dictionary<string, int>
            {{"channel",1 },{"mkt",2},{"order",3},{"speed",5},{"province",6},{"start",7},{"end",8} };

            Dictionary<string, int> indexHisp = new Dictionary<string, int>
            {{"mkt",1 },{"speed",2},{"subProfile",3},{"extra",4},{"order",6},{"channel",7},{"start",10},{"end",11},{"entry",12}, {"install",13} };

            Dictionary<string, int> indexCamp = new Dictionary<string, int>
            {{"type",0 },{"name",1},{"tolPack",2},{"tolDisc",3},{"tvsPack",4},{"tvsDisc",5} };

            if (type.Equals("VAS"))
            {
                int indexCol = indexVas[key];
                dataGridView1.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }
            else if (type.Equals("Disc"))
            {
                int indexCol = indexDisc[key];
                dataGridView1.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }
            else if (type.Equals("Hispeed"))
            {
                int indexCol = indexHisp[key];
                dataGridView1.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }
            else if (type.Equals("Campaign"))
            {
                int indexCol = indexCamp[key];
                dataGridView1.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }
        }
        #endregion
    }
}

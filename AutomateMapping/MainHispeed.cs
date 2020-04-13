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
using Excel = Microsoft.Office.Interop.Excel;

namespace AutomateMapping
{
    public partial class MainHispeed : Form
    {
        private OracleConnection ConnectionProd;
        private OracleConnection ConnectionTemp;
        private string filename;
        private string fileDesc;
        private string implementer;
        private string urNo;
        Dictionary<string, string> lstPname = new Dictionary<string, string>();
        List<string[]> lstChannel = new List<string[]>();
        List<string[]> lstSubProfile = new List<string[]>();
        Dictionary<int, string[]> lstSpeedMast = new Dictionary<int, string[]>();
        List<string[]> lstExtraProfile = new List<string[]>();
        DataTable tableContract = new DataTable();
        List<int> indexListbox = new List<int>();
        public MainHispeed(OracleConnection con, string file, string fDesc, string user, string ur)
        {
            InitializeComponent();
            ConnectionProd = con;
            filename = file;
            fileDesc = fDesc;
            implementer = user;
            urNo = ur;
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

        private void MainHispeed_Load(object sender, EventArgs e)
        {
            try
            {
                ConnectionTemp = new OracleConnection();
                string connStringTmp = "Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = 172.19.193.20)(PORT = 1560))" +
                       "(CONNECT_DATA = (SID = TEST03)));User Id= TRUREF71; Password= TRUREF71;";

                ConnectionTemp.ConnectionString = connStringTmp;
                ConnectionTemp.Open();

                DgvSettings dgvSettings = new DgvSettings();
                List<string> lstHeader = new List<string>();

                //Set header view hispeed promotion
                lstHeader.Add("Media");
                lstHeader.Add("MKTCode");
                lstHeader.Add("Speed");
                lstHeader.Add("Sub Profile");
                lstHeader.Add("Extra Msg");
                lstHeader.Add("Price");
                lstHeader.Add("Order Type");
                lstHeader.Add("Channel");
                lstHeader.Add("Modem Type");
                lstHeader.Add("Docsis Type");
                lstHeader.Add("Bundle Voice");
                lstHeader.Add("Effective");
                lstHeader.Add("Expire");
                lstHeader.Add("Entry Code");
                lstHeader.Add("Install Code");

                int i = dgvSettings.SetDgv(dataGridView1, filename, "HiSpeed Promotion$B3:P", lstHeader);

                if (i == -1)
                {
                    //set header view
                    lstHeader.Add("Type");
                    lstHeader.Add("Campaign Name");
                    lstHeader.Add("TOL Package");
                    lstHeader.Add("TOL Discount");
                    lstHeader.Add("TVS Package");
                    lstHeader.Add("TVS Discount");

                    i = dgvSettings.SetDgv(dataGridView1, filename, "Campaign Mapping$B2:G", lstHeader);

                    if (i == -1)
                    {
                        //show message not found both of the sheetname
                    }
                    else
                    {
                        ValidateFile("Campaign");
                    }
                }
                else
                {
                    ValidateFile("Hispeed");
                }
            }
            catch(Exception ex) 
            {
                //show message throw excep
                //show message ex.message
            }                   
        }

        private void ValidateFile(string type)
        {
            InitialValue();
            Validation validation = new Validation(ConnectionProd, ConnectionTemp);

            if (type == "Hispeed")
            {
                //Get Description
                if (String.IsNullOrEmpty(fileDesc) == false && lstPname.Count <= 0)
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

                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    string mkt = dataGridView1.Rows[i].Cells[1].Value.ToString().Trim();
                    string speed = dataGridView1.Rows[i].Cells[2].Value.ToString().Trim();
                    string subProfile = dataGridView1.Rows[i].Cells[3].Value.ToString().Trim();
                    string extra = dataGridView1.Rows[i].Cells[4].Value.ToString().ToUpper().Trim();
                    string order = dataGridView1.Rows[i].Cells[6].Value.ToString().Trim();
                    string channel = dataGridView1.Rows[i].Cells[7].Value.ToString().Trim();
                    string start = dataGridView1.Rows[i].Cells[11].Value.ToString().Trim();
                    string end = dataGridView1.Rows[i].Cells[12].Value.ToString().Trim();
                    string entry = dataGridView1.Rows[i].Cells[13].Value.ToString().Trim();
                    string install = dataGridView1.Rows[i].Cells[14].Value.ToString().Trim();

                    #region "Speed"
                    string[] msgSpeed = validation.CheckSpeed(lstSpeedMast, mkt, speed);
                    if(msgSpeed[0] != "Success")
                    {
                        listBox1.Items.Add(msgSpeed[0]);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "mkt", i);
                    }
                    else if (msgSpeed[1] != "Success")
                    {
                        listBox1.Items.Add(msgSpeed[1]);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "speed", i);
                    }
                    else if(msgSpeed[2] != "Success")
                    {
                        listBox1.Items.Add(msgSpeed[2]);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "mkt", i);
                        hilightRow("Hispeed", "speed", i);
                    }
                    #endregion

                    #region"Extra Profile"
                    string msgExtra = validation.CheckExtra(lstExtraProfile, extra);
                    if(msgExtra != "Success")
                    {
                        listBox1.Items.Add(msgExtra);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "extra", i);
                    }
                    #endregion

                    #region "SubProfile"
                    string msgSub = validation.CheckSubProfile(lstSubProfile, subProfile);
                    if (msgSub != "Success")
                    {
                        listBox1.Items.Add(msgSub);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "subProfile", i);
                    }
                    #endregion

                    #region "OrderType"
                    string msgOrder = validation.CheckOrderType(order);
                    if (msgOrder != "Success")
                    {
                        listBox1.Items.Add(msgOrder);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "order", i);
                    }
                    #endregion

                    #region "Channel"
                    string msgChannel = validation.CheckChannel(lstChannel, channel, end);
                    if (msgChannel != "Success")
                    {
                        listBox1.Items.Add(msgChannel);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "channel", i);
                    }
                    #endregion

                    #region "Date"
                    string msgDate = validation.CheckDate(start, end);
                    if (msgDate != "Success")
                    {
                        if(msgDate == "Start Date fotmat is not supported")
                        {
                            listBox1.Items.Add(msgDate);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "start", i);
                        }
                        else if (msgDate == "End Date fotmat is not supported")
                        {
                            listBox1.Items.Add(msgDate);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "end", i);
                        }
                        else
                        {
                            listBox1.Items.Add(msgDate);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "start", i);

                            listBox1.Items.Add(msgDate);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "end", i);
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
                    }
                    #endregion
                }
            }
            else
            {
                //Campaign
            }
        }

        private void MappingHiSpeed() 
        { 

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
            }

            //Clear list index
            indexListbox.Clear();
            //Clear listbox
            listBox1.Items.Clear();
        }

        private void hilightRow(string type, string key, int indexRow)
        {
            Dictionary<string, int> indexDisc = new Dictionary<string, int>
            { {"month",1}, {"channel",2 },{"mkt",3},{"order",4},{"speed",6},{"province",7},{"start",8},{"end",9} };

            Dictionary<string, int> indexVas = new Dictionary<string, int>
            {{"channel",1 },{"mkt",2},{"order",3},{"speed",5},{"province",6},{"start",7},{"end",8} };

            Dictionary<string, int> indexHisp = new Dictionary<string, int>
            {{"mkt",1 },{"speed",2},{"subProfile",3},{"extra",4},{"order",6},{"channel",7},{"start",11},{"end",12},{"entry",13}, {"install",14} };

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
            else
            {
                int indexCol = indexHisp[key];
                dataGridView1.Rows[indexRow].Cells[indexCol].Style.BackColor = Color.Red;
            }

        }

        private void MainHispeed_SizeChanged(object sender, EventArgs e)
        {
            int w = this.Size.Width;
            int h = this.Size.Height;

            btnClose.Location = new Point(w - 22, 13);
            btnMaximize.Location = new Point(w - 46, 13);
            btnMinimize.Location = new Point(w - 75, 13);


            btnExe.Location = new Point(w - 125, h - 87);
            btnLog.Location = new Point(w - 330, h - 87);

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

        private void btnMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
            InputHispeed inputHispeed = new InputHispeed(ConnectionProd, "");
            inputHispeed.Show();
        }

        private void listBox1_Click(object sender, EventArgs e)
        {
            dataGridView1.ClearSelection();
            if (listBox1.SelectedItem != null)
            {
                int selected = listBox1.SelectedIndex;
                dataGridView1.Rows[indexListbox[selected]].Selected = true;
            }
        }

        private void btnExe_Click(object sender, EventArgs e)
        {
           
        }
    }
}

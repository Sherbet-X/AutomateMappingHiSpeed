using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OracleClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutomateMapping
{
    public partial class MainHispeed : Form
    {
        private OracleConnection ConnectionProd;
        private OracleConnection ConnectionTemp;
        private string filename , fileDesc, implementer, urNo, outputPath, validateLog, sysLog;
        Dictionary<string, string> lstPname = new Dictionary<string, string>();
        List<string[]> lstChannel = new List<string[]>();
        List<string[]> lstSubProfile = new List<string[]>();
        Dictionary<int, string[]> lstSpeedMast = new Dictionary<int, string[]>();
        List<string[]> lstExtraProfile = new List<string[]>();
        DataTable tableContract = new DataTable();
        DataTable tableProdType = new DataTable();
        List<int> indexListbox = new List<int>();

        Validation validation;
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

                int i = dgvSettings.SetDgv(dataGridView1, filename, "HiSpeed Promotion$B3:O", lstHeader);

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
                        MessageBox.Show("Sheet name : 'HiSpeed Promotion' and 'Campaign Mapping' Not Found!!");
                        Application.Exit();
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
            validation = new Validation(ConnectionProd, ConnectionTemp);

            if (type == "Hispeed")
            {
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
                    if(msgSpeed[0] != "Success")
                    {
                        listBox1.Items.Add(msgSpeed[0]);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "mkt", i);

                        validateLog += msgSpeed[0] + "\r\n";
                    }

                    if (msgSpeed[1] != "Success")
                    {
                        listBox1.Items.Add(msgSpeed[1]);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "speed", i);

                        validateLog += msgSpeed[1] + "\r\n";
                    }

                    if(msgSpeed[2] != "Success")
                    {
                        listBox1.Items.Add(msgSpeed[2]);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "mkt", i);
                        hilightRow("Hispeed", "speed", i);

                        validateLog += msgSpeed[2] + "\r\n";
                    }
                    #endregion

                    #region"Extra Profile"
                    string msgExtra = validation.CheckExtra(lstExtraProfile, extra);
                    if(msgExtra != "Success")
                    {
                        listBox1.Items.Add(msgExtra);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "extra", i);

                        validateLog += msgExtra + "\r\n";
                    }
                    #endregion

                    #region "SubProfile"
                    string msgSub = validation.CheckSubProfile(lstSubProfile, subProfile);
                    if (msgSub != "Success")
                    {
                        listBox1.Items.Add(msgSub);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "subProfile", i);

                        validateLog += msgSub + "\r\n";
                    }
                    #endregion

                    #region "OrderType"
                    string msgOrder = validation.CheckOrderType(order);
                    if (msgOrder != "Success")
                    {
                        listBox1.Items.Add(msgOrder);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "order", i);

                        validateLog += msgOrder + "\r\n";
                    }
                    #endregion

                    #region "Channel"
                    string msgChannel = validation.CheckChannel(lstChannel, channel, end);
                    if (msgChannel != "Success")
                    {
                        listBox1.Items.Add(msgChannel);
                        indexListbox.Add(i);
                        hilightRow("Hispeed", "channel", i);

                        validateLog += msgChannel + "\r\n";
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

                            validateLog += msgDate + "\r\n";
                        }
                        else if (msgDate == "End Date fotmat is not supported")
                        {
                            listBox1.Items.Add(msgDate);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "end", i);

                            validateLog += msgDate + "\r\n";
                        }
                        else
                        {
                            listBox1.Items.Add(msgDate);
                            indexListbox.Add(i);
                            hilightRow("Hispeed", "start", i);
                            hilightRow("Hispeed", "end", i);

                            validateLog += msgDate + "\r\n";
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

                        validateLog += msgContract + "\r\n";
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
            ReserveID reserveID = new ReserveID();
            reserveID.Reserve(ConnectionProd, ConnectionTemp, "Hispeed", implementer, urNo);

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
                string modem = dataGridView1.Rows[i].Cells[8].Value.ToString().Trim();
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

                string[] lstOrder;
                if (order.Contains(","))
                {
                    lstOrder = order.Split('-');
                }
                else
                {
                    lstOrder = new string[1];
                    lstOrder[0] = order;
                }

                //Get P_Name
                string p_name;
                if (lstPname.Count <= 0)
                {
                    p_name = GetPName(dataGridView1.Rows[i].Cells[1].Value.ToString());
                }
                else
                {
                    p_name = lstPname[dataGridView1.Rows[i].Cells[1].Value.ToString()];
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

                    string txt = "SELECT P.P_ID, P.P_CODE, P.ORDER_TYPE,C.SALE_CHANNEL ,P.START_DATE,P.END_DATE,P.STATUS, " +
                        "S.SPEED_ID DOWNLOAD, S.UPLOAD_SPEED / 1024 UPLOAD,S.PRICE FROM HISPEED_PROMOTION P, " +
                        "HISPEED_SPEED_PROMOTION S,HISPEED_CHANNEL_PROMOTION C WHERE P.P_ID = S.P_ID AND P.P_ID = C.P_ID " +
                        "AND P_CODE = '" + mkt + "' AND SPEED_ID = '" + suffix + "' AND ORDER_TYPE = '" + lstOrder[j] + "'";
                    OracleCommand cmd = new OracleCommand(txt, ConnectionProd);
                    OracleDataReader reader = cmd.ExecuteReader();

                    if (reader.HasRows)
                    {
                        //Existing
                        reader.Read();
                        int id = Convert.ToInt32(reader["P_ID"]);

                        ExistingData(id, suffix, uploadK, channel, price, start,
                            end, dataGridView1.Rows[i].Cells[1].Value.ToString());


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
                        lstData.Add(modem);
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
            string modem = data[8];
            string docsis = data[9];
            string start = data[10];
            string end = data[11];
            string entryF = data[12];
            string installF = data[13];
            string pName = data[14];
            string suffix = data[15];

            if (String.IsNullOrEmpty(channel))
            {
                //alert msg and write log
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

                    ////Get Contract
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

                    //insert new data
                    cmd = ConnectionProd.CreateCommand();
                    OracleTransaction transaction = null;
                    using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
                    {
                        try
                        {
                            //Insert into hispeed_promotion
                            cmd.CommandText = "INSERT INTO HISPEED_PROMOTION VALUES (" + minID + ", '" + mkt + "', '" + mkt + "', '" + pName +
                                "', '" + pName + "', '" + order + "', 'Active','"+ extra +"','',0,0,'Y','Y','',0,'N','0','Y','Y','N','" + prodType +
                                "', sysdate, sysdate, '" + term + "',0,'TI', TO_DATE('" + start + "','dd/mm/yyyy'), " +
                                "TO_DATE('" + end + "','dd/mm/yyyy'), 'M', '" + mkt + "','N','N','Y', '" + entry + "', '" +
                                install + "','" + modem + "','N','" + sub + "','')";
                            cmd.ExecuteNonQuery();

                            //Insert into hispeed_speed_promotion
                            cmd.CommandText = "INSERT INTO HISPEED_SPEED_PROMOTION  VALUES (" + suffix + ", " + minID + ", " + 
                                price + ", null, 'Y', '" + suffix + "', '" + modem + "', " +"'" + uploadK + "', '" + docsis + "')";
                            cmd.ExecuteNonQuery();

                            string[] arrChannel;
                            if(channel.Contains(","))
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
                                cmd.ExecuteNonQuery();
                            }
     
                            transaction.Commit();
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            //write log cannot insert data
                        }
                    }
                }
                catch (Exception ex)
                {
                    //write log cannot insert data
                }
            }
        }

        private void ExistingData(int id, int suffix, int upload, string channel, double price, string start, string end, string mkt)
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
            OracleTransaction transaction = null;
            using (transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted))
            {
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
                                    cmd.ExecuteNonQuery();                                    
                                }
                                else
                                {
                                    Dictionary<string, string[]> lstChannelDB = new Dictionary<string, string[]>();
                                    OracleCommand command = new OracleCommand("SELECT * FROM HISPEED_CHANNEL_PROMOTION WHERE P_ID = " + 
                                        id, ConnectionProd);
                                    OracleDataReader dataReader = command.ExecuteReader();

                                    while (dataReader.Read())
                                    {
                                        string[] date = new string[2];
                                        date[0] = reader["START_DATE"].ToString();
                                        date[1] = reader["END_DATE"].ToString();
                                        lstChannelDB.Add(reader["SALE_CHANNEL"].ToString(), date);
                                    }

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
                                        if (lstChannelDB.Keys.Contains(ch))
                                        {
                                            string[] date = lstChannelDB[ch];
                                            DateTime startDB = Convert.ToDateTime(date[0]);
                                            DateTime endDB = Convert.ToDateTime(date[1]);

                                            DateTime startF = Convert.ToDateTime(start);
                                            DateTime endF = Convert.ToDateTime(end);

                                            if (endDB < DateTime.Now)
                                            {
                                                if (startF == DateTime.Now)
                                                {
                                                    //update startdate == datetime.now
                                                    cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET START_DATE = sysdate "+
                                                        "WHERE P_ID = " + id;
                                                    cmd.ExecuteNonQuery();
                                                }
                                                else
                                                {
                                                    if (startF > DateTime.Now)
                                                    {
                                                        //update start date by date on file
                                                        cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET START_DATE = TO_DATE('" + 
                                                            start + "', 'dd/MM/yyyy') WHERE P_ID = " + id;
                                                        cmd.ExecuteNonQuery();
                                                    }
                                                    else
                                                    {
                                                        if (endF == DateTime.Now)
                                                        {
                                                            //update enddate = datetime sysdate
                                                            cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET END_DATE = sysdate " +
                                                                "WHERE P_ID = " + id;
                                                            cmd.ExecuteNonQuery();
                                                        }
                                                        else
                                                        {
                                                            //update enddate = end on file
                                                            cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET END_DATE = TO_DATE('" +
                                                            end + "', 'dd/MM/yyyy') WHERE P_ID = " + id;
                                                            cmd.ExecuteNonQuery();
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (startDB > DateTime.Now)
                                                {
                                                    //update start = date sysdate
                                                    cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET START_DATE = sysdate " +
                                                                "WHERE P_ID = " + id;
                                                    cmd.ExecuteNonQuery();
                                                }
                                                else
                                                {
                                                    if (endF == DateTime.Now)
                                                    {
                                                        //update enddate = datetime sysdate
                                                        cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET END_DATE = sysdate " +
                                                                "WHERE P_ID = " + id;
                                                        cmd.ExecuteNonQuery();
                                                    }
                                                    else
                                                    {
                                                        //update enddate = end on file
                                                        cmd.CommandText = "UPDATE HISPEED_CHANNEL_PROMOTION SET END_DATE = TO_DATE('" +
                                                            end + "', 'dd/MM/yyyy') WHERE P_ID = " + id;
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
                                    cmd.ExecuteNonQuery();
                                }
                                else
                                {
                                    //write log
                                    //"MKT: " + mkt + ", price[" + price + "] on file is not matching price[" + priceDB + "] on DB";
                                }
                            }
                        }
                        else
                        {
                            if (price == priceDB)
                            {
                                //update active price = y
                                cmd.CommandText = "UPDATE HISPEED_SPEED_PROMOTION SET ACTIVE_PRICE = 'Y' WHERE P_ID = " + id;
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
                                    cmd.ExecuteNonQuery();
                                }
                                else
                                {
                                    //write log
                                    //"MKT: " + mkt + ", price[" + price + "] on file is not matching price[" + priceDB + "] on DB";
                                }
                            }
                        }
                    }
                    else
                    {
                        ///write log
                        // Download or Upload Speed of "+mkt+" not matching on Database!!
                    }

                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                }
            }
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

            validateLog = "";
        }

        /// <summary>
        /// Get PName (Description of package) from file excel
        /// </summary>
        private string GetPName(string mkt)
        {
            string pName = "";
            string txt = "SELECT X.ATTRIB_04 MKT, S.NAME FROM SIEBEL.S_PROD_INT S , SIEBEL.S_PROD_INT_X  X WHERE S.ROW_ID " +
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

            return pName;
        }

        private void hilightRow(string type, string key, int indexRow)
        {
            Dictionary<string, int> indexDisc = new Dictionary<string, int>
            { {"month",1}, {"channel",2 },{"mkt",3},{"order",4},{"speed",6},{"province",7},{"start",8},{"end",9} };

            Dictionary<string, int> indexVas = new Dictionary<string, int>
            {{"channel",1 },{"mkt",2},{"order",3},{"speed",5},{"province",6},{"start",7},{"end",8} };

            Dictionary<string, int> indexHisp = new Dictionary<string, int>
            {{"mkt",1 },{"speed",2},{"subProfile",3},{"extra",4},{"order",6},{"channel",7},{"start",10},{"end",11},{"entry",12}, {"install",13} };

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

            btnValidate.Location = new Point(w-35, h - 327);
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
            MappingHiSpeed();             
        }

        private void btnLog_Click(object sender, EventArgs e)
        {
            if(String.IsNullOrEmpty(validateLog))
            {
                MessageBox.Show("The verification process is complete. No errors occurred during process.");
            }
            else
            {
                string strFilePath = outputPath + "\\ValidateLog_" + urNo +"_"+DateTime.Now +".txt";
                using (StreamWriter writer = new StreamWriter(strFilePath, true))
                {
                    writer.Write(validateLog);
                }

                Application.Exit();
            }
        }
    }
}

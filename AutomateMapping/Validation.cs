using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.OracleClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutomateMapping
{
    class Validation
    {
        private OracleConnection ConnectionProd;
        private OracleConnection ConnectionTemp;

        public Validation(OracleConnection connProd, OracleConnection connTemp)
        {
            this.ConnectionProd = connProd;
            this.ConnectionTemp = connTemp;
        }

        #region "Get Data"
        //Get Description from file New MKT
        public Dictionary<string, string> GetDescription(string file)
        {
            Dictionary<string, string> lstPname = new Dictionary<string, string>();

            if (System.IO.File.Exists(file))
            {
                string connString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;""", file);
                string query = string.Format("select * from [VCARE-MKT$B3:D]", connString);
                OleDbDataAdapter dtAdapter = new OleDbDataAdapter(query, connString);
                DataSet ds = new DataSet();
                dtAdapter.Fill(ds);
                DataTable dt = ds.Tables[0];

                foreach (DataRow dr in dt.Rows)
                {
                    string mkt = dr[1].ToString();
                    string name = dr[3].ToString();

                    if (String.IsNullOrEmpty(mkt) == false && String.IsNullOrEmpty(name) == false)
                    {
                        lstPname.Add(mkt, name);
                    }
                }
            }

            return lstPname;
        }

        List<string[]> _lstChannel;
        public List<string[]> GetChannelFromDB
        {
            get
            {
                if (this._lstChannel is null || this._lstChannel.Count <= 0)
                {
                    this._lstChannel = getChannelFromDB();
                }

                return this._lstChannel;
            }
            set
            {
                this._lstChannel = value;
            }
        }

        //Get sale channel from database
        public List<string[]> getChannelFromDB()
        {
            List<string[]> list = new List<string[]>();

            try
            {
                string query = "SELECT VALUE1 FROM TRUE9_BPT_VALIDATE WHERE TYPE = 'SALE CHANNEL' AND NAME1 = 'SALE CHANNEL'";
                OracleCommand cmd = new OracleCommand(query, ConnectionTemp);
                OracleDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    string[] arr = new string[2];
                    arr[0] = reader.GetValue(0).ToString();
                    arr[1] = "DB";

                    list.Add(arr);
                }

                reader.Close();
            }
            catch (Exception)
            {
                string msg = "Cannot get sale channel from database.";

                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                list.Clear();
            }

            return list;
        }

        List<string[]> _lstSubProfile;
        public List<string[]> GetSubProfile
        {
            get
            {
                if (this._lstSubProfile is null || this._lstSubProfile.Count <= 0)
                {
                    this._lstSubProfile = getSubProfile();
                }

                return this._lstSubProfile;
            }
            set
            {
                this._lstSubProfile = value;
            }
        }
        private List<string[]> getSubProfile()
        {
            List<string[]> list = new List<string[]>();

            try
            {
                string query = "SELECT VALUE1 FROM TRUE9_BPT_VALIDATE WHERE TYPE = 'SUB_PROFILE' AND NAME1 = 'SUB_PROFILE'";
                OracleCommand cmd = new OracleCommand(query, ConnectionTemp);
                OracleDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    string[] arr = new string[2];
                    arr[0] = reader.GetValue(0).ToString();
                    arr[1] = "DB";

                    list.Add(arr);
                }

                reader.Close();
            }
            catch (Exception)
            {
                string msg = "Cannot get SubProfile from database.";

                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                list.Clear();
            }

            return list;
        }

        List<string[]> _lstExtra;
        public List<string[]> GetExtraProfile
        {
            get
            {
                if (this._lstExtra is null || this._lstExtra.Count <= 0)
                {
                    this._lstExtra = getExtraProfile();
                }

                return this._lstExtra;
            }
            set
            {
                this._lstExtra = value;
            }
        }
        private List<string[]> getExtraProfile()
        {
            List<string[]> list = new List<string[]>();

            try
            {
                string query = "SELECT VALUE1 FROM TRUE9_BPT_VALIDATE WHERE TYPE = 'EXTRA_PROFILE' AND NAME1 = 'EXTRA_PROFILE'";
                OracleCommand cmd = new OracleCommand(query, ConnectionTemp);
                OracleDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    string[] arr = new string[2];
                    arr[0] = reader.GetValue(0).ToString();
                    arr[1] = "DB";

                    list.Add(arr);
                }

                reader.Close();
            }
            catch (Exception)
            {
                string msg = "Cannot get Extra Profile from database.";

                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                list.Clear();
            }
            return list;
        }

        Dictionary<int, string[]> _speedFromDB;
        public Dictionary<int, string[]> GetSpeedFromDB
        {
            get
            {
                if(this._speedFromDB is null || this._speedFromDB.Count <= 0 )
                {
                    this._speedFromDB = GetSpeed();
                }

                return this._speedFromDB;
            }
            set
            {
                this._speedFromDB = value;
            }
        }
        private Dictionary<int, string[]> GetSpeed()
        {
            Dictionary<int, string[]> list = new Dictionary<int, string[]>();

            try
            {
                string query = "SELECT SPEED_ID, SPEED_DESC FROM HISPEED_SPEED";
                OracleCommand cmd = new OracleCommand(query, ConnectionProd);
                OracleDataReader reader = cmd.ExecuteReader();
                
                while (reader.Read())
                {
                    string[] arr = new string[2];
                    arr[0] = reader.GetValue(1).ToString();
                    arr[1] = "DB";

                    list.Add(Convert.ToInt32(reader.GetValue(0)), arr);
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                string msg = "Cannot get speed from database.";

                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                list.Clear();
            }

            return list;
        }
        public DataTable GetContract()
        {
            DataTable dataTable = new DataTable();

            try
            {
                string query = "SELECT NAME1,VALUE1,NAME2,VALUE2,NAME3,VALUE3 FROM TRUE9_BPT_VALIDATE WHERE TYPE = 'CONTRACT'";
                OracleDataAdapter adapter = new OracleDataAdapter(query, ConnectionTemp);
                adapter.Fill(dataTable);

            }
            catch (Exception)
            {
                string msg = "Cannot get contract from database.";

                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return dataTable;
        }

        public DataTable GetProdType()
        {
            DataTable dataTable = new DataTable();

            try
            {
                string query = "SELECT VALUE1, VALUE2,VALUE3 FROM TRUE9_BPT_VALIDATE WHERE TYPE = 'MEDIA' AND NAME1 = 'MEDIA'";
                OracleDataAdapter adapter = new OracleDataAdapter(query, ConnectionTemp);
                adapter.Fill(dataTable);

            }
            catch (Exception)
            {
                string msg = "Cannot get prodtype from database.";

                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return dataTable;
        }
        #endregion

        public string[] CheckSpeed(Dictionary<int, string[]> lstSpeed4Chk, string mkt, string speed)
        {
            string[] msg = new string[3];
            int speedID = -1;
            int suffixID = -1;

            Dictionary<int, string> resultCheckSuffixID = CheckFormatMKT(lstSpeed4Chk, mkt, speed);
            Dictionary<int, string> resultCheckUOM = CheckUOMSpeed(lstSpeed4Chk, speed);

            if (resultCheckSuffixID.ContainsKey(-1))
            {
                msg[0] = resultCheckSuffixID[-1];
            }
            else
            {
                suffixID = resultCheckSuffixID.FirstOrDefault(x => x.Value == "Success").Key;
                msg[0] = "Success";
            }

            if (resultCheckUOM.ContainsKey(-1))
            {
                msg[1] = resultCheckUOM[-1];
            }
            else
            {
                speedID = resultCheckUOM.FirstOrDefault(x => x.Value == "Success").Key;
                msg[1] = "Success";
            }

            if(speedID != -1 && suffixID != -1)
            {
                if(speedID != suffixID)
                {
                    msg[2] = "Suffix and download speed of MKT: "+mkt+" not matching!!";
                }
                else
                {
                    msg[2] = "Success";
                }
            }
            return msg;
        }

        private Dictionary<int,string> CheckUOMSpeed(Dictionary<int, string[]> lstSpeed4Chk, string speed)
        {
            Dictionary<int, string> result = new Dictionary<int, string>();
            int speedID = -1;
            string status = "";

            if (speed.Contains('/'))
            {
                string[] splitSpeed = speed.Split('/');
                string download = splitSpeed[0].Trim();
                string upload = splitSpeed[1].Trim();

                if(int.TryParse(upload, out _) == false)
                {
                    string uomDownload;

                    if (int.TryParse(download, out _) == false)
                    {
                        uomDownload = Regex.Replace(download, "[0-9]", "");
                    }
                    else
                    {
                        uomDownload = Regex.Replace(upload, "[0-9]", "");
                    }

                    int speed2K = ConvertUOM2K(download, uomDownload);

                    if(speed2K == -1)
                    {
                        //write log invalid UOM
                        result.Add(speedID, "Invalid UOM of speed: " + speed + ".");
                    }
                    else
                    {
                        foreach (KeyValuePair<int, string[]> keyValuePair in lstSpeed4Chk)
                        {
                            string[] val = keyValuePair.Value;

                            if (val[0] == speed2K.ToString())
                            {
                                speedID = keyValuePair.Key;
                                status = val[1];
                                break;
                            }
                        }

                        if (speedID == -1)
                        {
                            DialogResult dialog = MessageBox.Show("Do you want to insert new speed into DB[Hispeed Speed]?"+
                                "Detail :"+"\r\n"+"SpeedID = "+speedID+" Desc = "+speed2K, "Confirmation",
                                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                            if (dialog == DialogResult.Yes)
                            {
                                //insert speed into table
                                OracleTransaction transaction = null;
                                speedID = Convert.ToInt32(Regex.Replace(download, "[^0-9]", ""));
                                string cmd = "INSERT INTO HISPEED_SPEED VALUES (" + speedID + ",'" + speed2K + "','" +
                                speed2K + "K','" + speed2K + "')";
                              
                                try
                                {
                                    transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted);
                                    OracleCommand command = new OracleCommand(cmd, ConnectionProd);
                                    command.Transaction = transaction;
                                    command.ExecuteNonQuery();

                                    transaction.Commit();

                                    this.GetSpeedFromDB = GetSpeed();

                                    result.Add(speedID, "Success")
;                                }
                                catch (Exception ex)
                                {
                                    transaction.Rollback();

                                    string[] arr = { speed2K.ToString(), "ignore" };
                                    this.GetSpeedFromDB.Add(speedID,arr);

                                    result.Add(-1, "Not found speedID: " + speedID + " on Master Data.");
                                }
                            }
                            else if (dialog == DialogResult.No)
                            {
                                string[] arr = { speed2K.ToString(), "ignore" };
                                this.GetSpeedFromDB.Add(speedID, arr);

                                result.Add(-1, "Not found speedID: " + speedID + " on Master Data.");
                            }
                            else
                            {
                                Environment.Exit(0);
                            }
                        }
                        else
                        {
                            if(status == "ignore")
                            {
                                //write log not found speedID in database
                                result.Add(-1, "Not found speedID: " + speedID + " on Master Data.");
                            }
                            else
                            {
                                result.Add(speedID, "Success");
                            }
                        }
                    }
                }
                else
                {
                    //write log not found UOM
                    result.Add(speedID, "Not found UOM of speed: " + speed + ".");
                }              
            }
            else
            {
                //write log wrong format speed
                result.Add(speedID, "Speed:"+speed+" format is not supported");
            }

            return result;
        }
        private Dictionary<int, string> CheckFormatMKT(Dictionary<int, string[]> lstSpeed4Chk, string mkt, string speed)
        {
            int speedID = -1;
            Dictionary<int, string> result = new Dictionary<int, string>();

            if (mkt.Contains("-"))
            {
                string[] lstmkt = mkt.Split('-');
                string suffixMkt = lstmkt[1].Trim();

                if (suffixMkt == "00" || suffixMkt == "01")
                {
                    Dictionary<int, string> resultCheckUOM = CheckUOMSpeed(lstSpeed4Chk, speed);

                    if (resultCheckUOM.ContainsKey(-1))
                    {
                        result.Add(-1, resultCheckUOM[-1]);
                    }
                    else
                    {
                        speedID = resultCheckUOM.FirstOrDefault(x => x.Value == "Success").Key;

                        result.Add(speedID, "Success");
                    }
                }
                else
                {
                    int speed2K = 0;
                    string status = "";

                    if (suffixMkt.EndsWith("G"))
                    {
                        suffixMkt = suffixMkt.Substring(0, suffixMkt.Length - 1);
                        speed2K = Convert.ToInt32(suffixMkt) * 1000 * 1024;
                    }
                    else
                    {
                        if (int.TryParse(suffixMkt, out _))
                        {
                            speed2K = Convert.ToInt32(suffixMkt) * 1024;                          
                        }
                        else
                        {
                            result.Add(-1, "Suffix of MKT: "+mkt+" is Wrong!!");
                        }
                    }

                    //Searching speedID from list
                    if(speed2K != 0)
                    {
                        foreach (KeyValuePair<int, string[]> keyValuePair in lstSpeed4Chk)
                        {
                            string[] val = keyValuePair.Value;

                            if (val[0] == speed2K.ToString())
                            {
                                speedID = keyValuePair.Key;
                                status = val[1];
                                break;
                            }
                        }

                        if (speedID == -1)
                        {
                            DialogResult dialog = MessageBox.Show("Do you want to insert new speed into DB[Hispeed Speed] ? "+
                                "Detail :" + "\r\n" + "SpeedID = " +suffixMkt + " Desc = " + speed2K, "Confirmation",
                                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                            if (dialog == DialogResult.Yes)
                            {
                                //insert new speedID
                                OracleTransaction transaction = null;
                                string cmd = "INSERT INTO HISPEED_SPEED VALUES (" + Convert.ToInt32(suffixMkt) + ",'" + speed2K + "','" +
                                speed2K + "K','" + speed2K + "')";

                                try
                                {
                                    transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted);
                                    OracleCommand command = new OracleCommand(cmd, ConnectionProd);
                                    command.Transaction = transaction;
                                    command.ExecuteNonQuery();

                                    transaction.Commit();

                                    result.Add(Convert.ToInt32(suffixMkt), "Success");
                                    this.GetSpeedFromDB = GetSpeed();
                                }
                                catch (Exception ex)
                                {
                                    transaction.Rollback();

                                    string[] arr = { speed2K.ToString(), "ignore" };
                                    this.GetSpeedFromDB.Add(Convert.ToInt32(suffixMkt), arr);
                                    result.Add(Convert.ToInt32(suffixMkt), "Not found SpeedID: "+suffixMkt+" on Master Data.");
                                }
                            }
                            else if (dialog == DialogResult.No)
                            {
                                string[] arr = { speed2K.ToString(), "ignore" };
                                this.GetSpeedFromDB.Add(Convert.ToInt32(suffixMkt), arr);
                                result.Add(Convert.ToInt32(suffixMkt), "Not found SpeedID: " + suffixMkt + " on Master Data.");
                            }
                            else
                            {
                                Environment.Exit(0);
                            }
                        }
                        else
                        {
                            if (status == "ignore")
                            {
                                result.Add(-1, "Not found SpeedID: "+suffixMkt+" on Master Data.");
                            }
                            else
                            {
                                result.Add(speedID, "Success");
                            }
                        }
                    }
                }
            }
            else
            {
                result.Add(-1, "MKT Code:"+mkt+" format is not supported");
            }

            return result;
        }     
        public int ConvertUOM2K(string speed, string uom)
        {
            int convSpeed = Convert.ToInt32(Regex.Replace(speed.Trim(), "[^0-9]", ""));
            uom = uom.ToUpper().Trim();

            if (uom == "G")
            {
                convSpeed = convSpeed * 1024000;
            }
            else if (uom == "M")
            {
                convSpeed = convSpeed * 1024;
            }
            else
            {
               if(uom != "K")
                {
                    convSpeed = -1;
                }
            }

            return convSpeed;
        }
        public string CheckExtra(List<string[]> lstExtraProf, string extra)
        {
            string msg = "Success";

            if (extra == "NA" || extra == "N/A" || extra == "-" || extra == "" || extra == "NULL")
            {
                extra = null;
            }

            if (String.IsNullOrEmpty(extra) == false)
            {
                bool hasExtra = lstExtraProf.Any(p => p.SequenceEqual(new string[] { extra, "DB" }));
                if (hasExtra == false)
                {
                    DialogResult dialog = MessageBox.Show("Do you want to insert new extra profile?" + "\r\n" +
                                "Detail :" + "\r\n" + "Extra Profile = " + extra, "Confirmation",
                                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    if (dialog == DialogResult.Yes)
                    {
                        OracleTransaction transaction = null;
                        string cmd = "INSERT INTO TRUE9_BPT_VALIDATE (TYPE, NAME1, VALUE1) VALUES('EXTRA_PROFILE', " +
                            "'EXTRA_PROFILE', '" + extra + "')";
                        try
                        {
                            transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted);
                            OracleCommand command = new OracleCommand(cmd, ConnectionTemp);
                            command.Transaction = transaction;
                            command.ExecuteNonQuery();

                            transaction.Commit();

                            string[] arr = { extra, "DB" };
                            this.GetExtraProfile.Add(arr);
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();

                            string[] arr = { extra, "ignore" };
                            this.GetExtraProfile.Add(arr);

                            msg = "Cannot insert new extra profile : " + extra + " on Master Date";
                        }
                    }
                    else if (dialog == DialogResult.No)
                    {
                        string[] arr = { extra, "ignore" };
                        this.GetExtraProfile.Add(arr);

                        msg = "Not found extra profile : " + extra + " on Master Data";
                    }
                    else
                    {
                        Environment.Exit(0);
                    }
                }
            }

            return msg;
        }

        public string CheckSubProfile(List<string[]> lstSubProfile, string subProfile)
        {
            string msg = "Success";

            if (subProfile == "NA" || subProfile == "N/A" || subProfile == "-" 
                || subProfile == "" || subProfile == "NULL")
            {
                subProfile = null;
            }

            if(subProfile.StartsWith("STL"))
            {
                subProfile = "STL (stand alone)";
            }

            if (String.IsNullOrEmpty(subProfile) == false)
            {
                bool hasSub= lstSubProfile.Any(p => p.SequenceEqual(new string[] { subProfile, "DB" }));
                if (hasSub == false)
                {
                    DialogResult dialog = MessageBox.Show("Do you want to insert new SubProfile?" + "\r\n" +
                                "Detail :" + "\r\n" + "SubProfile = " + subProfile, "Confirmation",
                                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    if (dialog == DialogResult.Yes)
                    {
                        OracleTransaction transaction = null;
                        string cmd = "INSERT INTO TRUE9_BPT_VALIDATE (TYPE, NAME1, VALUE1) VALUES('SUB_PROFILE', " +
                            "'SUB_PROFILE', '" + subProfile + "')";
                        try
                        {
                            transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted);
                            OracleCommand command = new OracleCommand(cmd, ConnectionTemp);
                            command.Transaction = transaction;
                            command.ExecuteNonQuery();

                            transaction.Commit();

                            string[] arr = { subProfile, "DB" };
                            this.GetSubProfile.Add(arr);
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();

                            string[] arr = { subProfile, "ignore" };
                            this.GetSubProfile.Add(arr);

                            msg = "Cannot insert new SubProfile : " + subProfile + " on Master Data";
                        }
                    }
                    else if (dialog == DialogResult.No)
                    {
                        string[] arr = { subProfile, "ignore" };
                        this.GetSubProfile.Add(arr);

                        msg = "Not found SubProfile : " + subProfile + " on Master Data";
                    }
                    else
                    {
                        Environment.Exit(0);
                    }
                }
            }

            return msg;
        }

        public string CheckOrderType(string order)
        {
            string[] lstOrder = null;
            string msg = "Success";

            if (order.Contains(','))
            {
                lstOrder = new string[2];
                lstOrder = order.Split(',');
            }
            else
            {
                lstOrder = new string[1];
                lstOrder[0] = order;
            }

            for (int i = 0; i < lstOrder.Length; i++)
            {
                if(lstOrder[i].ToUpper().Trim() != "NEW" &&
                    lstOrder[i].ToUpper().Trim() != "CHANGE")
                {
                    msg = "Order type :"+ lstOrder[i]+" is invalid";
                }
            }

            return msg;
        }

        public string CheckChannel(List<string[]> lstChannelFromDB, string channel, string endDate)
        {
            string msg = "Success";
            string[] lstChannel = null;
            if (String.IsNullOrEmpty(channel) && String.IsNullOrEmpty(endDate))
            {
                msg = "Channel and End Date is empty";
            }
            else
            {
                channel = Regex.Replace(channel, "ALL", "DEFAULT", RegexOptions.IgnoreCase);

                if (channel.Contains(","))
                {
                    lstChannel = channel.Split(',');
                    if (lstChannel.Contains("DEFAULT"))
                    {
                        msg = "Channel 'ALL/DEFAULT' included with other channel in same MKT code";
                    }
                }
                else
                {
                    lstChannel = new string[1];
                    lstChannel[0] = channel;
                }

                for (int i = 0; i < lstChannel.Length; i++)
                {
                    string ch = lstChannel[i].Trim();
                    bool hasChannel = lstChannelFromDB.Any(p => p.SequenceEqual(new string[] { ch, "DB" }));
                    if (hasChannel == false)
                    {
                        DialogResult dialog = MessageBox.Show("Do you want to insert new channel to master table?" + "\r\n" +
                                    "Detail :" + "\r\n" + "Sale_Channel = " + ch, "Confirmation",
                                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                        if (dialog == DialogResult.Yes)
                        {
                            OracleTransaction transaction = null;
                            string cmd = "INSERT INTO TRUE9_BPT_VALIDATE (TYPE, NAME1, VALUE1) VALUES('SALE_CHANNEL', " +
                                "'SALE_CHANNEL', '" + ch + "')";
                            try
                            {
                                transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted);
                                OracleCommand command = new OracleCommand(cmd, ConnectionTemp);
                                command.Transaction = transaction;
                                command.ExecuteNonQuery();

                                transaction.Commit();

                                string[] arr = { ch, "DB" };
                                this.GetChannelFromDB.Add(arr);
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();

                                string[] arr = { ch, "ignore" };
                                this.GetChannelFromDB.Add(arr);

                                msg = "Cannot insert new sale channel : " + ch + " on master data";
                            }
                        }
                        else if (dialog == DialogResult.No)
                        {
                            string[] arr = { ch, "ignore" };
                            this.GetChannelFromDB.Add(arr);

                            msg = "Not found sale channel : " + ch + " on master data";
                        }
                        else
                        {
                            Environment.Exit(0);
                        }
                    }
                }
            }
            return msg;
        }

        public string CheckDate(string start, string end)
        {
            string msg = "Success";
            string dateEnd, dateStr;

            if (String.IsNullOrEmpty(end) || end == "-")
            {
                if (String.IsNullOrEmpty(start) || start == "-")
                {
                    msg = "Start Date and End Date are empty";
                }
                else
                {
                    //format start date
                    dateStr = this.ChangeFormatDate(start);

                    if (dateStr == "Invalid")
                    {
                        msg = "Start Date fotmat is not supported";
                    }
                }
            }
            else
            {
                if (String.IsNullOrEmpty(start) || start == "-")
                {
                    dateEnd = this.ChangeFormatDate(end);

                    if (dateEnd == "Invalid")
                    {
                        msg = "End Date fotmat is not supported";
                    }
                }
                else
                {
                    dateStr = this.ChangeFormatDate(start);
                    dateEnd = this.ChangeFormatDate(end);

                    if(dateStr != "Invalid" && dateEnd != "Invalid")
                    {
                        if (Convert.ToDateTime(dateStr) < Convert.ToDateTime(dateEnd) &&
                            Convert.ToDateTime(dateStr) >= DateTime.Now)
                        { }
                        else
                        {
                            msg = "Date ia invalid";
                        }
                    }
                    else
                    {
                        msg = "Start Date and End Date fotmat are not supported";
                    }
                }
            }

            return msg;
        }
        public string ChangeFormatDate(string date)
        {
            double d;
            DateTime dDate;

            if (DateTime.TryParse(date, out dDate))
            {
                date = dDate.ToString("dd/MM/yyyy");
            }
            else
            {
                if (double.TryParse(date, out d))
                {
                    dDate = DateTime.FromOADate(d);
                    date = dDate.ToString("dd/MM/yyyy");
                }
                else if (date == "-")
                {
                    date = null;
                }
                else
                {
                    date = "Invalid";
                }
            }

            return date;
        }

        public string CheckContract(DataTable tableContract, string entry, string install)
        {
            string msg = "";
            foreach (DataRow row in tableContract.Rows)
            {
                string entName = row[0].ToString();
                string insName = row[2].ToString();

                if (entry == entName && install == insName)
                {
                    msg = "Success";
                    break;
                }
                else
                {
                    msg = "Both entry and install didn't match";
                }
            }

            return msg;
        }
    }
}

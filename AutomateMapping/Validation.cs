using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.OracleClient;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace AutomateMapping
{
    class Validation
    {
        /// <summary>
        /// Connection of Production
        /// </summary>
        private OracleConnection ConnectionProd;
        /// <summary>
        /// Connection of CVMDev
        /// </summary>
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
                    if (dr[0] != null && dr[2] != null)
                    {
                        string mkt = dr[0].ToString();
                        string name = dr[2].ToString();

                        if (String.IsNullOrEmpty(mkt) == false && String.IsNullOrEmpty(name) == false)
                        {
                            lstPname.Add(mkt, name);
                        }
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
                string query = "SELECT VALUE1 FROM TRUE9_BPT_VALIDATE WHERE TYPE = 'SALE_CHANNEL' AND NAME1 = 'SALE_CHANNEL'";
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
                string msg = "Cannot get sub profile from database.";

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
                string msg = "Cannot get extra profile from database.";

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

        List<string[]> _provFromDB;
        public List<string[]> GetProvFromDB
        {
            get
            {
                if (this._provFromDB is null || this._provFromDB.Count <= 0)
                {
                    this._provFromDB = GetProvince();
                }

                return this._provFromDB;
            }
            set
            {
                this._provFromDB = value;
            }
        }
        private List<string[]> GetProvince()
        {
            List<string[]> list = new List<string[]>();

            try
            {
                string query = "SELECT DP_PROVINCE FROM DISCOUNT_CRITERIA_PROVINCE";
                OracleCommand cmd = new OracleCommand(query, ConnectionProd);
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
            catch (Exception ex)
            {
                string msg = "Cannot get province from database.";

                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                list.Clear();
            }

            return list;
        }

        List<string[]> _lstType;
        public List<string[]> GetVasType
        {
            get
            {
                if (this._lstType is null || this._lstType.Count <= 0)
                {
                    this._lstType = getVasType();
                }

                return this._lstType;
            }
            set
            {
                this._lstType = value;
            }
        }
        private List<string[]> getVasType()
        {
            List<string[]> list = new List<string[]>();

            try
            {
                string query = "SELECT VALUE1 FROM TRUE9_BPT_VALIDATE WHERE TYPE = 'VAS_TYPE' AND NAME1 = 'VAS_TYPE'";
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
                string msg = "Cannot get VAS_TYPE from database.";

                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                list.Clear();
            }
            return list;
        }

        List<string[]> _lstGroup;
        public List<string[]> GetVasGroup
        {
            get
            {
                if (this._lstGroup is null || this._lstGroup.Count <= 0)
                {
                    this._lstGroup = getVasGroup();
                }

                return this._lstGroup;
            }
            set
            {
                this._lstGroup = value;
            }
        }
        private List<string[]> getVasGroup()
        {
            List<string[]> list = new List<string[]>();

            try
            {
                string query = "SELECT VALUE1 FROM TRUE9_BPT_VALIDATE WHERE TYPE = 'VAS_GROUP' AND NAME1 = 'VAS_GROUP'";
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
                string msg = "Cannot get VAS_GROUP from database.";

                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                list.Clear();
            }
            return list;
        }

        List<string[]> _lstVasChannel;
        public List<string[]> GetVasChannel
        {
            get
            {
                if (this._lstVasChannel is null || this._lstVasChannel.Count <= 0)
                {
                    this._lstVasChannel = getVasChannel();
                }

                return this._lstVasChannel;
            }
            set
            {
                this._lstVasChannel = value;
            }
        }
        private List<string[]> getVasChannel()
        {
            List<string[]> list = new List<string[]>();

            try
            {
                string query = "SELECT VALUE1 FROM TRUE9_BPT_VALIDATE WHERE TYPE = 'VAS_CHANNEL' AND NAME1 = 'VAS_CHANNEL'";
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
                string msg = "Cannot get VAS_Channel from database.";

                MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                list.Clear();
            }
            return list;
        }
        #endregion

        #region "Check Data"
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

            try
            {
                if (speed.Contains('/'))
                {
                    string[] splitSpeed = speed.Split('/');
                    string download = splitSpeed[0].Trim();
                    string upload = splitSpeed[1].Trim();

                    if (int.TryParse(upload, out _) == false)
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

                        if (speed2K == -1)
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
                                speedID = Convert.ToInt32(Regex.Replace(download, "[^0-9]", ""));
                                if (uomDownload == "G")
                                {
                                    speedID = speedID * 1000;
                                }

                                DialogResult dialog = MessageBox.Show("Do you want to insert new speed into DB[Hispeed Speed]?" +
                                    "Detail : SpeedID = " + speedID + " : Desc = " + speed2K, "Confirmation",
                                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                                if (dialog == DialogResult.Yes)
                                {
                                    //insert speed into table
                                    OracleTransaction transaction = null;
                                    string cmd = "INSERT INTO HISPEED_SPEED VALUES (" + speedID + ",'" + speed2K + "','" +
                                    speed2K + "K','" + speed2K + "')";

                                    try
                                    {
                                        if (result.ContainsKey(speedID) == false)
                                        {
                                            transaction = ConnectionProd.BeginTransaction(IsolationLevel.ReadCommitted);
                                            OracleCommand command = new OracleCommand(cmd, ConnectionProd);
                                            command.Transaction = transaction;
                                            command.ExecuteNonQuery();

                                            string[] arr = { speed2K.ToString(), "DB" };
                                            this.GetSpeedFromDB.Add(speedID, arr);

                                            transaction.Commit();

                                            result.Add(speedID, "Success");
                                        }
                                        else
                                        {
                                            MessageBox.Show("SpeedID : " + speedID + " already exists in the database", "Existing SpeedID",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            result.Add(-1, "SpeedID : " + speedID + " already exists in the database");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        transaction.Rollback();

                                        string[] arr = { speed2K.ToString(), "ignore" };
                                        this.GetSpeedFromDB.Add(speedID, arr);

                                        result.Add(-1, "Cannot insert new speedID : " + speedID + " on Master Data.");
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
                                if (status == "ignore")
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
                    result.Add(speedID, "Speed:" + speed + " format is not supported");
                }
            }
            catch(Exception ex)
            {
                result.Add(-1, "An error occurred while working with the database[Hispeed_Speed] : " + ex.Message);
            }

            return result;
        }
        private Dictionary<int, string> CheckFormatMKT(Dictionary<int, string[]> lstSpeed4Chk, string mkt, string speed)
        {
            int speedID = -1;
            Dictionary<int, string> result = new Dictionary<int, string>();

            try
            {
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
                            suffixMkt = (Convert.ToInt32(suffixMkt.Substring(0, suffixMkt.Length - 1)) * 1000).ToString();
                            speed2K = Convert.ToInt32(suffixMkt) * 1024;
                        }
                        else
                        {
                            if (int.TryParse(suffixMkt, out _))
                            {
                                speed2K = Convert.ToInt32(suffixMkt) * 1024;
                            }
                            else
                            {
                                result.Add(-1, "Suffix of MKT: " + mkt + " is Wrong!!");
                            }
                        }

                        //Searching speedID from list
                        if (speed2K != 0)
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
                                DialogResult dialog = MessageBox.Show("Do you want to insert new speed into DB[Hispeed_Speed]? " +
                                    "Detail : SpeedID = " + suffixMkt + " : Desc = " + speed2K, "Confirmation",
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

                                        string[] arr = { speed2K.ToString(), "DB" };
                                        this.GetSpeedFromDB.Add(Convert.ToInt32(suffixMkt), arr);

                                        transaction.Commit();

                                        result.Add(Convert.ToInt32(suffixMkt), "Success");
                                    }
                                    catch (Exception ex)
                                    {
                                        transaction.Rollback();

                                        string[] arr = { speed2K.ToString(), "ignore" };
                                        this.GetSpeedFromDB.Add(Convert.ToInt32(suffixMkt), arr);
                                        result.Add(Convert.ToInt32(suffixMkt), "Cannot insert new SpeedID: " + suffixMkt + " on Master Data.");
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
                                    result.Add(-1, "Not found SpeedID: " + suffixMkt + " on Master Data.");
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
                    result.Add(-1, "MKT Code:" + mkt + " format is not supported");
                }
            }
            catch(Exception ex)
            {
                result.Add(-1, "An error occurred while working with the database[Hispeed_Speed] : " + ex.Message);
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
                                "Detail : Extra Profile = " + extra, "Confirmation",
                                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    if (dialog == DialogResult.Yes)
                    {
                        OracleTransaction transaction = null;
                        string cmd = "INSERT INTO TRUE9_BPT_VALIDATE (TYPE, NAME1, VALUE1) VALUES('EXTRA_PROFILE', " +
                            "'EXTRA_PROFILE', '" + extra + "')";
                        try
                        {
                            transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted);
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

        public string CheckType(List<string[]> lstVasType, string type)
        {
            string msg = "Success";

            bool hasType = lstVasType.Any(p => p.SequenceEqual(new string[] { type, "DB" }));
            if (hasType == false)
            {
                DialogResult dialog = MessageBox.Show("Do you want to insert new vas type?" + "\r\n" +
                            "Detail : VAS_TYPE = " + type, "Confirmation",
                            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (dialog == DialogResult.Yes)
                {
                    OracleTransaction transaction = null;
                    string cmd = "INSERT INTO TRUE9_BPT_VALIDATE (TYPE, NAME1, VALUE1) VALUES('VAS_TYPE', " +
                        "'VAS_TYPE', '" + type + "')";
                    try
                    {
                        transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted);
                        OracleCommand command = new OracleCommand(cmd, ConnectionTemp);
                        command.Transaction = transaction;
                        command.ExecuteNonQuery();

                        transaction.Commit();

                        string[] arr = { type, "DB" };
                        this.GetVasType.Add(arr);
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();

                        string[] arr = { type, "ignore" };
                        this.GetVasType.Add(arr);

                        msg = "Cannot insert new vas type : " + type + " on Master Date";
                    }
                }
                else if (dialog == DialogResult.No)
                {
                    string[] arr = { type, "ignore" };
                    this.GetVasType.Add(arr);

                    msg = "Not found vas type : " + type + " on Master Data";
                }
                else
                {
                    Environment.Exit(0);
                }
            }

            return msg;
        }

        public string CheckGroup(List<string[]> lstVasGroup, string group)
        {
            string msg = "Success";

            bool hasGroup = lstVasGroup.Any(p => p.SequenceEqual(new string[] { group, "DB" }));
            if (hasGroup == false)
            {
                DialogResult dialog = MessageBox.Show("Do you want to insert new vas group?" + "\r\n" +
                            "Detail : VAS_GROUP = " + group, "Confirmation",
                            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (dialog == DialogResult.Yes)
                {
                    OracleTransaction transaction = null;
                    string cmd = "INSERT INTO TRUE9_BPT_VALIDATE (TYPE, NAME1, VALUE1) VALUES('VAS_GROUP', " +
                        "'VAS_GROUP', '" + group + "')";
                    try
                    {
                        transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted);
                        OracleCommand command = new OracleCommand(cmd, ConnectionTemp);
                        command.Transaction = transaction;
                        command.ExecuteNonQuery();

                        transaction.Commit();

                        string[] arr = { group, "DB" };
                        this.GetVasGroup.Add(arr);
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();

                        string[] arr = { group, "ignore" };
                        this.GetVasGroup.Add(arr);

                        msg = "Cannot insert new vas group : " + group + " on Master Date";
                    }
                }
                else if (dialog == DialogResult.No)
                {
                    string[] arr = { group, "ignore" };
                    this.GetVasGroup.Add(arr);

                    msg = "Not found vas group : " + group + " on Master Data";
                }
                else
                {
                    Environment.Exit(0);
                }
            }

            return msg;
        }

        public string CheckVasChannel(List<string[]> lstVasChannel, string channel)
        {
            string msg = "Success";

            bool hasChannel = lstVasChannel.Any(p => p.SequenceEqual(new string[] { channel, "DB" }));
            if (hasChannel == false)
            {
                DialogResult dialog = MessageBox.Show("Do you want to insert new vas channel?" + "\r\n" +
                            "Detail : VAS_CHANNEL = " + channel, "Confirmation",
                            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (dialog == DialogResult.Yes)
                {
                    OracleTransaction transaction = null;
                    string cmd = "INSERT INTO TRUE9_BPT_VALIDATE (TYPE, NAME1, VALUE1) VALUES('VAS_CHANNEL', " +
                        "'VAS_CHANNEL', '" + channel + "')";
                    try
                    {
                        transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted);
                        OracleCommand command = new OracleCommand(cmd, ConnectionTemp);
                        command.Transaction = transaction;
                        command.ExecuteNonQuery();

                        transaction.Commit();

                        string[] arr = { channel, "DB" };
                        this.GetVasChannel.Add(arr);
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();

                        string[] arr = { channel, "ignore" };
                        this.GetVasChannel.Add(arr);

                        msg = "Cannot insert new vas channel : " + channel + " on Master Date";
                    }
                }
                else if (dialog == DialogResult.No)
                {
                    string[] arr = { channel, "ignore" };
                    this.GetVasChannel.Add(arr);

                    msg = "Not found vas channel : " + channel + " on Master Data";
                }
                else
                {
                    Environment.Exit(0);
                }
            }

            return msg;
        }

        public string CheckSubProfile(List<string[]> lstSubProfile, string subProfile)
        {
            string msg = "Success";

            if(subProfile.StartsWith("STL"))
            {
                subProfile = "STL (stand alone)";
            }

            if(String.IsNullOrEmpty(subProfile))
            {
                msg = "SubProfile is empty";
            }
            else
            {
                bool hasSub= lstSubProfile.Any(p => p.SequenceEqual(new string[] { subProfile, "DB" }));
                if (hasSub == false)
                {
                    DialogResult dialog = MessageBox.Show("Do you want to insert new SubProfile?" + "\r\n" +
                                "Detail : SubProfile = " + subProfile, "Confirmation",
                                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    if (dialog == DialogResult.Yes)
                    {
                        OracleTransaction transaction = null;
                        string cmd = "INSERT INTO TRUE9_BPT_VALIDATE (TYPE, NAME1, VALUE1) VALUES('SUB_PROFILE', " +
                            "'SUB_PROFILE', '" + subProfile + "')";
                        try
                        {
                            transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted);
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

            if (String.IsNullOrEmpty(order))
            {
                msg = "Order type is empty";
            }
            else
            {
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
                    if (lstOrder[i].ToUpper().Trim() != "NEW" &&
                        lstOrder[i].ToUpper().Trim() != "CHANGE")
                    {
                        msg = "Order type :" + lstOrder[i] + " is invalid";
                    }
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
                if (channel.Contains(","))
                {
                    if (channel.Contains("DEFAULT") || channel.Contains("ALL"))
                    {
                        msg = "Channel 'ALL/DEFAULT' included with other channel in same MKT code";
                    }
                    else
                    {
                        lstChannel = channel.Split(',');
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
                    if (String.IsNullOrEmpty(ch) == false)
                    {
                        if (ch != "ALL")
                        {
                            bool hasChannel = lstChannelFromDB.Any(p => p.SequenceEqual(new string[] { ch, "DB" }));
                            if (hasChannel == false)
                            {
                                DialogResult dialog = MessageBox.Show("Do you want to insert new channel to master table?" + "\r\n" +
                                            "Detail : Sale_Channel = " + ch, "Confirmation",
                                            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                                if (dialog == DialogResult.Yes)
                                {
                                    OracleTransaction transaction = null;
                                    string cmd = "INSERT INTO TRUE9_BPT_VALIDATE (TYPE, NAME1, VALUE1) VALUES('SALE_CHANNEL', " +
                                        "'SALE_CHANNEL', '" + ch + "')";
                                    try
                                    {
                                        transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted);
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
                    msg = "StartDate and EndDate are empty";
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
                            msg = "Date is invalid";
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
                else if (date == "-" || String.IsNullOrEmpty(date))
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

        public string CheckAllowOffer(string offers)
        {
            string msg = "Success";
            string[] lstOffer = null;

            offers = Regex.Replace(offers, "ALL", "ALL", RegexOptions.IgnoreCase);

            //check special character
            if (Regex.IsMatch(offers, @"^[a-zA-Z0-9\,\-]+$"))
            {
                if (offers.Contains(","))
                {
                    if (lstOffer.Contains("ALL"))
                    {
                        //write log conflict mkt
                        msg = "There is 'ALL' included with the list of main offer";
                    }
                    else
                    {
                        lstOffer = offers.Split(',');
                    }
                }
                else
                {
                    lstOffer = new string[1];
                    lstOffer[0] = offers;
                }

                for (int i = 0; i < lstOffer.Length; i++)
                {
                    if (lstOffer[i] != "ALL")
                    {
                        if (lstOffer[i].Contains('-') == false)
                        {
                            if(String.IsNullOrEmpty(lstOffer[i]) == false)
                            {
                                //write log wrong format
                                msg = "The offer format is not supported";
                            }
                        }
                    }
                }
            }
            else
            {
                //write log format not supported
                msg = "The format is not support (other special characters are not supported)";
            }

            return msg;
        }

        public string CheckProvince(string prov)
        {
            string msg = "Success";
            string[] lstProv = null;

            prov = Regex.Replace(prov, "ALL", "ALL", RegexOptions.IgnoreCase);

            if (Regex.IsMatch(prov, @"^[a-zA-Z\,]+$"))
            {
                if (prov.Contains(','))
                {
                    if(lstProv.Contains("ALL"))
                    {
                        //write log conflict mkt
                        msg = "There is 'ALL' included with the list of province";
                    }
                    else
                    {
                        lstProv = prov.Split(',');
                    }
                }
                else
                {
                    lstProv = new string[1];
                    lstProv[0] = prov;
                }

                for (int i = 0; i < lstProv.Length; i++)
                {
                    if (lstProv[i] != "ALL")
                    {
                        bool hasProv = _provFromDB.Any(p => p.SequenceEqual(new string[] { lstProv[i], "DB" }));

                        if (hasProv == false)
                        {
                            //write log not found province
                            msg = "Not found Province: " + lstProv[i] + " on Master Data.";
                        }
                    }
                }

            }
            else
            {
                //write log contain special charecter
                msg = "The format is not support (other special characters are not supported)";
            }

            return msg;
        }

        public string CheckAllowAdvMonth(string value)
        {
            string msg = "Success";

            value = value.ToUpper().Trim();
            if (!value.Equals("Y") && !value.Equals("N") && !value.Equals("ALL"))
            {
                //write log invalid data
                msg = "Invalid value of Advance Month";
            }

            return msg;
        }

        public string[] CheckSpeedVAS(string from, string to)
        {
            string[] msg = { "Success", "Success", "Success" };

            string uom;
            int speedF = 99999, speedT = 99999;
            bool hasLogF = false, hasLogT = false;

            from = from.ToUpper();
            to = to.ToUpper();

            if(from != "ALL")
            {
                uom = Regex.Replace(from, "[0-9]", "");

                if (!uom.Equals("K") && !uom.Equals("M") && !uom.Equals("G"))
                {
                    //write log
                    hasLogF = true;
                    msg[0] = "Invalid UOM of Speed(From): " + from + "." + "\r\n";
                }
                else
                {
                    speedF = ConvertUOM2K(from, uom);
                }
            }

            if(to != "ALL")
            {
                uom = Regex.Replace(to, "[0-9]", "");

                if (!uom.Equals("K") && !uom.Equals("M") && !uom.Equals("G"))
                {
                    //write log
                    hasLogT = true;
                    msg[1] = "Invalid UOM of Speed(To): " + to + "." + "\r\n";
                }
                else
                {
                    speedT = ConvertUOM2K(to, uom);
                }
            }

            if (hasLogF == false && hasLogT == false)
            {
                if (from != "ALL" && to != "ALL")
                {
                    if (speedF <= speedT)
                    { }
                    else
                    {
                        //write log
                        msg[2] = "Speed(To) must be more than Speed(From)";
                    }

                }
            }

            return msg;
        }

        public string[] CheckPrice(string from, string to)
        {
            string[] msg = { "Success", "Success", "Success" };
            bool hasLogF = false, hasLogT = false;

            from = from.ToUpper();
            to = to.ToUpper();

            if (from != "ALL")
            {
                if (double.TryParse(from, out _) == false)
                {
                    //write log
                    hasLogF = true;
                    msg[0] = "Price(From) is not a numeric";
                }
            }

            if(to != "ALL")
            {
                if (double.TryParse(to, out _) == false)
                {
                    //write log
                    hasLogT = true;
                    msg[1] = "Price(To) is not a numeric";
                }
            }

            if (hasLogF == false && hasLogT == false)
            {
                if (from != "ALL" && to != "ALL")
                {
                    if(Convert.ToDouble(from) <= Convert.ToDouble(to))
                    { }
                    else
                    {
                        //write log
                        msg[2] = "Price(To) must be more than Price(From)";
                    }
                }
            }

            return msg;
        }

        #endregion

        /// <summary>
        /// Get SheetName from file
        /// </summary>
        /// <param name="excelFilePath"></param>
        /// <returns></returns>
        public List<string> ToExcelsSheetList(string excelFilePath)
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
    }
}

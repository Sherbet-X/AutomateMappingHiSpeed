using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutomateMapping
{
    class Validation
    {
        private DataGridView dataGridView;
        private OracleConnection ConnectionProd;
        private OracleConnection ConnectionTemp;

        public Validation(DataGridView dgv, OracleConnection connProd, OracleConnection connTemp)
        {
            this.dataGridView = dgv;
            this.ConnectionProd = connProd;
            this.ConnectionTemp = connTemp;
        }

        //Get Description from file New MKT
        private Dictionary<string, string> GetDescription(string file)
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

        //Get sale channel from master
        private List<string> GetChannelFromDB()
        {
            List<string> list = new List<string>();

            try
            {
                //Get all channel in DB
                string query = "SELECT DISTINCT(SALE_CHANNEL) FROM HISPEED_CHANNEL_PROMOTION";
                OracleCommand cmd = new OracleCommand(query, ConnectionProd);
                OracleDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    list.Add(reader["SALE_CHANNEL"].ToString());
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

        private void CheckSpeed(int index, string mkt , string speed, string uom)
        {
            
        }
    }
}

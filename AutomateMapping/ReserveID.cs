using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OracleClient;
using System.Windows.Forms;
using System.Data;

namespace AutomateMapping
{
    class ReserveID
    {
        public void Reserve(OracleConnection ConnectionProd, OracleConnection ConnectionTemp,
            string type, string implementer, string urNo)
        {
            OracleCommand cmd = null;

            try
            {
                int minID = GetMinID(ConnectionTemp, ConnectionProd, type);
                minID = minID + 1;

                string query = "SELECT * FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = '" + type + "' AND COMPLETE_FLAG = 'N'";
                cmd = new OracleCommand(query, ConnectionTemp);
                OracleDataReader reader = cmd.ExecuteReader();
                reader.Read();
                if (reader.HasRows)
                {
                    string user = reader["USERNAME"].ToString();
                    string typeName = reader["TYPE_NAME"].ToString();
                    string ur = reader["UR_NO"].ToString();

                    if (user == implementer && typeName == type && urNo == ur)
                    {
                        using (OracleTransaction transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted))
                        {
                            cmd.Transaction = transaction;
                            try
                            {
                                cmd.CommandText = "UPDATE TRUE9_BPT_RESERVE_ID SET MIN_ID = '" + minID + "', CREATE_DATE = SYSDATE"+
                                    "' WHERE TYPE_NAME = '" + type + "' AND UR_NO = '" + urNo + "' AND USERNAME = '" + implementer + "'";

                                cmd.CommandType = CommandType.Text;
                                cmd.ExecuteNonQuery();
                                transaction.Commit();
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();
                                throw new Exception(ex.Message);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("UserName : " + user + " is in the process of inserting." + "\r\n" + "Please try again later");
                        ConnectionTemp.Close();
                        ConnectionProd.Close();
                        Environment.Exit(0);
                    }
                }
                else
                {
                    cmd = ConnectionTemp.CreateCommand();

                    using (OracleTransaction transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted))
                    {
                        cmd.Transaction = transaction;

                        try
                        {
                            cmd.CommandText = "INSERT INTO TRUE9_BPT_RESERVE_ID VALUES('" + type + "', 'N', '"+minID+"', null, '" + 
                                urNo + "', '" + implementer + "', sysdate)";

                            cmd.CommandType = CommandType.Text;

                            cmd.ExecuteNonQuery();
                            transaction.Commit();
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            throw new Exception(ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot reserve UR into table[TRUE9_BPT_RESERVE_ID] " + "\r\n" + "Error Detail : " + ex.Message);
                ConnectionTemp.Close();
                Environment.Exit(0);
            }
        }

        private int GetMinID(OracleConnection ConnectionTemp, OracleConnection ConnectionProd, string type)
        {
            OracleCommand cmd = null;
            string prefixID, col , table;
            if (type == "Hispeed")
            {
                prefixID = "20";
                col = "P_ID";
                table = "HISPEED_PROMOTION";
            }
            else if (type == "Disc")
            {
                prefixID = "DC";
                col = "DC_ID";
                table = "DISCOUNT_CRITERIA_MAPPING";
            }
            else
            {
                prefixID = "VAS";
                col = "DC_ID";
                table = "DISCOUNT_CRITERIA_MAPPING";
            }

            cmd = new OracleCommand("SELECT MAX(" + col + ") FROM " + table + " WHERE " + col + " LIKE '" + prefixID + "%'", ConnectionProd);
            OracleDataReader reader = cmd.ExecuteReader();
            reader.Read();

            cmd = new OracleCommand("SELECT MAX(MAX_ID) FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = '" + type + "'", ConnectionTemp);
            OracleDataReader readerReserve = cmd.ExecuteReader();
            readerReserve.Read();

            int max, min;
            if (type == "Hispeed")
            {
                min = Convert.ToInt32(reader[0]);
                max = Convert.ToInt32(readerReserve[0]);
            }
            else
            {
                string minid = Convert.ToString(reader[0]).Substring(prefixID.Length);
                string maxid = Convert.ToString(readerReserve[0]).Substring(prefixID.Length);
                min = Convert.ToInt32(minid);
                max = Convert.ToInt32(maxid);
            }

            if (min <= max)
            {
                MessageBox.Show("There is a conflict ID between production and reserve table[TRUE9_BPT_RESERVE_ID]" + "\r\n"
                    + "Please review and confirm the information");

                string qryDel = "DELETE FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = '" + type + "' AND COMPLETE_FLAG = 'N'";
                OracleCommand command = new OracleCommand(qryDel, ConnectionTemp);
                command.ExecuteNonQuery();

                ConnectionProd.Close();
                ConnectionTemp.Close();

                Environment.Exit(0);
            }

            return min;
        }

        public void updateID(OracleConnection ConnectionTemp, string minID, string maxID, string type, string implementer, string urNO)
        {
            OracleCommand cmd = ConnectionTemp.CreateCommand();

            using (OracleTransaction transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted))
            {
                cmd.Transaction = transaction;

                try
                {
                    cmd.CommandText = "UPDATE TRUE9_BPT_RESERVE_ID SET COMPLETE_FLAG = 'Y', MIN_ID = '" + minID + "', MAX_ID = '" +
                        maxID + "' WHERE TYPE_NAME = '" + type + "' AND UR_NO = '" + urNO + "' AND USERNAME = '" + implementer + "'";

                    cmd.CommandType = CommandType.Text;

                    cmd.ExecuteNonQuery();
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    MessageBox.Show("Cannot reserve MinID into table[TRUE9_BPT_RESERVE_ID]" + "\r\n" +
                        "Please check error message and manual reserve ID into table[TRUE9_BPT_RESERVE_ID]" + "Error Detail : " + ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}


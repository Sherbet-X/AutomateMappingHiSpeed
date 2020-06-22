using System;
using System.Data;
using System.Data.OracleClient;
using System.Windows.Forms;

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
                                cmd.CommandText = "UPDATE TRUE9_BPT_RESERVE_ID SET MIN_ID = '" + minID + "', CREATE_DATE = SYSDATE" +
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
                            cmd.CommandText = "INSERT INTO TRUE9_BPT_RESERVE_ID VALUES('" + type + "', 'N', '" + minID + "', null , '" +
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
            string prefixID, col, table;
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

            int maxReserv, maxProd;
            if (type == "Hispeed")
            {
                if(reader.IsDBNull(0))
                {
                    maxProd = 0;
                }
                else
                {
                    maxProd = Convert.ToInt32(reader[0]);
                }

                if(readerReserve.IsDBNull(0))
                {
                    maxReserv = 0;
                }
                else
                {
                    maxReserv = Convert.ToInt32(readerReserve[0]);
                }

            }
            else
            {
                string minid = Convert.ToString(reader[0]).Substring(prefixID.Length);
                string maxid = Convert.ToString(readerReserve[0]).Substring(prefixID.Length);
                maxProd = Convert.ToInt32(minid);
                maxReserv = Convert.ToInt32(maxid);
            }

            if ((maxProd + 1) <= maxReserv)
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

            return maxProd;
        }

        public void UpdateReserveID(OracleConnection ConnectionTemp, OracleConnection ConnectionProd, string type, string ur)
        {
            OracleCommand cmd = null;
            string prefixID, col, table;
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

            cmd = new OracleCommand("SELECT MIN_ID FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = '" + type + "' AND COMPLETE_FLAG = 'N' " +
                                "AND UR_NO = '" + ur + "'", ConnectionTemp);
            OracleDataReader readerReserve = cmd.ExecuteReader();
            readerReserve.Read();

            int minReserv, maxProd;
            if (type == "Hispeed")
            {
                if (reader.IsDBNull(0))
                {
                    maxProd = 0;
                }
                else
                {
                    maxProd = Convert.ToInt32(reader[0]);
                }

                if (readerReserve.IsDBNull(0))
                {
                    minReserv = 0;
                }
                else
                {
                    minReserv = Convert.ToInt32(readerReserve[0]);
                }

                if(minReserv != 0)
                {
                    if ((maxProd + 1) == minReserv)
                    {
                        //delete
                        string qryDel = "DELETE FROM TRUE9_BPT_RESERVE_ID WHERE TYPE_NAME = '" + type + "' AND COMPLETE_FLAG = 'N' " +
                                    "AND MIN_ID = '" + minReserv + "' AND UR_NO = '" + ur + "'";
                        cmd = new OracleCommand(qryDel, ConnectionTemp);
                        cmd.ExecuteNonQuery();
                    }
                    else
                    {
                        //update
                        OracleTransaction transaction = ConnectionTemp.BeginTransaction(IsolationLevel.ReadCommitted);
                        try
                        {
                            cmd = ConnectionTemp.CreateCommand();
                            cmd.Transaction = transaction;
                            cmd.CommandText = "UPDATE TRUE9_BPT_RESERVE_ID SET COMPLETE_FLAG = 'Y' , MAX_ID = '" + maxProd + "' " +
                                        "WHERE TYPE_NAME = '" + type + "' AND MIN_ID = '" + minReserv + "' AND UR_NO = '" + ur + "'";
                            cmd.ExecuteNonQuery();
                            transaction.Commit();
                        }
                        catch(Exception ex)
                        {
                            transaction.Rollback();
                            MessageBox.Show("Cannot reserve MIN_ID into table[TRUE9_BPT_RESERVE_ID]" + "\r\n" +
                                "Please manual update data into table[TRUE9_BPT_RESERVE_ID]" + "\r\n" + "Error Detail : "
                                + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }                    
                }
                else
                {
                    MessageBox.Show("Cannot reserve MIN_ID into table[TRUE9_BPT_RESERVE_ID]" + "\r\n" +
                                "Please manual update data into table[TRUE9_BPT_RESERVE_ID]", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                }             
            }
        }
    }
}

﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;

namespace AutomateMapping
{
    public partial class MainHispeed : Form
    {
        private OracleConnection ConnectionProd;
        private string filename;
        public MainHispeed(OracleConnection con, string file)
        {
            InitializeComponent();
            ConnectionProd = con;
            filename = file;
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
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook wb = xlApp.Workbooks.Open(filename);
            DgvSettings dgvSettings = new DgvSettings();
            List<string> lstHeader = new List<string>();

            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name == "HiSpeed Promotion")
                {
                    //set header view
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

                    dgvSettings.SetDgv(dataGridView1, filename, "HiSpeed Promotion$B3:P", lstHeader);

                    break;
                }
                else if (sheet.Name == "Campaign Mapping")
                {
                    //set header view
                    lstHeader.Add("Type");
                    lstHeader.Add("Campaign Name");
                    lstHeader.Add("TOL Package");
                    lstHeader.Add("TOL Discount");
                    lstHeader.Add("TVS Package");
                    lstHeader.Add("TVS Discount");

                    dgvSettings.SetDgv(dataGridView1, filename, "Campaign Mapping$B2:G", lstHeader);

                    break;
                }

               
            }
        }
    }
}

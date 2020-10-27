using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutomateMapping
{
    public partial class UserControlVASProduct : UserControl
    {
        public UserControlVASProduct()
        {
            InitializeComponent();
        }

        string _fileName;
        public string GetfileName
        {
            get
            {
                return this._fileName;
            }
            set
            {
                this._fileName = value;
            }
        }

        public DataGridView GetDataGridView
        {
            get
            {
                return this.dataGridView1;
            }
            set
            {
                this.dataGridView1 = value;
            }
        }

        public void ShowView()
        {
            //dataGridView1.Refresh();
            //dataGridView1.Show();
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.CellStyle.Font = new Font("Microsoft Sans Serif", 11, FontStyle.Bold);
        }
    }
}

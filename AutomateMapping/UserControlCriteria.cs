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
    public partial class UserControlCriteria : UserControl
    {
        public UserControlCriteria()
        {
            InitializeComponent();
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

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.CellStyle.Font = new Font("Microsoft Sans Serif", 11, FontStyle.Bold);
        }
    }
}

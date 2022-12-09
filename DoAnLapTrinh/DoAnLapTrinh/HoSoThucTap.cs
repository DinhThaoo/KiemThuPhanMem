using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DoAnLapTrinh
{
    public partial class HoSoThucTap : Form
    {
        public HoSoThucTap()
        {
            InitializeComponent();
        }

        private void btnThoatCT_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dgvthuctap_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        public void getData()
        {
            string query = "SELECT * FROM NhanSu WHERE TrangThai = N'Thực tập'";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "NhanSu");
            dgvthuctap.DataSource = ds.Tables["NhanSu"];
        }

        private void HoSoThucTap_Load(object sender, EventArgs e)
        {
            getData();
        }


    }
}

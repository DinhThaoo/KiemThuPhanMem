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
    public partial class HoSoNhanVien : Form
    {
        public HoSoNhanVien()
        {
            InitializeComponent();
        }

        private void btnThoatCT_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        public void getData()
        {
            string query = "SELECT * FROM NhanSu WHERE TrangThai = N'Nhân viên chính thức'";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "NhanSu");
            dgvnhanvien.DataSource = ds.Tables["NhanSu"];
        }

        private void HoSoNhanVien_Load(object sender, EventArgs e)
        {
            getData();
        }

        private void dgvnhanvien_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}

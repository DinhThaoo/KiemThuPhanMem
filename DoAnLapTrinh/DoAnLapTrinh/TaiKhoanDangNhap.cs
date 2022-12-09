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
    public partial class TaiKhoanDangNhap : Form
    {
        public TaiKhoanDangNhap()
        {
            InitializeComponent();
        }
        public void getData()
        {
            string query = "select * from TaiKhoan";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "TaiKhoan");
            dgvTKDN.DataSource = ds.Tables["TaiKhoan"];
        }

        private void QuanLyTaiKhoan_Load(object sender, EventArgs e)
        {
            getData();
             
        }

        private void btnthoat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnTaomoi_Click(object sender, EventArgs e)
        {
            txttendn.Text = "";
            txttenht.Text = "";
            txtmk.Text = "";
            txtltk.Text = "";
            
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            string tendn = txttendn.Text;
            string tenht = txttenht.Text;
            string mk = txtmk.Text;
            string ltk = txtltk.Text;

            string query = "insert into TaiKhoan(TenDangNhap, TenHienThi, MatKhau, LoaiTaiKhoan)"
                + " VALUES " + "(N'" + tendn + "', N'" + tenht + "','" + mk + "','" + ltk + "')";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm mới tài khoản thành công!");
                getData();
            }
            else
            {
                MessageBox.Show("Thêm mới tài khoản thất bại!");
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string tendn = txttendn.Text;
            string tenht = txttenht.Text;
            string mk = txtmk.Text;
            string ltk = txtltk.Text;

            string query = "update TaiKhoan set  TenHienThi=N'" + tenht + "', MatKhau='" + mk + "', LoaiTaiKhoan='" + ltk + "'where TenDangNhap ='" + tendn + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thông tin tài khoản thành công!");
                getData();
            }
            else
                MessageBox.Show("Sửa thông tin tài khoản thất bại!");
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string tendn = txttendn.Text;
            string tenht = txttenht.Text;
            string mk = txtmk.Text;
            string ltk = txtltk.Text;

            string query = "delete TaiKhoan where TenDangNhap = '" + tendn + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thông tin tài khoản thành công!");
                getData();
            }
            else
            {
                MessageBox.Show("Xóa thông tin tài khoản thất bại!");
            }

        }

        private void btnTimkiem_Click(object sender, EventArgs e)
        {

            string tendn = txttendn.Text;
            string query = "select * from TaiKhoan where TenDangNhap like N'%" + tendn + "%'";
            DataSet ds = new DataSet();
            KetNoi cn = new KetNoi();
            ds = cn.laydulieu(query, "TaiKhoan");

            dgvTKDN.DataSource = ds.Tables["TaiKhoan"];
        }

        private void dgvTKDN_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txttendn.Text = dgvTKDN.Rows[row].Cells["TenDangNhap"].Value.ToString();
                txttenht.Text = dgvTKDN.Rows[row].Cells["TenHienThi"].Value.ToString();
                txtmk.Text = dgvTKDN.Rows[row].Cells["MatKhau"].Value.ToString();
                txtltk.Text = dgvTKDN.Rows[row].Cells["LoaiTaiKhoan"].Value.ToString();

            }
        }
    }
}

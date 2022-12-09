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
    public partial class QuanLyKhachHang : Form
    {
        public QuanLyKhachHang()
        {
            InitializeComponent();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void getData()
        {
            string makh = txtmakh.Text;
            string query = "EXEC TichDiem; EXEC DoiDiem; EXEC KhachBac; EXEC KhachVang; EXEC KhachKimCuong;EXEC TongSoTien; select * from KhachHang;";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "KhachHang");
            dgvKhachHang.DataSource = ds.Tables["KhachHang"];
        }

        private void QuanLyKhachHang_Load(object sender, EventArgs e)
        {
            getData();
        }

        private void btnTaomoi_Click(object sender, EventArgs e)
        {
            txtmakh.Text = "";
            txthoten.Text = "";
            txtsdt.Text = "";
            txtsohoadon.Text = "";
            txttongtien.Text = "";
            txttrangthai.Text = "";
            txttichdiem.Text = "";
            txtdoidiem.Text = "";         
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            string makh = txtmakh.Text;
            string hoten = txthoten.Text;
            string sdt = txtsdt.Text;
            string sohoadon = txtsohoadon.Text;
            string tongtien = txttongtien.Text;
            string trangthai = txttrangthai.Text;
            string tichdiem = txttichdiem.Text;
            string doidiem = txtdoidiem.Text;
            

            string query = "INSERT INTO KhachHang(HoTen, SDT) VALUES (N'" + hoten + "', '" + sdt + "')";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm mới khách hàng thành công!");
                getData();
            }
            else
            {
                MessageBox.Show("Thêm mới khách hàng thất bại!");
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string makh = txtmakh.Text;
            string hoten = txthoten.Text;
            string sdt = txtsdt.Text;
            string sohoadon = txtsohoadon.Text;
            string tongtien = txttongtien.Text;
            string trangthai = txttrangthai.Text;
            string tichdiem = txttichdiem.Text;
            string doidiem = txtdoidiem.Text;

            string query = "update KhachHang set  HoTen=N'" + hoten + "', SDT=N'" + sdt + "', SoHoaDon='" + sohoadon + "', TongSoTien='" + tongtien + "', TrangThai=N'" + trangthai + "', TichDiem=N'" + tichdiem + "', DoiDiem='" + doidiem + "'where MaKhachHang ='" + makh + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thông tin khách hàng thành công!");
                getData();
            }
            else
                MessageBox.Show("Sửa thông tin khách hàng thất bại!");
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string makh = txtmakh.Text;

            string query = "delete KhachHang where MaKhachHang = '" + makh + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thông tin khách hàng thành công!");
                getData();
            }
            else
            {
                MessageBox.Show("Xóa thông tin khách hàng thất bại!");
            }
        }

        private void btnTimkiem_Click(object sender, EventArgs e)
        {
            string hoten = txthoten.Text;          
            string query = "select * from KhachHang where HoTen like '%" + hoten + "%'";
           
            DataSet ds = new DataSet();
            KetNoi cn = new KetNoi();
            ds = cn.laydulieu(query, "NhanSu");
            dgvKhachHang.DataSource = ds.Tables["NhanSu"];
        }

        private void dgvKhachHang_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmakh.Text = dgvKhachHang.Rows[row].Cells["MaKhachHang"].Value.ToString();
                txthoten.Text = dgvKhachHang.Rows[row].Cells["HoTen"].Value.ToString();
                txtsdt.Text = dgvKhachHang.Rows[row].Cells["SDT"].Value.ToString();
                txtsohoadon.Text = dgvKhachHang.Rows[row].Cells["SoHoaDon"].Value.ToString();
                txttongtien.Text = dgvKhachHang.Rows[row].Cells["TongSoTien"].Value.ToString();
                txttrangthai.Text = dgvKhachHang.Rows[row].Cells["TrangThai"].Value.ToString();
                txttichdiem.Text = dgvKhachHang.Rows[row].Cells["TichDiem"].Value.ToString();
                txtdoidiem.Text = dgvKhachHang.Rows[row].Cells["DoiDiem"].Value.ToString();
             
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            QuanLyBanHang frm = new QuanLyBanHang();
            frm.Show();
        }
    }
}

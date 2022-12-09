using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using app = Microsoft.Office.Interop.Excel.Application;

namespace DoAnLapTrinh
{
    public partial class QuanLyBanHang : Form
    {
        public QuanLyBanHang()
        {
            InitializeComponent();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void getData()
        {
            string query = "select * from Ban";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "Ban");
            dgvBan.DataSource = ds.Tables["Ban"];
            
            // combobox bảng HÓA ĐƠN//
            cbmaban.DataSource = ds.Tables["Ban"];
            cbmaban.DisplayMember = "TenBan";
            cbmaban.ValueMember = "MaBan";
        }

        //HÓA ĐƠN//
        public void getDataHD()
        {
            string query = "EXEC TongTienHoaDon; select * from HoaDon";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "HoaDon");
            dgvHoadon.DataSource = ds.Tables["HoaDon"];

            // combobox bảng CHI TIẾT HÓA ĐƠN
            cbmahd.DataSource = ds.Tables["HoaDon"];
            cbmahd.DisplayMember = "MaHoaDon";
            cbmahd.ValueMember = "MaHoaDon";
        }

        public void getDataNS()
        {
            string query = "select * from NhanSu";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "NhanSu");
           

            // combobox bảng HÓA ĐƠN//
            cbmanv.DataSource = ds.Tables["NhanSu"];
            cbmanv.DisplayMember = "HoTen";
            cbmanv.ValueMember = "MaNhanSu";
        }

        public void getDataKH()
        {
            string query = "select * from KhachHang";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "KhachHang");
           

            // combobox bảng HÓA ĐƠN//
            cbmakh.DataSource = ds.Tables["KhachHang"];
            cbmakh.DisplayMember = "HoTen";
            cbmakh.ValueMember = "MaKhachHang";
        }

        //CHI TIẾT HÓA ĐƠN//
        public void getDataCTHD()
        {
            string query = "EXEC CapNhatGiaCTHoaDon; EXEC ThanhTien; select * from CTHoaDon";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "CTHoaDon");
            dgvCTHD.DataSource = ds.Tables["CTHoaDon"];

        }

        public void getDataSP()
        {
            string query = "select * from SanPham";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "SanPham");


            // combobox bảng CHI TIẾT HÓA ĐƠN//
            cbmasp.DataSource = ds.Tables["SanPham"];
            cbmasp.DisplayMember = "TenSanPham";
            cbmasp.ValueMember = "MaSanPham";
        }

     
        private void QuanLyBanHang_Load(object sender, EventArgs e)
        {
            getData();
            getDataHD();
            getDataCTHD();
            getDataNS();
            getDataSP();
            getDataKH();
        }

        //BÀN//
        private void btnTaomoi_Click(object sender, EventArgs e)
        {
            txtmaban.Text = "";
            txttenban.Text = "";
            txttrangthai.Text = "";
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            string maban = txtmaban.Text;
            string tenban = txttenban.Text;
            string trangthai = txttrangthai.Text;

            string query = "INSERT INTO Ban(TenBan, TrangThai) VALUES (N'" + tenban + "', N'" + trangthai + "')";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm thông tin bàn thành công!");
                getData();
            }
            else
                MessageBox.Show("Thêm thông tin bàn thất bại!");
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string maban = txtmaban.Text;
            string tenban = txttenban.Text;
            string trangthai = txttrangthai.Text;

            string query = "update Ban set  TenBan=N'" + tenban + "', TrangThai=N'" + trangthai + "'where MaBan ='" + maban + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thông tin bàn thành công!");
                getData();
            }
            else
                MessageBox.Show("Sửa thông tin bàn thất bại!");
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string maban = txtmaban.Text;
            string tenban = txttenban.Text;
            string trangthai = txttrangthai.Text;

            string query = "delete Ban where MaBan = '" + maban + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thông tin bàn thành công!");
                getData();
            }
            else
            {
                MessageBox.Show("Xóa thông tin bàn thất bại!");
            }
        }

        private void dgvBan_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmaban.Text = dgvBan.Rows[row].Cells["MaBan"].Value.ToString();
                txttenban.Text = dgvBan.Rows[row].Cells["TenBan"].Value.ToString();
                txttrangthai.Text = dgvBan.Rows[row].Cells["TrangThai"].Value.ToString();
            }
        }

        private void btnTimkiem_Click(object sender, EventArgs e)
        {
            string timkiem = txtimkiem.Text;
            string query = string.Format("select * from Ban where TenBan like N'%{0}%'", timkiem);

            DataSet ds = new DataSet();
            KetNoi cn = new KetNoi();
            ds = cn.laydulieu(query, "Ban");
            dgvBan.DataSource = ds.Tables["Ban"];

        }

        //HÓA ĐƠN//

        private void btnTaomoi_HD_Click(object sender, EventArgs e)
        {
            txtmahoadon.Text = "";
            cbmanv.Text = "";
            cbmakh.Text = "";
            cbmaban.Text = "";
            txtthoigianbatdau.Text = "";
            txtthoigianketthuc.Text = "";
            txttrangthai_hd.Text = "";
            txttongtien.Text = "";
        }

        private void btnThem_HD_Click(object sender, EventArgs e)
        {
            string mahd = txtmahoadon.Text;
            string manv = cbmanv.SelectedValue.ToString();
            string makh = cbmakh.SelectedValue.ToString();
            string maban = cbmaban.SelectedValue.ToString();
            string tgbatdau = txtthoigianbatdau.Text;
            string tgketthuc = txtthoigianketthuc.Text;
            string trangthai = txttrangthai_hd.Text;
            string tongtien = txttongtien.Text;
            

            string query = "INSERT INTO HoaDon(MaNhanVien,MaKhachHang, MaBan, ThoiGianBatDau, ThoiGianKetThuc, TrangThai, TongTien) VALUES ('" + manv + "','" + makh + "', '" + maban + "', '" + tgbatdau + "', '" + tgketthuc + "', N'" + trangthai + "', '" + tongtien + "')";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm thông tin hóa đơn thành công!");
                getDataHD();
            }
            else
                MessageBox.Show("Sửa thông tin hóa đơn  thất bại!");
        }

        private void dgvHoadon_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmahoadon.Text = dgvHoadon.Rows[row].Cells["MaHoaDon"].Value.ToString();
                cbmanv.SelectedValue = dgvHoadon.Rows[row].Cells["MaNhanVien"].Value.ToString();
                cbmakh.SelectedValue = dgvHoadon.Rows[row].Cells["MaKhachHang"].Value.ToString();
                cbmaban.SelectedValue = dgvHoadon.Rows[row].Cells["MaBan"].Value.ToString();
                txtthoigianbatdau.Text = dgvHoadon.Rows[row].Cells["ThoiGianBatDau"].Value.ToString();
                txtthoigianketthuc.Text = dgvHoadon.Rows[row].Cells["ThoiGianKetThuc"].Value.ToString();
                txttrangthai_hd.Text = dgvHoadon.Rows[row].Cells["TrangThai"].Value.ToString();
                txttongtien.Text = dgvHoadon.Rows[row].Cells["TongTien"].Value.ToString();
            }
        }

        private void btnSua_HD_Click(object sender, EventArgs e)
        {
            string mahd = txtmahoadon.Text;
            string manv = cbmanv.SelectedValue.ToString();
            string makh = cbmakh.SelectedValue.ToString();
            string maban = cbmaban.SelectedValue.ToString();
            string tgbatdau = txtthoigianbatdau.Text;
            string tgketthuc = txtthoigianketthuc.Text;
            string trangthai = txttrangthai.Text;
            string tongtien = txttongtien.Text;

            string query = "update HoaDon set  MaNhanVien=N'" + manv + "',MaKhachHang=N'" + makh + "',MaBan=N'" + maban + "', ThoiGianBatDau=N'" + tgbatdau + "', ThoiGianKetThuc=N'" + tgketthuc + "', TrangThai=N'" + trangthai + "', TongTien=N'" + tongtien + "'where MaHoaDon ='" + mahd + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thông tin hóa đơn thành công!");
                getDataHD();
            }
            else
                MessageBox.Show("Sửa thông tin hóa đơn thất bại!");
        }

        private void btnXoa_HD_Click(object sender, EventArgs e)
        {
            string mahd = txtmahoadon.Text;
            
            string query = "delete HoaDon where MaHoaDon = '" + mahd + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thông tin hóa đơn thành công!");
                getDataHD();
            }
            else
            {
                MessageBox.Show("Xóa thông tin hóa đơn thất bại!");
            }
        }

        private void btnTimkiem_HD_Click(object sender, EventArgs e)
        {
            string mahd = txtmahoadon.Text;

            string query = "select * from HoaDon where MaHoaDon like '"+ mahd + "'";

            DataSet ds = new DataSet();
            KetNoi cn = new KetNoi();
            ds = cn.laydulieu(query, "HoaDon");
            dgvHoadon.DataSource = ds.Tables["HoaDon"];
        }


        //CHI TIẾT HÓA ĐƠN//
        private void btnTaomoi_CTHD_Click(object sender, EventArgs e)
        {
            txtmachitiethd.Text = "";
            cbmahd.Text = "";
            cbmasp.Text = "";
            txtgiaban.Text = "";
            txtsoluong.Text = "";
            txtthanhtien.Text = "";
           
        }

        private void btnThem_CTHD_Click(object sender, EventArgs e)
        {
            string macthd = txtmachitiethd.Text;
            string mahd = cbmahd.SelectedValue.ToString();
            string masp = cbmasp.SelectedValue.ToString();
            string soluong = txtsoluong.Text;
            string thanhtien = txtthanhtien.Text;

            string query = "INSERT INTO CTHoaDon(MaHoaDon, MaSanPham, SoLuong, ThanhTien) VALUES ('" + mahd + "', '" + masp + "', '" + soluong + "', '" + thanhtien + "')";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm thông tin chi tiết hóa đơn thành công!");
                getDataCTHD();
            }
            else
                MessageBox.Show("Sửa thông tin chi tiết hóa đơn  thất bại!");
        }

        private void btnSua_CTHD_Click(object sender, EventArgs e)
        {
            string macthd = txtmachitiethd.Text;
            string mahd = cbmahd.SelectedValue.ToString();
            string masp = cbmasp.SelectedValue.ToString();
            string soluong = txtsoluong.Text;
            string thanhtien = txtthanhtien.Text;

            string query = "update CTHoaDon set  MaHoaDon=N'" + mahd + "', MaSanPham=N'" + masp + "', SoLuong=N'" + soluong + "', ThanhTien=N'" + thanhtien + "' where MaCTHoaDon ='" + macthd + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thông tin chi tiết  hóa đơn thành công!");
                getDataCTHD();
            }
            else
                MessageBox.Show("Sửa thông tin chi tiết hóa đơn thất bại!");
        }

        private void btnXoa_CTHD_Click(object sender, EventArgs e)
        {
            string macthd = txtmachitiethd.Text;

            string query = "delete CTHoaDon where MaCTHoaDon = '" + macthd + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thông tin chi tiết hóa đơn thành công!");
                getDataCTHD();
            }
            else
            {
                MessageBox.Show("Xóa thông tin chi tiết hóa đơn thất bại!");
            }
        }

        private void btnTimkiemCTHD_Click(object sender, EventArgs e)
        {
            string macthd = txtmachitiethd.Text;
            string query = "select * from CTHoaDon where MaCTHoaDon like '" + macthd + "'";

            DataSet ds = new DataSet();
            KetNoi cn = new KetNoi();
            ds = cn.laydulieu(query, "CTHoaDon");
            dgvCTHD.DataSource = ds.Tables["CTHoaDon"];
        }

        private void dgvCTHD_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmachitiethd.Text = dgvCTHD.Rows[row].Cells["MaCTHoaDon"].Value.ToString();
                cbmahd.SelectedValue = dgvCTHD.Rows[row].Cells["MaHoaDon"].Value.ToString();
                cbmasp.SelectedValue = dgvCTHD.Rows[row].Cells["MaSanPham"].Value.ToString();
                txtsoluong.Text = dgvCTHD.Rows[row].Cells["SoLuong"].Value.ToString();
                txtthanhtien.Text = dgvCTHD.Rows[row].Cells["ThanhTien"].Value.ToString();
                txtgiaban.Text = dgvCTHD.Rows[row].Cells["GiaBan"].Value.ToString();

            }
        }

        private void txtimkiem_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            QuanLyKhachHang frm = new QuanLyKhachHang();
            frm.Show();
        }

        /// XUẤT FILE EXCEL
        private void excel(DataGridView g, string duongdan, string tenfile)
        {
            app obj = new app();
            obj.Application.Workbooks.Add(Type.Missing);
            obj.Columns.ColumnWidth = 25;
            for (int i = 1; i < g.Columns.Count + 1; i++)
            {
                obj.Cells[1, i] = g.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < g.Rows.Count; i++)
            {
                for (int j = 0; j < g.Columns.Count; j++)
                {
                    if (g.Rows[i].Cells[j].Value != null)
                    {
                        obj.Cells[i + 2, j + 1] = g.Rows[i].Cells[j].Value.ToString();
                    }

                }
            }
            obj.ActiveWorkbook.SaveCopyAs(duongdan + tenfile + ".xlsx");
            obj.ActiveWorkbook.Saved = true;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Xuất file thành công!");
            excel(dgvHoadon, @"D:\", @"Danh_Sach_Hoa_Don");
        }
    }
}


     


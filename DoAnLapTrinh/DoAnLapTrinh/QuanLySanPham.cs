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
    public partial class QuanLySanPham : Form
    {
        public QuanLySanPham()
        {
            InitializeComponent();
        }
        public void getDataNCC()  
        {
            string query = "select * from NhaCungCap";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "NhaCungCap");
            dgvncc.DataSource = ds.Tables["NhaCungCap"];

            //cobobox bảng NGUYÊN LIỆU 
            cbmancc_nl.DataSource = ds.Tables["NhaCungCap"];
            cbmancc_nl.DisplayMember = "TenNhaCungCap";
            cbmancc_nl.ValueMember = "MaNhaCungCap";

            //combobox bảng PHIẾU NHẬP
            cbmancc_pn.DataSource = ds.Tables["NhaCungCap"];
            cbmancc_pn.DisplayMember = "TenNhaCungCap";
            cbmancc_pn.ValueMember = "MaNhaCungCap";
        }

        public void getDataLSP()
        {
            string query = "select * from LoaiSanPham";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "LoaiSanPham");
            dgvloaisp.DataSource = ds.Tables["LoaiSanPham"];

            cbmaloai.DataSource = ds.Tables["LoaiSanPham"];
            cbmaloai.DisplayMember = "TenLoai";
            cbmaloai.ValueMember = "MaLoai";
        }

        public void getDataSP()
        {
            string query = "select * from SanPham; EXEC CapNhatGiaCTHoaDon";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "SanPham");
            dgvSP.DataSource = ds.Tables["SanPham"];

            //combobox bảng CÔNG THỨC
            cbmasp_ct.DataSource = ds.Tables["SanPham"];
            cbmasp_ct.DisplayMember = "TenSanPham";
            cbmasp_ct.ValueMember = "MaSanPham";
        }

        public void getDataPN()
        {
            string query = " EXEC CapNhatGiaNhapNguyenLieu;select * from PhieuNhap ";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "PhieuNhap");
            dgvPN.DataSource = ds.Tables["PhieuNhap"];
        }

        public void getDataNL()
        {
            string query = " select * from NguyenLieu";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "NguyenLieu");
            dgvNL.DataSource = ds.Tables["NguyenLieu"];

            //combobox bảng PHIẾU NHẬP
            cbmanl_pn.DataSource = ds.Tables["NguyenLieu"];
            cbmanl_pn.DisplayMember = "TenNguyenLieu";
            cbmanl_pn.ValueMember = "MaNguyenLieu";

            //combobox bảng PHIẾU XUẤT
            cbmanl_px.DataSource = ds.Tables["NguyenLieu"];
            cbmanl_px.DisplayMember = "TenNguyenLieu";
            cbmanl_px.ValueMember = "MaNguyenLieu";

            //combobox bảng CÔNG THỨC
            cbmanl_ct.DataSource = ds.Tables["NguyenLieu"];
            cbmanl_ct.DisplayMember = "TenNguyenLieu";
            cbmanl_ct.ValueMember = "MaNguyenLieu";
        }

        public void getDataPX()
        {
            string query = "EXEC GiaPhieuXuat2; select * from PhieuXuat";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "PhieuXuat");
            dgvPX.DataSource = ds.Tables["PhieuXuat"];


        }

        public void getDataCT()
        {
            string query = "select * from CongThuc";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "CongThuc");
            dgvCT.DataSource = ds.Tables["CongThuc"];


        }

        private void QuanLySanPham_Load(object sender, EventArgs e)
        {
            getDataNCC();
            getDataLSP();
            getDataSP();
            getDataNL();
            getDataPN();
            getDataPX();
            getDataCT();
        }


        //NHÀ CUNG CẤP
        private void btnThoat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnTaomoi_Click(object sender, EventArgs e)
        {
            txtmancc.Text = "";
            txttenncc.Text = "";
            txtsdt.Text = "";
            txtemail.Text = "";
            txtdiachi.Text = ""; 
            txttrangthai.Text = "";
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            string mancc = txtmancc.Text;
            string tenncc = txttenncc.Text;
            string sdt = txtsdt.Text;
            string email = txtemail.Text;
            string diachi = txtdiachi.Text;
            string trangthai = txttrangthai.Text;
          
            string query = "INSERT INTO NhaCungCap(TenNhaCungCap, SDT, Email, DiaChi, TrangThai) VALUES (N'" + tenncc + "', '" + sdt + "', '" + email + "', N'" + diachi + "', N'" + trangthai + "')";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm mới nhà cung cấp thành công!");
                getDataNCC();
            }
            else
            {
                MessageBox.Show("Thêm mới nhà cung cấp thất bại!");
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string mancc = txtmancc.Text;
            string tenncc = txttenncc.Text;
            string sdt = txtsdt.Text;
            string email = txtemail.Text;
            string diachi = txtdiachi.Text;
            string trangthai = txttrangthai.Text;

            string query = "update NhaCungCap set  TenNhaCungCap=N'" + tenncc + "', SDT='" + sdt + "', Email='" + email + "', DiaChi='" + diachi + "', TrangThai=N'" + trangthai + "'where MaNhaCungCap ='" + mancc + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thông tin nhà cung cấp thành công!");
                getDataNCC();
            }
            else
                MessageBox.Show("Sửa thông tin nhà cung cấp thất bại!");
        }

        private void dgvncc_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmancc.Text = dgvncc.Rows[row].Cells["MaNhaCungCap"].Value.ToString();
                txttenncc.Text = dgvncc.Rows[row].Cells["TenNhaCungCap"].Value.ToString();
                txtsdt.Text = dgvncc.Rows[row].Cells["SDT"].Value.ToString();
                txtemail.Text = dgvncc.Rows[row].Cells["Email"].Value.ToString();
                txtdiachi.Text = dgvncc.Rows[row].Cells["DiaChi"].Value.ToString();
                txttrangthai.Text = dgvncc.Rows[row].Cells["TrangThai"].Value.ToString();
                
            }
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string mancc = txtmancc.Text;
            string tenncc = txttenncc.Text;
            string sdt = txtsdt.Text;
            string email = txtemail.Text;
            string diachi = txtdiachi.Text;
            string trangthai = txttrangthai.Text;

            string query = "delete NhaCungCap where MaNhaCungCap = '" + mancc + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thông tin nhà cung cấp thành công!");
                getDataNCC();
            }
            else
            {
                MessageBox.Show("Xóa thông tin nhà cung cấp thất bại!");
            }
        }

        private void txtTimkiem_Click(object sender, EventArgs e)
        {
            string mancc = txtmancc.Text;
            string tenncc = txttenncc.Text;
            /*string query = "select * from NhaCungCap where MaNhaCungCap like '%" + mancc + "%'";*/
            string query2 = string.Format("select * from NhaCungCap where TenNhaCungCap like N'%{0}%'", tenncc);
            DataSet ds = new DataSet();
            KetNoi cn = new KetNoi();
           /* ds = cn.laydulieu(query, "NhaCungCap");*/
            ds = cn.laydulieu(query2, "NhaCungCap");

            dgvncc.DataSource = ds.Tables["NhaCungCap"];
        }


       ///LOẠI SẢN PHẨM//

        private void btnTaomoiLSP_Click(object sender, EventArgs e)
        {
            txtmalsp.Text = "";
            txttenlsp.Text = "";
        }

        private void btnThemLSP_Click(object sender, EventArgs e)
        {
            string malsp = txtmalsp.Text;
            string tenlsp = txttenlsp.Text;

            string query = "INSERT INTO LoaiSanPham(TenLoai) VALUES (N'" + tenlsp + "')";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm mới loại sản phẩm thành công!");
                getDataLSP();
            }
            else
            {
                MessageBox.Show("Thêm mới loại sản phẩm thất bại!");
            }
        }

        private void btnSuaLSP_Click(object sender, EventArgs e)
        {
            string malsp = txtmalsp.Text;
            string tenlsp = txttenlsp.Text;

            string query = "update LoaiSanPham set  TenLoai=N'" + tenlsp + "' where MaLoai ='" + malsp + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thông tin loại sản phẩm thành công!");
                getDataLSP();
            }
            else
                MessageBox.Show("Sửa thông tin loại sản phẩm thất bại!");
        }

        private void btnXoaLSP_Click(object sender, EventArgs e)
        {
            string malsp = txtmalsp.Text;
            string tenlsp = txttenlsp.Text;

            string query = "delete LoaiSanPham where MaLoai = '" + malsp + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thông tin loại sản phẩm  thành công!");
                getDataLSP();
            }
            else
            {
                MessageBox.Show("Xóa thông tin  loại sản phẩm thất bại!");
            }
        }

        private void btnTimkiem_Click(object sender, EventArgs e)
        {
            string malsp = txtmalsp.Text;
            string tenlsp = txttenlsp.Text;
          /*  string query = "select * from LoaiSanPham where MaLoai like '%" + malsp + "%'";*/
            string query1 = string.Format("select * from LoaiSanPham where TenLoai like N'%{0}%'", tenlsp);
            DataSet ds = new DataSet();
            KetNoi cn = new KetNoi();
          /*  ds = cn.laydulieu(query, "LoaiSanPham");*/
            ds = cn.laydulieu(query1, "LoaiSanPham");

            dgvloaisp.DataSource = ds.Tables["LoaiSanPham"];
        }

        private void dgvloaisp_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmalsp.Text = dgvloaisp.Rows[row].Cells["MaLoai"].Value.ToString();
                txttenlsp.Text = dgvloaisp.Rows[row].Cells["TenLoai"].Value.ToString();
            }
        }

        private void btnThoatLSP_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// SẢN PHẨM ////

        private void btnTaomoiSP_Click(object sender, EventArgs e)
        {
            txtmasp.Text = "";
            txttensp.Text = "";
            cbmaloai.Text = "";
            txtsl.Text = "";
            txtgiaban.Text = "";
            
        }

        private void btnThemSP_Click(object sender, EventArgs e)
        {
            string masp = txtmasp.Text;
            string tensp = txttensp.Text;
            string maloai = cbmaloai.SelectedValue.ToString();
            string sl = txtsl.Text;
            string giaban = txtgiaban.Text;
         

            string query = "INSERT INTO SanPham(TenSanPham, MaLoai, SoLuong, GiaBan) VALUES (N'" + tensp +"', N'" + maloai + "', '" + sl + "', N'" + giaban + "')";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm mới sản phẩm thành công!");
                getDataSP();
            }
            else
            {
                MessageBox.Show("Thêm mới sản phẩm thất bại!");
            }
        }

        private void btnXoaSP_Click(object sender, EventArgs e)
        {
            string masp = txtmasp.Text;

            string query = "DELETE SanPham where MaSanPham = '" + masp + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa sản phẩm thành công!");
                getDataSP();
            }
            else
            {
                MessageBox.Show("Xóa sản phẩm thất bại!");
            }
        }

        private void btnSuaSP_Click(object sender, EventArgs e)
        {
            string masp = txtmasp.Text;
            string tensp = txttensp.Text;
            string maloai = cbmaloai.SelectedValue.ToString();
            string sl = txtsl.Text;
            string giaban = txtgiaban.Text;

            string query = "update SanPham set TenSanPham = N'" + tensp + "',MaLoai= '" + maloai + "', SoLuong='" + sl + "',GiaBan= '" + giaban + "'where MaSanPham =  '" + masp + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa sản phẩm thành công!");
                getDataSP();
            }
            else
            {
                MessageBox.Show("Sửa sản phẩm thất bại!");
            }
        }

        private void btnThoatSP_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnTimkiemSP_Click(object sender, EventArgs e)
        {
            string masp = txtmasp.Text;
            string tensp = txttensp.Text;
            string query = "select * from SanPham where MaSanPham like '%" + masp + "%'";
            string query2 = string.Format("select * from SanPham where TenSanPham like N'%{0}%'", tensp);
            DataSet ds = new DataSet();
            KetNoi cn = new KetNoi();
            ds = cn.laydulieu(query, "SanPham");
            ds = cn.laydulieu(query2, "SanPham");
            dgvSP.DataSource = ds.Tables["SanPham"];
        }

        private void dgvSP_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmasp.Text = dgvSP.Rows[row].Cells["MaSanPham"].Value.ToString();
                txttensp.Text = dgvSP.Rows[row].Cells["TenSanPham"].Value.ToString();
                cbmaloai.SelectedValue = dgvSP.Rows[row].Cells["Maloai"].Value.ToString();
                txtsl.Text = dgvSP.Rows[row].Cells["SoLuong"].Value.ToString();
                txtgiaban.Text = dgvSP.Rows[row].Cells["GiaBan"].Value.ToString();
            }
        }

        // PHIẾU NHẬP ///

        private void btnTaomoi_pn_Click(object sender, EventArgs e)
        {
            txtmapn.Text = "";
            cbmancc_pn.Text = "";
            cbmanl_pn.Text = "";
            txtsl_pn.Text = "";
            txtgianhap_pn.Text = "";
            txtngaynhap_pn.Text = "";
        }

        private void btnThem_pn_Click(object sender, EventArgs e)
        {
            string mapn = txtmapn.Text;
            string mancc = cbmancc_pn.SelectedValue.ToString();
            string manl = cbmanl_pn.SelectedValue.ToString();
            string sl = txtsl_pn.Text;
            string gianhap = txtgianhap_pn.Text;
            string ngaynhap = txtngaynhap_pn.Text;

            string query = "INSERT INTO PhieuNhap(MaNhaCungCap, MaNguyenLieu, SoLuong, GiaNhap, NgayNhap) VALUES ('" + mancc + "', '" + manl + "', '" + sl + "', '" + gianhap + "', '" + ngaynhap + "'); EXEC Nhap_NLC ;EXEC CapNhatNhaCungCap; ";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm mới phiếu nhập  thành công!");
                getDataPN();
            }
            else
            {
                MessageBox.Show("Thêm mới phiếu nhập thất bại!");
            }
        }

        private void btnSua_pn_Click(object sender, EventArgs e)
        {
            string mapn = txtmapn.Text;
            string mancc = cbmancc_pn.SelectedValue.ToString();
            string manl = cbmanl_pn.SelectedValue.ToString();
            string sl = txtsl_pn.Text;
            string gianhap = txtgianhap_pn.Text;
            string ngaynhap = txtngaynhap_pn.Text;

            string query = "update PhieuNhap set  MaNhaCungCap=N'" + mancc + "', MaNguyenLieu=N'" + manl + "', SoLuong='" + sl + "', GiaNhap='" + gianhap + "', NgayNhap='" + ngaynhap + "'where MaPhieuNhap ='" + mapn + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thông tin phiếu nhập thành công!");
                getDataPN();
            }
            else
                MessageBox.Show("Sửa thông tin phiếu nhập thất bại!");
        }

        private void btnXoa_pn_Click(object sender, EventArgs e)
        {
            string mapn = txtmapn.Text;

            string query = "delete PhieuNhap where MaPhieuNhap = '" + mapn + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thông tin phiếu nhập thành công!");
                getDataPN();
            }
            else
            {
                MessageBox.Show("Xóa thông tin phiếu nhập thất bại!");
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnTimkiem_pn_Click(object sender, EventArgs e)
        {
            string mapn = txtmapn.Text;
            string query = "select * from PhieuNhap where MaPhieuNhap like '%" + mapn + "%'";
            
            DataSet ds = new DataSet();
            KetNoi cn = new KetNoi();
            ds = cn.laydulieu(query, "PhieuNhap");

            dgvPN.DataSource = ds.Tables["PhieuNhap"];
        }

        private void dgvPN_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmapn.Text = dgvPN.Rows[row].Cells["MaPhieuNhap"].Value.ToString();
                cbmancc_pn.SelectedValue = dgvPN.Rows[row].Cells["MaNhaCungCap"].Value.ToString();
                cbmanl_pn.SelectedValue = dgvPN.Rows[row].Cells["MaNguyenLieu"].Value.ToString();
                txtsl_pn.Text = dgvPN.Rows[row].Cells["SoLuong"].Value.ToString();
                txtgianhap_pn.Text = dgvPN.Rows[row].Cells["GiaNhap"].Value.ToString();
                txtngaynhap_pn.Text = dgvPN.Rows[row].Cells["NgayNhap"].Value.ToString();
            }
        }


        //NGUYÊN LIỆU //////
        private void btnTaomoi_nl_Click(object sender, EventArgs e)
        {
            txtmanl.Text = "";
            txttennl.Text = "";
            cbmancc_nl.Text = "";
            txtsltonkho.Text = "";
            txtgianhap_nl.Text = "";           
        }

        private void btnThem_nl_Click(object sender, EventArgs e)
        {
            string manl = txtmanl.Text;
            string tennl = txttennl.Text;
            string mancc = cbmancc_nl.SelectedValue.ToString();
            string sltk = txtsltonkho.Text;
            string gianhap = txtgianhap_nl.Text;
            
            string query = "INSERT INTO NguyenLieu(MaNhaCungCap, TenNguyenLieu, SoLuongTonKho, GiaNhap) VALUES (N'" + tennl + "', '" + mancc + "', '" + sltk + "', '" + gianhap + "')";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm mới nguyên liệu  thành công!");
                getDataNL();
            }
            else
            {
                MessageBox.Show("Thêm mới nguyên liệu thất bại!");
            }
        }

        private void btnSua_nl_Click(object sender, EventArgs e)
        {
            string manl = txtmanl.Text;
            string tennl = txttennl.Text;
            string mancc = cbmancc_nl.SelectedValue.ToString();
            string sltk = txtsltonkho.Text;
            string gianhap = txtgianhap_nl.Text;

            string query = "update NguyenLieu set TenNguyenLieu=N'" + tennl + "', MaNhaCungCap='" + mancc + "', SoLuongTonKho='" + sltk + "', GiaNhap='" + gianhap + "'where MaNguyenLieu=N'" + manl + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thông tin nguyên liệu thành công!");
                getDataNL();
            }
            else
                MessageBox.Show("Sửa thông tin nguyên liệu thất bại!");
        }

        private void btnXoa_nl_Click(object sender, EventArgs e)
        {
            string manl = txtmanl.Text;

            string query = "delete NguyenLieu where MaNguyenLieu = '" + manl + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thông tin nguyên liệu thành công!");
                getDataNL();
            }
            else
            {
                MessageBox.Show("Xóa thông tin nguyên liệu thất bại!");
            }
        }

        private void btnThoat_nl_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dgvNL_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmanl.Text = dgvNL.Rows[row].Cells["MaNguyenLieu"].Value.ToString();
                txttennl.Text = dgvNL.Rows[row].Cells["TenNguyenLieu"].Value.ToString();
                cbmancc_nl.SelectedValue = dgvNL.Rows[row].Cells["MaNhaCungCap"].Value.ToString();
                txtsltonkho.Text = dgvNL.Rows[row].Cells["SoLuongTonKho"].Value.ToString();
                txtgianhap_nl.Text = dgvNL.Rows[row].Cells["GiaNhap"].Value.ToString();
               
            }
        }

        private void btnTimkiem_nl_Click(object sender, EventArgs e)
        {
            string manl = txtmanl.Text;
            string tennl = txttennl.Text;
            string query = "select * from NguyenLieu where MaNguyenLieu like '%" + manl + "%'";
            string query2 = string.Format("select * from NguyenLieu where TenNguyenLieu like N'%{0}%'", tennl);
            DataSet ds = new DataSet();
            DataSet ds1 = new DataSet();
            KetNoi cn = new KetNoi();
            ds = cn.laydulieu(query, "NguyenLieu");
            ds1 = cn.laydulieu(query2, "NguyenLieu");
            dgvNL.DataSource = ds.Tables["NguyenLieu"];
        }


        //PHIẾU XUẤT///
        private void btnTaomoiPX_Click(object sender, EventArgs e)
        {
            txtmapx.Text = "";
            cbmanl_px.Text = "";
            txtsl_px.Text = "";
            txtgia_px.Text = "";
            txtngayxuat.Text = "";
        }

        private void btnThemPX_Click(object sender, EventArgs e)
        {
            string mapx = txtmapx.Text;
            string manl = cbmanl_px.SelectedValue.ToString();
            string slpx = txtsl_px.Text;
            string gia = txtgia_px.Text;
            string ngayxuat = txtngayxuat.Text; 

            string query = "INSERT INTO PhieuXuat(MaNguyenLieu, SoLuong, Gia, NgayXuat) VALUES ('" + manl + "', '" + slpx + "', '" + gia + "', '" + ngayxuat + "'); EXEC Xuat_NLC";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm mới phiếu xuất  thành công!");
                getDataPX();
            }
            else
            {
                MessageBox.Show("Thêm mới phiếu xuất thất bại!");
            }
        }

        private void btnSuaPX_Click(object sender, EventArgs e)
        {
            string mapx = txtmapx.Text;
            string manl = cbmanl_px.SelectedValue.ToString();
            string slpx = txtsl_px.Text;
            string gia = txtgia_px.Text;
            string ngayxuat = txtngayxuat.Text;

            string query = "update PhieuXuat set MaNguyenLieu=N'" + manl + "', SoLuong='" + slpx + "', Gia='" + gia + "', NgayXuat='" + ngayxuat + "'where MaPhieuXuat=N'" + mapx + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thông tin phiếu xuất thành công!");
                getDataPX();
            }
            else
                MessageBox.Show("Sửa thông tin phiếu xuất thất bại!");
        }

        private void btnXoaPX_Click(object sender, EventArgs e)
        {
            string mapx = txtmapx.Text;

            string query = "delete PhieuXuat where MaPhieuXuat = '" + mapx + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thông tin phiếu xuất thành công!");
                getDataPX();
            }
            else
            {
                MessageBox.Show("Xóa thông tin phiếu xuất thất bại!");
            }
        }

        private void btnThoatPX_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtTimkiemPX_Click(object sender, EventArgs e)
        {
            string mapx = txtmapx.Text;
            string query = "select * from PhieuXuat where MaPhieuXuat like '%" + mapx + "%'";
            
            DataSet ds = new DataSet();
           
            KetNoi cn = new KetNoi();
            ds = cn.laydulieu(query, "PhieuXuat");
            
            dgvPX.DataSource = ds.Tables["PhieuXuat"];
        }

        private void dgvPX_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmapx.Text = dgvPX.Rows[row].Cells["MaPhieuXuat"].Value.ToString();
                cbmanl_px.SelectedValue = dgvPX.Rows[row].Cells["MaNguyenLieu"].Value.ToString();
                txtsl_px.Text = dgvPX.Rows[row].Cells["SoLuong"].Value.ToString();
                txtgia_px.Text = dgvPX.Rows[row].Cells["Gia"].Value.ToString();
                txtngayxuat.Text = dgvPX.Rows[row].Cells["NgayXuat"].Value.ToString();

            }
        }


        //CÔNG THỨC///

        private void btnTaomoiCT_Click(object sender, EventArgs e)
        {
            txtmact.Text = "";
            cbmasp_ct.Text = "";
            cbmanl_ct.Text = "";
            txtuoclong.Text = "";
            
        }

        private void btnThemCT_Click(object sender, EventArgs e)
        {
            string mact = txtmact.Text;
            string masp = cbmasp_ct.SelectedValue.ToString();
            string manl = cbmanl_ct.SelectedValue.ToString();
            string uoclong = txtuoclong.Text;
       
            string query = "INSERT INTO CongThuc(MaSanPham, MaNguyenLieu, UocLuongSLTP) VALUES ('" + masp + "', '" + manl + "', '" + uoclong + "')";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm mới công thức  thành công!");
                getDataCT();
            }
            else
            {
                MessageBox.Show("Thêm mới công thức thất bại!");
            }
        }

        private void btnSuaCT_Click(object sender, EventArgs e)
        {
            string mact = txtmact.Text;
            string masp = cbmasp_ct.SelectedValue.ToString();
            string manl = cbmanl_ct.SelectedValue.ToString();
            string uoclong = txtuoclong.Text;

            string query = "update CongThuc set MaSanPham=N'" + masp + "', MaNguyenLieu='" + manl + "', UocLuongSLTP='" + uoclong + "'where MaCongThuc=N'" + mact + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thông tin công thức thành công!");
                getDataCT();
            }
            else
                MessageBox.Show("Sửa thông tin công thức thất bại!");
        }

        private void btnXoaCT_Click(object sender, EventArgs e)
        {
            string mact = txtmact.Text;

            string query = "delete CongThuc where MaCongThuc = '" + mact + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thông tin công thức thành công!");
                getDataCT();
            }
            else
            {
                MessageBox.Show("Xóa thông tin công thức thất bại!");
            }
        }

        private void dgvCT_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmact.Text = dgvCT.Rows[row].Cells["MaCongThuc"].Value.ToString();
                cbmasp_ct.SelectedValue = dgvCT.Rows[row].Cells["MaSanPham"].Value.ToString();
                cbmanl_ct.SelectedValue = dgvCT.Rows[row].Cells["MaNguyenLieu"].Value.ToString();
                txtuoclong.Text = dgvCT.Rows[row].Cells["UocLuongSLTP"].Value.ToString();
               
            }
        }

        private void btnTimKiemCT_Click(object sender, EventArgs e)
        {
            string mact = txtmact.Text;
            string query = "select * from CongThuc where MaCongThuc like '%" + mact + "%'";

            DataSet ds = new DataSet();

            KetNoi cn = new KetNoi();
            ds = cn.laydulieu(query, "CongThuc");

            dgvCT.DataSource = ds.Tables["CongThuc"];
        }

        private void btnThoatCT_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        /// XUẤT FILE EXCEL  DANH SÁCH SẢN PHẨM 
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

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Xuất file thành công!");
            excel(dgvSP, @"D:\", @"Danh_Sach_San_Pham");
        }

        //XUẤT FILE EXCEL PHIẾU NHẬP 
        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Xuất file thành công!");
            excel(dgvPN, @"D:\", @"Danh_Sach_Phieu_Nhap");
        }

        //XUẤT FILE EXCEL PHIẾU XUẤT
        private void button5_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Xuất file thành công!");
            excel(dgvPX, @"D:\", @"Danh_Sach_Phieu_Xuat");
        }
    }
}

  

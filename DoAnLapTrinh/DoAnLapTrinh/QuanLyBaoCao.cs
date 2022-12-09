using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using app = Microsoft.Office.Interop.Excel.Application;

namespace DoAnLapTrinh
{
    public partial class QuanLyBaoCao : Form
    {
        SqlConnection conn = new KetNoi().conn;
        public QuanLyBaoCao()
        {
            InitializeComponent();
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        //BÁO CÁO THỐNG KÊ 
        public void getDataBCTK()
        {
            string query = " select * from BaoCaoThongKe";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "BaoCaoThongKe");
            dgvBaoCaoTK.DataSource = ds.Tables["BaoCaoThongKe"];

            //combobox bảng THỐNG KÊ THU CHI
            cbmabc_tktc.DataSource = ds.Tables["BaoCaoThongKe"];
            cbmabc_tktc.DisplayMember = "MaBaoCao";
            cbmabc_tktc.ValueMember = "MaBaoCao";
        }


        public void getDataNS()
        {
            string query = "  select * from NhanSu";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "NhanSu");

            //combobox bảng BÁO CÁO THỐNG KÊ 
            cbmans.DataSource = ds.Tables["NhanSu"];
            cbmans.DisplayMember = "HoTen";
            cbmans.ValueMember = "MaNhanSu";
        }

        //THỐNG KÊ NHẬP XUẤT   

        public void getDataTKNX()
        {
            string query = " EXEC Xuat; EXEC Nhap; select * from ThongKeNhapXuat ";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "ThongKeNhapXuat");
            dgvTKTC.DataSource = ds.Tables["ThongKeNhapXuat"];
        }
        //public void getDataNL()
        //{
        //    string query = "  select * from NguyenLieu";
        //    KetNoi kn = new KetNoi();
        //    DataSet ds = new DataSet();
        //    ds = kn.laydulieu(query, "NguyenLieu");

        //    //combobox bảng BÁO CÁO THỐNG KÊ 
        //    cbmanl.DataSource = ds.Tables["NguyenLieu"];
        //    cbmanl.DisplayMember = "TenNguyenLieu";
        //    cbmanl.ValueMember = "MaNguyenLieu";
        //}

        //THỐNG KÊ THU CHI
        public void getDataTKTC()
        {
            string query = " EXEC TongLuongNV; EXEC TienChi; EXEC TienThu; EXEC TienLai; select * from ThongKeThuChi";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "ThongKeThuChi");
            dgvTKTC.DataSource = ds.Tables["ThongKeThuChi"];
        }

        private void QuanLyBaoCao_Load(object sender, EventArgs e)
        {
            getDataBCTK();
            getDataTKNX();
            getDataTKTC();
            getDataNS();
            //getDataNL();
        }

        //BÁO CÁO THỐNG KÊ//
        private void btnTaomoi_Click(object sender, EventArgs e)
        {
            txtmabaocao.Text = "";
            txtthoigian.Text = "";
            cbmans.Text = "";
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            string mabaocao = txtmabaocao.Text;
            string thoigian = txtthoigian.Text;
            string mans = cbmans.SelectedValue.ToString();


            string query = "INSERT INTO BaoCaoThongKe(ThoiGian, MaNhanSu) VALUES ('" + thoigian + "', '" + mans + "')";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm mới báo cáo thống kê thành công!");
                getDataBCTK();
            }
            else
            {
                MessageBox.Show("Thêm mới báo cáo thống kê cấp thất bại!");
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string mabaocao = txtmabaocao.Text;
            string thoigian = txtthoigian.Text;
            string mans = cbmans.SelectedValue.ToString();

            string query = "update BaoCaoThongKe set ThoiGian='" + thoigian + "', MaNhanSu='" + mans + "'where MaBaoCao ='" + mabaocao + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thông tin báo cáo thống kê thành công!");
                getDataBCTK();
            }
            else
                MessageBox.Show("Sửa thông tin báo cáo thống kê thất bại!");
        }

        private void btnTimkiem_BCTK_Click(object sender, EventArgs e)
        {
            string thoigian = txtthoigian.Text;
            string query = "select * from BaoCaoThongKe where ThoiGian like '%" + thoigian + "%'";

            DataSet ds = new DataSet();
            KetNoi cn = new KetNoi();
            ds = cn.laydulieu(query, "BaoCaoThongKe");
            dgvBaoCaoTK.DataSource = ds.Tables["BaoCaoThongKe"];
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string mabaocao = txtmabaocao.Text;

            string query = "delete BaoCaoThongKe where MaBaoCao = '" + mabaocao + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thông tin báo cáo thống kê thành công!");
                getDataBCTK();
            }
            else
            {
                MessageBox.Show("Xóa thông tin báo cáo thống kê thất bại!");
            }
        }

        private void dgvBaoCaoTK_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmabaocao.Text = dgvBaoCaoTK.Rows[row].Cells["MaBaoCao"].Value.ToString();
                txtthoigian.Text = dgvBaoCaoTK.Rows[row].Cells["ThoiGian"].Value.ToString();
                cbmans.SelectedValue = dgvBaoCaoTK.Rows[row].Cells["MaNhanSu"].Value.ToString();
            }
        }

        // THỐNG KÊ THU CHI //
        private void btnTaomoi_TKCT_Click(object sender, EventArgs e)
        {
            txtmathuchi.Text = "";
            cbmabc_tktc.Text = "";
            txtsotienchi.Text = "";
            txtsotienthu.Text = "";
            txtslnv.Text = "";
            txtsotienlai.Text = "";
        }

        private void btnThem_TKCT_Click(object sender, EventArgs e)
        {
            string matc = txtmathuchi.Text;
            string mabc = cbmabc_tktc.SelectedValue.ToString();
            string tienchi = txtsotienchi.Text;
            string tienthu = txtsotienthu.Text;
            string slnv = txtslnv.Text;
            string sotienlai = txtsotienlai.Text;

            string query = "INSERT INTO ThongKeThuChi(MaBaoCao, SoTienChi, SoTienThu, TongLuongNhanVien, SoTienLai) VALUES ('" + mabc + "', '" + tienchi + "', '" + tienthu + "', '" + slnv + "', '" + sotienlai + "'); EXEC TienChi; EXEC TongLuongNV; EXEC TienThu; EXEC TienLai";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm mới thống kê thu chi thành công!");
                getDataTKTC();
            }
            else
            {
                MessageBox.Show("Thêm mới thống kê thu chi thất bại!");
            }
        }

        private void btnSua_TKCT_Click(object sender, EventArgs e)
        {
            string matc = txtmathuchi.Text;
            string mabc = cbmabc_tktc.SelectedValue.ToString();
            string tienchi = txtsotienchi.Text;
            string tienthu = txtsotienthu.Text;
            string slnv = txtslnv.Text;
            string sotienlai = txtsotienlai.Text;

            string query = "update ThongKeThuChi set MaBaoCao = N'" + mabc + "', SoTienChi='" + tienchi + "',SoTienThu= '" + tienthu + "',TongLuongNhanVien= '" + slnv + "',SoTienLai= '" + sotienlai + "'where MaThuChi =  '" + matc + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thống kê thu chi thành công!");
                getDataTKTC();
            }
            else
            {
                MessageBox.Show("Sửa thống kê thu chi thất bại!");
            }
        }

        private void btnXoa_TKCT_Click(object sender, EventArgs e)
        {
            string matc = txtmathuchi.Text;

            string query = "DELETE ThongKeThuChi where MaThuChi = '" + matc + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thống kê thu chi thành công!");
                getDataTKTC();
            }
            else
            {
                MessageBox.Show("Xóa thống kê thu chi thất bại!");
            }
        }

        private void btnThoat_TKCT_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnTimkiem_TKTC_Click(object sender, EventArgs e)
        {
            string mabc = cbmabc_tktc.Text;
            string query = "select * from ThongKeThuChi where MaBaoCao like '%" + mabc + "%'";

            DataSet ds = new DataSet();
            KetNoi cn = new KetNoi();
            ds = cn.laydulieu(query, "ThongKeThuChi");
            dgvTKTC.DataSource = ds.Tables["ThongKeThuChi"];
        }

        private void dgvTKTC_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmathuchi.Text = dgvTKTC.Rows[row].Cells["MaThuChi"].Value.ToString();
                cbmabc_tktc.SelectedValue = dgvTKTC.Rows[row].Cells["MaBaoCao"].Value.ToString();
                txtsotienchi.Text = dgvTKTC.Rows[row].Cells["SoTienChi"].Value.ToString();
                txtsotienthu.Text = dgvTKTC.Rows[row].Cells["SoTienThu"].Value.ToString();
                txtslnv.Text = dgvTKTC.Rows[row].Cells["TongLuongNhanVien"].Value.ToString();
                txtsotienlai.Text = dgvTKTC.Rows[row].Cells["SoTienLai"].Value.ToString();
            }
        }


        /// XUẤT FILE EXCEL  THỐNG KÊ 
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

        //THỐNG KÊ THU CHI 
        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Xuất file thành công!");
            excel(dgvTKTC, @"D:\", @"Thong_Ke_Thu_Chi");
        }

        private void btnhienthi_Click(object sender, EventArgs e)
        {
            string selectQuerydonhang = "Select COUNT(MaHoaDon) FROM HoaDon";
            string selectQuerydoanhthu = "Select SUM(TongTien) FROM HoaDon";
            string selectQuerysanpham = "Select SUM(SoLuong) FROM CTHoaDon";

            conn.Open();

            SqlCommand sqlcmd1 = new SqlCommand(selectQuerydonhang, conn);
            SqlCommand sqlcmd2 = new SqlCommand(selectQuerydoanhthu, conn);
            SqlCommand sqlcmd3 = new SqlCommand(selectQuerysanpham, conn);

            var result1 = sqlcmd1.ExecuteScalar();
            var result2 = sqlcmd2.ExecuteScalar();
            var result3 = sqlcmd3.ExecuteScalar();

            conn.Close();

            txtdonhang.Text = result1.ToString();
            txtdoanhthu.Text = result2.ToString();
            txtsanpham.Text = result3.ToString();

        }
    }
}

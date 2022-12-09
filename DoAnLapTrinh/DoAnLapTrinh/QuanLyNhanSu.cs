using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using app = Microsoft.Office.Interop.Excel.Application;

namespace DoAnLapTrinh
{
    public partial class QuanLyNhanSu : Form
    {
        public QuanLyNhanSu()
        {
            InitializeComponent();
        }
        public void getData()
        {
            string query = "select * from NhanSu";
            KetNoi kn = new KetNoi();
            DataSet ds = new DataSet();
            ds = kn.laydulieu(query, "NhanSu");
            dgvQLNS.DataSource = ds.Tables["NhanSu"];


        }

        private void QuanLyNhanSu_Load(object sender, EventArgs e)
        {
            getData();
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        private void btnTaomoi_Click(object sender, EventArgs e)
        {
            txtmans.Text = "";
            txthoten.Text = "";
            /*cbgioitinh.Text = "";*/
            txtngaysinh.Text = "";
            txtcmnd.Text = "";
            txtvitri.Text = "";
            txttrangthai.Text = "";
            txtthoigianvaolam.Text = "";
            txtsongaycong.Text = "";
            txtluong.Text = "";
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            string mans = txtmans.Text;
            string hoten = txthoten.Text;
            string gioitinh = txtgioitinh.Text;
            string ngaysinh = txtngaysinh.Text;
            string cmnd = txtcmnd.Text;
            string vitri = txtvitri.Text;
            string trangthai = txttrangthai.Text;
            string thoigianvaolam = txtthoigianvaolam.Text;
            string songaycong = txtsongaycong.Text;
            string luong = txtluong.Text;

            string query = "INSERT INTO NhanSu(HoTen, GioiTinh, NgaySinh, CMND, ViTri, TrangThai, ThoiGianVaoLam, SoNgayCong, LuongThang) VALUES (N'" + hoten + "', N'" + gioitinh + "','" + ngaysinh + "','" + cmnd + "',N'" + vitri + "',N'" + trangthai + "','" + thoigianvaolam + "','" + songaycong + "','" + luong + "')";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Thêm mới nhân sự thành công!");
                getData();
            }
            else
            {
                MessageBox.Show("Thêm mới nhân sự thất bại!");
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string mans = txtmans.Text;
            string hoten = txthoten.Text;
            string gioitinh = txtgioitinh.Text;
            string ngaysinh = txtngaysinh.Text;
            string cmnd = txtcmnd.Text;
            string vitri = txtvitri.Text;
            string trangthai = txttrangthai.Text;
            string thoigianvaolam = txtthoigianvaolam.Text;
            string songaycong = txtsongaycong.Text;
            string luong = txtluong.Text;

            string query = "update NhanSu set  HoTen=N'" + hoten + "', GioiTinh=N'" + gioitinh + "', NgaySinh='" + ngaysinh + "', CMND='" + cmnd + "', ViTri=N'" + vitri + "', TrangThai=N'" + trangthai + "', ThoiGianVaoLam='" + thoigianvaolam + "', SoNgayCong='" + songaycong + "', LuongThang='" + luong + "'where MaNhanSu ='" + mans + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Sửa thông tin nhân sự thành công!");
                getData();
            }
            else
                MessageBox.Show("Sửa thông tin nhân sự thất bại!");
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string mans = txtmans.Text;
            string hoten = txthoten.Text;
            string gioitinh = txtgioitinh.Text;
            string cmnd = txtcmnd.Text;
            string vitri = txtvitri.Text;
            string trangthai = txttrangthai.Text;
            string songaycong = txtsongaycong.Text;
            string luong = txtluong.Text;

            string query = "delete NhanSu where MaNhanSu = '" + mans + "'";
            KetNoi kn = new KetNoi();
            bool kt = kn.thucthi(query);
            if (kt == true)
            {
                MessageBox.Show("Xóa thông tin nhân sự thành công!");
                getData();
            }
            else
            {
                MessageBox.Show("Xóa thông tin nhân sự thất bại!");
            }
        }

        private void btnTimkiem_Click(object sender, EventArgs e)
        {
            string mans = txtmans.Text;
            string hoten = txthoten.Text;
            /*string query = "select * from NhanSu where MaNhanSu like '%" + mans + "%'";*/
            string query2 = string.Format("select * from NhanSu where HoTen like N'%{0}%'", hoten);
            DataSet ds = new DataSet();
            KetNoi cn = new KetNoi();
            ds = cn.laydulieu(query2, "NhanSu");

            dgvQLNS.DataSource = ds.Tables["NhanSu"];
        }

        private void dgvQLNS_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            if (row >= 0)
            {
                txtmans.Text = dgvQLNS.Rows[row].Cells["MaNhanSu"].Value.ToString();
                txthoten.Text = dgvQLNS.Rows[row].Cells["HoTen"].Value.ToString();
                txtgioitinh.Text = dgvQLNS.Rows[row].Cells["GioiTinh"].Value.ToString();
                txtngaysinh.Text = dgvQLNS.Rows[row].Cells["Ngaysinh"].Value.ToString();
                txtcmnd.Text = dgvQLNS.Rows[row].Cells["CMND"].Value.ToString();
                txtvitri.Text = dgvQLNS.Rows[row].Cells["ViTri"].Value.ToString();
                txttrangthai.Text = dgvQLNS.Rows[row].Cells["TrangThai"].Value.ToString();
                txtthoigianvaolam.Text = dgvQLNS.Rows[row].Cells["ThoiGianVaoLam"].Value.ToString();
                txtsongaycong.Text = dgvQLNS.Rows[row].Cells["SoNgayCong"].Value.ToString();
                txtluong.Text = dgvQLNS.Rows[row].Cells["LuongThang"].Value.ToString();
            }
        }


        ///Xuất File Excel 
        ///

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
            MessageBox.Show("Xuất file thành công!!");
            excel(dgvQLNS, @"D:\", @"Danh_Sach_Nhan_Vien");

            
        }

    }
}

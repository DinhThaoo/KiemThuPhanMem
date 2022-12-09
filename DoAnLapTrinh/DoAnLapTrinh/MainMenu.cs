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
    public partial class MainMenu : Form
    {
        public MainMenu()
        {
            InitializeComponent();
        }

        private void đăngXuấtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult tb = MessageBox.Show("Bạn có muốn đăng xuất hệ thống không?", "Thông báo",
               MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (tb == DialogResult.OK)
                Application.Exit();
        }

        private void tàiKhoảnĐăngNhậpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TaiKhoanDangNhap frm = new TaiKhoanDangNhap();
            frm.Show();
        }

        private void trởLạiTrangĐăngNhậpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DangNhap frm = new DangNhap();
            frm.Show();
            this.Hide();

        }

        private void nhânViênToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void nhânSựToolStripMenuItem_Click(object sender, EventArgs e)
        {
            QuanLyNhanSu frm = new QuanLyNhanSu();
            frm.Show();
        }

        private void kháchHàngToolStripMenuItem_Click(object sender, EventArgs e)
        {
            QuanLyKhachHang frm = new QuanLyKhachHang();
            frm.Show();
        }

        private void nhàCungCấpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            QuanLySanPham frm = new QuanLySanPham();
            frm.Show();
        }

        private void quảnLýSốBànToolStripMenuItem_Click(object sender, EventArgs e)
        {
            QuanLyBanHang frm = new QuanLyBanHang();
            frm.Show();
        }

        private void báoCáoThốngKêToolStripMenuItem_Click(object sender, EventArgs e)
        {
            QuanLyBaoCao frm = new QuanLyBaoCao();
            frm.Show();
        }

        private void hồSơNhânViênToolStripMenuItem_Click(object sender, EventArgs e)
        {
            HoSoNhanVien frm = new HoSoNhanVien();
            frm.Show();
        }

        private void hồSơThựcTậpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            HoSoThucTap frm = new HoSoThucTap();
            frm.Show();
        }

        private void sảnPhẩmToolStripMenuItem_Click(object sender, EventArgs e)
        {
            QuanLySanPham frm = new QuanLySanPham();
            frm.Show();
        }

        private void quảnLýBánHàngToolStripMenuItem_Click(object sender, EventArgs e)
        {
            QuanLyBanHang frm = new QuanLyBanHang();
            frm.Show();
        }

        private void quảnLýBáoCáoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            QuanLyBaoCao frm = new QuanLyBaoCao();
            frm.Show();
        }

        private void nhânViênToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            QuanLyNhanSu frm = new QuanLyNhanSu();
            frm.Show();
        }
    }
}

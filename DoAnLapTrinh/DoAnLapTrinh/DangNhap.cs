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
    public partial class DangNhap : Form
    {
        public DangNhap()
        {
            InitializeComponent();
        }

        private void DangNhap_Load(object sender, EventArgs e)
        {
            
        }

        private void btnthoat_Click(object sender, EventArgs e)
        {
            DialogResult rs = MessageBox.Show("Bạn có chắc muốn thoát không ? ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            if (rs == DialogResult.OK)
            {
                Application.Exit();

            }
        }

        private void btndn_Click(object sender, EventArgs e)
        {

            string tendn = txttendn.Text;
            string mk = txtmk.Text;
            KetNoi kn = new KetNoi();
            //Admin
            string query = "select count (*) as soluong from TaiKhoan " +
                            "where TenDangNhap ='" + tendn + "'and MatKhau='" + mk + "' and LoaiTaiKhoan='1' ";
            DataSet ds = kn.laydulieu(query, "TaiKhoan");
            //Nhân viên
            string querynv = "select count (*) as soluong from TaiKhoan " +
                            "where TenDangNhap ='" + tendn + "'and MatKhau='" + mk + "' and LoaiTaiKhoan='0' ";
            DataSet dsnv = kn.laydulieu(querynv, "TaiKhoan");

            if ((int)ds.Tables["TaiKhoan"].Rows[0].ItemArray[0] == 1)
            {
                MessageBox.Show("Đăng nhập thành công!");
                MainMenu frm = new MainMenu();
                frm.Show();
                this.Hide();
            }
            else
            {
                if ((int)dsnv.Tables["TaiKhoan"].Rows[0].ItemArray[0] == 1)
                {
                    MessageBox.Show("Đăng nhập thành công!");
                    MenuNhanVien frmnv = new MenuNhanVien();
                    frmnv.Show();
                    this.Hide();
                }
                else
                {
                    if (tendn == " " || mk == "")
                    {
                        this.lblstatus.ForeColor = Color.Red;
                        this.lblstatus.Text = "Tài khoản hoặc mật khẩu trống. Vui lòng nhập mật khẩu hoặc tài khoản!";
                        this.txttendn.Focus();

                    }
                    else
                    {
                        MessageBox.Show("Đăng nhập thất bại!");
                        this.lblstatus.ForeColor = Color.Red;
                        this.lblstatus.Text = "Tài khoản không tồn tại";
                        this.txttendn.Clear();
                        this.txtmk.Clear();
                        this.txttendn.Focus();
                    }
                }
            }
        }       
    }
}

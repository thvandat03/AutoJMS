using System;
using System.Windows.Forms;

namespace AutoJMS
{
    public partial class frmLogin : Form
    {
        public string EnteredKey { get; private set; } = "";
        private string _myHwid;
        public frmLogin(string hwid)
        {
            InitializeComponent();
            _myHwid = hwid;
            txt_hwid.Text = _myHwid;
        }


        private void btn_copy_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(_myHwid);
            MessageBox.Show("Đã copy mã yêu cầu vào bộ nhớ tạm! Để kích hoạt vui lòng liên hệ 0355520331", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btn_activate_Click(object sender, EventArgs e)
        {
            string key = txt_key.Text.Trim();
            if (string.IsNullOrEmpty(key))
            {
                MessageBox.Show("Vui lòng nhập Key kích hoạt!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            EnteredKey = key;

            this.DialogResult = DialogResult.OK;
            this.Close(); 
        }

        private void btn_exit_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
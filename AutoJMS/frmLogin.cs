using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Sunny.UI;


namespace AutoJMS
{
    public partial class frmLogin : UIForm
    {
        public string EnteredKey { get; private set; } = "";
        private string _myHwid;
        private string watermarkText = "Nhập mã kích hoạt của bạn vào đây...";
        public frmLogin(string hwid)
        {
            InitializeComponent();

            _myHwid = hwid;
            txt_hwid.Text = _myHwid;
            this.ShowRadius = true;
            this.ZoomScaleSize = new Size(this.Width, this.Height);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            txt_key.Text = watermarkText;
            txt_key.ForeColor = Color.Gray;

            txt_key.Enter += txt_key_Enter;
            txt_key.Leave += txt_key_Leave;
        }

        private void txt_hwid_ButtonClick(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txt_hwid.Text))
            {
                Clipboard.SetText(_myHwid);
                UIMessageTip.Show("Đã copy mã yêu cầu vào bộ nhớ tạm!");
            }
            else
            {
                UIMessageTip.ShowWarning("Ô dữ liệu đang trống, không có gì để copy!");
            }
        }
        private void txt_key_Enter(object sender, EventArgs e)
        {
            if (txt_key.Text == watermarkText)
            {
                txt_key.Text = "";
                txt_key.ForeColor = Color.Black; 
            }
        }

        private void txt_key_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_key.Text))
            {
                txt_key.Text = watermarkText;
                txt_key.ForeColor = Color.Gray;
            }
        }
        private void btn_activate_Click(object sender, EventArgs e)
        {
            string key = txt_key.Text.Trim();
            if (string.IsNullOrEmpty(key))
            { 
                MessageBox.Show("Vui lòng nhập Key kích hoạt để tiếp tục!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            EnteredKey = key;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

    }

}  
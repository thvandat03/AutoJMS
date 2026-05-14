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

        private async void btn_activate_Click(object sender, EventArgs e)
        {
            string key = txt_key.Text.Trim();
            if (string.IsNullOrEmpty(key))
            {
                MessageBox.Show("Vui lòng nhập Key kích hoạt để tiếp tục!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string originalText = btn_active.Text;

            btn_active.Enabled = false;
            btn_active.Text = "Đang kết nối...";

            try
            {
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(60));

                (bool success, string message, string token) = await LicenseApiService.VerifyLicenseSecureAsync(key, Program.HWID, cts.Token);

                if (success)
                {
                    EnteredKey = key;
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show(message, "Kết nối thất bại", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Đường truyền mạng đang khônng ổn định.\nVui lòng thử lại sau!", "Lỗi kết nối", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi hệ thống:\n" + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btn_active.Enabled = true;
                btn_active.Text = originalText;
            }
        }

    }

}  
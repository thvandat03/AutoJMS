namespace AutoJMS
{
    partial class frmLogin
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            uiLabel1 = new Sunny.UI.UILabel();
            uiLabel2 = new Sunny.UI.UILabel();
            txt_hwid = new Sunny.UI.UITextBox();
            uiLabel3 = new Sunny.UI.UILabel();
            txt_key = new Sunny.UI.UITextBox();
            btn_active = new Sunny.UI.UISymbolButton();
            uiLabel4 = new Sunny.UI.UILabel();
            btn_copy = new Sunny.UI.UISymbolButton();
            SuspendLayout();
            // 
            // uiLabel1
            // 
            uiLabel1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            uiLabel1.AutoSize = true;
            uiLabel1.Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold);
            uiLabel1.ForeColor = Color.FromArgb(48, 48, 48);
            uiLabel1.Location = new Point(20, 50);
            uiLabel1.Name = "uiLabel1";
            uiLabel1.Size = new Size(400, 42);
            uiLabel1.Style = Sunny.UI.UIStyle.Custom;
            uiLabel1.TabIndex = 0;
            uiLabel1.Text = "Bản quyền hiện tại không xác định hoặc không hợp lệ.\r\nVui lòng kích hoạt để sử dụng...";
            // 
            // uiLabel2
            // 
            uiLabel2.AutoSize = true;
            uiLabel2.Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold);
            uiLabel2.ForeColor = Color.FromArgb(48, 48, 48);
            uiLabel2.Location = new Point(20, 130);
            uiLabel2.Name = "uiLabel2";
            uiLabel2.Size = new Size(126, 21);
            uiLabel2.TabIndex = 1;
            uiLabel2.Text = "REQUEST CODE:";
            // 
            // txt_hwid
            // 
            txt_hwid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txt_hwid.ButtonWidth = 35;
            txt_hwid.Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold);
            txt_hwid.Location = new Point(20, 160);
            txt_hwid.Margin = new Padding(4, 5, 4, 5);
            txt_hwid.MinimumSize = new Size(1, 16);
            txt_hwid.Name = "txt_hwid";
            txt_hwid.Padding = new Padding(5);
            txt_hwid.Radius = 10;
            txt_hwid.ReadOnly = true;
            txt_hwid.ShowText = false;
            txt_hwid.Size = new Size(373, 40);
            txt_hwid.SymbolSize = 32;
            txt_hwid.TabIndex = 2;
            txt_hwid.TextAlignment = ContentAlignment.MiddleCenter;
            txt_hwid.Watermark = "";
            txt_hwid.WordWarp = false;
            txt_hwid.ButtonClick += txt_hwid_ButtonClick;
            // 
            // uiLabel3
            // 
            uiLabel3.AutoSize = true;
            uiLabel3.Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold);
            uiLabel3.ForeColor = Color.FromArgb(48, 48, 48);
            uiLabel3.Location = new Point(20, 210);
            uiLabel3.Name = "uiLabel3";
            uiLabel3.Size = new Size(105, 21);
            uiLabel3.TabIndex = 3;
            uiLabel3.Text = "LICENSE KEY:";
            // 
            // txt_key
            // 
            txt_key.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txt_key.Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold);
            txt_key.Location = new Point(20, 240);
            txt_key.Margin = new Padding(4, 5, 4, 5);
            txt_key.MinimumSize = new Size(1, 16);
            txt_key.Name = "txt_key";
            txt_key.Padding = new Padding(5);
            txt_key.Radius = 10;
            txt_key.ShowText = false;
            txt_key.Size = new Size(460, 40);
            txt_key.TabIndex = 4;
            txt_key.TextAlignment = ContentAlignment.MiddleCenter;
            txt_key.Watermark = "";
            txt_key.WordWarp = false;
            // 
            // btn_active
            // 
            btn_active.Anchor = AnchorStyles.Top;
            btn_active.Cursor = Cursors.Hand;
            btn_active.Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold);
            btn_active.Location = new Point(173, 300);
            btn_active.MinimumSize = new Size(1, 1);
            btn_active.Name = "btn_active";
            btn_active.Radius = 20;
            btn_active.Size = new Size(135, 45);
            btn_active.TabIndex = 5;
            btn_active.Text = "ACTIVE";
            btn_active.TipsFont = new Font("Microsoft Sans Serif", 9F);
            btn_active.Click += btn_activate_Click;
            // 
            // uiLabel4
            // 
            uiLabel4.Dock = DockStyle.Bottom;
            uiLabel4.Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold);
            uiLabel4.ForeColor = Color.FromArgb(48, 48, 48);
            uiLabel4.Location = new Point(0, 348);
            uiLabel4.Name = "uiLabel4";
            uiLabel4.Size = new Size(500, 72);
            uiLabel4.TabIndex = 6;
            uiLabel4.Text = "Mọi thắc mắc xin liên hệ:\r\nZalo: 0355520331 - FS: 01525852";
            uiLabel4.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // btn_copy
            // 
            btn_copy.Anchor = AnchorStyles.None;
            btn_copy.Font = new Font("Microsoft Sans Serif", 12F);
            btn_copy.Location = new Point(400, 160);
            btn_copy.MinimumSize = new Size(1, 1);
            btn_copy.Name = "btn_copy";
            btn_copy.Size = new Size(80, 40);
            btn_copy.Symbol = 61637;
            btn_copy.TabIndex = 7;
            btn_copy.Text = "Copy";
            btn_copy.TipsFont = new Font("Microsoft Sans Serif", 9F);
            btn_copy.Click += txt_hwid_ButtonClick;
            // 
            // frmLogin
            // 
            AutoScaleMode = AutoScaleMode.None;
            BackColor = Color.FromArgb(244, 242, 251);
            ClientSize = new Size(500, 420);
            ControlBoxFillHoverColor = Color.FromArgb(133, 97, 198);
            Controls.Add(btn_copy);
            Controls.Add(uiLabel4);
            Controls.Add(btn_active);
            Controls.Add(txt_key);
            Controls.Add(uiLabel3);
            Controls.Add(txt_hwid);
            Controls.Add(uiLabel2);
            Controls.Add(uiLabel1);
            MaximizeBox = false;
            Name = "frmLogin";
            RectColor = Color.FromArgb(102, 58, 183);
            ShowIcon = false;
            ShowInTaskbar = false;
            Style = Sunny.UI.UIStyle.Custom;
            Text = "Xác thực bản quyền";
            ZoomScaleRect = new Rectangle(15, 15, 800, 450);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Sunny.UI.UILabel uiLabel1;
        private Sunny.UI.UILabel uiLabel2;
        private Sunny.UI.UITextBox txt_hwid;
        private Sunny.UI.UILabel uiLabel3;
        private Sunny.UI.UITextBox txt_key;
        private Sunny.UI.UISymbolButton btn_active;
        private Sunny.UI.UILabel uiLabel4;
        private Sunny.UI.UISymbolButton btn_copy;
    }
}
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace AutoJMS.CustomControl
{
    public class CustomButton : CheckBox
    {
        //Fields
        private Color onBackColor = Color.MediumSlateBlue;
        private Color onToggleColor = Color.WhiteSmoke;
        private Color offBackColor = Color.Gray;
        private Color offToggleColor = Color.Gainsboro;

        public Color OnBackColor { get => onBackColor; set { onBackColor = value; this.Invalidate(); } }
        public Color OnToggleColor { get => onToggleColor; set { onToggleColor = value; this.Invalidate(); } }
        public Color OffBackColor { get => offBackColor; set { offBackColor = value; this.Invalidate(); } }
        public Color OffToggleColor
        {
            get { return offToggleColor; }
            set { offToggleColor = value; this.Invalidate(); }
        }

        public override string Text
        {
            get { return base.Text; }
            set { }
        }

        protected override bool ShowFocusCues => false;

        //Constructor
        public CustomButton()
        {
            this.MinimumSize = new Size(45, 22);

            this.SetStyle(ControlStyles.UserPaint |
                          ControlStyles.AllPaintingInWmPaint |
                          ControlStyles.OptimizedDoubleBuffer |
                          ControlStyles.ResizeRedraw, true);

            this.AutoSize = false;
            this.Cursor = Cursors.Hand;

            this.FlatStyle = FlatStyle.Flat;
            this.FlatAppearance.BorderSize = 0;
        }

        //Methods
        private GraphicsPath GetFigurePath()
        {
            int arcSize = this.Height - 1;
            Rectangle leftArc = new Rectangle(0, 0, arcSize, arcSize);
            Rectangle rightArc = new Rectangle(this.Width - arcSize - 2, 0, arcSize, arcSize);

            GraphicsPath path = new GraphicsPath();
            path.StartFigure();
            path.AddArc(leftArc, 90, 180);
            path.AddArc(rightArc, 270, 180);
            path.CloseFigure();

            return path;
        }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            if (this.Parent != null)
            {
                if (this.Parent.BackgroundImage != null)
                {
                    ButtonRenderer.DrawParentBackground(pevent.Graphics, this.ClientRectangle, this);
                }
                else 
                {
                    pevent.Graphics.Clear(this.Parent.BackColor);
                }
            }
            else
            {
                pevent.Graphics.Clear(this.BackColor);
            }

            // Bật làm mượt
            pevent.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            int toggleSize = this.Height - 5;

            using (GraphicsPath path = GetFigurePath())
            {
                if (this.Checked) // ON
                {
                    
                    using (SolidBrush brush = new SolidBrush(onBackColor))
                    using (Pen pen = new Pen(onBackColor, 1.5f)) 
                    {
                        pevent.Graphics.FillPath(brush, path);
                        pevent.Graphics.DrawPath(pen, path); 
                    }

                    using (SolidBrush toggleBrush = new SolidBrush(onToggleColor))
                    {
                        pevent.Graphics.FillEllipse(toggleBrush,
                            new Rectangle(this.Width - this.Height + 1, 2, toggleSize, toggleSize));
                    }
                }
                else // OFF
                {
                    using (SolidBrush brush = new SolidBrush(offBackColor))
                    using (Pen pen = new Pen(offBackColor, 1.5f))
                    {
                        pevent.Graphics.FillPath(brush, path);
                        pevent.Graphics.DrawPath(pen, path);
                    }

                    using (SolidBrush toggleBrush = new SolidBrush(offToggleColor))
                    {
                        pevent.Graphics.FillEllipse(toggleBrush,
                            new Rectangle(2, 2, toggleSize, toggleSize));
                    }
                }
            }
        }
    }
}
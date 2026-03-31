using System;
using System.Drawing;
using System.Windows.Forms;

namespace AutoJMS
{
    public class FocusModeHelper
    {
        private readonly Control _leftContainer;
        private readonly Button _focusButton; 

        private bool _isCollapsed = false;
        private int _savedLeftWidth;

        public bool IsCollapsed => _isCollapsed;

        public Action OnFocusToggled;

        public FocusModeHelper(Control leftContainer, Button focusButton = null)
        {
            _leftContainer = leftContainer;
            _focusButton = focusButton;

            _savedLeftWidth = _leftContainer.Width;
            

            UpdateButtonUI();
        }

        public void Toggle()
        {
            if (_isCollapsed) ExpandAll();
            else CollapseAll();

            UpdateButtonUI();

            OnFocusToggled?.Invoke();
        }

        private void CollapseAll()
        {
            if (_isCollapsed) return;

            _savedLeftWidth = _leftContainer.Width;

            _leftContainer.Visible = false;
            _isCollapsed = true;
        }

        private void ExpandAll()
        {
            if (!_isCollapsed) return;

            _leftContainer.Visible = true;

            _leftContainer.Width = _savedLeftWidth;
            _isCollapsed = false;
        }

        private void UpdateButtonUI()
        {
            if (_focusButton != null)
            {
                _focusButton.Image = _isCollapsed ? Properties.Resources.resize : Properties.Resources.fullscreen;

            }
        }
    }
}
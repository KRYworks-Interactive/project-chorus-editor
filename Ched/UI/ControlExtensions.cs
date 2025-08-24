using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reactive.Linq;

namespace Ched.UI
{
    internal static class ControlExtensions
    {
        public static LayoutManager WorkWithLayout(this Control control)
        {
            return new LayoutManager(control);
        }

        public static void InvokeIfRequired(this Control control, Action action)
        {
            if (control.InvokeRequired) control.Invoke((MethodInvoker)(() => action()));
            else action();
        }

        public static IObservable<MouseEventArgs> MouseDownAsObservable(this Control control)
        {
            return Observable.FromEvent<MouseEventHandler, MouseEventArgs>(
                     h => (o, e) => h(e),
                     h => control.MouseDown += h,
                     h => control.MouseDown -= h);
        }

        public static IObservable<MouseEventArgs> MouseMoveAsObservable(this Control control)
        {
            return Observable.FromEvent<MouseEventHandler, MouseEventArgs>(
                     h => (o, e) => h(e),
                     h => control.MouseMove += h,
                     h => control.MouseMove -= h);
        }

        public static IObservable<MouseEventArgs> MouseUpAsObservable(this Control control)
        {
            return Observable.FromEvent<MouseEventHandler, MouseEventArgs>(
                     h => (o, e) => h(e),
                     h => control.MouseUp += h,
                     h => control.MouseUp -= h);
        }

        public static int GetMaximumValue(this ScrollBar scrollbar)
        {
            return scrollbar.Maximum - scrollbar.LargeChange + 1;
        }

        public static void SelectAll(this NumericUpDown control)
        {
            control.Select(0, control.Text.Length);
        }
    }

    internal class LayoutManager : IDisposable
    {
        protected Control _control;

        public LayoutManager(Control control)
        {
            control.SuspendLayout();
            _control = control;
        }

        public void Dispose()
        {
            _control.ResumeLayout(false);
            _control.PerformLayout();
        }
    }

    public class DarkDialogForm : Form
    {
        protected readonly Color DialogBackColor = Color.FromArgb(45, 45, 48);
        protected readonly Color ControlBackColor = Color.FromArgb(63, 63, 70);
        protected readonly Color ControlForeColor = Color.White;
        protected readonly Color BorderColor = Color.FromArgb(80, 80, 80);
        protected readonly Color HighlightColor = Color.FromArgb(0, 120, 215);

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.BackColor = DialogBackColor;
            this.ForeColor = ControlForeColor;
            ApplyThemeAndBehavior(this.Controls);
        }

        private void ApplyThemeAndBehavior(Control.ControlCollection controls)
        {
            foreach (Control control in controls)
            {
                if (control is Button button)
                {
                    button.FlatStyle = FlatStyle.Flat;
                    button.FlatAppearance.BorderColor = BorderColor;
                    button.BackColor = ControlBackColor;
                    button.ForeColor = ControlForeColor;

                    if (button.Name == "buttonOK") button.DialogResult = DialogResult.OK;
                    if (button.Name == "buttonCancel") button.DialogResult = DialogResult.Cancel;
                }
                else if (control is NumericUpDown numericUpDown)
                {
                    numericUpDown.BackColor = ControlBackColor;
                    numericUpDown.ForeColor = ControlForeColor;
                }
                else if (control is ComboBox comboBox)
                {
                    comboBox.FlatStyle = FlatStyle.Flat;
                    comboBox.BackColor = ControlBackColor;
                    comboBox.ForeColor = ControlForeColor;
                    if (comboBox.DropDownStyle == ComboBoxStyle.DropDownList)
                    {
                        comboBox.DrawMode = DrawMode.OwnerDrawFixed;
                        comboBox.DrawItem += ComboBox_DrawItem;
                        comboBox.DropDown += ComboBox_DropDown;
                    }
                }

                if (control.HasChildren)
                {
                    ApplyThemeAndBehavior(control.Controls);
                }
            }
        }

        private void ComboBox_DropDown(object sender, EventArgs e)
        {
            var comboBox = (ComboBox)sender;
            int maxWidth = 0;
            using (Graphics g = comboBox.CreateGraphics())
            {
                foreach (var item in comboBox.Items)
                {
                    int width = (int)g.MeasureString(comboBox.GetItemText(item), comboBox.Font).Width;
                    if (width > maxWidth)
                    {
                        maxWidth = width;
                    }
                }
            }

            if (comboBox.Items.Count > comboBox.MaxDropDownItems)
            {
                maxWidth += SystemInformation.VerticalScrollBarWidth;
            }

            comboBox.DropDownWidth = Math.Max(comboBox.Width, maxWidth + 4);
        }

        private void ComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            var comboBox = (ComboBox)sender;

            Color backColor = (e.State & DrawItemState.Selected) == DrawItemState.Selected ? HighlightColor : this.ControlBackColor;

            using (var b = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(b, e.Bounds);
            }

            TextRenderer.DrawText(e.Graphics, comboBox.GetItemText(comboBox.Items[e.Index]), e.Font, e.Bounds, this.ControlForeColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);

            e.DrawFocusRectangle();
        }
    }

    public class DarkColorTable : ProfessionalColorTable
    {
        private readonly Color _backColor = Color.FromArgb(45, 45, 48);
        private readonly Color _borderColor = Color.FromArgb(67, 67, 70);
        private readonly Color _menuItemSelectedColor = Color.FromArgb(51, 51, 55);
        private readonly Color _buttonCheckedColor = Color.FromArgb(85, 85, 85);
        private readonly Color _buttonSelectedColor = Color.FromArgb(81, 81, 81);
        private readonly Color _buttonPressedColor = Color.FromArgb(40, 40, 40);

        public override Color MenuBorder => _borderColor;
        public override Color MenuItemBorder => Color.Transparent;
        public override Color MenuItemSelected => _menuItemSelectedColor;
        public override Color MenuStripGradientBegin => _backColor;
        public override Color MenuStripGradientEnd => _backColor;
        public override Color ToolStripGradientBegin => _backColor;
        public override Color ToolStripGradientMiddle => _backColor;
        public override Color ToolStripGradientEnd => _backColor;
        public override Color MenuItemSelectedGradientBegin => _menuItemSelectedColor;
        public override Color MenuItemSelectedGradientEnd => _menuItemSelectedColor;
        public override Color MenuItemPressedGradientBegin => _menuItemSelectedColor;
        public override Color MenuItemPressedGradientEnd => _menuItemSelectedColor;
        public override Color ToolStripDropDownBackground => _backColor;
        public override Color ImageMarginGradientBegin => _backColor;
        public override Color ImageMarginGradientMiddle => _backColor;
        public override Color ImageMarginGradientEnd => _backColor;
        public override Color ToolStripBorder => _borderColor;
        public override Color ButtonSelectedHighlight => _buttonSelectedColor;
        public override Color ButtonSelectedGradientBegin => _buttonSelectedColor;
        public override Color ButtonSelectedGradientEnd => _buttonSelectedColor;
        public override Color ButtonSelectedBorder => _borderColor;
        public override Color ButtonCheckedGradientBegin => _buttonCheckedColor;
        public override Color ButtonCheckedGradientEnd => _buttonCheckedColor;
        public override Color ButtonCheckedHighlight => _buttonSelectedColor;
        public override Color ButtonPressedGradientBegin => _buttonPressedColor;
        public override Color ButtonPressedGradientEnd => _buttonPressedColor;
        public override Color ButtonPressedBorder => _borderColor;
        public override Color GripDark => _borderColor;
        public override Color GripLight => _borderColor;
        public override Color ToolStripContentPanelGradientBegin => _backColor;
        public override Color ToolStripContentPanelGradientEnd => _backColor;
        public override Color ToolStripPanelGradientBegin => _backColor;
        public override Color ToolStripPanelGradientEnd => _backColor;
        public override Color OverflowButtonGradientBegin => _backColor;
        public override Color OverflowButtonGradientEnd => _backColor;
        public override Color OverflowButtonGradientMiddle => _backColor;
        public override Color SeparatorDark => _borderColor;
        public override Color SeparatorLight => _borderColor;
        public override Color StatusStripGradientBegin => _backColor;
        public override Color StatusStripGradientEnd => _backColor;
    }

    public class DarkScrollBar : Control
    {
        // Properties
        private int minimum = 0;
        public int Minimum
        {
            get => minimum;
            set { minimum = value; Invalidate(); }
        }

        private int maximum = 100;
        public int Maximum
        {
            get => maximum;
            set { maximum = value; Invalidate(); }
        }

        private int value;
        public int Value
        {
            get => this.value;
            set
            {
                int newValue = Math.Max(Minimum, Math.Min(GetMaximumValue(), value));
                if (this.value == newValue) return;
                this.value = newValue;
                ValueChanged?.Invoke(this, EventArgs.Empty);
                Invalidate();
            }
        }

        private int largeChange = 10;
        public int LargeChange
        {
            get => largeChange;
            set { largeChange = value; Invalidate(); }
        }

        private int smallChange = 1;
        public int SmallChange
        {
            get => smallChange;
            set { smallChange = value; Invalidate(); }
        }

        // Events
        public event EventHandler ValueChanged;
        public event EventHandler<ScrollEventArgs> Scroll;

        // Theming
        private readonly Color _backColor = Color.FromArgb(45, 45, 48);
        private readonly Color _thumbColor = Color.FromArgb(104, 104, 104);
        private readonly Color _arrowColor = Color.FromArgb(160, 160, 160);
        private readonly int _arrowHeight = 18;

        private bool isDragging = false;
        private int dragOffset;
        private Timer arrowTimer = new Timer() { Interval = 100 };
        private ScrollEventType currentScrollType;

        public DarkScrollBar()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
            this.Width = 17; // Standard scrollbar width
            arrowTimer.Tick += (s, e) =>
            {
                if (currentScrollType == ScrollEventType.SmallDecrement)
                    SetValue(this.Value - SmallChange, ScrollEventType.SmallDecrement);
                else if (currentScrollType == ScrollEventType.SmallIncrement)
                    SetValue(this.Value + SmallChange, ScrollEventType.SmallIncrement);
            };
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.Clear(_backColor);

            DrawArrow(e.Graphics, new Rectangle(0, 0, Width, _arrowHeight), true); // Up arrow
            DrawArrow(e.Graphics, new Rectangle(0, Height - _arrowHeight, Width, _arrowHeight), false); // Down arrow

            Rectangle thumbRect = GetThumbRectangle();
            using (var b = new SolidBrush(_thumbColor))
            {
                e.Graphics.FillRectangle(b, thumbRect);
            }
        }

        private void DrawArrow(Graphics g, Rectangle bounds, bool isUp)
        {
            using (var b = new SolidBrush(_arrowColor))
            {
                PointF[] arrow = new PointF[3];
                if (isUp)
                {
                    arrow[0] = new PointF(bounds.Width / 2f, bounds.Top + 5);
                    arrow[1] = new PointF(bounds.Width / 2f - 4, bounds.Top + 11);
                    arrow[2] = new PointF(bounds.Width / 2f + 4, bounds.Top + 11);
                }
                else // Down
                {
                    arrow[0] = new PointF(bounds.Width / 2f, bounds.Bottom - 5);
                    arrow[1] = new PointF(bounds.Width / 2f - 4, bounds.Bottom - 11);
                    arrow[2] = new PointF(bounds.Width / 2f + 4, bounds.Bottom - 11);
                }
                g.FillPolygon(b, arrow);
            }
        }

        private Rectangle GetThumbRectangle()
        {
            float trackHeight = Height - (2 * _arrowHeight);
            if (trackHeight <= 0) return Rectangle.Empty;

            float valueRange = (Maximum - Minimum);
            if (valueRange <= 0 || LargeChange <= 0) return new Rectangle(1, _arrowHeight, Width - 2, (int)trackHeight);

            float thumbHeight = Math.Max(10, trackHeight * LargeChange / (valueRange + LargeChange));
            if (thumbHeight >= trackHeight) return new Rectangle(1, _arrowHeight, Width - 2, (int)trackHeight);

            float scrollableRange = GetMaximumValue() - Minimum;
            if (scrollableRange <= 0) return new Rectangle(1, _arrowHeight, Width - 2, (int)thumbHeight);

            float thumbY = _arrowHeight + ((trackHeight - thumbHeight) * (this.Value - Minimum) / scrollableRange);

            return new Rectangle(1, (int)thumbY, Width - 2, (int)thumbHeight);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Rectangle thumbRect = GetThumbRectangle();
            if (thumbRect.Contains(e.Location))
            {
                isDragging = true;
                dragOffset = e.Y - thumbRect.Y;
                currentScrollType = ScrollEventType.ThumbTrack;
                Scroll?.Invoke(this, new ScrollEventArgs(currentScrollType, this.Value));
            }
            else if (new Rectangle(0, 0, Width, _arrowHeight).Contains(e.Location))
            {
                currentScrollType = ScrollEventType.SmallDecrement;
                SetValue(this.Value - SmallChange, currentScrollType);
                arrowTimer.Start();
            }
            else if (new Rectangle(0, Height - _arrowHeight, Width, _arrowHeight).Contains(e.Location))
            {
                currentScrollType = ScrollEventType.SmallIncrement;
                SetValue(this.Value + SmallChange, currentScrollType);
                arrowTimer.Start();
            }
            else
            {
                currentScrollType = e.Y < thumbRect.Y ? ScrollEventType.LargeDecrement : ScrollEventType.LargeIncrement;
                SetValue(this.Value + (e.Y < thumbRect.Y ? -LargeChange : LargeChange), currentScrollType);
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (!isDragging) return;

            float trackHeight = Height - (2 * _arrowHeight);
            if (trackHeight <= 0) return;

            Rectangle thumbRect = GetThumbRectangle();
            float thumbHeight = thumbRect.Height;
            if (thumbHeight >= trackHeight) return;

            float scrollableTrackHeight = trackHeight - thumbHeight;
            float scrollableRange = GetMaximumValue() - Minimum;
            if (scrollableRange <= 0) return;

            float newThumbY = e.Y - dragOffset;
            float relativeY = newThumbY - _arrowHeight;

            int newValue = (int)(Minimum + (scrollableRange * relativeY / scrollableTrackHeight));
            SetValue(newValue, ScrollEventType.ThumbTrack);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            arrowTimer.Stop();
            isDragging = false;
            Scroll?.Invoke(this, new ScrollEventArgs(ScrollEventType.EndScroll, this.Value));
        }

        private void SetValue(int newValue, ScrollEventType type)
        {
            int clampedValue = Math.Max(Minimum, Math.Min(GetMaximumValue(), newValue));
            if (this.Value == clampedValue) return;

            this.Value = clampedValue;
            Scroll?.Invoke(this, new ScrollEventArgs(type, this.Value));
        }

        public int GetMaximumValue()
        {
            if (Maximum - LargeChange < Minimum) return Minimum;
            return Maximum - LargeChange + 1;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                arrowTimer.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}

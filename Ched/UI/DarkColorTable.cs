using System.Drawing;
using System.Windows.Forms;

namespace Ched.UI
{
    public class DarkColorTable : ProfessionalColorTable
    {
        private readonly Color _backColor = Color.FromArgb(45, 45, 48);
        private readonly Color _borderColor = Color.FromArgb(67, 67, 70);
        private readonly Color _menuItemSelectedColor = Color.FromArgb(51, 51, 55);
        private readonly Color _buttonCheckedColor = Color.FromArgb(0, 122, 204);

        public override Color MenuBorder => _borderColor;
        public override Color MenuItemBorder => Color.Transparent;
        public override Color MenuItemSelected => _menuItemSelectedColor;
        public override Color MenuStripGradientBegin => _backColor;
        public override Color MenuStripGradientEnd => _backColor;
        public override Color MenuItemSelectedGradientBegin => _menuItemSelectedColor;
        public override Color MenuItemSelectedGradientEnd => _menuItemSelectedColor;
        public override Color MenuItemPressedGradientBegin => _menuItemSelectedColor;
        public override Color MenuItemPressedGradientEnd => _menuItemSelectedColor;
        public override Color ToolStripDropDownBackground => _backColor;
        public override Color ImageMarginGradientBegin => _backColor;
        public override Color ImageMarginGradientMiddle => _backColor;
        public override Color ImageMarginGradientEnd => _backColor;
        public override Color ToolStripBorder => _borderColor;
        public override Color ButtonSelectedHighlight => _menuItemSelectedColor;
        public override Color ButtonSelectedGradientBegin => _menuItemSelectedColor;
        public override Color ButtonSelectedGradientEnd => _menuItemSelectedColor;
        public override Color ButtonSelectedBorder => _borderColor;
        public override Color ButtonCheckedGradientBegin => _buttonCheckedColor;
        public override Color ButtonCheckedGradientEnd => _buttonCheckedColor;
        public override Color ButtonCheckedHighlight => _buttonCheckedColor;
        public override Color ButtonPressedGradientBegin => _menuItemSelectedColor;
        public override Color ButtonPressedGradientEnd => _menuItemSelectedColor;
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
}

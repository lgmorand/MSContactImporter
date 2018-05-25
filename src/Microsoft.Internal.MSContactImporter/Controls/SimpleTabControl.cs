using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace Microsoft.Internal.MSContactImporter.Controls
{
    /// <summary>
    /// Summary description for SimpleTabControl.
    /// </summary>
	public class SimpleTabControl : TabControl
    {
        private bool simpleMode = false;

        [DefaultValue(false)]
        public bool SimpleMode
        {
            get { return simpleMode; }
            set
            {
                simpleMode = value;
                RecreateHandle();
            }
        }

        private bool simpleModeInDesign = false;

        [DefaultValue(false)]
        public bool SimpleModeInDesign
        {
            get { return simpleModeInDesign; }
            set
            {
                simpleModeInDesign = value;
                RecreateHandle();
            }
        }

        public override Rectangle DisplayRectangle
        {
            get
            {
                if ((simpleMode) && (!DesignMode || simpleModeInDesign))
                {
                    return new Rectangle(0, 0, Width, Height);
                }
                else
                    return base.DisplayRectangle;
            }
        }
    }
}
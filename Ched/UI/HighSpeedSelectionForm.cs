﻿﻿﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ched.UI
{
    public partial class HighSpeedSelectionForm : DarkDialogForm
    {
        public decimal SpeedRatio
        {
            get { return speedRatioBox.Value; }
            set
            {
                speedRatioBox.Value = value;
                speedRatioBox.SelectAll();
            }
        }

        public HighSpeedSelectionForm()
        {
            InitializeComponent();
            AcceptButton = buttonOK;
            CancelButton = buttonCancel;

            speedRatioBox.Minimum = 0.01m;
            speedRatioBox.Maximum = 10000m;
            speedRatioBox.Increment = 0.01m;
            speedRatioBox.DecimalPlaces = 2;
            speedRatioBox.Value = 1;
        }
    }
}

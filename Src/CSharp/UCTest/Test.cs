using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace UCTest
{
    public partial class Test : ActiveXControl
    {
        public Test()
        {
            InitializeComponent();
        }

        public string GetDateTime()
        {
            InitializeComponent();
            return DateTime.Now.ToString();
        }
    }
}

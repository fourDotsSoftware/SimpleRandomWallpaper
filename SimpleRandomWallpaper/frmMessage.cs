using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SimpleRandomWallpaper
{
    public partial class frmMessage : CustomForm
    {
        public frmMessage()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.ShowMessage = !chkShow.Checked;

            this.DialogResult = DialogResult.OK;

            try
            {
                this.Close();
            }
            catch { }
        }

        private void frmMessage_Load(object sender, EventArgs e)
        {

        }
    }
}


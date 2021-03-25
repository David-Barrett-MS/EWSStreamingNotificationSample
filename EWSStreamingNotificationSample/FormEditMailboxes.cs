/*
 * By David Barrett, Microsoft Ltd. 2013-2021. Use at your own risk.  No warranties are given.
 * 
 * DISCLAIMER:
 * THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
 * MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
 * A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
 * MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
 * BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
 * SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
 * OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
 * */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace EWSStreamingNotificationSample
{
    public partial class FormEditMailboxes : Form
    {
        bool _cancel = false;
        public FormEditMailboxes()
        {
            InitializeComponent();
        }

        public string EditMailboxes(string mailboxes)
        {
            // Allow the editing of the list of mailboxes
            textBoxMailboxes.Text = mailboxes;
            _cancel = false;
            this.ShowDialog();
            if (!_cancel)
                return textBoxMailboxes.Text;
            return mailboxes;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            _cancel = true;
            this.Hide();
        }
    }
}

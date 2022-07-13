/*
 * By David Barrett, Microsoft Ltd. 2013-2022. Use at your own risk.  No warranties are given.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 * */

using System;
using System.Windows.Forms;

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

using System;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace XboxDownload
{
    public partial class FormStore : Form
    {
        public FormStore()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            string email = tbEMail.Text.Trim();
            if (string.IsNullOrEmpty(email))
            {
                MessageBox.Show("邮箱地址不能空。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbEMail.Focus();
                return;
            }
            if (!Regex.IsMatch(email, @"^[a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$")) //^[A-Za-z0-9\u4e00-\u9fa5]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$
            {
                MessageBox.Show("邮箱地址不正确。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tbEMail.Focus();
                return;

            }
            button1.Text = "请稍候...";
            tbEMail.Enabled = button1.Enabled = false;
            ThreadPool.QueueUserWorkItem(delegate { AddWhitelist(email); });
        }

        private void AddWhitelist(string email)
        {
            string url = "https://" + ClassWeb.HostToIP() + ":45678/Store/Whitelist/" + ClassWeb.UrlEncode(email);
            SocketPackage socketPackage = ClassWeb.HttpRequest(url, "GET", null, null, true, false, true, null, null, null, null, null, null, null, null, 0, null);
            if (socketPackage.Html == "True")
            {
                Application.OpenForms[0].Invoke(new MethodInvoker(() => {
                    MessageBox.Show("已成功加入白名单。", "加入白名单", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    button1.Text = "成功";
                }));
            }
            else
            {
                string err = socketPackage.Headers.StartsWith("HTTP/1.1 200 OK") ? "加入白名单失败 (" + socketPackage.Html + ")" : "加入白名单失败，请稍候再试。";
                Application.OpenForms[0].Invoke(new MethodInvoker(() => {
                    MessageBox.Show(err, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    button1.Text = "失败";
                }));
            }
        }
    }
}

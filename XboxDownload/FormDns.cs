using System;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace XboxDownload
{
    public partial class FormDns : Form
    {
        public FormDns()
        {
            InitializeComponent();
        }

        Thread thread = null;
        private void ButTest_Click(object sender, EventArgs e)
        {
            string domainName = cbDomainName.Text;
            if (!string.IsNullOrEmpty(domainName))
            {
                butTest.Enabled = false;
                textBox1.Text = ">nslookup " + domainName + " " + Properties.Settings.Default.LocalIP + "\r\n";
                thread = new Thread(new ThreadStart(() =>
                {
                    Test(domainName);
                }))
                {
                    IsBackground = true
                };
                thread.Start();
            }
        }

        private void FormDns_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (thread != null && thread.IsAlive) thread.Abort();
        }

        private void Test(string domainName)
        {
            string resultInfo = string.Empty;
            using (Process p = new Process())
            {
                p.StartInfo = new ProcessStartInfo("nslookup", domainName + " " + Properties.Settings.Default.LocalIP)
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    RedirectStandardOutput = true
                };
                p.Start();
                resultInfo = p.StandardOutput.ReadToEnd();
                p.Close();
            }
            SetMsg(resultInfo);
            MatchCollection mc = Regex.Matches(resultInfo, @":\s*(?<ip>\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})");
            if (mc.Count == 2)
            {
                string ip = mc[1].Groups["ip"].Value;
                SetMsg("\n//IP地址: " + ip + " " + QueryLocation(ip));
            }
            else if(Regex.IsMatch(resultInfo, @"timeout", RegexOptions.IgnoreCase))
            {
                SetMsg("*** 请求超时");
            }
            else
            {
                SetMsg("*** 找不到 " + domainName + ": Non-existent domain");
            }
            SetButEnable(true);
        }

        private string QueryLocation(string ip)
        {
            SocketPackage socketPackage = ClassWeb.HttpRequest("https://www.ip138.com/iplookup.asp?ip=" + ip + "&action=2", "GET", null, null, true, false, true, "GBK", null, null, ClassWeb.useragent, null, null, null, null, 0, null);
            Match result = Regex.Match(socketPackage.Html, @"""ASN归属地"":""(?<location>[^""]+)""");
            if (result.Success)
            {
                return result.Groups["location"].Value.Trim() + " (数据来源：ip138)";
            }
            else
            {
                socketPackage = ClassWeb.HttpRequest("https://www.baidu.com/s?wd=" + ip, "GET", null, null, true, false, true, "utf-8", null, null, ClassWeb.useragent, null, null, null, null, 0, null);
                result = Regex.Match(socketPackage.Html, @"<span [^>]+>IP地址:[^<]+</span>(?<location>.+)");
                if (result.Success)
                {
                    return result.Groups["location"].Value.Trim() + " (数据来源：百度)";
                }
            }
            return "";
        }

        delegate void CallbackButEnable(bool enabled);
        private void SetButEnable(bool enabled)
        {
            if (butTest.InvokeRequired)
            {
                CallbackButEnable d = new CallbackButEnable(SetButEnable);
                this.Invoke(d, new object[] { enabled });
            }
            else
            {
                butTest.Enabled = enabled;
            }
        }

        delegate void CallbackMsg(string str);
        private void SetMsg(string str)
        {
            if (textBox1.InvokeRequired)
            {
                CallbackMsg d = new CallbackMsg(SetMsg);
                Invoke(d, new object[] { str });
            }
            else textBox1.AppendText(str);
        }
    }
}

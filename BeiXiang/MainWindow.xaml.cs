using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace BeiXiang
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        DispatcherTimer timer = new DispatcherTimer();

        public MainWindow()
        {
            InitializeComponent();

            SuppressScriptErrors(web, true);
            web.Navigate("http://data.eastmoney.com/hsgtcg/default.html");

            web.Navigated += (o, ex) => {
                Task.Run(() => {
                    Thread.Sleep(4000);
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        compute();
                    }));
                });
            };

            timer.Interval = new TimeSpan(0, 0, 6);
            timer.Tick += (o, ex) => {
                SuppressScriptErrors(web, true);
                web.Navigate($"http://data.eastmoney.com/hsgtcg/default.html?_={DateTime.Now.Ticks}");
            };
            timer.Start();
        }

        /// <summary>
        /// 在加载页面之前调用此方法设置hide为true就能抑制错误的弹出了。
        /// </summary>
        /// <param name="webBrowser"></param>
        /// <param name="hide"></param>
        static void SuppressScriptErrors(WebBrowser webBrowser, bool hide)
        {
            webBrowser.Navigating += (s, e) =>
            {
                var fiComWebBrowser = typeof(WebBrowser).GetField("_axIWebBrowser2", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                if (fiComWebBrowser == null)
                    return;

                object objComWebBrowser = fiComWebBrowser.GetValue(webBrowser);
                if (objComWebBrowser == null)
                    return;

                objComWebBrowser.GetType().InvokeMember("Silent", System.Reflection.BindingFlags.SetProperty, null, objComWebBrowser, new object[] { hide });
            };
        }
        private void compute()
        {
            var doc = web.Document as mshtml.HTMLDocument;
            if (doc == null)
            {
                return;
            }

            IHTMLElement hgtElement = doc.getElementById("zjlx_hgt");
            IHTMLElement sgtElement = doc.getElementById("zjlx_sgt");
            IHTMLElement _hgt = hgtElement.children[4].children[0];
            txtHGT.Content = _hgt.innerText.Replace("亿元", string.Empty);

            IHTMLElement _sgt = sgtElement.children[4].children[0];
            txtSGT.Content = _sgt.innerText.Replace("亿元", string.Empty);

            txtTotal.Content = Math.Round(Convert.ToSingle(txtHGT.Content) + Convert.ToSingle(txtSGT.Content), 2);
        }

        public static void DoEvents()
        {
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Background,
              new Action(delegate { }));
        }
    }
}

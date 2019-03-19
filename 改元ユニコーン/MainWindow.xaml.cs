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

namespace 改元ユニコーン
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private CancellationTokenSource tokenSource = null;
        private CancellationToken token;

		private static int winnerTime = 41;

		public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 表示されたとき
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void Window_ContentRendered(object sender, EventArgs e)
        {
            //スレッド開始
            tokenSource = null;
            try
            {
                tokenSource = new CancellationTokenSource();
                token = tokenSource.Token;
                await Task.Run(new Action(MainLoop), token);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (tokenSource != null)
                {
                    tokenSource.Dispose();
                    //別メソッドでtokenSourceがnullかどうか参照しているので、nullの代入をする。
                    tokenSource = null;

                    this.Close();
                }
            }
        }

        /// <summary>
        /// メインループ
        /// 現在時刻を監視して、指定時刻になったら、ユニコーンを再生する。
        /// </summary>
        private void MainLoop()
        {
            while (!token.IsCancellationRequested)
            {
				DateTime kaigenTime = detePicker1.SelectedDate.Value;
				TimeSpan timeSpan = new TimeSpan(0, 0, winnerTime);
				DateTime startUnicornTime = kaigenTime - timeSpan;
			}
        }

        /// <summary>
        /// メインループにキャンセル要求をします。
        /// </summary>
        private void CancelMainLoop()
        {
            //スレッド終了
            if (tokenSource != null)
            {
                tokenSource.Cancel();
            }
        }

        /// <summary>
        /// ウィンドウが閉じようとしているとき
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            CancelMainLoop();
            if (tokenSource != null)
            {
                e.Cancel = true;
            }
        }
    }
}

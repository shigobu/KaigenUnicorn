using CoreAudioApi;
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
        //キャンセルトークン
        private CancellationTokenSource tokenSource = null;
        private CancellationToken token;

        //ユニコーンが完全勝利するまでの秒数
		private static int winnerTime = 41;

        //フラグ
        /// <summary>
        /// Unicorn再生済フラグ
        /// </summary>
        private bool IsStartUnicorn { get; set; } = false;

        /// <summary>
        /// Unicorn読み込み済フラグ
        /// </summary>
        private bool IsLoadUnicorn { get; set; } = false;

        /// <summary>
        /// フェードアウト開始済フラグ
        /// </summary>
        private bool IsFadeoutStart { get; set; } = false;


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
                //改元日取得
				DateTime kaigenTime = GetKaigenTime();
                //Unicorn再生開始時間
				DateTime startUnicornTime = kaigenTime - new TimeSpan(0, 0, winnerTime);
                //Unicornロード開始時間
                DateTime loadUnicornTime = startUnicornTime - new TimeSpan(0, 0, 10);
                //フェードアウト開始時間
                DateTime fadeoutStartTime = startUnicornTime - new TimeSpan(0, 0, 3);
                //現在時刻取得
                DateTime NowTime = DateTime.Now;

                if (NowTime >= startUnicornTime)
                {
                    //ユニコーン再生開始
                    StartUnicorn();
                }
                if (NowTime >= loadUnicornTime)
                {
                    //ユニコーン読み込み開始
                    LoadUnicorn();
                }
                if (NowTime >= fadeoutStartTime)
                {
                    //フェードアウト開始
                    FadeOutStart();
                }
            }
        }

        /// <summary>
        /// Unicornの再生を開始します。
        /// </summary>
        private void StartUnicorn()
        {
            //一回のみ実行させる
            if (IsStartUnicorn)
            {
                return;
            }

            IsStartUnicorn = true;
        }

        /// <summary>
        /// Unicornの読み込みを開始します。
        /// </summary>
        private void LoadUnicorn()
        {
            //一回のみ実行させる
            if (IsLoadUnicorn)
            {
                return;
            }

            IsLoadUnicorn = true;
        }

        /// <summary>
        /// フェードアウトを開始します。
        /// </summary>
        private void FadeOutStart()
        {
            //一回のみ実行させる
            if (IsFadeoutStart)
            {
                return;
            }
            //フェードアウトタスク実行
            Task task = Task.Run(new Action(FadeOutThread));

            IsFadeoutStart = true;
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
        /// フェードアウトを実行するタスク
        /// </summary>
        private void FadeOutThread()
        {
            //自分自身のプロセスを取得する
            System.Diagnostics.Process p = System.Diagnostics.Process.GetCurrentProcess();
            int pid = p.Id;

            MMDevice device = null;
            try
            {
                using (MMDeviceEnumerator DevEnum = new MMDeviceEnumerator())
                {
                    device = DevEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
                }
                float MasterVolumeLevel = device.AudioEndpointVolume.MasterVolumeLevelScalar;
                AudioSessionManager sessionManager = device.AudioSessionManager;
                for (float i = 100; i >= 0; i -= 1f)
                {
                    foreach (var item in sessionManager.Sessions)
                    {
                        if (item.ProcessID != (uint)pid)
                        {
                            item.SimpleAudioVolume.MasterVolume = (i / 100.0f);
                        }
                    }
                    Thread.Sleep(50);
                }
            }
            finally
            {
                if (device != null)
                {
                    device.Dispose();
                }
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

        /// <summary>
        /// 入力されている改元日時を取得します。
        /// </summary>
        /// <returns></returns>
        private DateTime GetKaigenTime()
        {
            if (datePicker1.Dispatcher.CheckAccess())
            {
                return datePicker1.SelectedDate.Value;
            }
            else
            {
                return datePicker1.Dispatcher.Invoke<DateTime>(new Func<DateTime>(GetKaigenTime));
            }
        }
    }
}

using CoreAudioApi;
using System;
using System.Media;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace 改元ユニコーン
{
	/// <summary>
	/// MainWindow.xaml の相互作用ロジック
	/// </summary>
	public partial class MainWindow : Window
    {
		//サウンドを再生するWin32 APIの宣言
		[Flags]
		public enum PlaySoundFlags : int
		{
			SND_SYNC = 0x0000,
			SND_ASYNC = 0x0001,
			SND_NODEFAULT = 0x0002,
			SND_MEMORY = 0x0004,
			SND_LOOP = 0x0008,
			SND_NOSTOP = 0x0010,
			SND_NOWAIT = 0x00002000,
			SND_ALIAS = 0x00010000,
			SND_ALIAS_ID = 0x00110000,
			SND_FILENAME = 0x00020000,
			SND_RESOURCE = 0x00040004,
			SND_PURGE = 0x0040,
			SND_APPLICATION = 0x0080
		}
		[System.Runtime.InteropServices.DllImport("winmm.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
		private static extern bool PlaySound(string pszSound, IntPtr hmod, PlaySoundFlags fdwSound);

		//キャンセルトークン
		private CancellationTokenSource tokenSource = null;
        private CancellationToken token;

        //ユニコーンが完全勝利するまでの秒数
		private static int winnerTime = 41;

        //フェードアウトタスク
        private Task fadeOutTask = null;

        //Unicorn再生関連
        private static string unicornWavePath = @"D:\茂信\Music\青と白の逆方向.wav";
        SoundPlayer soundPlayer = null;

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
				datePicker1.SelectedDate = new DateTime(2019, 3, 21, 15, 30, 0);
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
                //フェードアウト開始時間
                DateTime fadeoutStartTime = startUnicornTime - new TimeSpan(0, 0, 8);
                //現在時刻取得
                DateTime NowTime = DateTime.Now;

                if (NowTime >= startUnicornTime)
                {
                    //ユニコーン再生開始
                    StartUnicorn();
                }
                if (NowTime >= fadeoutStartTime)
                {
                    //フェードアウト開始
                    FadeOutStart();
                }
				Thread.Sleep(10);
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
			PlaySound(unicornWavePath, IntPtr.Zero, PlaySoundFlags.SND_FILENAME | PlaySoundFlags.SND_ASYNC);

			IsStartUnicorn = true;
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
            fadeOutTask = Task.Run(new Action(FadeOutThread));

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
                    Thread.Sleep(100);
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
        private async void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //メインループキャンセルリクエスト
            CancelMainLoop();
            if (tokenSource != null)
            {
                e.Cancel = true;
                return;
            }

            //フェードアウトタスク終了待機
            if (fadeOutTask != null)
            {
                await fadeOutTask;
                fadeOutTask.Dispose();
                fadeOutTask = null;
            }

			//再生しているWAVを停止する
			PlaySound(null, IntPtr.Zero, PlaySoundFlags.SND_PURGE);
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

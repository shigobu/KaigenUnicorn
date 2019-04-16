using NAudio.CoreAudioApi;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;
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
        //Audio再生オブジェクト
        private WaveOutEvent outputDevice;
        private AudioFileReader audioFile;

        //キャンセルトークン
        private CancellationTokenSource tokenSource = null;
        private CancellationToken token;

        //ユニコーンが完全勝利するまでの秒数
		private static int winnerTime = 41;

        //フェードアウトタスク
        private Task fadeOutTask = null;

        //Unicorn再生関連
        private static string unicornWavePath = @"D:\nc112577.mp3";

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
#if DEBUG
				datePicker1.SelectedDate = new DateTime(2019, 4, 12, 21, 35, 0);
#endif
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
                //Unicorn読み込み開始時間
                DateTime loadUnicornTime = startUnicornTime - new TimeSpan(0, 1, 0);
                //フェードアウト開始時間
                DateTime fadeoutStartTime = startUnicornTime - new TimeSpan(0, 0, 7);
                //現在時刻取得
                DateTime NowTime = DateTime.Now;

                if (NowTime >= loadUnicornTime)
                {
                    //読み込み開始
                    LoadUnicorn();
                }
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
        /// オーディオファイルを読み込みます。
        /// </summary>
        private void LoadUnicorn()
        {
            //一回のみ実行させる
            if (IsLoadUnicorn)
            {
                return;
            }

            if (outputDevice == null)
            {
                outputDevice = new WaveOutEvent();
                outputDevice.PlaybackStopped += OnPlaybackStopped;
            }
            if (audioFile == null)
            {
                audioFile = new AudioFileReader(unicornWavePath);
                outputDevice.Init(audioFile);
            }

            IsLoadUnicorn = true;
        }

        /// <summary>
        /// 再生終了時の発動するイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnPlaybackStopped(object sender, StoppedEventArgs e)
        {
            outputDevice.Dispose();
            outputDevice = null;
            audioFile.Dispose();
            audioFile = null;
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

            //再生
            outputDevice?.Play();

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
                    device = DevEnum.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                }
                AudioSessionManager sessionManager = device.AudioSessionManager;
                var sessions = sessionManager.Sessions;
                for (float i = 100; i >= 0; i -= 1f)
                {
                    for (int j = 0; j < sessions.Count; j++)
                    {
                        if (sessions[j].GetProcessID != (uint)pid)
                        {
                            sessions[j].SimpleAudioVolume.Volume = (i / 100.0f);
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

            outputDevice?.Stop();

            ResetVolume();
        }

        /// <summary>
        /// ボリュームを100に戻す
        /// </summary>
        private void ResetVolume()
        {
            MMDevice device = null;
            try
            {
                using (MMDeviceEnumerator DevEnum = new MMDeviceEnumerator())
                {
                    device = DevEnum.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                }
                AudioSessionManager sessionManager = device.AudioSessionManager;
                var sessions = sessionManager.Sessions;
                for (int j = 0; j < sessions.Count; j++)
                {
                    sessions[j].SimpleAudioVolume.Volume = 1.0f;
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

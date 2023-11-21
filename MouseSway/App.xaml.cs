using MouseSway.Dtos;
using MouseSway.Helpers;
using MouseSway.Logs;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Navigation;
using Application = System.Windows.Application;
using Point = System.Drawing.Point;
using Timer = System.Timers.Timer;

namespace MouseSway
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// 設定ソースの要素名
        /// </summary>
        private static readonly string CONFIG_SOURCE_ELEMENT = "appSettings";

        /// <summary>
        /// 移動値の設定キー
        /// </summary>
        private static readonly string MOVE_POINT_KEY = "movePoint";

        /// <summary>
        /// 起動時のマウス移動の設定キー
        /// </summary>
        private static readonly string SWAY_ON_START_KEY = "swayOnStart";

        /// <summary>
        /// 一時停止タイマーの活性値の設定キー
        /// </summary>
        private static readonly string SET_PAUSE_TIMER_KEY = "setPauseTimer";

        /// <summary>
        /// 監視インターバル（秒）の設定キー
        /// </summary>
        private static readonly string MONITOR_INTERVAL_SECONDS_KEY = "monitorIntervalSeconds";

        /// <summary>
        /// 一時停止開始値の設定キー
        /// </summary>
        private static readonly string PAUSE_START_TIME_KEY = "pauseStartTime";

        /// <summary>
        /// 一時停止時間（分）の設定キー
        /// </summary>
        private static readonly string PAUSE_MINUTES_KEY = "pauseMinutes";

        /// <summary>
        /// 一時停止の開始乱数（秒）の設定キー
        /// </summary>
        public static readonly string RANDOM_PAUSE_START_SECONDS_KEY = "randomPauseStartSeconds";

        /// <summary>
        /// 一時停止の終了乱数（秒）の設定キー
        /// </summary>
        public static readonly string RANDOM_PAUSE_END_SECONDS_KEY = "randomPauseEndSeconds";

        /// <summary>
        /// ログの活性値の設定キー
        /// </summary>
        public static readonly string IS_ENABLED_LOG_KEY = "isEnabledLog";

        /// <summary>
        /// MouseSway開始メニュー名
        /// </summary>
        private static readonly string START_SWAY_MOUSE_MENU_NAME = "SwayOn";

        /// <summary>
        /// MouseSway停止メニュー名
        /// </summary>
        private static readonly string STOP_SWAY_MOUSE_MENU_NAME = "SwayOff";

        /// <summary>
        /// 終了メニュー名
        /// </summary>
        private static readonly string EXIT_MENU_NAME = "Exit";

        /// <summary>
        /// 監視インターバル（秒）の既定値
        /// </summary>
        private static readonly int DEFAULT_MONITOR_INTERVAL_SECONDS = 1;

        /// <summary>
        /// ロガー
        /// </summary>
        private Logger _logger = null;

        /// <summary>
        /// MouseSway開始処理
        /// </summary>
        private Action _startSwayMouseAction = null;

        /// <summary>
        /// MouseSway停止処理
        /// </summary>
        private Action stopSwayMouseAction = null;

        /// <summary>
        /// タスクトレイ非表示処理
        /// </summary>
        private Action _hideNotifyIconAction = null;

        /// <summary>
        /// 破棄対象オブジェクトのリスト
        /// </summary>
        private List<IDisposable> _disposables = new List<IDisposable>();

        /// <summary>
        /// 設定値Dto
        /// </summary>
        private SettingDto _settingDto = null;

        /// <summary>
        /// 一時停止の開始時刻（実効値）
        /// </summary>
        private TimeSpan _actualPauseStartAt;

        /// <summary>
        /// 一時停止時間（実効値）
        /// </summary>
        private TimeSpan _actualPauseTicks;

        /// <summary>
        /// 動作確認用の一時停止値
        /// </summary>
        private bool _isStatePause = false;

        /// <summary>
        /// 終了処理の判定
        /// </summary>
        private bool _doneExit = false;

        /// <summary>
        /// デバッグ用ログファイルのパス
        /// </summary>
        private static readonly string LOG_FILE_PATH = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            $"MouseSway_Debug.log");

        /// <summary>
        /// アプリケーション開始
        /// </summary>
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            try
            {
                // 画面終了時にアプリケーションが終了しないよう設定
                Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;

                // 設定ファイルを読込
                this.LoadSetting();

                // ロガーを生成
                this._logger = new Logger(this._settingDto.IsEnabledLog);
                this._disposables.Add(this._logger);

                // アイコンを取得
                var iconStream = GetResourceStream(new Uri("./Icons/MouseSway.ico", UriKind.Relative)).Stream;
                this._disposables.Add(iconStream);

                // タスクトレイに常駐させるアイコンを生成
                var notifyIcon = new NotifyIcon
                {
                    Visible = true,
                    Icon = new Icon(iconStream),
                    Text = Assembly.GetExecutingAssembly().GetName().Name,
                    ContextMenuStrip = this.CreateContextMenuStrip(),
                };
                this._disposables.Add(notifyIcon);

                // // タスクトレイ常駐のアイコンを非表示処理を設定
                this._hideNotifyIconAction = () => notifyIcon.Visible = false;

                // 設定値から初期処理を判定して実行
                if (this._settingDto.SwayOnStart)
                {
                    // MouseSway開始を実行
                    this._startSwayMouseAction?.Invoke();
                }
                else
                {
                    // MouseSway停止を実行
                    this.stopSwayMouseAction?.Invoke();
                }

                // タスクトレイアイコンのマウスクリックイベントを定義
                notifyIcon.MouseClick += this.NotifyIcon_MouseClick;

                // 開始ログ
                this._logger.PutBeginLog();
            }
            catch(Exception ex)
            {
                // 例外ログを出力
                this._logger.PutExceptionLog(ex);

                // 例外が発生した場合は終了処理
                this.ExitApplication();
            }
        }

        /// <summary>
        /// タスクトレイアイコンのマウスクリックイベント
        /// </summary>
        private void NotifyIcon_MouseClick(object sender, MouseEventArgs e)
        {
            // 左クリック時にコンテキストメニューを表示
            if (e.Button == MouseButtons.Left)
                if (sender is NotifyIcon notifyIcon)
                {
                    var method = typeof(NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic);
                    method?.Invoke(notifyIcon, null);
                }
        }

        /// <summary>
        /// アプリケーション終了イベント
        /// </summary>
        private void Application_Exit(object sender, ExitEventArgs e) => this.ExitApplication();

        /// <summary>
        /// アプリケーション終了
        /// </summary>
        private void ExitApplication()
        {
            // 終了処理実施済の場合はスキップ
            if (this._doneExit)
                return;

            // 終了処理実施の判定値を設定
            this._doneExit = true;

            // 終了ログ
            this._logger.PutEndLog();

            // タスクトレイ常駐のアイコンを非表示に設定
            this._hideNotifyIconAction?.Invoke();

            // オブジェクトの破棄
            this._disposables?.ForEach(x => x.Dispose());

            // アプリケーション終了
            this.Shutdown();
        }

        /// <summary>
        /// 設定ファイル読込
        /// </summary>
        private void LoadSetting()
        {
            // 設定ファイルから値を取得
            ConfigurationManager.RefreshSection(CONFIG_SOURCE_ELEMENT);
            var setting = (NameValueCollection)ConfigurationManager.GetSection(CONFIG_SOURCE_ELEMENT);
            this._settingDto = new SettingDto()
            {
                MovePoint = this.ParseOrDefault<int>(setting[MOVE_POINT_KEY]),
                SwayOnStart = this.ParseOrDefault<bool>(setting[SWAY_ON_START_KEY]),
                SetPauseTimer = this.ParseOrDefault<bool>(setting[SET_PAUSE_TIMER_KEY]),
                MonitorIntervalSeconds = this.ParseOrDefault<int>(
                    setting[MONITOR_INTERVAL_SECONDS_KEY],
                    DEFAULT_MONITOR_INTERVAL_SECONDS),
                PauseStartAt = this.ParseOrDefault<TimeSpan>(setting[PAUSE_START_TIME_KEY]),
                PauseMinutes = this.ParseOrDefault<int>(setting[PAUSE_MINUTES_KEY]),
                RandomPauseStartSeconds = this.ParseOrDefault<int>(setting[RANDOM_PAUSE_START_SECONDS_KEY]),
                RandomPauseEndSeconds = this.ParseOrDefault<int>(setting[RANDOM_PAUSE_END_SECONDS_KEY]),
                IsEnabledLog = this.ParseOrDefault<bool>(setting[IS_ENABLED_LOG_KEY]),
            };

            // 実効値用の乱数を生成
            var seed = this.GetSeed();
            var random = new Random(seed);

            // 一時停止の開始時刻（実効値）
            var addPauseStartSeconds = random.Next(this._settingDto.RandomPauseStartSeconds);
            this._actualPauseStartAt = this._settingDto.PauseStartAt + (new TimeSpan(0, 0, addPauseStartSeconds));

            // 一時停止時間（実効値）
            var addPauseTicksSeconds = random.Next(this._settingDto.RandomPauseEndSeconds);
            this._actualPauseTicks = new TimeSpan(
                0,
                this._settingDto.PauseMinutes,
                addPauseTicksSeconds);
        }

        /// <summary>
        /// コンテキストメニューを生成
        /// </summary>
        private ContextMenuStrip CreateContextMenuStrip()
        {
            var result = new ContextMenuStrip();
            var timer = this.CreateMouseSwayTimer();
            var swayOnMenuItemIndex = 0;
            var swayOffMenuItemIndex = 0;

            // MouseSway開始
            this._startSwayMouseAction = new Action(() =>
            {
                result.Items[swayOnMenuItemIndex].Enabled = false;
                result.Items[swayOffMenuItemIndex].Enabled = true;
                timer.Start();
            });
            result.Items.Add(
                START_SWAY_MOUSE_MENU_NAME,
                null,
                (sender, e) => this._startSwayMouseAction?.Invoke());
            swayOnMenuItemIndex = result.Items.Count - 1;

            // MouseSway停止
            this.stopSwayMouseAction = new Action(() =>
            {
                result.Items[swayOnMenuItemIndex].Enabled = true;
                result.Items[swayOffMenuItemIndex].Enabled = false;
                timer.Stop();
            });
            result.Items.Add(
                STOP_SWAY_MOUSE_MENU_NAME,
                null,
                (sender, e) => this.stopSwayMouseAction?.Invoke());
            swayOffMenuItemIndex = result.Items.Count - 1;

            // セパレーター
            result.Items.Add(new ToolStripSeparator());

            // 常駐終了
            result.Items.Add(
                EXIT_MENU_NAME,
                null,
                (sender, e) => this.ExitApplication());

            return result;
        }

        /// <summary>
        /// マウス動作タイマーを生成
        /// </summary>
        private Timer CreateMouseSwayTimer()
        {
            // タイマーの間隔を設定
            var result = new Timer(1000);
            this._disposables.Add(result);

            var isMovePositive = false;
            result.Elapsed += (sender, e) =>
            {
                try
                {
                    // 一時停止を判定
                    if (this._settingDto.SetPauseTimer &&
                    this.IsPause())
                        return;

                    // 正方向と負方向を交互に移動
                    var pointValue = isMovePositive ? this._settingDto.MovePoint : this._settingDto.MovePoint * -1;
                    isMovePositive = !isMovePositive;

                    // マウス移動
                    var helper = MouseHelper.Instance;
                    helper.Move(new Point(pointValue));
                }
                catch (Exception ex)
                {
                    // Message TextBlockの文言を変更します。
                    this.Dispatcher.Invoke((Action)(() =>
                    {
                        // 例外ログを出力
                        this._logger.PutExceptionLog(ex);

                        // 例外が発生した場合は終了処理
                        this.ExitApplication();
                    }));
                }
            };

            return result;
        }

        /// <summary>
        /// 一時停止を判定
        /// </summary>
        private bool IsPause()
        {
            // 一時停止時間を算出
            var timestamp = DateTime.Now;
            var pauseStartTimestamp = timestamp.Date.Add(this._actualPauseStartAt);
            var pauseEndTimestamp = pauseStartTimestamp.Add(this._actualPauseTicks);
            var result = (timestamp > pauseStartTimestamp && pauseEndTimestamp > timestamp);

            // 動作確認用にログ出力
            if (this._settingDto.IsEnabledLog &&
                this._isStatePause != result)
            {
                // 一時停止状態を保持
                this._isStatePause = result;

                // 一時停止変更ログを出力
                var changeAt = result ? pauseStartTimestamp : pauseEndTimestamp;
                this._logger.PutChangePauseLog(
                    result,
                    changeAt);
            }

            return result;
        }

        /// <summary>
        /// 型変換
        /// </summary>
        /// <typeparam name="T">変換する型</typeparam>
        /// <param name="value">変換元の値</param>
        /// <returns>変換した値</returns>
        private T ParseOrDefault<T>(string value) where T : struct => this.ParseOrDefault<T>(
            value, 
            default(T));

        /// <summary>
        /// 型変換
        /// </summary>
        /// <typeparam name="T">変換する型</typeparam>
        /// <param name="value">変換元の値</param>
        /// <param name="defaultValue">既定値</param>
        /// <returns>変換した値</returns>
        private T ParseOrDefault<T>(
            string value,
            T defaultValue) where T : struct
        {
            try
            {
                return (T)TypeDescriptor.GetConverter(typeof(T)).ConvertFrom(value);
            }
            catch
            {
                return default;
            }
        }

        /// <summary>
        /// 乱数のシード値を取得
        /// </summary>
        /// <returns>シード値を取得</returns>
        private int GetSeed()
        {
            var timestampTicks = DateTime.Now.Ticks;
            while (timestampTicks > int.MaxValue)
            {
                timestampTicks -= int.MaxValue;
            }

            return (int)timestampTicks;
        }
    }
}

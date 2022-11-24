using MouseSway.Helpers;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
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
        /// 開始処理の設定キー
        /// </summary>
        private static readonly string SWAY_ON_START_KEY = "swayOnStart";

        /// <summary>
        /// 移動値の設定キー
        /// </summary>
        private static readonly string MOVE_POINT_KEY = "movePoint";

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
        /// MouseSway開始処理
        /// </summary>
        private Action? _StartSwayMouseAction = null;

        /// <summary>
        /// MouseSway停止処理
        /// </summary>
        private Action? _StopSwayMouseAction = null;

        /// <summary>
        /// タスクトレイ非表示処理
        /// </summary>
        private Action? _HideNotifyIconAction = null;

        /// <summary>
        /// 破棄対象オブジェクトのリスト
        /// </summary>
        private List<IDisposable> _Disposables = new List<IDisposable>();

        /// <summary>
        /// MouseSway設定値
        /// </summary>
        private NameValueCollection _MouseSwaySettings = (NameValueCollection)ConfigurationManager.GetSection("mouseSwaySettings");

        /// <summary>
        /// アプリケーション開始
        /// </summary>
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            // 画面終了時にアプリケーションが終了しないよう設定
            Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;

            // アイコンを取得
            var iconStream = GetResourceStream(new Uri("./Icons/MouseSway.ico", UriKind.Relative)).Stream;
            this._Disposables.Add(iconStream);

            // タスクトレイに常駐させるアイコンを生成
            var notifyIcon = new NotifyIcon
            {
                Visible = true,
                Icon = new Icon(iconStream),
                Text = Assembly.GetExecutingAssembly().GetName().Name,
                ContextMenuStrip = this.CreateContextMenuStrip(),
            };
            this._Disposables.Add(notifyIcon);

            // // タスクトレイ常駐のアイコンを非表示処理を設定
            this._HideNotifyIconAction = () => notifyIcon.Visible = false;

            // 設定値から初期処理を判定して実行
            var swayOnStartStr = this._MouseSwaySettings[SWAY_ON_START_KEY];
            if (bool.TryParse(swayOnStartStr, out var swayOnStart) && swayOnStart)
            {
                // MouseSway開始を実行
                this._StartSwayMouseAction?.Invoke();
            }
            else
            {
                // MouseSway停止を実行
                this._StopSwayMouseAction?.Invoke();
            }

            // タスクトレイアイコンのマウスクリックイベントを定義
            notifyIcon.MouseClick += (sender, e) =>
            {
                // 左クリック時にコンテキストメニューを表示
                if (e.Button == MouseButtons.Left)
                    if (sender is NotifyIcon notifyIcon)
                    {
                        var method = typeof(NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic);
                        method?.Invoke(notifyIcon, null);
                    }
            };
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
            this._StartSwayMouseAction = new Action(() =>
            {
                result.Items[swayOnMenuItemIndex].Enabled = false;
                result.Items[swayOffMenuItemIndex].Enabled = true;
                timer.Start();
            });
            result.Items.Add(
                START_SWAY_MOUSE_MENU_NAME,
                null,
                (sender, e) => this._StartSwayMouseAction?.Invoke());
            swayOnMenuItemIndex = result.Items.Count - 1;

            // MouseSway停止
            this._StopSwayMouseAction = new Action(() =>
            {
                result.Items[swayOnMenuItemIndex].Enabled = true;
                result.Items[swayOffMenuItemIndex].Enabled = false;
                timer.Stop();
            });
            result.Items.Add(
                STOP_SWAY_MOUSE_MENU_NAME,
                null,
                (sender, e) => this._StopSwayMouseAction?.Invoke());
            swayOffMenuItemIndex = result.Items.Count - 1;

            // セパレーター
            result.Items.Add(new ToolStripSeparator());

            // 常駐終了
            result.Items.Add(
                EXIT_MENU_NAME,
                null,
                (sender, e) =>
                {
                    // タスクトレイ常駐のアイコンを非表示に設定
                    this._HideNotifyIconAction?.Invoke();

                    // オブジェクトの破棄
                    this._Disposables?.ForEach(x => x.Dispose());

                    // アプリケーション終了
                    this.Shutdown();
                });

            return result;
        }

        /// <summary>
        /// マウス動作タイマーを生成
        /// </summary>
        /// <returns></returns>
        private Timer CreateMouseSwayTimer()
        {
            // タイマーの間隔を設定
            var result = new Timer(1000);
            this._Disposables.Add(result);

            var isMovePositive = false;
            result.Elapsed += (sender, e) =>
            {
                // 移動値を取得
                var movePointStr = this._MouseSwaySettings[MOVE_POINT_KEY];
                if (!int.TryParse(movePointStr, out var movePoint))
                    movePoint = default;

                // 正方向と負方向を交互に移動
                var pointValue = isMovePositive ? movePoint : movePoint * - 1;
                isMovePositive = !isMovePositive;

                // マウス移動
                var helper = MouseHelper.Instance;
                helper.Move(new Point(pointValue));
            };

            return result;
        }
    }
}

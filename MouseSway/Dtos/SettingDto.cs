using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MouseSway.Dtos
{
    /// <summary>
    /// 設定値Dto
    /// </summary>
    public class SettingDto
    {
        /// <summary>
        /// 移動値
        /// </summary>
        public int MovePoint { get; set; }

        /// <summary>
        /// 起動時のマウス移動
        /// </summary>
        public bool SwayOnStart { get; set; }

        /// <summary>
        /// 一時停止タイマーの活性値
        /// </summary>
        public bool SetPauseTimer { get; set; }

        /// <summary>
        /// 監視インターバル（秒）
        /// </summary>
        public int MonitorIntervalSeconds { get; set; }

        /// <summary>
        /// 一時停止開始値
        /// </summary>
        public TimeSpan PauseStartAt { get; set; }

        /// <summary>
        /// 一時停止時間（分）
        /// </summary>
        public int PauseMinutes { get; set; }

        /// <summary>
        /// 一時停止の開始乱数（秒）
        /// </summary>
        public int RandomPauseStartSeconds { get; set; }

        /// <summary>
        /// 一時停止の終了乱数（秒）
        /// </summary>
        public int RandomPauseEndSeconds { get; set; }

        /// <summary>
        /// ログの活性値
        /// </summary>
        public bool IsEnabledLog { get; set; }
    }
}

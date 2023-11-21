using MouseSway.Properties;
using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;
using System.Windows.Shapes;
using Path = System.IO.Path;

namespace MouseSway.Logs
{
    public class Logger : IDisposable
    {
        /// <summary>
        /// 活性値
        /// </summary>
        public bool IsEnabled { get; private set; } = false;

        /// <summary>
        /// ログファイル名
        /// </summary>
        private static readonly string FILE_NAME = @"MouseSway.log";

        /// <summary>
        /// ログファイル名
        /// </summary>
        private static readonly string DIRECTORY_NAME = @"Logs";

        /// <summary>
        /// ログファイルの最大サイズ(5MB)
        /// </summary>
        private static readonly long MAX_FILE_SIZE = 5 * 1024 * 1024;

        /// <summary>
        /// ログディレクトリのパス
        /// </summary>
        private string _directoryPath = string.Empty;

        /// <summary>
        /// ログファイルのパス
        /// </summary>
        private string _filePath = string.Empty;

        /// <summary>
        /// ログファイル用のストリーム
        /// </summary>
        private StreamWriter _streamWriter = null;

        /// <summary>
        /// ファイル排他用オブジェクト
        /// </summary>
        private object _lockObj = new object();

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public Logger(bool isEnaled)
        {
            // 活性値を設定
            this.IsEnabled = isEnaled;

            // ログファイルのパスを設定
            var assembly = Assembly.GetEntryAssembly();
            var currentDirectoryPath = Path.GetDirectoryName(assembly.Location);
            this._directoryPath = Path.Combine(
                currentDirectoryPath,
                DIRECTORY_NAME);
            this._filePath = Path.Combine(
                this._directoryPath,
                FILE_NAME);

            // ログディレクトリを生成
            if (!Directory.Exists(this._directoryPath))
                Directory.CreateDirectory(this._directoryPath);

            // ログファイル用のストリームを生成
            this._streamWriter = new StreamWriter(
                this._filePath,
                true,
                Encoding.UTF8)
            { AutoFlush = true };
        }

        /// <summary>
        /// オブジェクトの破棄
        /// </summary>
        public void Dispose()
        {
            lock (this._lockObj)
            {
                // ログファイル用のストリームを破棄
                this._streamWriter?.Close();
                this._streamWriter?.Dispose();
                this._streamWriter = null;

                // ログのローテート処理
                var logFile = new FileInfo(this._filePath);
                if (MAX_FILE_SIZE < logFile.Length)
                    RotateLog();
            }
        }

        /// <summary>
        /// 開始ログを出力
        /// </summary>
        public void PutBeginLog() => this.OutPut($"MouseSwayを起動しました。");

        /// <summary>
        /// 終了ログを出力
        /// </summary>
        public void PutEndLog() => this.OutPut($"MouseSwayを終了しました。{Environment.NewLine}");

        /// <summary>
        /// 一時停止変更ログを出力
        /// </summary>
        public void PutChangePauseLog(
            bool isPause,
            DateTime changeAt)
        {
            var sb = new StringBuilder();
            if (isPause)
                sb.AppendLine($"MouseSwayを一時停止しました。");
            else
                sb.AppendLine($"MouseSwayを再開しました。");

            sb.Append($"変更設定日時：{changeAt:HH:mm:ss}");
            this.OutPut(sb);
        }

        /// <summary>
        /// 例外ログを出力
        /// </summary>
        public void PutExceptionLog(Exception exception)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"例外が発生しました。");
            sb.Append(exception);

            this.OutPut(sb);
        }

        /// <summary>
        /// ログを出力する
        /// </summary>
        /// <param name="message">メッセージ</param>
        private void OutPut(object message) => this.OutPut(message?.ToString() ?? string.Empty);

        /// <summary>
        /// ログを出力する
        /// </summary>
        /// <param name="message">メッセージ</param>
        private void OutPut(string message)
        {
            // 非活性の場合またはメッセージが無い場合は処理をスキップ
            if (!this.IsEnabled ||
                string.IsNullOrEmpty(message))
                return;

            var logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}]{Environment.NewLine}{message}";
            lock (this._lockObj)
                this._streamWriter.WriteLine(logMessage);
            
        }

        /// <summary>
        /// ログファイルをローテートする
        /// </summary>
        private void RotateLog()
        {
            // ローテート先のファイルパスを生成
            var rotateFileName = $"{Path.GetFileNameWithoutExtension(this._filePath)}_{DateTime.Now:yyyyMMddHHmmss}.log";
            var rotateFilePath = Path.Combine(
                this._directoryPath,
                rotateFileName);

            // ローテートファイルへリネーム
            File.Move(
                this._filePath,
                rotateFilePath);
        }
    }
}

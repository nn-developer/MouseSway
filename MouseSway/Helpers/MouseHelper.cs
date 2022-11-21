using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MouseSway.Helpers
{
    public class MouseHelper
    {
        /// <summary>
        /// 単一のインスタンス
        /// </summary>
        private static MouseHelper _Instance = new MouseHelper();

        /// <summary>
        /// シングルトン用の隠蔽コンストラクタ
        /// </summary>
        private MouseHelper() { }

        /// <summary>
        /// 単一のインスタンスを取得
        /// </summary>
        public static MouseHelper Instance => _Instance;

        /// <summary>
        /// マウスカーソルを移動
        /// </summary>
        /// <param name="p"></param>
        public void Move(Point p)
        {
            Cursor.Position = new Point(
                Cursor.Position.X + p.X,
                Cursor.Position.Y + p.Y);
        }
    }
}

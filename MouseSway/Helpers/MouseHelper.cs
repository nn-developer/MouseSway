using PInvoke;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
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
        public void Move(Point p)
        {
            var inp = new User32.INPUT
            {
                type = User32.InputType.INPUT_MOUSE,
                Inputs = new User32.INPUT.InputUnion
                {
                    mi = new User32.MOUSEINPUT
                    {
                        dx = p.X,
                        dy = p.Y,
                        mouseData = 0,
                        dwFlags = User32.MOUSEEVENTF.MOUSEEVENTF_MOVE,
                        time = 0,
                        dwExtraInfo_IntPtr = IntPtr.Zero,
                    },
                },
            };

            User32.SendInput(nInputs: 1, pInputs: new[] { inp, }, cbSize: Marshal.SizeOf<User32.INPUT>());
        }
    }
}

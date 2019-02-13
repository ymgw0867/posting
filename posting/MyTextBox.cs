using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace posting
{
    public class TextBoxEx : TextBox
    {
        const int WM_IME_COMPOSITION = 0x010F;
        const int GCS_RESULTREADSTR = 0x0200;

        [DllImport("Imm32.dll")]
        public static extern int ImmGetContext(IntPtr hWnd);
        [DllImport("Imm32.dll")]
        public static extern int ImmGetCompositionString(
            int hIMC,
            int dwIndex,
            StringBuilder lpBuf,
            int dwBufLen
            );
        [DllImport("Imm32.dll")]
        public static extern bool ImmReleaseContext(IntPtr hWnd, int hIMC);

        [Serializable]
        public class CompositionEventArgs : EventArgs
        {
            public string ImeStr;
        }
        public delegate void CompositionEventHandler(object sender, CompositionEventArgs e);
        public event CompositionEventHandler CompositionEvent;

        protected override void WndProc(ref Message m)
        {
            if (CompositionEvent != null && m.Msg == WM_IME_COMPOSITION)
            {
                if (((int)m.LParam & GCS_RESULTREADSTR) > 0)
                {
                    int Imc = ImmGetContext(this.Handle);
                    int sz = ImmGetCompositionString(Imc, GCS_RESULTREADSTR, null, 0);
                    StringBuilder str = new StringBuilder(sz);
                    ImmGetCompositionString(Imc, GCS_RESULTREADSTR, str, str.Capacity);
                    ImmReleaseContext(this.Handle, Imc);
                    CompositionEventArgs args = new CompositionEventArgs();
                    args.ImeStr = str.ToString();
                    CompositionEvent(this, args);
                }
            }
            base.WndProc(ref m);
        }
    }
}


using System;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ExtratoComprovantes
{
    static class Program
    {
        static long ultimoAviso;

        [STAThread]
        static void Main()
        {
            var br = new CultureInfo("pt-BR");
            CultureInfo.DefaultThreadCurrentCulture = br;
            CultureInfo.DefaultThreadCurrentUICulture = br;
            Thread.CurrentThread.CurrentCulture = br;
            Thread.CurrentThread.CurrentUICulture = br;
#if !MONO
            Application.SetHighDpiMode(HighDpiMode.DpiUnawareGdiScaled);   // telas com escala 125%/150%: o Windows amplia a janela inteira, sem desmontar o layout
#endif
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            // erro inesperado: grava no log e avisa uma vez (o trabalho na tela não se perde)
            Application.ThreadException += (s, e) => Avisar(e.Exception);
            AppDomain.CurrentDomain.UnhandledException += (s, e) => { var ex = e.ExceptionObject as Exception; if (ex != null) Log.Erro("erro não tratado", ex); };
            Log.Info("início " + TelaPrincipal.VERSAO);
            Application.Run(new TelaPrincipal());
        }

        static void Avisar(Exception ex)
        {
            Log.Erro("erro inesperado", ex);
            long agora = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
            if (agora - ultimoAviso < 5000) return;
            ultimoAviso = agora;
            string msg = ex is OutOfMemoryException ? "Memória insuficiente. Feche outros programas e tente de novo."
                : "Ocorreu um erro inesperado: " + ex.Message + "\r\n\r\nO trabalho na tela não foi perdido. Detalhes gravados em:\r\n" + Path.Combine(Log.PastaApp(), "log.txt");
            try { Dialogos.Aviso(null, msg); } catch { MessageBox.Show(msg, "Extrato → Lançamentos"); }
        }
    }
}

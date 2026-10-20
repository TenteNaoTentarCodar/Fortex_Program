using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace ExtratoComprovantes
{
    /// <summary>Janelas de mensagem com botões em português (independem do idioma do Windows) e espera de leitura em segundo plano.</summary>
    public static class Dialogos
    {
        public static readonly Font Fonte = new Font("Segoe UI", 9f);

        /// <summary>Só para testes automáticos: responde às janelas sem mostrá-las (título, mensagem, botões → índice).</summary>
        public static Func<string, string, string[], int> Automatico;

        static Form Base(string titulo)
        {
            return new Form
            {
                Text = titulo,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowInTaskbar = false,
                StartPosition = FormStartPosition.CenterParent,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Font = Fonte,
                Padding = new Padding(12)
            };
        }

        /// <summary>Mensagem com botões; devolve o índice do botão clicado ou -1 (fechou/Esc).</summary>
        public static int Escolher(IWin32Window dono, string titulo, string msg, int padrao, params string[] botoes)
        {
            if (Automatico != null) return Automatico(titulo, msg, botoes);
            using (var f = Base(titulo))
            {
                var lay = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, Dock = DockStyle.Fill };
                var l = new Label { Text = msg, AutoSize = true, MaximumSize = new Size(720, 0), Margin = new Padding(0, 0, 0, 12) };
                lay.Controls.Add(l, 0, 0);
                var bar = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill, WrapContents = false };
                int esc = -1;
                var bs = new List<Button>();
                for (int i = 0; i < botoes.Length; i++)
                {
                    int k = i;
                    var b = new Button { Text = botoes[i], AutoSize = true, MinimumSize = new Size(90, 28), Margin = new Padding(6, 0, 0, 0) };
                    b.Click += (s, e) => { f.Tag = k; f.Close(); };
                    bs.Add(b);
                }
                for (int i = bs.Count - 1; i >= 0; i--) bar.Controls.Add(bs[i]);
                lay.Controls.Add(bar, 0, 1);
                f.Controls.Add(lay);
                if (padrao >= 0 && padrao < bs.Count) f.AcceptButton = bs[padrao];
                f.KeyPreview = true;
                f.KeyDown += (s, e) => { if (e.KeyCode == Keys.Escape) { f.Tag = esc; f.Close(); } };
                f.Shown += (s, e) => { if (padrao >= 0 && padrao < bs.Count) bs[padrao].Focus(); };
                f.ShowDialog(dono);
                return f.Tag is int ? (int)f.Tag : -1;
            }
        }

        public static void Aviso(IWin32Window dono, string msg) { Escolher(dono, "Atenção", msg, 0, "OK"); }

        public static void Info(IWin32Window dono, string titulo, string msg) { Escolher(dono, titulo, msg, 0, "OK"); }

        public static bool SimNao(IWin32Window dono, string titulo, string msg) { return Escolher(dono, titulo, msg, 0, "Sim", "Não") == 0; }

        /// <summary>Texto longo (relatório) em fonte de largura fixa, com botões.</summary>
        public static int Texto(IWin32Window dono, string titulo, string texto, params string[] botoes)
        {
            if (Automatico != null) return Automatico(titulo, texto, botoes);
            using (var f = new Form { Text = titulo, StartPosition = FormStartPosition.CenterParent, Size = new Size(820, 560), Font = Fonte, ShowInTaskbar = false, MinimizeBox = false })
            {
                var tb = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false, Dock = DockStyle.Fill, Font = new Font(FontFamily.GenericMonospace, 9.5f), Text = texto.Replace("\r\n", "\n").Replace("\n", "\r\n"), BackColor = SystemColors.Window };
                var bar = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(6) };
                var bs = new List<Button>();
                for (int i = 0; i < botoes.Length; i++)
                {
                    int k = i;
                    var b = new Button { Text = botoes[i], AutoSize = true, MinimumSize = new Size(100, 28) };
                    b.Click += (s, e) => { f.Tag = k; f.Close(); };
                    bs.Add(b);
                }
                for (int i = bs.Count - 1; i >= 0; i--) bar.Controls.Add(bs[i]);
                f.Controls.Add(tb);
                f.Controls.Add(bar);
                if (bs.Count > 0) f.AcceptButton = bs[0];
                f.Shown += (s, e) => { tb.SelectionStart = 0; tb.SelectionLength = 0; if (bs.Count > 0) bs[0].Focus(); };
                f.ShowDialog(dono);
                return f.Tag is int ? (int)f.Tag : -1;
            }
        }

        /// <summary>Escolha numa lista (caixa de combinação). Devolve o índice ou -1.</summary>
        public static int Lista(IWin32Window dono, string titulo, string msg, IList<string> itens, int pre)
        {
            if (Automatico != null) { int a = Automatico(titulo, msg, itens.ToArray()); return a < 0 ? pre : a; }
            using (var f = Base(titulo))
            {
                var lay = new TableLayoutPanel { AutoSize = true, ColumnCount = 1, Dock = DockStyle.Fill };
                lay.Controls.Add(new Label { Text = msg, AutoSize = true, MaximumSize = new Size(720, 0), Margin = new Padding(0, 0, 0, 8) }, 0, 0);
                var cb = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 640, MaxDropDownItems = 25 };
                foreach (var s in itens) cb.Items.Add(s);
                if (itens.Count > 0) cb.SelectedIndex = Math.Max(0, Math.Min(pre, itens.Count - 1));
                lay.Controls.Add(cb, 0, 1);
                var bar = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 12, 0, 0) };
                var ok = new Button { Text = "OK", MinimumSize = new Size(90, 28), AutoSize = true };
                var ca = new Button { Text = "Cancelar", MinimumSize = new Size(90, 28), AutoSize = true };
                ok.Click += (s, e) => { f.Tag = cb.SelectedIndex; f.Close(); };
                ca.Click += (s, e) => { f.Tag = -1; f.Close(); };
                bar.Controls.Add(ca); bar.Controls.Add(ok);
                lay.Controls.Add(bar, 0, 2);
                f.Controls.Add(lay);
                f.AcceptButton = ok; f.CancelButton = ca;
                f.ShowDialog(dono);
                return f.Tag is int ? (int)f.Tag : -1;
            }
        }

        /// <summary>Pede um texto curto. null = cancelado.</summary>
        public static string PedirTexto(IWin32Window dono, string titulo, string msg, string inicial)
        {
            if (Automatico != null) { Automatico(titulo, msg, new string[0]); return inicial; }
            using (var f = Base(titulo))
            {
                var lay = new TableLayoutPanel { AutoSize = true, ColumnCount = 1, Dock = DockStyle.Fill };
                lay.Controls.Add(new Label { Text = msg, AutoSize = true, MaximumSize = new Size(640, 0), Margin = new Padding(0, 0, 0, 8) }, 0, 0);
                var tb = new TextBox { Width = 420, Text = inicial ?? "" };
                lay.Controls.Add(tb, 0, 1);
                var bar = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 12, 0, 0) };
                var ok = new Button { Text = "OK", MinimumSize = new Size(90, 28), AutoSize = true };
                var ca = new Button { Text = "Cancelar", MinimumSize = new Size(90, 28), AutoSize = true };
                string r = null;
                ok.Click += (s, e) => { r = tb.Text; f.Close(); };
                ca.Click += (s, e) => f.Close();
                bar.Controls.Add(ca); bar.Controls.Add(ok);
                lay.Controls.Add(bar, 0, 2);
                f.Controls.Add(lay);
                f.AcceptButton = ok; f.CancelButton = ca;
                f.ShowDialog(dono);
                return r;
            }
        }

        /// <summary>
        /// Executa uma leitura demorada fora da tela (a janela não "congela") mostrando "Lendo ...".
        /// Devolve o resultado ou relança o erro da leitura.
        /// </summary>
        public static T Pesado<T>(IWin32Window dono, string msg, Func<T> trabalho)
        {
            T res = default(T);
            Exception erro = null;
            using (var f = new Form
            {
                Text = "Aguarde",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                ControlBox = false,
                ShowInTaskbar = false,
                StartPosition = FormStartPosition.CenterParent,
                Size = new Size(460, 120),
                Font = Fonte
            })
            {
                var l = new Label { Text = msg, Dock = DockStyle.Top, Height = 34, AutoEllipsis = true, Padding = new Padding(10, 10, 10, 0) };
                var pb = new ProgressBar { Style = ProgressBarStyle.Marquee, Dock = DockStyle.Top, Height = 18, MarqueeAnimationSpeed = 30 };
                var pad = new Panel { Dock = DockStyle.Top, Height = 6 };
                f.Controls.Add(pb); f.Controls.Add(pad); f.Controls.Add(l);
                bool fim = false;
                f.Shown += (s, e) =>
                {
                    var t = new Thread(() =>
                    {
                        try { res = trabalho(); }
                        catch (Exception ex) { erro = ex; }
                        try { f.BeginInvoke((Action)(() => { fim = true; f.Close(); })); } catch { }
                    });
                    t.IsBackground = true;
                    t.Start();
                };
                f.FormClosing += (s, e) => { if (!fim) e.Cancel = true; };
                f.ShowDialog(dono);
            }
            if (erro != null)
            {
                if (erro is LeituraException || erro is System.IO.IOException) throw erro;
                throw new LeituraException(erro.Message, erro);
            }
            return res;
        }
    }
}

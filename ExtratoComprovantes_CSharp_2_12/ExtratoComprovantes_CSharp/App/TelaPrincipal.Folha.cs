using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Cliente &gt; Folha e adiantamento: relações de líquidos do Domínio (Folha Mensal e Adiantamento), contas e intervalos de pagamento.
    /// </summary>
    public partial class TelaPrincipal
    {
        FolhaPagamento folha;

        /// <summary>Janela aberta e escolha dos PDFs (usados pelo teste automático).</summary>
        Form FolhaAberta;
        Func<string[]> ArquivosFolha;

        /// <summary>Configuração inicial: contas achadas no Plano pelo nome.</summary>
        FolhaPagamento.Config ConfigPadraoFolha()
        {
            var c = new FolhaPagamento.Config();
            if (plano == null) return c;
            foreach (var k in plano.lista)
            {
                if (!k.Analitica()) continue;
                string n = Regras.Normalizar(k.nome);
                if (c.contaFolha == null && k.Grupo() == '2' && n.Contains("SALARI") && (n.Contains("PAGAR") || n.Contains("ORDENADO"))) c.contaFolha = k.codigo;
                if (c.contaAdiantamento == null && k.Grupo() == '1' && n.Contains("ADIANTAMENTO") && n.Contains("SALARI")) c.contaAdiantamento = k.codigo;
                if (c.contaProlabore == null && k.Grupo() == '2' && n.Replace("-", " ").Contains("PRO LABORE") && n.Contains("PAGAR")) c.contaProlabore = k.codigo;
            }
            return c;
        }

        /// <summary>Relê as relações de líquidos gravadas nas regras do cliente.</summary>
        void CarregarFolhaDasRegras(StringBuilder err)
        {
            folha = null;
            if (regras.folhaArquivos.Count == 0) return;
            var f = new FolhaPagamento { cfg = regras.folhaCfg ?? ConfigPadraoFolha() };
            foreach (var a in regras.folhaArquivos)
            {
                try
                {
                    if (!File.Exists(a)) { err.Append("• Relação de líquidos não encontrada: ").Append(a).Append("\r\n"); continue; }
                    Dialogos.Pesado(this, "Lendo " + Path.GetFileName(a) + "...", () => f.Carregar(a));
                }
                catch (Exception ex) { err.Append("• ").Append(ex.Message).Append("\r\n"); }
            }
            if (f.itens.Count > 0) folha = f;
        }

        static string ResumoFolha(FolhaPagamento f)
        {
            if (f == null || f.itens.Count == 0) return "Nenhuma relação carregada.";
            var sb = new StringBuilder();
            foreach (var g in f.itens.GroupBy(x => (x.tipo == FolhaPagamento.Tipo.Folha ? "Folha" : "Adiantamento")).OrderBy(g => g.Key))
            {
                var comps = g.Select(x => x.ano * 100 + x.mes).Distinct().OrderBy(x => x).ToList();
                sb.Append(g.Key).Append(": ").Append(comps.Count).Append(" competência(s) — ")
                  .Append((comps[0] % 100).ToString("00")).Append('/').Append(comps[0] / 100)
                  .Append(comps.Count > 1 ? " a " + (comps[comps.Count - 1] % 100).ToString("00") + "/" + comps[comps.Count - 1] / 100 : "")
                  .Append(" | ").Append(g.Count()).Append(" líquido(s)").Append("\r\n");
            }
            int pl = f.itens.Count(x => x.contribuinte);
            if (pl > 0) sb.Append("Pró-labore (contribuintes): ").Append(pl).Append(" líquido(s)\r\n");
            if (f.ignorados.Count > 0) sb.Append("Páginas de outros cálculos ignoradas: ").Append(string.Join(", ", f.ignorados.Distinct().Take(6))).Append(f.ignorados.Count > 6 ? "..." : "").Append("\r\n");
            return sb.ToString().TrimEnd();
        }

        void JanelaFolha()
        {
            if (!SincronizarCampos(true)) return;
            var cfg = regras.folhaCfg ?? ConfigPadraoFolha();
            var arqs = new List<string>(regras.folhaArquivos);
            var f = new Form
            {
                Text = "Folha e adiantamento (relação de líquidos)", Font = Dialogos.Fonte, FormBorderStyle = FormBorderStyle.Sizable, MinimizeBox = false,
                ShowInTaskbar = false, StartPosition = FormStartPosition.CenterParent, Size = new Size(820, 640), MinimumSize = new Size(700, 560)
            };
            var lay = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(8) };
            lay.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            lay.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            lay.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            lay.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            lay.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            lay.Controls.Add(new Label
            {
                AutoSize = true, MaximumSize = new Size(780, 0), Padding = new Padding(0, 0, 0, 6),
                Text = "Anexe a \"Relação Geral dos Líquidos\" do Domínio — uma de Folha Mensal e uma de Adiantamento (pode ser o ano todo, uma competência por página).\r\n"
                     + "Cada saída do extrato com o nome de um empregado, paga dentro do intervalo abaixo, vai para a conta da folha ou do adiantamento. "
                     + "Valor igual ao líquido: confere. Valor diferente: usa a conta e marca \"conferir\" na Observação. Fora dos intervalos, vale a classificação normal."
            }, 0, 0);

            // ---------- arquivos ----------
            var g1 = new GroupBox { Text = "1. Relações de líquidos", Dock = DockStyle.Fill, Padding = new Padding(8) };
            var lista = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true, IntegralHeight = false };
            var resumo = new Label { Dock = DockStyle.Bottom, AutoSize = true, MaximumSize = new Size(760, 0), Padding = new Padding(0, 6, 0, 0) };
            var bAdd = new Button { Name = "bAddFolha", Text = "Adicionar PDF(s)...", AutoSize = true };
            var bRem = new Button { Text = "Remover selecionado", AutoSize = true };
            var barA = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
            barA.Controls.Add(bAdd); barA.Controls.Add(bRem);
            g1.Controls.Add(lista); g1.Controls.Add(resumo); g1.Controls.Add(barA);
            lay.Controls.Add(g1, 0, 1);

            FolhaPagamento previa = null;
            Action reler = () =>
            {
                var p = new FolhaPagamento();
                var err = new StringBuilder();
                foreach (var a in arqs)
                {
                    try
                    {
                        if (!File.Exists(a)) { err.Append("⚠ não encontrado: ").Append(a).Append("\r\n"); continue; }
                        Dialogos.Pesado(f, "Lendo " + Path.GetFileName(a) + "...", () => p.Carregar(a));
                    }
                    catch (Exception ex) { err.Append("⚠ ").Append(ex.Message).Append("\r\n"); }
                }
                previa = p;
                lista.Items.Clear();
                foreach (var a in arqs) lista.Items.Add(a);
                resumo.Text = (err.Length > 0 ? err.ToString() : "") + ResumoFolha(p);
            };

            // ---------- contas ----------
            var g2 = new GroupBox { Text = "2. Contas", Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 4, 8, 4) };
            var t2 = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3 };
            TextBox fFol = new TextBox { Name = "fContaFolha", Width = 80, Text = cfg.contaFolha == null ? "" : cfg.contaFolha.ToString() },
                    fAd = new TextBox { Name = "fContaAdiant", Width = 80, Text = cfg.contaAdiantamento == null ? "" : cfg.contaAdiantamento.ToString() },
                    fPl = new TextBox { Name = "fContaProlabore", Width = 80, Text = cfg.contaProlabore == null ? "" : cfg.contaProlabore.ToString() };
            Label nFol = new Label { AutoSize = true, Padding = new Padding(0, 6, 0, 0) }, nAd = new Label { AutoSize = true, Padding = new Padding(0, 6, 0, 0) },
                  nPl = new Label { AutoSize = true, Padding = new Padding(0, 6, 0, 0) };
            int lin = 0;
            Action<string, TextBox, Label> campo = (r, tb, nl) =>
            {
                t2.Controls.Add(new Label { Text = r, AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, lin);
                t2.Controls.Add(CampoConta.ComBotao(tb, () => plano), 1, lin);
                t2.Controls.Add(nl, 2, lin);
                lin++;
            };
            campo("Folha (salários a pagar):", fFol, nFol);
            campo("Adiantamento de salário:", fAd, nAd);
            campo("Pró-labore a pagar (contribuintes):", fPl, nPl);
            Action<TextBox, Label> nome = (tb, lb) =>
            {
                int v;
                if (tb.Text.Trim().Length == 0) { lb.Text = tb == fPl ? "em branco: pró-labore segue a classificação normal" : ""; lb.ForeColor = Color.Gray; return; }
                if (!int.TryParse(tb.Text.Trim(), out v)) { lb.Text = "⚠ código inválido"; lb.ForeColor = Color.FromArgb(180, 0, 0); return; }
                var pc = plano == null ? null : plano.Get(v);
                lb.Text = plano == null ? "" : pc == null ? "⚠ não existe no Plano" : pc.Analitica() ? pc.nome : "⚠ " + pc.nome + " é sintética";
                lb.ForeColor = pc != null && pc.Analitica() ? Color.FromArgb(0, 110, 0) : Color.FromArgb(180, 0, 0);
            };
            foreach (var x in new[] { Tuple.Create(fFol, nFol), Tuple.Create(fAd, nAd), Tuple.Create(fPl, nPl) })
            {
                var tb = x.Item1; var lb = x.Item2;
                tb.TextChanged += (s, e) => nome(tb, lb);
                nome(tb, lb);
            }
            g2.Controls.Add(t2);
            lay.Controls.Add(g2, 0, 2);

            // ---------- intervalos ----------
            var g3 = new GroupBox { Text = "3. Intervalos de pagamento", Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 4, 8, 4) };
            var t3 = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1 };
            Func<string, int, NumericUpDown> dia = (n, v) => new NumericUpDown { Name = n, Minimum = 1, Maximum = 31, Value = Math.Max(1, Math.Min(31, v)), Width = 55 };
            var adI = dia("nAdIni", cfg.adIni); var adF = dia("nAdFim", cfg.adFim); var foI = dia("nFolIni", cfg.folIni); var foF = dia("nFolFim", cfg.folFim);
            var l1 = new FlowLayoutPanel { AutoSize = true };
            l1.Controls.Add(Rot("Adiantamento: pago do dia")); l1.Controls.Add(adI); l1.Controls.Add(Rot("ao dia")); l1.Controls.Add(adF); l1.Controls.Add(Rot("do MÊS DA COMPETÊNCIA (ex.: adiantamento 01/2026 pago de 15/01 a 25/01)"));
            var l2 = new FlowLayoutPanel { AutoSize = true };
            l2.Controls.Add(Rot("Folha:            pago do dia")); l2.Controls.Add(foI); l2.Controls.Add(Rot("ao dia")); l2.Controls.Add(foF); l2.Controls.Add(Rot("do MÊS SEGUINTE à competência (ex.: folha 01/2026 paga de 01/02 a 10/02)"));
            t3.Controls.Add(l1); t3.Controls.Add(l2);
            g3.Controls.Add(t3);
            lay.Controls.Add(g3, 0, 3);

            var bts = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(0, 6, 0, 0) };
            var ok = new Button { Name = "bOkFolha", Text = "OK", MinimumSize = new Size(100, 30), AutoSize = true };
            var ca = new Button { Text = "Cancelar", MinimumSize = new Size(100, 30), AutoSize = true };
            var limpar = new Button { Text = "Não usar folha neste cliente", AutoSize = true, MinimumSize = new Size(0, 30) };
            bts.Controls.Add(ca); bts.Controls.Add(ok); bts.Controls.Add(limpar);
            lay.Controls.Add(bts, 0, 4);
            f.Controls.Add(lay);
            f.AcceptButton = ok; f.CancelButton = ca;

            bAdd.Click += (s, e) =>
            {
                string[] novos = ArquivosFolha != null ? ArquivosFolha() : null;
                if (novos == null)
                    using (var fc = new OpenFileDialog { Multiselect = true, Filter = "Relação de líquidos (PDF)|*.pdf|Todos os arquivos|*.*", Title = "Relação Geral dos Líquidos — Folha Mensal e/ou Adiantamento", InitialDirectory = Prefs.Pasta("dirFolha", PastaBase()) })
                    {
                        if (fc.ShowDialog(f) != DialogResult.OK) return;
                        Prefs.Put("dirFolha", Path.GetDirectoryName(fc.FileNames[0]));
                        novos = fc.FileNames;
                    }
                foreach (var a in novos) { string c = Path.GetFullPath(a); if (!arqs.Contains(c)) arqs.Add(c); }
                reler();
            };
            bRem.Click += (s, e) => { if (lista.SelectedIndex >= 0) { arqs.RemoveAt(lista.SelectedIndex); reler(); } };
            ca.Click += (s, e) => f.Close();
            bool aplicar = false, desligar = false;
            limpar.Click += (s, e) => { desligar = true; f.Close(); };
            ok.Click += (s, e) =>
            {
                int v;
                Func<TextBox, int?> ler = tb => int.TryParse(tb.Text.Trim(), out v) ? (int?)v : null;
                var nc = new FolhaPagamento.Config
                {
                    contaFolha = ler(fFol), contaAdiantamento = ler(fAd), contaProlabore = ler(fPl),
                    adIni = (int)adI.Value, adFim = (int)adF.Value, folIni = (int)foI.Value, folFim = (int)foF.Value
                };
                if (arqs.Count > 0 && (nc.contaFolha == null || nc.contaAdiantamento == null)) { Dialogos.Aviso(f, "Informe a conta da folha e a do adiantamento."); return; }
                foreach (var c in new[] { nc.contaFolha, nc.contaAdiantamento, nc.contaProlabore })
                    if (c != null && plano != null && !plano.Valida(c)) { Dialogos.Aviso(f, "A conta " + c + " não existe no Plano de Contas ou é sintética."); return; }
                if (nc.adIni > nc.adFim || nc.folIni > nc.folFim) { Dialogos.Aviso(f, "Intervalo inválido: o dia inicial não pode ser maior que o final."); return; }
                cfg = nc;
                aplicar = true;
                f.Close();
            };
            f.Shown += (s, e) => reler();
            FolhaAberta = f;
            try { f.ShowDialog(this); }
            finally { FolhaAberta = null; f.Dispose(); }

            if (desligar)
            {
                regras.folhaArquivos.Clear(); regras.folhaCfg = null; folha = null;
                if (arquivoRegras != null) SalvarRegras();
                Revincular();
                status.Text = "Folha e adiantamento: desligado para este cliente.";
                return;
            }
            if (!aplicar) return;
            regras.folhaArquivos.Clear(); regras.folhaArquivos.AddRange(arqs);
            regras.folhaCfg = cfg;
            if (previa != null) previa.cfg = cfg;
            folha = previa != null && previa.itens.Count > 0 ? previa : null;
            if (arquivoRegras != null || arqs.Count > 0) SalvarRegras();
            Revincular();
            if (folha != null) Dialogos.Info(this, "Folha e adiantamento", RelatorioFolha());
        }

        /// <summary>Resultado no extrato: conferidos, a conferir e líquidos sem pagamento localizado no período do extrato.</summary>
        string RelatorioFolha()
        {
            var sb = new StringBuilder();
            var usados = new HashSet<FolhaPagamento.Item>();
            int ok = 0, conf = 0;
            foreach (var l in extrato)
            {
                if (l.descartada || !l.saida) continue;
                var r = folha.Encontrar(l, w => Classificador.TokSet(TextoBusca(l)).Any(t => t == w || w.Length >= 5 && t.Length >= 5 && (t.StartsWith(w, StringComparison.Ordinal) || w.StartsWith(t, StringComparison.Ordinal))));
                if (r == null) continue;
                usados.Add(r.item);
                if (r.valorConfere) ok++; else conf++;
            }
            sb.Append("Relações carregadas:\r\n").Append(ResumoFolha(folha)).Append("\r\n\r\n");
            sb.Append("No extrato: ").Append(ok).Append(" pagamento(s) com o líquido conferido");
            if (conf > 0) sb.Append(" | ").Append(conf).Append(" com valor diferente do líquido (marcados \"conferir\" na Observação)");
            sb.Append(".\r\n");
            // líquidos cujo intervalo de pagamento cai no período do extrato e que não foram localizados
            var datas = extrato.Select(l => J.Data(l.data)).Where(d => d != null).Select(d => d.Value).ToList();
            if (datas.Count > 0)
            {
                DateTime ini = datas.Min(), fim = datas.Max();
                var falta = new List<string>();
                foreach (var it in folha.itens)
                {
                    if (usados.Contains(it)) continue;
                    bool noPeriodo = false;
                    for (var d = ini; d <= fim; d = d.AddDays(1)) if (folha.NoIntervalo(it, d)) { noPeriodo = true; break; }
                    if (noPeriodo) falta.Add("   • " + it.Rotulo() + " – " + it.nome + ": R$ " + Moeda(it.valor));
                }
                if (falta.Count > 0)
                {
                    sb.Append("\r\nNão localizados no extrato (intervalo de pagamento dentro do período do extrato):\r\n");
                    foreach (var x in falta.Take(25)) sb.Append(x).Append("\r\n");
                    if (falta.Count > 25) sb.Append("   ... e mais ").Append(falta.Count - 25).Append("\r\n");
                    sb.Append("(pagos em outra conta, em dinheiro, ainda não pagos ou com nome diferente no extrato)\r\n");
                }
            }
            return sb.ToString().TrimEnd();
        }
    }
}

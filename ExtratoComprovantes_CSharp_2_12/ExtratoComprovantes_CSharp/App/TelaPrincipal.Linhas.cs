using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExtratoComprovantes
{
    /// <summary>Inserir, duplicar e excluir lançamentos na tela principal.</summary>
    public partial class TelaPrincipal
    {
        /// <summary>Janela de inclusão aberta (usada pelo teste automático).</summary>
        Form InclusaoAberta;

        void ExcluirSelecionados()
        {
            if (grade.IsCurrentCellInEditMode) return;
            var rem = Selecionados();
            if (rem.Count == 0) { Dialogos.Aviso(this, "Selecione na tabela o(s) lançamento(s) a excluir."); return; }
            double ent = rem.Where(l => !l.saida).Sum(l => l.valor), sai = rem.Where(l => l.saida).Sum(l => l.valor);
            int exp = rem.Count(l => l.exportadoEm.Length > 0);
            string msg = rem.Count == 1
                ? "Excluir o lançamento abaixo?\r\n\r\n" + rem[0].data + "  " + rem[0].historico + "  " + (rem[0].saida ? "D " : "C ") + Moeda(rem[0].valor)
                : "Excluir " + rem.Count + " lançamento(s) selecionado(s)?\r\n\r\nEntradas (C): R$ " + Moeda(ent) + "   |   Saídas (D): R$ " + Moeda(sai);
            if (exp > 0) msg += "\r\n\r\n" + exp + " já foi(ram) exportado(s): a exclusão não desfaz o TXT já gerado (continua guardado no progresso para não importar de novo).";
            msg += "\r\n\r\nComprovantes, notas e itens da planilha ligados a eles voltam a ficar livres.";
            if (Dialogos.Escolher(this, "Excluir lançamentos", msg, 1, "Excluir", "Cancelar") != 0) return;
            var set = new HashSet<Lancamento>(rem);
            int primeiro = vis.FindIndex(i => i < extrato.Count && set.Contains(extrato[i]));
            extrato.RemoveAll(set.Contains);
            Revincular();
            if (grade.RowCount > 0 && primeiro >= 0)
            {
                int r = Math.Min(primeiro, grade.RowCount - 1);
                grade.ClearSelection();
                grade.CurrentCell = grade.Rows[r].Cells[Math.Max(0, grade.CurrentCell == null ? 0 : grade.CurrentCell.ColumnIndex)];
            }
            status.Text = status.Text + "   |   Excluído(s): " + rem.Count;
        }

        void InserirLancamento() { NovoLancamento(false); }
        void DuplicarLancamento()
        {
            if (Atual() == null) { Dialogos.Aviso(this, "Selecione na tabela o lançamento a duplicar."); return; }
            NovoLancamento(true);
        }

        void NovoLancamento(bool duplicar)
        {
            if (grade.IsCurrentCellInEditMode) grade.EndEdit();
            SincronizarCampos(false);
            abas.SelectedIndex = 0;
            var ref0 = Atual();
            var f = new Form
            {
                Text = duplicar ? "Duplicar lançamento" : "Inserir lançamento", Font = Dialogos.Fonte, FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false, ShowInTaskbar = false, StartPosition = FormStartPosition.CenterParent,
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(12)
            };
            var t = new TableLayoutPanel { AutoSize = true, ColumnCount = 3, Dock = DockStyle.Fill };
            var fData = new MaskedTextBox { Name = "fData", Mask = "00/00/0000", Width = 90, TextMaskFormat = MaskFormat.IncludeLiterals, PromptChar = '_' };
            var fHist = new TextBox { Name = "fHist", Width = 460 };
            var fValor = new TextBox { Name = "fValor", Width = 120, TextAlign = HorizontalAlignment.Right };
            var cbDC = new ComboBox { Name = "cbDC", DropDownStyle = ComboBoxStyle.DropDownList, Width = 330 };
            cbDC.Items.AddRange(new object[] { "D - Débito (saída do banco)", "C - Crédito (entrada no banco)" });
            var fContra = new TextBox { Name = "fContra", Width = 80 };
            var fBanco = new TextBox { Name = "fBanco", Width = 80 };
            Label nContra = new Label { AutoSize = true, Padding = new Padding(0, 6, 0, 0) }, nBanco = new Label { AutoSize = true, Padding = new Padding(0, 6, 0, 0) };

            if (ref0 != null)
            {
                fData.Text = ref0.data;
                if (BancoDe(ref0) != null) fBanco.Text = BancoDe(ref0).ToString();
            }
            else if (regras.contaBanco != null) fBanco.Text = regras.contaBanco.ToString();
            cbDC.SelectedIndex = ref0 != null && !ref0.saida ? 1 : 0;
            if (duplicar && ref0 != null)
            {
                fHist.Text = ref0.historico;
                fValor.Text = Moeda(ref0.valor);
                int? c = ref0.saida ? ref0.debito : ref0.credito;
                if (c != null) fContra.Text = c.ToString();
            }

            int lin = 0;
            Action<string, Control, Control> campo = (r, c, extra) =>
            {
                t.Controls.Add(new Label { Text = r, AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, lin);
                t.Controls.Add(c, 1, lin);
                if (extra != null) t.Controls.Add(extra, 2, lin);
                lin++;
            };
            campo("Data:", fData, null);
            campo("Histórico:", fHist, null); t.SetColumnSpan(fHist, 2);
            campo("Valor (R$):", fValor, null);
            campo("D/C:", cbDC, null); t.SetColumnSpan(cbDC, 2);
            campo("Contrapartida:", CampoConta.ComBotao(fContra, () => plano), nContra);
            campo("Banco/caixa:", CampoConta.ComBotao(fBanco, () => plano), nBanco);
            var dica = new Label
            {
                AutoSize = true, MaximumSize = new Size(620, 0), ForeColor = Color.FromArgb(0, 90, 160), Padding = new Padding(0, 8, 0, 0),
                Text = "D: a contrapartida vai no Débito e o banco no Crédito. C: o banco vai no Débito e a contrapartida no Crédito.\r\n"
                     + "Contrapartida em branco: o programa classifica pelas regras, Razão e Plano, como as linhas do extrato. F2 busca a conta no Plano."
            };
            t.Controls.Add(dica, 0, lin); t.SetColumnSpan(dica, 3); lin++;
            var bar = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 12, 0, 0) };
            var ok = new Button { Name = "bInserir", Text = duplicar ? "Duplicar" : "Inserir", MinimumSize = new Size(100, 30), AutoSize = true };
            var ca = new Button { Text = "Cancelar", MinimumSize = new Size(100, 30), AutoSize = true };
            bar.Controls.Add(ca); bar.Controls.Add(ok);
            t.Controls.Add(bar, 0, lin); t.SetColumnSpan(bar, 3);
            f.Controls.Add(t);
            f.AcceptButton = ok; f.CancelButton = ca;

            Action<TextBox, Label> nome = (tb, lb) =>
            {
                string s = tb.Text.Trim();
                int v;
                if (s.Length == 0) { lb.Text = ""; return; }
                if (!int.TryParse(s, out v)) { lb.Text = "⚠ código inválido"; lb.ForeColor = Color.FromArgb(180, 0, 0); return; }
                if (plano == null) { lb.Text = ""; return; }
                var pc = plano.Get(v);
                lb.Text = pc == null ? "⚠ não existe no Plano" : pc.Analitica() ? pc.nome : "⚠ " + pc.nome + " é sintética";
                lb.ForeColor = pc != null && pc.Analitica() ? Color.FromArgb(0, 110, 0) : Color.FromArgb(180, 0, 0);
            };
            fContra.TextChanged += (s, e) => nome(fContra, nContra);
            fBanco.TextChanged += (s, e) => nome(fBanco, nBanco);
            nome(fContra, nContra); nome(fBanco, nBanco);
            ca.Click += (s, e) => f.Close();

            Lancamento novo = null;
            ok.Click += (s, e) =>
            {
                string data = fData.Text.Trim();
                var dt = J.Data(data);
                if (!data.Matches(@"[0-9]{2}/[0-9]{2}/[0-9]{4}") || dt == null) { Dialogos.Aviso(f, "Informe a data (dd/mm/aaaa)."); fData.Focus(); return; }
                string hist = fHist.Text.Trim();
                if (hist.Length == 0) { Dialogos.Aviso(f, "Informe o histórico."); fHist.Focus(); return; }
                double? v = Numero(fValor.Text);
                if (v == null || Math.Abs(v.Value) < 0.005) { Dialogos.Aviso(f, "Informe o valor."); fValor.Focus(); return; }
                int? contra = null, banco = null;
                int x;
                if (fContra.Text.Trim().Length > 0)
                {
                    if (!int.TryParse(fContra.Text.Trim(), out x)) { Dialogos.Aviso(f, "Código da contrapartida inválido."); return; }
                    if (plano != null && !plano.Valida(x)) { Dialogos.Aviso(f, "A conta " + x + " não existe no Plano de Contas ou é sintética."); return; }
                    contra = x;
                }
                if (fBanco.Text.Trim().Length > 0)
                {
                    if (!int.TryParse(fBanco.Text.Trim(), out x)) { Dialogos.Aviso(f, "Código do banco/caixa inválido."); return; }
                    if (plano != null && !plano.Valida(x)) { Dialogos.Aviso(f, "A conta " + x + " não existe no Plano de Contas ou é sintética."); return; }
                    banco = x;
                }
                if (banco == null) { Dialogos.Aviso(f, "Informe a conta do banco/caixa."); fBanco.Focus(); return; }
                bool saida = cbDC.SelectedIndex == 0;
                var l = new Lancamento(data, hist, Math.Abs(v.Value), saida) { origem = "Inserido manualmente" };
                l.contaBanco = banco == regras.contaBanco ? null : banco;
                if (ref0 != null && BancoDe(ref0) == banco) l.contaBanco = ref0.contaBanco;
                if (contra != null)
                {
                    if (saida) { l.debito = contra; l.credito = banco; } else { l.debito = banco; l.credito = contra; }
                    l.manual = true;
                    l.obs = "Inserido manualmente";
                }
                l.chave = "MANUAL|" + ChaveBase(l) + "|" + DateTime.Now.Ticks;
                novo = l;
                f.Close();
            };
            f.Shown += (s, e) => (duplicar ? (Control)fValor : fHist).Focus();
            InclusaoAberta = f;
            try { f.ShowDialog(this); }
            finally { InclusaoAberta = null; f.Dispose(); }
            if (novo == null) return;

            // entra logo depois da linha selecionada; sem seleção, na ordem da data
            int pos = ref0 != null ? extrato.IndexOf(ref0) + 1 : -1;
            if (pos <= 0)
            {
                var dn = J.Data(novo.data);
                pos = extrato.FindLastIndex(l => { var d = J.Data(l.data); return d != null && dn != null && d.Value <= dn.Value; }) + 1;
            }
            extrato.Insert(Math.Max(0, Math.Min(pos, extrato.Count)), novo);
            for (int i = 0; i < extrato.Count; i++) extrato[i].seq = i + 1;
            sequencia = extrato.Count;
            Revincular();
            int ix = vis.FindIndex(i => i < extrato.Count && extrato[i] == novo);
            if (ix < 0 && (cbFiltro.SelectedIndex != 0 || cbSentido.SelectedIndex != 0 || txtPesquisa.Text.Length > 0))
            {
                Dialogos.Info(this, "Lançamento inserido", "Lançamento inserido, mas fora dos filtros/pesquisa aplicados na tela.");
            }
            else if (ix >= 0)
            {
                grade.ClearSelection();
                grade.CurrentCell = grade.Rows[ix].Cells[1];
                grade.Rows[ix].Selected = true;
            }
            status.Text = status.Text + "   |   Inserido: " + novo.data + " " + novo.historico + " " + Moeda(novo.valor);
        }
    }
}

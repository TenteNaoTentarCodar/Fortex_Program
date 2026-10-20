using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Alteração de contas em lote: escolhe os lançamentos (selecionados, por texto no histórico/favorecido, sem banco/caixa
    /// ou todos da tela), confere na prévia e aplica a nova conta como contrapartida, débito/crédito ou banco/caixa.
    /// Opcionalmente grava como regra do cliente.
    /// </summary>
    public partial class TelaPrincipal
    {
        /// <summary>Lado do banco/caixa vazio: saída sem conta no Crédito, entrada sem conta no Débito.</summary>
        static bool SemBancoCaixa(Lancamento l) { return (l.saida ? l.credito : l.debito) == null; }

        int?[] NovasContas(Lancamento l, int modo, int?[] novo)
        {
            int? d = l.debito, c = l.credito;
            if (modo == 2)
            {
                if (novo[3] != null) { if (l.saida) c = novo[3]; else d = novo[3]; }
            }
            else if (modo == 0)
            {
                if (novo[0] != null)
                {
                    int? banco = BancoDe(l);
                    if (l.saida) { d = novo[0]; if (c == null) c = banco; }
                    else { c = novo[0]; if (d == null) d = banco; }
                }
            }
            else
            {
                if (novo[1] != null) d = novo[1];
                if (novo[2] != null) c = novo[2];
            }
            return new[] { d, c };
        }

        void AlterarEmLote()
        {
            if (extrato.Count == 0) { Dialogos.Aviso(this, "Adicione o extrato primeiro."); return; }
            if (!SincronizarCampos(true)) return;
            abas.SelectedIndex = 0;
            var selecionados = Selecionados().Where(l => !l.descartada).ToList();
            var naTela = vis.Where(i => i < extrato.Count).Select(i => extrato[i]).Where(l => !l.descartada).ToList();
            var semBanco = extrato.Where(l => !l.descartada && SemBancoCaixa(l)).ToList();

            var f = new Form
            {
                Text = "Alterar contas em lote", Font = Dialogos.Fonte, Size = new Size(1020, 800), StartPosition = FormStartPosition.CenterParent,
                ShowInTaskbar = false, MinimizeBox = false, MinimumSize = new Size(760, 560)
            };
            // ---------- 1. quais ----------
            var g1 = new GroupBox { Text = "1. Quais lançamentos", Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 4, 8, 4) };
            var t1 = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 2 };
            var rbSel = new RadioButton { Text = "Linhas selecionadas na tabela (" + selecionados.Count + ")", AutoSize = true, Enabled = selecionados.Count > 0 };
            var rbTexto = new RadioButton { Text = "Todas as linhas cujo histórico (ou favorecido) contém:", AutoSize = true };
            var fTexto = new TextBox { Name = "fTexto", Width = 320, Text = selecionados.Count == 0 ? "" : SugestaoTexto(selecionados[0]) };
            var rbSemBanco = new RadioButton { Text = "Todos os lançamentos SEM BANCO/CAIXA identificado (" + semBanco.Count + ")", AutoSize = true, Enabled = semBanco.Count > 0 };
            var rbTela = new RadioButton { Text = "Todos os lançamentos da tela (" + naTela.Count + (naTela.Count < extrato.Count ? ", com os filtros aplicados" : "") + ")", AutoSize = true, Enabled = naTela.Count > 0 };
            dicas.SetToolTip(rbSemBanco, "Saídas sem conta no Crédito e entradas sem conta no Débito (o lado do banco/caixa está vazio).");
            dicas.SetToolTip(rbTela, "Todas as linhas que aparecem agora na tabela do Extrato, respeitando a pesquisa e os filtros.");
            // com linha(s) selecionada(s) na tabela, começa nelas (antes, 1 linha selecionada abria por texto e pegava todas as parecidas)
            if (selecionados.Count > 0) rbSel.Checked = true; else rbTexto.Checked = true;
            t1.Controls.Add(rbSel, 0, 0); t1.SetColumnSpan(rbSel, 2);
            t1.Controls.Add(rbTexto, 0, 1); t1.Controls.Add(fTexto, 1, 1);
            t1.Controls.Add(rbSemBanco, 0, 2); t1.SetColumnSpan(rbSemBanco, 2);
            t1.Controls.Add(rbTela, 0, 3); t1.SetColumnSpan(rbTela, 2);
            var filt = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            var cbS = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 150 };
            cbS.Items.AddRange(new object[] { "Saídas e entradas", "Só saídas (D)", "Só entradas (C)" });
            cbS.SelectedIndex = selecionados.Count == 1 ? (selecionados[0].saida ? 1 : 2) : 0;
            var chk685 = new CheckBox { Text = "Somente os que estão em Valores a Identificar (" + regras.contaVAI + ")", AutoSize = true, Padding = new Padding(8, 3, 0, 0) };
            filt.Controls.Add(Rot("Sentido:")); filt.Controls.Add(cbS); filt.Controls.Add(chk685);
            t1.Controls.Add(filt, 0, 4); t1.SetColumnSpan(filt, 2);
            // período (como no e-CAC): intervalo ou apenas uma data
            var periodo = new FiltroPeriodo(dicas);
            t1.Controls.Add(periodo.Painel, 0, 5); t1.SetColumnSpan(periodo.Painel, 2);
            g1.Controls.Add(t1);

            // ---------- 2. nova conta ----------
            var g2 = new GroupBox { Text = "2. Nova conta", Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 4, 8, 4) };
            var t2 = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 3 };
            var rbContra = new RadioButton { Text = "Contrapartida (vai no Débito das saídas e no Crédito das entradas; o banco não muda)", AutoSize = true, Checked = true };
            var rbDC = new RadioButton { Text = "Débito e/ou Crédito (deixe em branco o que não muda)", AutoSize = true };
            var rbBanco = new RadioButton { Text = "Banco/caixa (vai no Crédito das saídas e no Débito das entradas; a contrapartida não muda)", AutoSize = true };
            TextBox fContra = new TextBox { Name = "fContra", Width = 80 }, fDeb = new TextBox { Name = "fDeb", Width = 80 }, fCred = new TextBox { Name = "fCred", Width = 80 }, fBanco = new TextBox { Name = "fBanco", Width = 80 };
            Label nContra = new Label { AutoSize = true, Padding = new Padding(0, 6, 0, 0) }, nDeb = new Label { AutoSize = true, Padding = new Padding(0, 6, 0, 0) },
                  nCred = new Label { AutoSize = true, Padding = new Padding(0, 6, 0, 0) }, nBanco = new Label { AutoSize = true, Padding = new Padding(0, 6, 0, 0) };
            var chkRegra = new CheckBox { Text = "Gravar também como regra do cliente (os próximos extratos já vêm classificados)", AutoSize = true };
            int lin = 0;
            Action<Control> span = c => { t2.Controls.Add(c, 0, lin); t2.SetColumnSpan(c, 3); lin++; };
            Action<string, TextBox, Label> campo = (r, tb, nl) =>
            {
                t2.Controls.Add(new Label { Text = "      " + r, AutoSize = true, Padding = new Padding(0, 6, 0, 0) }, 0, lin);
                t2.Controls.Add(CampoConta.ComBotao(tb, () => plano), 1, lin);
                t2.Controls.Add(nl, 2, lin);
                lin++;
            };
            span(rbContra); campo("Contrapartida:", fContra, nContra);
            span(rbDC); campo("Débito:", fDeb, nDeb); campo("Crédito:", fCred, nCred);
            span(rbBanco); campo("Banco/caixa:", fBanco, nBanco);
            span(new Label { Text = "      Dica: com o cursor no campo, pressione F2 (ou clique em \"F2\") para buscar a conta no Plano de Contas.", AutoSize = true, ForeColor = Color.FromArgb(0, 90, 160) });
            span(chkRegra);
            g2.Controls.Add(t2);

            // ---------- 3. prévia ----------
            var g3 = new GroupBox { Text = "3. Prévia (desmarque o que não deve mudar)", Dock = DockStyle.Fill, Padding = new Padding(8) };
            var tp = new DataGridView
            {
                Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false, BackgroundColor = SystemColors.Window,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect, AllowUserToResizeRows = false, EditMode = DataGridViewEditMode.EditProgrammatically,
                ReadOnly = true, MultiSelect = true
            };
            // a marcação fica numa lista própria (não no estado de edição da célula), para valer exatamente o que aparece
            tp.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Alterar", Width = 55, ReadOnly = true });
            string[] cols = { "Data", "Histórico", "Valor", "D/C", "Débito atual", "Crédito atual", "Novo Débito", "Novo Crédito" };
            int[] lw = { 78, 360, 90, 40, 80, 80, 85, 85 };
            for (int i = 0; i < cols.Length; i++) tp.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = cols[i], Width = lw[i], ReadOnly = true });
            tp.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            var lblQtd = new Label { Dock = DockStyle.Bottom, Height = 22, Padding = new Padding(2, 5, 2, 0) };
            // filtro da prévia e marcar/desmarcar
            var barra = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true, Padding = new Padding(0, 0, 0, 4) };
            var fFiltro = new TextBox { Name = "fFiltroPrevia", Width = 260 };
            var bLimpar = new Button { Text = "Limpar filtro", AutoSize = true };
            var bMarcar = new Button { Name = "bMarcarTodos", Text = "Marcar todos", AutoSize = true };
            var bDesmarcar = new Button { Name = "bDesmarcarTodos", Text = "Desmarcar todos", AutoSize = true };
            var bInverter = new Button { Text = "Inverter marcação", AutoSize = true };
            var bDestacadas = new Button { Name = "bSoDestacadas", Text = "Marcar só as destacadas", AutoSize = true };
            barra.Controls.Add(Rot("Filtrar a prévia:")); barra.Controls.Add(fFiltro); barra.Controls.Add(bLimpar);
            barra.Controls.Add(new Label { Text = "   ", AutoSize = true });
            barra.Controls.Add(bMarcar); barra.Controls.Add(bDesmarcar); barra.Controls.Add(bInverter); barra.Controls.Add(bDestacadas);
            dicas.SetToolTip(bDestacadas, "Destaque na prévia as linhas desejadas (clique, Ctrl+clique ou Shift+clique) e clique aqui: só elas ficam marcadas. Barra de espaço marca/desmarca as destacadas.");
            dicas.SetToolTip(fFiltro, "Mostra só as linhas da prévia que contêm o texto (data, histórico, valor ou conta). Vários termos separados por espaço: todos precisam aparecer.");
            foreach (var b in new[] { bMarcar, bDesmarcar, bInverter })
                dicas.SetToolTip(b, "Vale para as linhas visíveis na prévia (com o filtro, só as filtradas). Clicar no cabeçalho \"Alterar\" também marca/desmarca todas.");
            g3.Controls.Add(tp); g3.Controls.Add(barra); g3.Controls.Add(lblQtd);

            var bts = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(6) };
            var bOk = new Button { Text = "Aplicar", MinimumSize = new Size(100, 30), AutoSize = true };
            var bCa = new Button { Text = "Fechar", MinimumSize = new Size(100, 30), AutoSize = true };
            var lblFeito = new Label { Name = "lblFeito", AutoSize = true, Padding = new Padding(0, 8, 12, 0), ForeColor = Color.FromArgb(0, 110, 0), Font = new Font(Dialogos.Fonte, FontStyle.Bold) };
            bts.Controls.Add(bCa); bts.Controls.Add(bOk); bts.Controls.Add(lblFeito);
            var lay = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(6) };
            lay.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            lay.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            lay.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            lay.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            lay.Controls.Add(g1, 0, 0); lay.Controls.Add(g2, 0, 1); lay.Controls.Add(g3, 0, 2); lay.Controls.Add(bts, 0, 3);
            bts.Dock = DockStyle.Fill;
            f.Controls.Add(lay);
            f.CancelButton = bCa;

            Func<int> modo = () => rbContra.Checked ? 0 : rbDC.Checked ? 1 : 2;
            var alvo = new List<Lancamento>();
            Func<int?[]> lerContas = () =>
            {
                var r = new int?[4];
                var fs = new[] { fContra, fDeb, fCred, fBanco };
                var ls = new[] { nContra, nDeb, nCred, nBanco };
                int m = modo();
                for (int i = 0; i < 4; i++)
                {
                    fs[i].Enabled = m == 0 ? i == 0 : m == 1 ? (i == 1 || i == 2) : i == 3;
                    string t = fs[i].Text.Trim();
                    if (t.Length == 0) { ls[i].Text = ""; continue; }
                    int v;
                    if (!int.TryParse(t, out v)) { ls[i].Text = "⚠ código inválido"; ls[i].ForeColor = Color.FromArgb(180, 0, 0); continue; }
                    r[i] = v;
                    if (plano == null) { ls[i].Text = ""; continue; }
                    var c = plano.Get(v);
                    ls[i].Text = c == null ? "⚠ não existe no Plano" : c.Analitica() ? c.nome : "⚠ " + c.nome + " é sintética";
                    ls[i].ForeColor = c != null && c.Analitica() ? Color.FromArgb(0, 110, 0) : Color.FromArgb(180, 0, 0);
                }
                return r;
            };
            var desmarcados = new HashSet<Lancamento>();
            Func<int, bool> marc = i => i >= 0 && i < alvo.Count && !desmarcados.Contains(alvo[i]);
            // só altera o que está marcado E visível na prévia (o filtro da prévia esconde = não altera)
            Func<List<Lancamento>> paraAplicar = () =>
            {
                var r = new List<Lancamento>();
                for (int i = 0; i < alvo.Count && i < tp.Rows.Count; i++) if (tp.Rows[i].Visible && marc(i)) r.Add(alvo[i]);
                return r;
            };
            Action contar = () =>
            {
                int n = paraAplicar().Count;
                int vis = tp.Rows.Cast<DataGridViewRow>().Count(r => r.Visible);
                lblQtd.Text = n + " lançamento(s) marcado(s) para alterar" + (vis < tp.Rows.Count ? " (das " + vis + " linhas visíveis no filtro; as ocultas não são alteradas)" : " de " + tp.Rows.Count);
            };
            Action<int, bool> definir = (i, v) =>
            {
                if (i < 0 || i >= alvo.Count) return;
                if (v) desmarcados.Remove(alvo[i]); else desmarcados.Add(alvo[i]);
                tp.Rows[i].Cells[0].Value = v;
            };
            Action filtrar = () =>
            {
                var termos = Regras.Normalizar(fFiltro.Text).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                tp.CurrentCell = null;
                foreach (DataGridViewRow r in tp.Rows)
                {
                    var sb = new System.Text.StringBuilder();
                    for (int c = 1; c < tp.Columns.Count; c++) sb.Append(' ').Append(Convert.ToString(r.Cells[c].Value));
                    string tx = Regras.Normalizar(sb.ToString());
                    tx = tx + " " + tx.Replace(".", "");          // valor com ou sem ponto de milhar
                    r.Visible = termos.All(tx.Contains);
                }
                contar();
            };
            Action<Func<bool, bool>> marcar = fn =>
            {
                for (int i = 0; i < tp.Rows.Count; i++) if (tp.Rows[i].Visible) definir(i, fn(marc(i)));
                contar();
            };
            Action previa = () =>
            {
                var novo = lerContas();
                int sentido = cbS.SelectedIndex;
                string txt = Regras.Normalizar(fTexto.Text).Trim();
                var origem = rbSel.Checked ? selecionados : rbSemBanco.Checked ? semBanco : rbTela.Checked ? naTela : extrato;
                var baseL = new List<Lancamento>();
                foreach (var l in origem)
                {
                    if (l.descartada) continue;
                    if (sentido == 1 && !l.saida || sentido == 2 && l.saida) continue;
                    if (chk685.Checked && !EmVAI(l)) continue;
                    if (rbSemBanco.Checked && !SemBancoCaixa(l)) continue;
                    if (!periodo.Contem(l.data)) continue;
                    if (rbTexto.Checked)
                    {
                        if (txt.Length < 2) continue;
                        string a = Regras.Normalizar(l.historico + " " + (l.comprovante == null ? "" : l.comprovante.favorecido) + " " + string.Join(" ", l.planilha.Select(it => it.favorecido)));
                        if (!a.Contains(txt)) continue;
                    }
                    baseL.Add(l);
                }
                tp.CurrentCell = null;
                alvo.Clear(); alvo.AddRange(baseL);
                tp.Rows.Clear();
                foreach (var l in baseL)
                {
                    var dc = NovasContas(l, modo(), novo);
                    int r = tp.Rows.Add(!desmarcados.Contains(l), l.data, l.historico, Moeda(l.valor), l.saida ? "D" : "C",
                        l.debito == null ? "" : l.debito.ToString(), l.credito == null ? "" : l.credito.ToString(),
                        dc[0] == null ? "" : dc[0].ToString(), dc[1] == null ? "" : dc[1].ToString());
                    if (dc[0] != l.debito) tp.Rows[r].Cells[7].Style.ForeColor = Color.FromArgb(0, 90, 160);
                    if (dc[1] != l.credito) tp.Rows[r].Cells[8].Style.ForeColor = Color.FromArgb(0, 90, 160);
                }
                tp.CurrentCell = null;
                filtrar();
                chkRegra.Enabled = rbTexto.Checked && rbContra.Checked;
                if (!chkRegra.Enabled) chkRegra.Checked = false;
            };
            foreach (var tb in new[] { fTexto, fContra, fDeb, fCred, fBanco }) tb.TextChanged += (s, e) => previa();
            foreach (var rb in new[] { rbSel, rbTexto, rbSemBanco, rbTela, rbContra, rbDC, rbBanco }) rb.CheckedChanged += (s, e) => { if (((RadioButton)s).Checked) previa(); };
            rbSemBanco.CheckedChanged += (s, e) => { if (rbSemBanco.Checked) rbBanco.Checked = true; };
            chk685.CheckedChanged += (s, e) => previa();
            periodo.Mudou += () => previa();
            cbS.SelectedIndexChanged += (s, e) => previa();
            // clique na caixa "Alterar" marca/desmarca a linha; barra de espaço vale para as linhas destacadas
            tp.CellMouseClick += (s, e) =>
            {
                if (e.Button != MouseButtons.Left || e.RowIndex < 0 || e.ColumnIndex != 0 || (ModifierKeys & (Keys.Control | Keys.Shift)) != 0) return;
                definir(e.RowIndex, !marc(e.RowIndex));
                contar();
            };
            tp.KeyDown += (s, e) =>
            {
                if (e.KeyCode != Keys.Space) return;
                e.Handled = true; e.SuppressKeyPress = true;
                var sel = tp.SelectedRows.Cast<DataGridViewRow>().Where(r => r.Visible).Select(r => r.Index).ToList();
                if (sel.Count == 0) return;
                bool v = !sel.All(marc);
                foreach (int i in sel) definir(i, v);
                contar();
            };
            bDestacadas.Click += (s, e) =>
            {
                var sel = new HashSet<int>(tp.SelectedRows.Cast<DataGridViewRow>().Where(r => r.Visible).Select(r => r.Index));
                if (sel.Count == 0) { Dialogos.Aviso(f, "Destaque na prévia as linhas que devem ser alteradas (clique, Ctrl+clique ou Shift+clique)."); return; }
                for (int i = 0; i < tp.Rows.Count; i++) definir(i, sel.Contains(i));
                contar();
            };
            fFiltro.TextChanged += (s, e) => filtrar();
            fFiltro.KeyDown += (s, e) => { if (e.KeyCode == Keys.Escape && fFiltro.Text.Length > 0) { fFiltro.Text = ""; e.Handled = true; e.SuppressKeyPress = true; } };
            bLimpar.Click += (s, e) => fFiltro.Text = "";
            bMarcar.Click += (s, e) => marcar(m => true);
            bDesmarcar.Click += (s, e) => marcar(m => false);
            bInverter.Click += (s, e) => marcar(m => !m);
            tp.ColumnHeaderMouseClick += (s, e) =>
            {
                if (e.ColumnIndex != 0) return;
                bool todos = tp.Rows.Cast<DataGridViewRow>().Where(r => r.Visible).All(r => marc(r.Index));
                marcar(m => !todos);
            };
            bCa.Click += (s, e) => f.Close();
            int total = 0;
            // ao mexer nos filtros/contas, some o aviso da última aplicação
            foreach (var tb in new[] { fTexto, fContra, fDeb, fCred, fBanco }) tb.TextChanged += (s, e) => lblFeito.Text = "";
            bOk.Click += (s, e) =>
            {
                var novo = lerContas();
                string erro = ValidarLote(modo(), fDeb, fCred, novo);
                if (erro != null) { Dialogos.Aviso(f, erro); return; }
                var aplicar = paraAplicar();
                if (aplicar.Count == 0) { Dialogos.Aviso(f, "Nenhum lançamento marcado."); return; }
                string conta = modo() == 0 ? "contrapartida " + novo[0] : modo() == 2 ? "banco/caixa " + novo[3]
                             : (novo[1] != null ? "Débito " + novo[1] : "") + (novo[1] != null && novo[2] != null ? " e " : "") + (novo[2] != null ? "Crédito " + novo[2] : "");
                var lista = string.Join("\r\n", aplicar.Take(8).Select(l => "   " + l.data + "  " + (l.historico.Length > 60 ? l.historico.Substring(0, 60) + "…" : l.historico) + "  " + (l.saida ? "D " : "C ") + Moeda(l.valor)));
                if (aplicar.Count > 8) lista += "\r\n   … e mais " + (aplicar.Count - 8);
                if (Dialogos.Escolher(f, "Aplicar alteração em lote", "Aplicar " + conta + " em " + aplicar.Count + " lançamento(s)?\r\n\r\n" + lista, 0, "Aplicar", "Cancelar") != 0) return;
                foreach (var l in aplicar)
                {
                    var dc = NovasContas(l, modo(), novo);
                    l.debito = dc[0]; l.credito = dc[1];
                    l.manual = true;
                    l.conferirSinal = false;
                }
                string extra = "";
                if (chkRegra.Checked && chkRegra.Enabled && novo[0] != null)
                {
                    string tipo = cbS.SelectedIndex == 1 ? "S" : cbS.SelectedIndex == 2 ? "E" : "A";
                    regras.lista.Add(new Regras.Regra(fTexto.Text.Trim(), novo[0].Value, tipo));
                    if (SalvarRegras()) extra = " | regra gravada: \"" + fTexto.Text.Trim() + "\" → " + novo[0];
                }
                // aplica e continua na tela: reclassifica, atualiza a tabela principal e a prévia
                string feito = aplicar.Count + " lançamento(s)" + extra;
                total += aplicar.Count;
                if (extra.Length > 0) chkRegra.Checked = false;   // não grava a mesma regra de novo
                foreach (var l in aplicar) desmarcados.Add(l);    // já alterados: ficam desmarcados para a próxima aplicação
                Revincular();
                status.Text = status.Text + "   |   Alteração em lote: " + feito;
                previa();
                lblFeito.Text = "✔ Aplicado: " + feito + (total > aplicar.Count ? "  (total nesta tela: " + total + ")" : "");
            };
            f.Shown += (s, e) => { previa(); (rbTexto.Checked ? (Control)fTexto : fContra).Focus(); };
            LoteAberto = f;
            try { f.ShowDialog(this); }
            finally { LoteAberto = null; f.Dispose(); }
        }

        static bool Marcado(DataGridViewCell c)
        {
            object v = c.Value;
            if (v is bool) return (bool)v;
            if (v is CheckState) return (CheckState)v == CheckState.Checked;
            bool b;
            return v != null && bool.TryParse(v.ToString(), out b) && b;
        }

        /// <summary>Janela de lote aberta (usada pelo teste automático).</summary>
        Form LoteAberto;

        string ValidarLote(int modo, TextBox fDeb, TextBox fCred, int?[] novo)
        {
            if (modo == 0 || modo == 2)
            {
                int? c = modo == 0 ? novo[0] : novo[3];
                if (c == null) return modo == 0 ? "Informe o código da conta de contrapartida." : "Informe o código da conta do banco/caixa.";
                if (plano != null && !plano.Valida(c)) return "A conta " + c + " não existe no Plano de Contas ou é sintética.";
            }
            else
            {
                if (novo[1] == null && novo[2] == null) return "Informe a conta de Débito e/ou de Crédito.";
                if (fDeb.Text.Trim().Length > 0 && novo[1] == null || fCred.Text.Trim().Length > 0 && novo[2] == null) return "Código de conta inválido.";
                foreach (var c in new[] { novo[1], novo[2] })
                    if (c != null && plano != null && !plano.Valida(c)) return "A conta " + c + " não existe no Plano de Contas ou é sintética.";
            }
            return null;
        }
    }
}

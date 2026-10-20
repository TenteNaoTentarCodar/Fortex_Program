using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExtratoComprovantes
{
    /// <summary>F2: busca de conta no Plano de Contas (código, classificação ou parte do nome).</summary>
    public class BuscaConta : Form
    {
        readonly List<Plano.Conta> todas, vis = new List<Plano.Conta>();
        readonly TextBox fBusca = new TextBox { Width = 320 };
        readonly CheckBox chkSint = new CheckBox { Text = "Mostrar também contas sintéticas", AutoSize = true };
        readonly DataGridView t = new DataGridView();
        readonly Label lblN = new Label { AutoSize = true, Padding = new Padding(0, 8, 0, 0) };
        public int? Escolhida;

        public BuscaConta(Plano plano, string inicial)
        {
            Text = "Plano de Contas — buscar conta (F2)";
            Font = Dialogos.Fonte;
            Size = new Size(860, 540);
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            KeyPreview = true;
            todas = plano.lista.OrderBy(c => c.classificacao, StringComparer.Ordinal).ToList();

            var topo = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(6), WrapContents = true };
            topo.Controls.Add(new Label { Text = "Buscar (código, classificação ou nome):", AutoSize = true, Padding = new Padding(0, 4, 0, 0) });
            topo.Controls.Add(fBusca);
            topo.Controls.Add(chkSint);

            t.Dock = DockStyle.Fill;
            t.ReadOnly = true;
            t.AllowUserToAddRows = false;
            t.AllowUserToDeleteRows = false;
            t.AllowUserToResizeRows = false;
            t.RowHeadersVisible = false;
            t.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            t.MultiSelect = false;
            t.VirtualMode = true;
            t.BackgroundColor = SystemColors.Window;
            t.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            string[] cols = { "Código", "Classificação", "Nome", "Tipo" };
            int[] lw = { 70, 150, 420, 80 };
            for (int i = 0; i < cols.Length; i++)
                t.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = cols[i], FillWeight = lw[i], SortMode = DataGridViewColumnSortMode.NotSortable });
            t.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            t.CellValueNeeded += (s, e) =>
            {
                if (e.RowIndex < 0 || e.RowIndex >= vis.Count) return;
                var k = vis[e.RowIndex];
                e.Value = e.ColumnIndex == 0 ? (object)k.codigo : e.ColumnIndex == 1 ? k.classificacao : e.ColumnIndex == 2 ? k.nome : (k.Analitica() ? "Analítica" : "Sintética");
            };
            t.CellFormatting += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.RowIndex < vis.Count && !vis[e.RowIndex].Analitica()) e.CellStyle.ForeColor = Color.Gray;
            };
            t.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0) EscolherAgora(); };

            var sul = new Panel { Dock = DockStyle.Bottom, Height = 40, Padding = new Padding(6) };
            var bOk = new Button { Text = "Escolher", MinimumSize = new Size(90, 28), AutoSize = true };
            var bCa = new Button { Text = "Cancelar", MinimumSize = new Size(90, 28), AutoSize = true };
            bOk.Click += (s, e) => EscolherAgora();
            bCa.Click += (s, e) => Close();
            var bts = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Right, AutoSize = true };
            bts.Controls.Add(bCa); bts.Controls.Add(bOk);
            bts.Controls.Add(new Label { Text = "Enter ou duplo clique escolhe · Esc fecha", AutoSize = true, Padding = new Padding(0, 8, 8, 0) });
            sul.Controls.Add(lblN);
            sul.Controls.Add(bts);

            Controls.Add(t);
            Controls.Add(topo);
            Controls.Add(sul);

            fBusca.TextChanged += (s, e) => Filtrar();
            chkSint.CheckedChanged += (s, e) => Filtrar();
            fBusca.KeyDown += (s, e) =>
            {
                int r = t.CurrentCell == null ? -1 : t.CurrentCell.RowIndex;
                if (e.KeyCode == Keys.Down && r < vis.Count - 1) { Selecionar(r + 1); e.Handled = true; }
                else if (e.KeyCode == Keys.Up && r > 0) { Selecionar(r - 1); e.Handled = true; }
                else if (e.KeyCode == Keys.Enter) { EscolherAgora(); e.Handled = true; e.SuppressKeyPress = true; }
            };
            t.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { EscolherAgora(); e.Handled = true; } };
            KeyDown += (s, e) => { if (e.KeyCode == Keys.Escape) Close(); };

            fBusca.Text = (inicial ?? "").Trim();
            Filtrar();
            if (vis.Count == 0 && fBusca.Text.Length > 0) { fBusca.Text = ""; Filtrar(); }
            Shown += (s, e) => { fBusca.Focus(); fBusca.SelectAll(); };
        }

        void Selecionar(int r)
        {
            if (r < 0 || r >= vis.Count) return;
            t.CurrentCell = t.Rows[r].Cells[0];
            t.Rows[r].Selected = true;
        }

        void Filtrar()
        {
            vis.Clear();
            string q = Regras.Normalizar(fBusca.Text).Trim();
            var ws = q.Length == 0 ? new string[0] : Regex.Split(q, @"\s+");
            foreach (var c in todas)
            {
                if (!chkSint.Checked && !c.Analitica()) continue;
                string alvo = c.codigo + " " + c.classificacao + " " + Regras.Normalizar(c.nome);
                bool ok = true;
                foreach (var w in ws)
                {
                    if (w.Matches("[0-9]+")) { if (!(c.codigo.ToString().StartsWith(w) || c.classificacao.Replace(".", "").StartsWith(w) || alvo.Contains(w))) { ok = false; break; } }
                    else if (w.Matches("[0-9.]+")) { if (!c.classificacao.StartsWith(w)) { ok = false; break; } }
                    else if (!alvo.Contains(w)) { ok = false; break; }
                }
                if (ok) vis.Add(c);
            }
            t.RowCount = 0;
            t.RowCount = vis.Count;
            t.Invalidate();
            if (vis.Count > 0) Selecionar(0);
            lblN.Text = vis.Count + " conta(s)";
        }

        void EscolherAgora()
        {
            int r = t.CurrentCell == null ? -1 : t.CurrentCell.RowIndex;
            if (r < 0 && vis.Count == 1) r = 0;
            if (r < 0 || r >= vis.Count) return;
            var c = vis[r];
            if (!c.Analitica()) { lblN.Text = "⚠ " + c.nome + " é sintética — escolha uma conta analítica"; return; }
            Escolhida = c.codigo;
            Close();
        }

        /// <summary>Abre a busca e devolve o código escolhido (null = cancelado).</summary>
        public static int? Abrir(IWin32Window dono, Plano plano, string inicial)
        {
            if (plano == null || plano.contas.Count == 0) { Dialogos.Aviso(dono, "Carregue o Plano de Contas do cliente para buscar as contas (F2)."); return null; }
            using (var f = new BuscaConta(plano, inicial)) { f.ShowDialog(dono); return f.Escolhida; }
        }
    }

    /// <summary>Campo de código de conta com F2 (abre a busca no Plano).</summary>
    public static class CampoConta
    {
        public static void InstalarF2(TextBox f, Func<Plano> plano)
        {
            f.KeyDown += (s, e) =>
            {
                if (e.KeyCode != Keys.F2) return;
                e.Handled = true;
                var c = BuscaConta.Abrir(f.FindForm(), plano(), f.Text);
                if (c != null) { f.Text = c.ToString(); f.Focus(); }
            };
        }

        /// <summary>Campo + botão "F2" ao lado.</summary>
        public static Control ComBotao(TextBox f, Func<Plano> plano)
        {
            InstalarF2(f, plano);
            var p = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
            var b = new Button { Text = "F2", Width = 36, Height = f.Height + 2, TabStop = false };
            b.Click += (s, e) => { var c = BuscaConta.Abrir(f.FindForm(), plano(), f.Text); if (c != null) f.Text = c.ToString(); };
            p.Controls.Add(f); p.Controls.Add(b);
            return p;
        }
    }

    /// <summary>Cliente &gt; Editar regras: texto | conta | tipo (A/E/S).</summary>
    public class EditorRegras : Form
    {
        readonly DataGridView t = new DataGridView();
        public List<Regras.Regra> Resultado;

        public EditorRegras(Regras regras, Func<Plano> plano)
        {
            Text = "Regras do cliente";
            Font = Dialogos.Fonte;
            Size = new Size(860, 520);
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            t.Dock = DockStyle.Fill;
            t.AllowUserToAddRows = false;
            t.RowHeadersVisible = false;
            t.BackgroundColor = SystemColors.Window;
            t.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            t.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Texto no histórico / favorecido / CNPJ", FillWeight = 420 });
            t.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Cód. conta (F2 busca)", FillWeight = 110 });
            var tipo = new DataGridViewComboBoxColumn { HeaderText = "Tipo (D/C/A)", FillWeight = 80, FlatStyle = FlatStyle.Flat };
            tipo.Items.AddRange("A", "D", "C");
            t.Columns.Add(tipo);
            // no arquivo continua S/E/A (compatível com a versão Java); na tela, D/C como na coluna D/C do extrato
            foreach (var r in regras.lista) t.Rows.Add(r.texto, r.conta.ToString(), r.tipo == "S" ? "D" : r.tipo == "E" ? "C" : "A");
            t.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.F2 && t.CurrentCell != null && t.CurrentCell.ColumnIndex == 1)
                {
                    e.Handled = true;
                    var c = BuscaConta.Abrir(this, plano(), Convert.ToString(t.CurrentCell.Value));
                    if (c != null) t.CurrentCell.Value = c.ToString();
                }
            };
            t.EditingControlShowing += (s, e) =>
            {
                var tb = e.Control as TextBox;
                if (tb == null) return;
                tb.KeyDown -= F2Celula;
                if (t.CurrentCell.ColumnIndex == 1) tb.KeyDown += F2Celula;
            };
            var sul = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(6), WrapContents = true };
            var add = new Button { Text = "Adicionar", AutoSize = true, MinimumSize = new Size(90, 28) };
            var rem = new Button { Text = "Remover selecionadas", AutoSize = true, MinimumSize = new Size(90, 28) };
            add.Click += (s, e) =>
            {
                int i = t.Rows.Add("", "", "A");
                t.CurrentCell = t.Rows[i].Cells[0];
                t.BeginEdit(true);
            };
            rem.Click += (s, e) =>
            {
                var sel = t.SelectedCells.Cast<DataGridViewCell>().Select(c => c.RowIndex).Concat(t.SelectedRows.Cast<DataGridViewRow>().Select(r => r.Index)).Distinct().OrderByDescending(i => i).ToList();
                foreach (int i in sel) if (i >= 0 && i < t.Rows.Count) t.Rows.RemoveAt(i);
            };
            sul.Controls.Add(add); sul.Controls.Add(rem);
            sul.Controls.Add(new Label { Text = "D = Débito (saídas do banco), C = Crédito (entradas), A = ambos. O texto mais longo que casar vence.", AutoSize = true, Padding = new Padding(8, 7, 0, 0) });
            var bts = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(6) };
            var ok = new Button { Text = "OK", MinimumSize = new Size(90, 28), AutoSize = true };
            var ca = new Button { Text = "Cancelar", MinimumSize = new Size(90, 28), AutoSize = true };
            ok.Click += (s, e) => Confirmar();
            ca.Click += (s, e) => Close();
            bts.Controls.Add(ca); bts.Controls.Add(ok);
            Controls.Add(t);
            Controls.Add(sul);
            Controls.Add(bts);
            CancelButton = ca;
            this.plano = plano;
        }

        readonly Func<Plano> plano;

        void F2Celula(object s, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.F2) return;
            e.Handled = true;
            var tb = (TextBox)s;
            var c = BuscaConta.Abrir(this, plano(), tb.Text);
            if (c != null) { tb.Text = c.ToString(); t.EndEdit(); }
        }

        void Confirmar()
        {
            t.EndEdit();
            var nova = new List<Regras.Regra>();
            foreach (DataGridViewRow r in t.Rows)
            {
                string tx = Convert.ToString(r.Cells[0].Value).Trim(), ct = Convert.ToString(r.Cells[1].Value).Trim();
                string tp = Convert.ToString(r.Cells[2].Value);
                if (tx.Length == 0 || ct.Length == 0) continue;
                int conta;
                if (!int.TryParse(ct, out conta)) { Dialogos.Aviso(this, "Conta inválida na regra \"" + tx + "\": " + ct); return; }
                nova.Add(new Regras.Regra(tx, conta, tp == "D" ? "S" : tp == "C" ? "E" : "A"));
            }
            Resultado = nova;
            Close();
        }
    }

    /// <summary>Extrato &gt; Criar regra da linha.</summary>
    public class NovaRegra : Form
    {
        public readonly TextBox fTexto = new TextBox { Width = 520 };
        public readonly TextBox fConta = new TextBox { Width = 90 };
        public readonly ComboBox fTipo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 330 };
        static readonly string[] TIPOS = { "D - Débito (saídas do banco: a conta vai no Débito)", "C - Crédito (entradas no banco: a conta vai no Crédito)", "A - Ambos" };

        /// <summary>Tipo gravado na regra (mesmo formato da versão Java): S = saídas (D), E = entradas (C), A = ambos.</summary>
        public string Tipo { get { int i = fTipo.SelectedIndex; return i == 0 ? "S" : i == 1 ? "E" : "A"; } }
        public bool Ok;

        public NovaRegra(Lancamento l, string sugestao, int? conta, Func<Plano> plano)
        {
            Text = "Nova regra";
            Font = Dialogos.Fonte;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(12);
            var lay = new TableLayoutPanel { AutoSize = true, ColumnCount = 1, Dock = DockStyle.Fill };
            Action<Control> add = c => { c.Margin = new Padding(0, 2, 0, 2); lay.Controls.Add(c); };
            add(new Label { Text = "Histórico: " + l.historico, AutoSize = true, MaximumSize = new Size(620, 0) });
            if (l.comprovante != null) add(new Label { Text = "Comprovante: " + l.comprovante.favorecido + "  " + l.comprovante.docFavorecido, AutoSize = true, MaximumSize = new Size(620, 0) });
            foreach (var it in l.planilha)
                add(new Label { Text = "Planilha: " + it.Resumo() + "  " + it.documento + (it.categoria.Length == 0 ? "" : "  [" + it.categoria + "]"), AutoSize = true, MaximumSize = new Size(620, 0) });
            add(new Label { Text = "Texto que identifica (nome do fornecedor, CNPJ ou parte do histórico):", AutoSize = true, Margin = new Padding(0, 10, 0, 0) });
            fTexto.Text = sugestao;
            add(fTexto);
            add(new Label { Text = "Cód. conta contrapartida (Plano de Contas) — F2 busca no Plano:", AutoSize = true });
            fConta.Text = conta == null ? "" : conta.ToString();
            add(CampoConta.ComBotao(fConta, plano));
            add(new Label { Text = "Aplicar a (D = Débito, C = Crédito, A = ambos):", AutoSize = true });
            fTipo.Items.AddRange(TIPOS);
            fTipo.SelectedIndex = l.saida ? 0 : 1;
            add(fTipo);
            var bar = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 12, 0, 0) };
            var ok = new Button { Text = "OK", MinimumSize = new Size(90, 28), AutoSize = true };
            var ca = new Button { Text = "Cancelar", MinimumSize = new Size(90, 28), AutoSize = true };
            ok.Click += (s, e) => { Ok = true; Close(); };
            ca.Click += (s, e) => Close();
            bar.Controls.Add(ca); bar.Controls.Add(ok);
            add(bar);
            Controls.Add(lay);
            AcceptButton = ok; CancelButton = ca;
        }
    }

    /// <summary>Cliente &gt; Aprendizado: o que o programa memorizou, com opção de apagar.</summary>
    public class JanelaAprendizado : Form
    {
        public readonly DataGridView t = new DataGridView();
        public bool Apagar;

        public JanelaAprendizado(List<string[]> linhas)
        {
            Text = "Aprendizado do cliente";
            Font = Dialogos.Fonte;
            Size = new Size(1000, 560);
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false; MinimizeBox = false;
            var topo = new Label
            {
                Dock = DockStyle.Top, Height = 44, Padding = new Padding(6),
                Text = "O programa memoriza a conta de cada lançamento exportado (por histórico, CNPJ/CPF e categoria da planilha) e usa nos próximos extratos.\r\n"
                     + "Se algo foi aprendido errado, selecione e clique em \"Apagar selecionados\". " + linhas.Count + " item(ns)."
            };
            t.Dock = DockStyle.Fill;
            t.ReadOnly = true;
            t.AllowUserToAddRows = false;
            t.RowHeadersVisible = false;
            t.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            t.BackgroundColor = SystemColors.Window;
            t.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            string[] cols = { "Tipo", "Sentido", "Palavras / CNPJ / categoria", "Conta", "Nome da conta", "Vezes" };
            int[] lw = { 120, 60, 380, 60, 260, 50 };
            for (int i = 0; i < cols.Length; i++) t.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = cols[i], FillWeight = lw[i] });
            foreach (var l in linhas) { int i = t.Rows.Add(l[0], l[1], l[2], l[3], l[4], l[5]); t.Rows[i].Tag = l; }
            var bts = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(6) };
            var fe = new Button { Text = "Fechar", MinimumSize = new Size(90, 28), AutoSize = true };
            var ap = new Button { Text = "Apagar selecionados", MinimumSize = new Size(90, 28), AutoSize = true };
            fe.Click += (s, e) => Close();
            ap.Click += (s, e) => { Apagar = true; Close(); };
            bts.Controls.Add(fe); bts.Controls.Add(ap);
            Controls.Add(t); Controls.Add(topo); Controls.Add(bts);
            CancelButton = fe;
        }

        public List<string[]> Selecionadas()
        {
            return t.SelectedRows.Cast<DataGridViewRow>().Select(r => (string[])r.Tag).Where(x => x != null).ToList();
        }
    }
}

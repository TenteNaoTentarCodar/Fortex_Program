using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Extrato + Comprovantes + Planilha Financeira + XML → Lançamentos Contábeis (versão em C#): lê o extrato, concilia com
    /// comprovantes, notas fiscais e planilha financeira, classifica pelas regras, Razão e Plano de Contas e exporta o TXT
    /// de importação (layout lanctos.txt). Todos os comandos ficam nos menus do cabeçalho.
    /// </summary>
    public partial class TelaPrincipal : Form
    {
        public const string VERSAO = "C# 2.12";
        const string TITULO = "Extrato + Comprovantes + Planilha Financeira + XML → Lançamentos Contábeis (TXT)";

        // ---------------- dados ----------------
        readonly List<Lancamento> extrato = new List<Lancamento>();
        readonly List<Lancamento> exportados = new List<Lancamento>();
        readonly List<Lancamento> ultimaExportacao = new List<Lancamento>();
        /// <summary>Conta banco de cada arquivo de extrato (null = a conta banco do topo). Mantém a ordem de inclusão.</summary>
        readonly List<KeyValuePair<string, int?>> bancoPorArquivo = new List<KeyValuePair<string, int?>>();
        Regras regras = new Regras();
        string arquivoRegras;
        Plano plano;
        Razao razao;
        readonly Classificador classif = new Classificador();
        readonly List<string> funcionarios = new List<string>();
        string arquivoProgresso;
        bool alterado;
        long sequencia;
        bool carregandoCampos;
        /// <summary>Índices (em extrato) das linhas mostradas na tabela, na ordem.</summary>
        List<int> vis = new List<int>();
        readonly List<Comprovante> comps = new List<Comprovante>();
        /// <summary>Arquivo de origem de cada PDF de comprovante (nome → caminho), para guardar no progresso.</summary>
        readonly Dictionary<string, string> origemComp = new Dictionary<string, string>();
        readonly List<NotaFiscal> notas = new List<NotaFiscal>();
        readonly List<ItemPlanilha> planilha = new List<ItemPlanilha>();
        readonly List<string> arquivosPlanilha = new List<string>();
        List<int> visComp = new List<int>(), visNF = new List<int>();

        // ---------------- tela ----------------
        readonly TextBox txtRegras = new TextBox { ReadOnly = true, Dock = DockStyle.Fill, BackColor = SystemColors.Window };
        readonly TextBox txtCliente = new TextBox { Width = 320 };
        readonly TextBox txtBanco = new TextBox { Width = 70 };
        readonly TextBox txtVAI = new TextBox { Width = 70, Text = "685" };
        readonly Label lblBanco = new Label { AutoSize = true, ForeColor = Color.FromArgb(60, 60, 60), Padding = new Padding(0, 5, 0, 0) };
        readonly TextBox txtBase = new TextBox { ReadOnly = true, Dock = DockStyle.Fill, BackColor = SystemColors.Window };
        readonly TextBox txtPlanilha = new TextBox { ReadOnly = true, Dock = DockStyle.Fill, BackColor = SystemColors.Window };
        readonly TextBox txtCnpjEmp = new TextBox { Width = 150 };
        readonly NumericUpDown spDias = new NumericUpDown { Minimum = 0, Maximum = 15, Value = 3, Width = 50 };
        readonly NumericUpDown spDiasComp = new NumericUpDown { Minimum = 0, Maximum = 15, Value = 3, Width = 50 };
        readonly NumericUpDown spVenc = new NumericUpDown { Minimum = 0, Maximum = 30, Value = 5, Width = 50 };
        readonly NumericUpDown spPrazo = new NumericUpDown { Minimum = 0, Maximum = 365, Value = 90, Increment = 5, Width = 60 };
        readonly TabControl abas = new TabControl { Dock = DockStyle.Fill };
        readonly DataGridView gradeComp = new GradeDupla(), gradeNF = new GradeDupla();
        readonly Label lblNF = new Label { Dock = DockStyle.Top, Height = 22, Padding = new Padding(4, 4, 4, 0), AutoEllipsis = true };
        readonly TextBox txtPesquisa = new TextBox { Width = 300 };
        readonly Label lblPesquisa = new Label { AutoSize = true, ForeColor = Color.FromArgb(0, 90, 160), Padding = new Padding(0, 5, 0, 0) };
        readonly ComboBox cbFiltro = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 210 };
        readonly ComboBox cbSentido = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 130 };
        readonly DataGridView grade = new GradeDupla();
        readonly Label lblSoma = new Label { Dock = DockStyle.Bottom, Height = 22, Padding = new Padding(4, 4, 4, 0), AutoEllipsis = true };
        readonly Label status = new Label { Dock = DockStyle.Bottom, Height = 20, Padding = new Padding(2, 3, 2, 0), AutoEllipsis = true };
        readonly ToolTip dicas = new ToolTip { AutoPopDelay = 20000, InitialDelay = 400 };
        readonly Timer timerPesquisa = new Timer { Interval = 220 };
        readonly Panel rolagem = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        readonly Panel raiz = new Panel();
        readonly ToolStripMenuItem menuRecentes = new ToolStripMenuItem("Abrir recente");
        readonly ToolStripMenuItem chkCont = new ToolStripMenuItem("PDF: juntar linhas de continuação") { CheckOnClick = true };
        readonly ToolStripMenuItem chkExigirNF = new ToolStripMenuItem("Conciliação exige NF em todos os itens") { CheckOnClick = true };
        readonly ToolStripMenuItem chkIncluirVAI = new ToolStripMenuItem("Exportar também os de Valores a Identificar (685)") { CheckOnClick = true };
        readonly ContextMenuStrip menuLinha = new ContextMenuStrip();
        static readonly Size MINIMO = new Size(900, 520);

        static readonly string[] COLUNAS = { "Data", "Histórico", "Valor", "D/C", "Débito", "Crédito", "Comprovante (favorecido)", "Nota fiscal (XML)",
                                             "Planilha financeira", "Categoria / Conta (planilha)", "Conciliação (planilha)", "Observação", "Banco", "Arquivo" };
        static readonly int[] LARGURAS = { 82, 330, 95, 44, 70, 70, 220, 200, 220, 170, 230, 330, 60, 170 };
        static readonly Color VERDE = Color.FromArgb(226, 244, 226), AMARELO = Color.FromArgb(255, 244, 196), ROSA = Color.FromArgb(255, 228, 228),
                              CINZA = Color.FromArgb(232, 232, 232), AZUL = Color.FromArgb(222, 234, 250);

        public TelaPrincipal()
        {
            Text = TITULO;
            Font = Dialogos.Fonte;
            var util = Screen.PrimaryScreen.WorkingArea;
            Size = new Size(Math.Min(1360, util.Width), Math.Min(800, util.Height));
            MinimumSize = new Size(640, 420);
            StartPosition = FormStartPosition.CenterScreen;
            if (util.Width < 1300) WindowState = FormWindowState.Maximized;
            KeyPreview = true;

            chkCont.Checked = Prefs.GetBool("juntarContinuacao", true);
            chkCont.CheckedChanged += (s, e) => Prefs.PutBool("juntarContinuacao", chkCont.Checked);
            chkExigirNF.Checked = Prefs.GetBool("conciliarExigeNF", false);
            chkExigirNF.ToolTipText = "Marcado: item da planilha sem número de NF não é conciliado (fica como divergente). Desmarcado: a NF é conferida quando a planilha informa.";
            chkExigirNF.CheckedChanged += (s, e) => { Prefs.PutBool("conciliarExigeNF", chkExigirNF.Checked); Atualizar(); };
            chkIncluirVAI.Checked = Prefs.GetBool("incluirVAI", true);
            chkIncluirVAI.ToolTipText = "Marcado: exporta tudo, inclusive os lançamentos em Valores a Identificar (685). Desmarcado: os de 685 ficam na tela para corrigir e exportar depois.";
            chkIncluirVAI.CheckedChanged += (s, e) => Prefs.PutBool("incluirVAI", chkIncluirVAI.Checked);

            // ---------- topo ----------
            var topo = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, Padding = new Padding(0, 0, 0, 4) };
            topo.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 175));
            topo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Action<string, Control> linha = (r, c) =>
            {
                topo.Controls.Add(new Label { Text = r, AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(0, 3, 0, 0) });
                c.Margin = new Padding(0, 2, 0, 2);
                topo.Controls.Add(c);
            };
            linha("Regras do cliente:", txtRegras);
            var contas = new FlowLayoutPanel { AutoSize = true, WrapContents = true, Dock = DockStyle.Fill, Margin = new Padding(0) };
            contas.Controls.Add(txtCliente);
            contas.Controls.Add(new Label { Text = "    Cód. conta banco:", AutoSize = true, Padding = new Padding(0, 5, 0, 0) });
            contas.Controls.Add(txtBanco);
            contas.Controls.Add(lblBanco);
            contas.Controls.Add(new Label { Text = "    Cód. Valores a Identificar:", AutoSize = true, Padding = new Padding(0, 5, 0, 0) });
            contas.Controls.Add(txtVAI);
            linha("Cliente:", contas);
            linha("Razão / Plano de Contas:", txtBase);
            linha("Planilha financeira:", txtPlanilha);
            CampoConta.InstalarF2(txtBanco, () => plano);
            dicas.SetToolTip(txtBanco, "Código da conta do banco no Plano de Contas. F2: buscar no Plano.");
            txtBanco.TextChanged += (s, e) => CampoAlterado();
            txtVAI.TextChanged += (s, e) => CampoAlterado();
            txtCliente.TextChanged += (s, e) => { if (!carregandoCampos) regras.cliente = txtCliente.Text.Trim(); };

            // ---------- pesquisa e filtros ----------
            var barPesq = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true, Padding = new Padding(0, 2, 0, 2) };
            barPesq.Controls.Add(new Label { Text = "Pesquisar (CNPJ, data, valor, nome ou conta):", AutoSize = true, Font = new Font(Font, FontStyle.Bold), Padding = new Padding(0, 5, 0, 0) });
            barPesq.Controls.Add(txtPesquisa);
            barPesq.Controls.Add(lblPesquisa);
            dicas.SetToolTip(txtPesquisa, "Digite CNPJ/CPF (com ou sem pontos), data (09/10, 09/10/2026), valor (1.234,56 ou 1234,56), nome,\r\n"
                + "ou a conta contábil: código (ex.: 685), classificação (ex.: 2.1.1.01) ou nome da conta do Plano (ex.: fornecedores).\r\n"
                + "Várias palavras: mostra as linhas que têm todas. Esc limpa.");
            txtPesquisa.TextChanged += (s, e) => { timerPesquisa.Stop(); timerPesquisa.Start(); };
            txtPesquisa.KeyDown += (s, e) => { if (e.KeyCode == Keys.Escape) { txtPesquisa.Text = ""; e.Handled = true; e.SuppressKeyPress = true; } };
            timerPesquisa.Tick += (s, e) => { timerPesquisa.Stop(); Pesquisar(); };

            var barFiltro = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true, Padding = new Padding(0, 0, 0, 2) };
            barFiltro.Controls.Add(new Label { Text = "Mostrar:", AutoSize = true, Padding = new Padding(0, 5, 0, 0) });
            cbFiltro.Items.AddRange(new object[] { "Todos", "Valores a Identificar", "Sem nota fiscal", "Com nota fiscal", "Identificados", "Resgatados (já exportados)",
                                             "Conciliados com a planilha", "Planilha com divergência" });
            cbFiltro.SelectedIndex = 0;
            cbSentido.Items.AddRange(new object[] { "D e C", "Só D (saídas)", "Só C (entradas)" });
            cbSentido.SelectedIndex = 0;
            cbFiltro.SelectedIndexChanged += (s, e) => AplicarFiltro();
            cbSentido.SelectedIndexChanged += (s, e) => AplicarFiltro();
            dicas.SetToolTip(cbSentido, "D = saídas do banco | C = entradas (como na coluna D/C). Combina com \"Mostrar\" e com a pesquisa.");
            barFiltro.Controls.Add(cbFiltro);
            barFiltro.Controls.Add(cbSentido);
            var legenda = new Label { Text = "  Cores: ", AutoSize = true, Padding = new Padding(0, 5, 0, 0) };
            barFiltro.Controls.Add(legenda);
            foreach (var c in new[] { Tuple.Create("identificado", VERDE), Tuple.Create("Valores a Identificar", ROSA), Tuple.Create("conferir D/C", AMARELO),
                                      Tuple.Create("descartada", CINZA), Tuple.Create("já exportado", AZUL) })
                barFiltro.Controls.Add(new Label { Text = c.Item1, AutoSize = true, BackColor = c.Item2, Margin = new Padding(2, 4, 2, 0), Padding = new Padding(4, 2, 4, 2) });

            // ---------- tabela ----------
            ConfigurarGrade();
            var pExt = new TabPage("Extrato") { Padding = new Padding(3) };
            pExt.Controls.Add(grade);
            pExt.Controls.Add(barFiltro);
            pExt.Controls.Add(lblSoma);
            abas.TabPages.Add(pExt);
            abas.TabPages.Add(CriarAbaComprovantes());
            abas.TabPages.Add(CriarAbaNotas());
            abas.SelectedIndexChanged += (s, e) => AtualizarPesquisa();

            var vinc = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, WrapContents = true, Padding = new Padding(0, 2, 0, 0) };
            vinc.Controls.Add(new Label { Text = "Conciliar planilha financeira com até", AutoSize = true, Padding = new Padding(0, 5, 0, 0) });
            vinc.Controls.Add(spDias);
            vinc.Controls.Add(new Label { Text = "dia(s) de diferença", AutoSize = true, Padding = new Padding(0, 5, 0, 0) });
            spDias.ValueChanged += (s, e) => Revincular();

            var rodape = new Label
            {
                Dock = DockStyle.Bottom, Height = 18, ForeColor = Color.FromArgb(90, 90, 90), Padding = new Padding(4, 2, 4, 0), AutoEllipsis = true,
                Text = "Use os menus acima (Arquivo, Cliente, Extrato, Comprovantes, Notas fiscais, Planilha, Exportar) ou arraste os arquivos para a tabela. Botão direito na tabela: ações da linha. Insert: inserir lançamento | Delete: excluir | F2 na célula Débito/Crédito: buscar no Plano."
            };
            lblSoma.Font = new Font(Font, FontStyle.Bold);

            raiz.Padding = new Padding(8, 6, 8, 4);
            raiz.Controls.Add(abas);
            raiz.Controls.Add(barPesq);
            raiz.Controls.Add(topo);
            raiz.Controls.Add(status);
            raiz.Controls.Add(vinc);
            raiz.Controls.Add(rodape);
            rolagem.Controls.Add(raiz);
            Controls.Add(rolagem);
            var menu = CriarMenu();
            MainMenuStrip = menu;
            Controls.Add(menu);
            rolagem.Resize += (s, e) => AjustarRolagem();
            AjustarRolagem();

            AllowDrop = true;
            grade.AllowDrop = true;
            DragEnter += Arrastando;
            grade.DragEnter += Arrastando;
            DragDrop += Soltou;
            grade.DragDrop += Soltou;
            foreach (var g in new[] { gradeComp, gradeNF }) { g.AllowDrop = true; g.DragEnter += Arrastando; g.DragDrop += Soltou; }
            FormClosing += Fechando;

            string ultimo = Prefs.Get("regras", null);
            Shown += (s, e) =>
            {
                if (ultimo != null && File.Exists(ultimo)) { CarregarRegras(ultimo); Revincular(); alterado = false; AtualizarTitulo(); }
                status.Text = "Abra as regras do cliente, informe a conta banco e adicione o extrato (menu Extrato ou arraste o arquivo para a tabela).";
                TesteAuto();
            };
            Atualizar();
        }

        /// <summary>Gancho para o teste automático da tela (só existe na compilação de teste).</summary>
        partial void TesteAuto();

        /// <summary>Barra de rolagem na tela inteira quando a janela fica menor que o mínimo (em vez de espremer).</summary>
        void AjustarRolagem()
        {
            var cs = rolagem.ClientSize;
            raiz.Location = new Point(rolagem.AutoScrollPosition.X, rolagem.AutoScrollPosition.Y);
            raiz.Size = new Size(Math.Max(cs.Width, MINIMO.Width), Math.Max(cs.Height, MINIMO.Height));
            rolagem.AutoScrollMinSize = new Size(cs.Width < MINIMO.Width ? MINIMO.Width : 0, cs.Height < MINIMO.Height ? MINIMO.Height : 0);
        }

        // ===================== tabela =====================

        /// <summary>DataGridView com desenho em buffer duplo (sem piscar ao rolar).</summary>
        class GradeDupla : DataGridView
        {
            public GradeDupla() { DoubleBuffered = true; }
        }

        void ConfigurarGrade()
        {
            var g = grade;
            g.Dock = DockStyle.Fill;
            g.VirtualMode = true;
            g.AllowUserToAddRows = false;
            g.AllowUserToDeleteRows = false;
            g.AllowUserToResizeRows = false;
            g.AllowUserToOrderColumns = false;
            g.RowHeadersVisible = false;
            g.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            g.MultiSelect = true;
            g.EditMode = DataGridViewEditMode.EditOnKeystroke;
            g.BackgroundColor = SystemColors.Window;
            g.BorderStyle = BorderStyle.FixedSingle;
            g.RowTemplate.Height = 22;
            g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            g.ColumnHeadersHeight = 26;
            g.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            g.ShowCellToolTips = true;
            g.DefaultCellStyle.SelectionBackColor = Color.FromArgb(51, 153, 255);
            g.DefaultCellStyle.SelectionForeColor = Color.White;
            var larg = Larguras();
            for (int i = 0; i < COLUNAS.Length; i++)
            {
                DataGridViewColumn c;
                if (i == 3)
                {
                    var cb = new DataGridViewComboBoxColumn { DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing, FlatStyle = FlatStyle.Flat };
                    cb.Items.AddRange("C", "D");
                    c = cb;
                }
                else c = new DataGridViewTextBoxColumn();
                c.HeaderText = COLUNAS[i];
                c.Width = larg[i];
                c.MinimumWidth = 25;
                c.SortMode = DataGridViewColumnSortMode.NotSortable;
                c.ReadOnly = i > 5;
                g.Columns.Add(c);
            }
            g.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            foreach (int i in new[] { 3, 4, 5, 12 }) g.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            g.CellValueNeeded += ValorDaCelula;
            g.CellValuePushed += EditouCelula;
            g.CellFormatting += CorDaCelula;
            g.CellToolTipTextNeeded += (s, e) =>
            {
                var lt = Linha(e.RowIndex);
                if (lt == null || e.ColumnIndex < 0) return;
                var v = Convert.ToString(Valor(lt, e.ColumnIndex));
                e.ToolTipText = v.Length > 0 ? v : null;
            };
            g.SelectionChanged += (s, e) => AtualizarSoma();
            g.ColumnWidthChanged += (s, e) => LembrarLarguras();
            g.DataError += (s, e) => { e.ThrowException = false; e.Cancel = false; };
            g.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.ColumnIndex <= 5) g.BeginEdit(true); };
            g.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Delete && e.Modifiers == Keys.None && !g.IsCurrentCellInEditMode) { e.Handled = true; Executar(ExcluirSelecionados); return; }
                if (e.KeyCode == Keys.Insert && e.Modifiers == Keys.None && !g.IsCurrentCellInEditMode) { e.Handled = true; Executar(InserirLancamento); return; }
                if (e.KeyCode != Keys.F2 || g.CurrentCell == null) return;
                e.Handled = true;
                int c = g.CurrentCell.ColumnIndex;
                if (c == 4 || c == 5) F2NaCelula(g.CurrentCell.RowIndex, c, Convert.ToString(g.CurrentCell.Value));
                else if (c <= 3) g.BeginEdit(true);
            };
            g.EditingControlShowing += (s, e) =>
            {
                var tb = e.Control as TextBox;
                if (tb == null) return;
                tb.KeyDown -= F2NoEditor;
                int c = g.CurrentCell.ColumnIndex;
                if (c == 4 || c == 5) tb.KeyDown += F2NoEditor;
            };
            g.CellMouseDown += (s, e) =>
            {
                if (e.Button != MouseButtons.Right || e.RowIndex < 0) return;
                if (!g.Rows[e.RowIndex].Selected) { g.ClearSelection(); g.Rows[e.RowIndex].Selected = true; }
                if (e.ColumnIndex >= 0) g.CurrentCell = g.Rows[e.RowIndex].Cells[e.ColumnIndex];
            };
            menuLinha.Items.Add("Criar regra da linha", null, (s, e) => CriarRegraDaLinha());
            menuLinha.Items.Add("Inverter D/C", null, (s, e) => InverterDC());
            menuLinha.Items.Add("Alterar contas em lote...", null, (s, e) => Executar(AlterarEmLote));
            menuLinha.Items.Add("Buscar conta no Plano (F2)", null, (s, e) =>
            {
                if (g.CurrentCell == null) return;
                int c = g.CurrentCell.ColumnIndex;
                if (c != 4 && c != 5) { var l = Atual(); if (l == null) return; c = l.saida ? 4 : 5; }
                F2NaCelula(g.CurrentCell.RowIndex, c, Convert.ToString(g.Rows[g.CurrentCell.RowIndex].Cells[c].Value));
            });
            menuLinha.Items.Add(new ToolStripSeparator());
            menuLinha.Items.Add("Inserir lançamento... (Insert)", null, (s, e) => Executar(InserirLancamento));
            menuLinha.Items.Add("Duplicar lançamento... (Ctrl+D)", null, (s, e) => Executar(DuplicarLancamento));
            menuLinha.Items.Add("Excluir selecionados (Delete)", null, (s, e) => Executar(ExcluirSelecionados));
            g.ContextMenuStrip = menuLinha;
        }

        void F2NoEditor(object s, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.F2) return;
            e.Handled = true;
            var tb = (TextBox)s;
            var c = BuscaConta.Abrir(this, plano, tb.Text);
            if (c != null) { tb.Text = c.ToString(); grade.EndEdit(); }
        }

        void F2NaCelula(int linhaVis, int col, string atual)
        {
            var l = Linha(linhaVis);
            if (l == null) return;
            var c = BuscaConta.Abrir(this, plano, atual);
            if (c == null) return;
            DefinirValor(l, col, c.ToString());
        }

        int[] Larguras()
        {
            var s = Prefs.Get("larguraExt", "").Split(',');
            if (s.Length != LARGURAS.Length) return LARGURAS;
            var r = new int[s.Length];
            for (int i = 0; i < s.Length; i++) { int v; if (!int.TryParse(s[i], out v)) return LARGURAS; r[i] = Math.Max(25, v); }
            return r;
        }

        bool montando;
        void LembrarLarguras()
        {
            if (montando) return;
            Prefs.Put("larguraExt", string.Join(",", grade.Columns.Cast<DataGridViewColumn>().Select(c => c.Width.ToString())));
        }

        static string Moeda(double v) { return J.Moeda(v); }

        int? BancoDe(Lancamento l) { return l.contaBanco ?? regras.contaBanco; }

        object Valor(Lancamento l, int c)
        {
            switch (c)
            {
                case 0: return l.data;
                case 1: return l.historico;
                case 2: return Moeda(l.valor);
                case 3: return l.saida ? "D" : "C";
                case 4: return l.debito == null ? "" : l.debito.ToString();
                case 5: return l.credito == null ? "" : l.credito.ToString();
                case 6: { var k = l.comprovante; return k == null ? "" : k.favorecido + (k.docFavorecido.Length == 0 ? "" : "  " + k.docFavorecido) + (k.tipo.Length == 0 ? "" : "  [" + k.tipo + "]"); }
                case 7: return NotaTexto(l);
                case 8: return (l.planilha.Count > 1 ? "[" + l.planilha.Count + " itens] " : "") + Juntar(l.planilha.Select(it => it.Resumo() + (l.planilha.Count > 1 ? " (" + Moeda(it.valor) + ")" : "")));
                case 9: return Juntar(l.planilha.Select(it => it.categoria + (it.conta == null ? "" : (it.categoria.Length == 0 ? "" : " / ") + it.conta)));
                case 10: return l.conciliacao == null ? "" : l.conciliacao.Texto(ExigirNF());
                case 11: return (l.exportadoEm.Length == 0 ? "" : "Já exportado em " + l.exportadoEm + " | ") + (l.manual ? "Editado manualmente" : l.obs);
                case 12: return BancoDe(l) == null ? "" : BancoDe(l).ToString();
                case 13: return l.origem;
            }
            return "";
        }

        void ValorDaCelula(object s, DataGridViewCellValueEventArgs e)
        {
            var l = Linha(e.RowIndex);
            if (l != null) e.Value = Valor(l, e.ColumnIndex);
        }

        void EditouCelula(object s, DataGridViewCellValueEventArgs e)
        {
            var l = Linha(e.RowIndex);
            if (l == null) return;
            string v = Convert.ToString(e.Value);
            int col = e.ColumnIndex;
            // a tabela ainda está gravando a célula: reclassifica logo depois
            BeginInvoke((Action)(() => DefinirValor(l, col, v)));
        }

        /// <summary>Edição de uma célula (como o setValueAt do programa original).</summary>
        void DefinirValor(Lancamento l, int c, string v)
        {
            string s = (v ?? "").Trim();
            if (c == 0) { if (s.Matches(@"[0-9]{2}/[0-9]{2}/[0-9]{4}")) l.data = s; else if (s != l.data) Dialogos.Aviso(this, "Data inválida (use dd/mm/aaaa): " + s); }
            else if (c == 1) { if (s != l.historico) { l.historico = s; l.complementoSalvo = null; } }
            else if (c == 2)
            {
                double? n = Numero(s);
                if (n == null) { Dialogos.Aviso(this, "Valor inválido: " + s); return; }
                l.valor = Math.Abs(n.Value);
            }
            else if (c == 3)
            {
                bool nova = s.Equals("D", StringComparison.OrdinalIgnoreCase) || s.Equals("S", StringComparison.OrdinalIgnoreCase);
                if (nova != l.saida) { l.saida = nova; var d = l.debito; l.debito = l.credito; l.credito = d; }
                l.conferirSinal = false;
                l.obs = Regex.Replace(l.obs, @"^Sinal[^|]*\| ?", "");
            }
            else if (c == 4 || c == 5)
            {
                int? n = null;
                if (s.Length > 0) { int x; if (!int.TryParse(s, out x)) { Dialogos.Aviso(this, "Código de conta inválido: " + s); return; } n = x; }
                if (c == 4) { if (n == l.debito && l.manual) return; l.debito = n; }
                else { if (n == l.credito && l.manual) return; l.credito = n; }
                l.manual = true;
            }
            else return;
            Revincular();
        }

        static double? Numero(string s)
        {
            s = Regex.Replace(s.Replace("R$", ""), @"\s", "");
            if (s.Length == 0) return 0;
            if (s.Contains(",")) s = s.Replace(".", "").Replace(',', '.');
            double v;
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v) ? (double?)v : null;
        }

        void CorDaCelula(object s, DataGridViewCellFormattingEventArgs e)
        {
            var l = Linha(e.RowIndex);
            if (l == null) return;
            Color bg = Color.White, fg = Color.Black;
            if (l.descartada) { bg = CINZA; fg = Color.Gray; }
            else
            {
                if (l.conferirSinal) bg = AMARELO;
                else if (l.exportadoEm.Length > 0 && (l.saida ? l.debito : l.credito) != regras.contaVAI) bg = AZUL;
                else if (l.exportadoEm.Length > 0) bg = ROSA;
                else if (Em685(l)) bg = ROSA;
                else if (l.debito != null && l.credito != null) bg = VERDE;
                if (e.ColumnIndex == 3 && l.saida || e.ColumnIndex == 7 && "Sem nota".Equals(e.Value)) fg = Color.FromArgb(170, 0, 0);
                if (e.ColumnIndex == 10 && l.conciliacao != null) fg = Conciliado(l) ? Color.FromArgb(0, 110, 0) : Color.FromArgb(170, 85, 0);
            }
            e.CellStyle.BackColor = bg;
            e.CellStyle.ForeColor = fg;
        }

        /// <summary>Lançamento da linha r da tabela (null se a lista mudou e a tabela ainda não foi refeita).</summary>
        Lancamento Linha(int r)
        {
            if (r < 0 || r >= vis.Count) return null;
            int i = vis[r];
            return i >= 0 && i < extrato.Count ? extrato[i] : null;
        }

        Lancamento Atual()
        {
            return grade.CurrentCell == null ? null : Linha(grade.CurrentCell.RowIndex);
        }

        List<Lancamento> Selecionados()
        {
            var idx = new SortedSet<int>();
            foreach (DataGridViewRow r in grade.SelectedRows) idx.Add(r.Index);
            foreach (DataGridViewCell c in grade.SelectedCells) idx.Add(c.RowIndex);
            return idx.Select(Linha).Where(x => x != null).ToList();
        }

        // ===================== filtro e pesquisa =====================

        void Pesquisar()
        {
            AplicarFiltro();
            FiltrarComprovantes();
            FiltrarNotas();
            AtualizarPesquisa();
        }

        void AtualizarPesquisa()
        {
            lblPesquisa.Text = txtPesquisa.Text.Trim().Length == 0 ? ""
                : "Encontrados: " + vis.Count + " no extrato | " + visComp.Count + " comprov. | " + visNF.Count + " notas  (Ctrl+A seleciona os encontrados — use em Extrato > Alterar contas em lote)";
        }

        List<string> TermosPesquisa()
        {
            string q = txtPesquisa.Text.Trim();
            var r = new List<string>();
            if (q.Length == 0) return r;
            foreach (var t in Regex.Split(Regex.Replace(q, @"R\$\s*", "", RegexOptions.IgnoreCase), @"\s+")) if (t.Length > 0) r.Add(t);
            return r;
        }

        static readonly Regex TERMO_DATA = new Regex(@"^([0-9]{1,2})[/.-]([0-9]{1,2})(?:[/.-]([0-9]{2,4}))?$");

        static bool CasaTermo(string termo, string texto, string digitos, string data, double[] valores)
        {
            string t = Regras.Normalizar(termo).Trim();
            if (t.Length == 0) return true;
            if (Regras.Normalizar(texto).Contains(t)) return true;
            var md = TERMO_DATA.Match(termo);
            if (md.Success && data != null && data.Length >= 10)
            {
                string dd = int.Parse(md.Groups[1].Value).ToString("00") + "/" + int.Parse(md.Groups[2].Value).ToString("00");
                string aa = !md.Groups[3].Success ? null : (md.Groups[3].Value.Length == 2 ? "20" + md.Groups[3].Value : md.Groups[3].Value);
                if (data.StartsWith(dd, StringComparison.Ordinal) && (aa == null || data.EndsWith(aa, StringComparison.Ordinal))) return true;
            }
            string v = Regex.Replace(termo, @"[^0-9.,\-]", "");
            if (v.Matches(@"-?[0-9]{1,3}(\.[0-9]{3})*(,[0-9]{1,2})?|-?[0-9]+(,[0-9]{1,2})?|-?[0-9]+\.[0-9]{1,2}"))
            {
                double x;
                string n = v.Contains(",") || v.Matches(@"-?[0-9]{1,3}(\.[0-9]{3})+") ? v.Replace(".", "").Replace(',', '.') : v;
                if (double.TryParse(n, NumberStyles.Float, CultureInfo.InvariantCulture, out x))
                    foreach (var val in valores) if (Math.Abs(Math.Abs(val) - Math.Abs(x)) < 0.005) return true;
            }
            string dt = Regex.Replace(termo, @"[^0-9]", "");
            if (dt.Length >= 4 && Regex.Replace(termo, @"[0-9./\-\s]", "").Length == 0 && digitos.Contains(dt)) return true;
            return false;
        }

        bool PesquisaLancamento(Lancamento l, List<string> termos)
        {
            if (termos.Count == 0) return true;
            var tx = new StringBuilder(l.historico).Append(' ').Append(l.obs).Append(' ').Append(l.origem)
                .Append(' ').Append(l.debito).Append(' ').Append(l.credito);
            if (plano != null)
                foreach (var cod in new[] { l.debito, l.credito })
                {
                    var pc = plano.Get(cod);
                    if (pc != null) tx.Append(' ').Append(pc.nome).Append(' ').Append(pc.classificacao);
                }
            var dg = new StringBuilder(Regex.Replace(l.historico, @"[^0-9]", ""));
            if (l.comprovante != null)
            {
                tx.Append(' ').Append(l.comprovante.favorecido).Append(' ').Append(l.comprovante.descricao);
                dg.Append(' ').Append(Regex.Replace(l.comprovante.docFavorecido, "[^0-9]", ""));
            }
            foreach (var it in l.planilha)
            {
                tx.Append(' ').Append(it.favorecido).Append(' ').Append(it.descricao).Append(' ').Append(it.categoria);
                dg.Append(' ').Append(Regex.Replace(it.documento, "[^0-9]", ""));
            }
            foreach (var t in l.notas)
            {
                tx.Append(' ').Append(t.nf.emitNome).Append(' ').Append(t.nf.destNome).Append(" NF ").Append(t.nf.numero);
                dg.Append(' ').Append(t.nf.emitCnpj).Append(' ').Append(t.nf.destCnpj);
            }
            string texto = tx.ToString(), dgs = dg.ToString();
            return termos.All(t => CasaTermo(t, texto, dgs, l.data, new[] { l.valor }));
        }

        /// <summary>Monta a lista de linhas visíveis (Mostrar + D/C + pesquisa) e redesenha a tabela.</summary>
        void AplicarFiltro()
        {
            int f = cbFiltro.SelectedIndex, sent = cbSentido.SelectedIndex;
            var termos = TermosPesquisa();
            var nova = new List<int>();
            for (int i = 0; i < extrato.Count; i++)
            {
                var l = extrato[i];
                if (sent == 1 && !l.saida || sent == 2 && l.saida) continue;
                if (!PesquisaLancamento(l, termos)) continue;
                if (f == 1 && !Em685(l)) continue;
                if (f == 2 && !(l.notas.Count == 0 && !l.descartada)) continue;
                if (f == 3 && l.notas.Count == 0) continue;
                if (f == 4 && !(!l.descartada && l.debito != null && l.credito != null && !Em685(l))) continue;
                if (f == 5 && l.exportadoEm.Length == 0) continue;
                if (f == 6 && !Conciliado(l)) continue;
                if (f == 7 && !(l.conciliacao != null && !Conciliado(l))) continue;
                nova.Add(i);
            }
            // mantém a linha atual e a rolagem quando possível
            Lancamento atual = Atual();
            int topo = grade.RowCount > 0 ? grade.FirstDisplayedScrollingRowIndex : -1;
            if (grade.IsCurrentCellInEditMode) grade.CancelEdit();
            vis = nova;
            montando = true;
            try
            {
                grade.RowCount = 0;
                grade.RowCount = vis.Count;
                if (atual != null)
                {
                    int ix = vis.FindIndex(i => extrato[i] == atual);
                    if (ix >= 0)
                    {
                        int col = grade.CurrentCell != null ? grade.CurrentCell.ColumnIndex : 0;
                        grade.CurrentCell = grade.Rows[ix].Cells[Math.Max(0, col)];
                    }
                }
                if (topo >= 0 && topo < vis.Count) try { grade.FirstDisplayedScrollingRowIndex = topo; } catch { }
            }
            finally { montando = false; }
            grade.Invalidate();
            AtualizarSoma();
        }

        /// <summary>Soma do valor contábil das linhas na tela (com ou sem filtro) e das selecionadas.</summary>
        void AtualizarSoma()
        {
            try
            {
                int tot = extrato.Count;
                double ent = 0, sai = 0;
                int desc = 0;
                foreach (int i in vis)
                {
                    if (i >= extrato.Count) continue;
                    var l = extrato[i];
                    if (l.saida) sai += l.valor; else ent += l.valor;
                    if (l.descartada) desc++;
                }
                bool filtrado = vis.Count != tot;
                var b = new StringBuilder(filtrado ? "Filtrado: " + vis.Count + " de " + tot + " lançamento(s)" : "Todos: " + tot + " lançamento(s)");
                b.Append("  |  Soma do valor: R$ ").Append(Moeda(ent + sai));
                b.Append("  |  Entradas (C): R$ ").Append(Moeda(ent)).Append("  |  Saídas (D): R$ ").Append(Moeda(sai));
                b.Append("  |  Entradas − Saídas: R$ ").Append(Moeda(ent - sai));
                if (desc > 0) b.Append("  (inclui ").Append(desc).Append(" descartada(s))");
                var sel = Selecionados();
                if (sel.Count > 1) b.Append("  |  Selecionadas: ").Append(sel.Count).Append(" = R$ ").Append(Moeda(sel.Sum(x => x.valor)));
                lblSoma.Text = b.ToString();
                lblSoma.ForeColor = filtrado ? Color.FromArgb(0, 80, 200) : SystemColors.ControlText;
            }
            catch (Exception ex) { Log.Erro("soma", ex); }
        }

        // ===================== estado =====================

        bool Em685(Lancamento l)
        {
            if (l.descartada) return false;
            int? contra = l.saida ? l.debito : l.credito;
            return contra != null && contra == regras.contaVAI;
        }

        bool EmVAI(Lancamento l) { return Em685(l) && !l.manual; }

        void AtualizarTitulo()
        {
            Text = (arquivoProgresso == null ? "" : Path.GetFileName(arquivoProgresso) + (alterado ? " *" : "") + " — ")
                   + (string.IsNullOrEmpty(regras.cliente) ? "" : regras.cliente + " — ") + TITULO;
        }

        void Atualizar()
        {
            AtualizarTitulo();
            AplicarFiltro();
            FiltrarComprovantes();
            FiltrarNotas();
            AtualizarPesquisa();
            double ent = 0, sai = 0;
            int conc = 0, conf = 0;
            int div = 0;
            foreach (var l in extrato)
            {
                if (l.saida) sai += l.valor; else ent += l.valor;
                if (Conciliado(l)) conc++; else if (l.conciliacao != null) div++;
                if (l.conferirSinal) conf++;
            }
            int vai = extrato.Count(Em685);
            int pend = planilha.Count(it => it.exportado.Length == 0);
            int naoLoc = planilha.Count(it => it.vinculo == null && it.exportado.Length == 0);
            txtPlanilha.Text = arquivosPlanilha.Count == 0 ? "Nenhuma — anexe a planilha financeira do cliente (opcional, menu Planilha)"
                : string.Join(", ", arquivosPlanilha.Select(Path.GetFileName)) + "  (" + planilha.Count + " itens | vinculados ao extrato: " + (pend - naoLoc) + " | não localizados no extrato: " + naoLoc + ")";
            dicas.SetToolTip(txtPlanilha, txtPlanilha.Text);
            int comNF = extrato.Count(l => l.notas.Count > 0), semNF = extrato.Count(l => l.notas.Count == 0 && !l.descartada);
            int nfAtivas = 0, nfLoc = 0, nfParc = 0, nfNao = 0, nfCanc = 0, nfExp = 0;
            foreach (var nf in notas)
            {
                if (nf.cancelada) { nfCanc++; continue; }
                nfAtivas++;
                int v = nf.titulos.Count(t => t.vinculo != null), e = nf.titulos.Count(t => t.exportado.Length > 0);
                if (e == nf.titulos.Count) nfExp++;
                else if (v + e == nf.titulos.Count) nfLoc++;
                else if (v + e > 0) nfParc++;
                else nfNao++;
            }
            int compVinc = comps.Count(c => c.vinculo != null);
            abas.TabPages[0].Text = "Extrato (" + extrato.Count + ")";
            abas.TabPages[1].Text = "Comprovantes (" + comps.Count + (comps.Count == 0 ? "" : " | vinculados " + compVinc) + ")";
            abas.TabPages[2].Text = "Notas fiscais (XML) (" + notas.Count + ")";
            lblNF.Text = notas.Count == 0 ? "Adicione os XMLs (NF-e, NFC-e, CT-e, NFS-e) de entradas e saídas do período — menu Notas fiscais ou arraste os arquivos/pasta aqui."
                : string.Format("Notas: {0} | Localizadas no extrato: {1} | Parciais: {2} | Não localizadas: {3} | Já exportadas: {4} | Canceladas: {5}   ||   Extrato: {6} com nota | {7} sem nota",
                    nfAtivas, nfLoc, nfParc, nfNao, nfExp, nfCanc, comNF, semNF);
            status.Text = string.Format("Extrato: {0} lançamentos | Entradas R$ {1} | Saídas R$ {2}{3} | Valores a Identificar: {4}",
                              extrato.Count, Moeda(ent), Moeda(sai), conf > 0 ? " | Conferir D/C: " + conf : "", vai)
                          + (planilha.Count == 0 ? "" : "   ||   Conciliados com a planilha: " + conc + (div > 0 ? " | com divergência: " + div : ""))
                          + (notas.Count == 0 ? "" : "   ||   Com NF: " + comNF + " | Sem NF: " + semNF)
                          + (comps.Count == 0 ? "" : "   ||   Comprovantes: " + comps.Count + " | vinculados: " + compVinc)
                          + (exportados.Count > 0 ? "   ||   Exportados guardados no progresso: " + exportados.Count : "");
        }
    }
}

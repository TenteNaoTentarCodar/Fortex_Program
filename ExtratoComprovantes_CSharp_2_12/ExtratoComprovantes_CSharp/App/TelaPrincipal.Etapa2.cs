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
    /// <summary>Etapa 2: comprovantes (PDF), notas fiscais (XML/ZIP) e planilha financeira.</summary>
    public partial class TelaPrincipal
    {
        static readonly string[] COLS_COMP = { "Arquivo", "Pág.", "Tipo", "Data", "Valor", "Favorecido", "CPF/CNPJ", "Pagador", "Cód. barras / ID", "Descrição", "Linha do extrato" };
        static readonly int[] LARG_COMP = { 150, 40, 75, 80, 90, 230, 130, 160, 120, 200, 260 };
        static readonly string[] COLS_NF = { "Tipo", "Número", "Emissão", "Emitente", "CNPJ emitente", "Destinatário / Tomador", "CNPJ destinatário",
                                             "Operação", "Valor", "Parcelas (vencimento)", "Situação no extrato", "Arquivo" };
        static readonly int[] LARG_NF = { 55, 70, 78, 230, 120, 230, 120, 110, 95, 150, 330, 160 };
        int ordemNF = -1;
        bool ordemNFDesc;

        static DataGridView PrepararGrade(DataGridView g, string[] cols, int[] larg)
        {
            g.Dock = DockStyle.Fill;
            g.VirtualMode = true;
            g.AllowUserToAddRows = false;
            g.AllowUserToDeleteRows = false;
            g.AllowUserToResizeRows = false;
            g.RowHeadersVisible = false;
            g.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            g.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            g.BackgroundColor = SystemColors.Window;
            g.BorderStyle = BorderStyle.FixedSingle;
            g.RowTemplate.Height = 22;
            g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            g.ColumnHeadersHeight = 26;
            g.ShowCellToolTips = true;
            g.DefaultCellStyle.SelectionBackColor = Color.FromArgb(51, 153, 255);
            g.DefaultCellStyle.SelectionForeColor = Color.White;
            for (int i = 0; i < cols.Length; i++)
                g.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = cols[i], Width = larg[i], MinimumWidth = 25, SortMode = DataGridViewColumnSortMode.NotSortable });
            g.DataError += (s, e) => { e.ThrowException = false; };
            return g;
        }

        static Label Rot(string t) { return new Label { Text = t, AutoSize = true, Padding = new Padding(0, 5, 0, 0) }; }

        // ===================== aba Comprovantes =====================

        TabPage CriarAbaComprovantes()
        {
            var p = new TabPage("Comprovantes") { Padding = new Padding(3) };
            var bar = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true };
            bar.Controls.Add(Rot("Vincular ao extrato com até"));
            bar.Controls.Add(spDiasComp);
            bar.Controls.Add(Rot("dia(s)   |   duplo clique no arquivo: ver o texto lido   |   células editáveis para correção   |   menu Comprovantes: adicionar, remover"));
            spDiasComp.ValueChanged += (s, e) => Revincular();
            var g = PrepararGrade(gradeComp, COLS_COMP, LARG_COMP);
            g.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            g.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            for (int i = 0; i < COLS_COMP.Length; i++) g.Columns[i].ReadOnly = !(i >= 2 && i <= 7 || i == 9);
            g.CellValueNeeded += (s, e) => { var c = Comp(e.RowIndex); if (c != null) e.Value = ValorComp(c, e.ColumnIndex); };
            g.CellToolTipTextNeeded += (s, e) => { var c = Comp(e.RowIndex); if (c != null && e.ColumnIndex >= 0) { string v = ValorComp(c, e.ColumnIndex); e.ToolTipText = v.Length > 0 ? v : null; } };
            g.CellValuePushed += (s, e) =>
            {
                var c = Comp(e.RowIndex);
                if (c == null) return;
                int col = e.ColumnIndex;
                string v = Convert.ToString(e.Value).Trim();
                BeginInvoke((Action)(() => EditarComprovante(c, col, v)));
            };
            g.CellFormatting += (s, e) =>
            {
                var c = Comp(e.RowIndex);
                if (c == null) return;
                Color bg = Color.White;
                if (c.exportado.Length > 0 || c.vinculo != null) bg = VERDE;
                else if (c.conferir) bg = AMARELO;
                else if (extrato.Count > 0) bg = ROSA;
                e.CellStyle.BackColor = bg;
                e.CellStyle.ForeColor = Color.Black;
            };
            g.CellDoubleClick += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                if (e.ColumnIndex <= 1) VerTexto();
                else if (!g.Columns[e.ColumnIndex].ReadOnly) g.BeginEdit(true);
            };
            var menu = new ContextMenuStrip();
            menu.Items.Add("Ver texto do comprovante", null, (s, e) => Executar(VerTexto));
            menu.Items.Add("Remover selecionados", null, (s, e) => Executar(RemoverComprovantes));
            g.ContextMenuStrip = menu;
            g.CellMouseDown += (s, e) => { if (e.Button == MouseButtons.Right && e.RowIndex >= 0 && !g.Rows[e.RowIndex].Selected) { g.ClearSelection(); g.Rows[e.RowIndex].Selected = true; } };
            p.Controls.Add(g);
            p.Controls.Add(bar);
            return p;
        }

        Comprovante Comp(int r)
        {
            if (r < 0 || r >= visComp.Count) return null;
            int i = visComp[r];
            return i >= 0 && i < comps.Count ? comps[i] : null;
        }

        static string SituacaoComp(Comprovante k)
        {
            return k.exportado.Length > 0 ? "Exportado — " + k.exportado : k.vinculo == null ? "Não localizado no extrato" : k.vinculo.data + " — " + k.vinculo.historico;
        }

        static string ValorComp(Comprovante k, int c)
        {
            switch (c)
            {
                case 0: return k.arquivo;
                case 1: return k.pagina.ToString();
                case 2: return k.tipo;
                case 3: return k.data;
                case 4: return k.valor == 0 ? "" : J.Moeda(k.valor);
                case 5: return k.favorecido;
                case 6: return k.docFavorecido;
                case 7: return k.pagador;
                case 8: return k.codigoBarras.Length > 0 ? k.codigoBarras : k.idTransacao;
                case 9: return k.descricao;
                case 10: return SituacaoComp(k);
            }
            return "";
        }

        void EditarComprovante(Comprovante k, int c, string s)
        {
            switch (c)
            {
                case 2: k.tipo = s; break;
                case 3: if (s.Matches(@"[0-9]{2}/[0-9]{2}/[0-9]{4}")) k.data = s; else { Dialogos.Aviso(this, "Data inválida (use dd/mm/aaaa): " + s); return; } break;
                case 4: { var n = Numero(s); if (n == null) { Dialogos.Aviso(this, "Valor inválido: " + s); return; } k.valor = Math.Abs(n.Value); break; }
                case 5: k.favorecido = s; break;
                case 6: k.docFavorecido = s; break;
                case 7: k.pagador = s; break;
                case 9: k.descricao = s; break;
                default: return;
            }
            k.conferir = k.valor == 0 || k.data.Length == 0 || k.favorecido.Length == 0;
            Revincular();
        }

        bool PesquisaComprovante(Comprovante c, List<string> termos)
        {
            if (termos.Count == 0) return true;
            string tx = c.favorecido + " " + c.pagador + " " + c.tipo + " " + c.descricao + " " + c.arquivo + " " + c.idTransacao;
            string dg = Regex.Replace(c.docFavorecido, "[^0-9]", "") + " " + Regex.Replace(c.docPagador, "[^0-9]", "") + " " + Regex.Replace(c.codigoBarras, "[^0-9]", "");
            return termos.All(t => CasaTermo(t, tx, dg, c.data, new[] { c.valor }));
        }

        void FiltrarComprovantes()
        {
            var termos = TermosPesquisa();
            var nova = new List<int>();
            for (int i = 0; i < comps.Count; i++) if (PesquisaComprovante(comps[i], termos)) nova.Add(i);
            if (gradeComp.IsCurrentCellInEditMode) gradeComp.CancelEdit();
            visComp = nova;
            gradeComp.RowCount = 0;
            gradeComp.RowCount = visComp.Count;
            gradeComp.Invalidate();
        }

        List<Comprovante> CompsSelecionados()
        {
            var idx = new SortedSet<int>();
            foreach (DataGridViewRow r in gradeComp.SelectedRows) idx.Add(r.Index);
            foreach (DataGridViewCell c in gradeComp.SelectedCells) idx.Add(c.RowIndex);
            return idx.Select(Comp).Where(x => x != null).ToList();
        }

        static List<string> ArquivosDaPasta(string dir)
        {
            try { return Directory.GetFiles(dir, "*", SearchOption.AllDirectories).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList(); }
            catch (Exception e) { Log.Erro("pasta " + dir, e); return new List<string>(); }
        }

        void AdicionarComprovantes()
        {
            using (var fc = new OpenFileDialog { Multiselect = true, Filter = "Comprovantes (*.pdf)|*.pdf|Todos os arquivos|*.*", Title = "Comprovantes bancários (PDF) — pode selecionar vários", InitialDirectory = Prefs.Pasta("dirComp", Prefs.Pasta("dirExt", null)) })
            {
                if (fc.ShowDialog(this) != DialogResult.OK) return;
                Prefs.Put("dirComp", Path.GetDirectoryName(fc.FileNames[0]));
                CarregarComprovantes(fc.FileNames.ToList());
            }
        }

        void CarregarComprovantes(List<string> arquivos)
        {
            var avisos = new StringBuilder();
            var duplicados = new StringBuilder();
            int antes = comps.Count;
            foreach (var f in arquivos)
            {
                string nome = Path.GetFileName(f);
                if (!nome.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) continue;
                try
                {
                    var lidos = Dialogos.Pesado(this, "Lendo comprovantes " + nome + "...", () => LeitorComprovante.Ler(f));
                    if (lidos.Count == 0) { avisos.Append("• ").Append(nome).Append(": nenhum comprovante reconhecido.\r\n"); continue; }
                    origemComp[nome] = Path.GetFullPath(f);
                    // mesma DATA de comprovante já anexado → arquivo provavelmente já anexado
                    var repetidas = new SortedDictionary<string, SortedSet<string>>(StringComparer.Ordinal);
                    foreach (var novo in lidos)
                    {
                        if (novo.data.Length == 0) continue;
                        foreach (var c in comps) if (novo.data == c.data) repetidas.ObterOuCriar(novo.data).Add(c.arquivo);
                    }
                    if (repetidas.Count > 0)
                    {
                        var msg = new StringBuilder("O arquivo " + nome + " já foi anexado.\r\n\r\nJá existem comprovantes anexados com a(s) mesma(s) data(s):\r\n");
                        foreach (var e in repetidas) msg.Append("   • ").Append(e.Key).Append("  (em ").Append(string.Join(", ", e.Value)).Append(")\r\n");
                        if (Dialogos.Escolher(this, "Comprovante já anexado", msg.ToString(), 0, "Não incluir", "Incluir mesmo assim") != 1) continue;
                    }
                    foreach (var c in lidos)
                    {
                        var igual = ComprovanteIgual(c);
                        if (igual != null)
                        {
                            duplicados.Append("• ").Append(c.arquivo).Append(" p.").Append(c.pagina).Append(" (").Append(c.favorecido.Length == 0 ? c.tipo : c.favorecido)
                                .Append(", R$ ").Append(Moeda(c.valor)).Append(", ").Append(c.data).Append(") é igual a ").Append(igual.arquivo)
                                .Append(" p.").Append(igual.pagina).Append(igual.exportado.Length == 0 ? "" : " — já exportado").Append("\r\n");
                            continue;
                        }
                        comps.Add(c);
                    }
                }
                catch (Exception ex)
                {
                    if (!(ex is IOException)) Log.Erro("comprovante " + nome, ex);
                    avisos.Append("• ").Append(nome).Append(": ").Append(ex.Message).Append("\r\n");
                }
            }
            abas.SelectedIndex = 1;
            Revincular();
            status.Text = status.Text + "   |   " + (comps.Count - antes) + " comprovante(s) lido(s)";
            if (duplicados.Length > 0) Dialogos.Escolher(this, "Comprovante já anexado", "Comprovante(s) já anexado(s) — não foram incluídos novamente:\r\n\r\n" + duplicados, 0, "OK");
            if (avisos.Length > 0) Dialogos.Escolher(this, "Avisos", avisos.ToString().TrimEnd(), 0, "OK");
        }

        /// <summary>Mesmo comprovante já anexado: mesmo ID, mesmo código de barras e valor, ou mesma data + valor + favorecido + tipo.</summary>
        Comprovante ComprovanteIgual(Comprovante novo)
        {
            foreach (var c in comps)
            {
                if (novo.idTransacao.Length > 0 && novo.idTransacao == c.idTransacao) return c;
                if (novo.codigoBarras.Length > 0 && novo.codigoBarras == c.codigoBarras && Math.Abs(novo.valor - c.valor) < 0.005 && novo.data == c.data) return c;
                if (novo.favorecido.Length > 0 && novo.data == c.data && Math.Abs(novo.valor - c.valor) < 0.005 && novo.tipo == c.tipo
                    && Regras.Normalizar(novo.favorecido) == Regras.Normalizar(c.favorecido)
                    && Regex.Replace(novo.docFavorecido, "[^0-9]", "") == Regex.Replace(c.docFavorecido, "[^0-9]", "")) return c;
            }
            return null;
        }

        void VerTexto()
        {
            var c = gradeComp.CurrentCell == null ? null : Comp(gradeComp.CurrentCell.RowIndex);
            if (c == null) { Dialogos.Aviso(this, "Selecione um comprovante na aba Comprovantes."); abas.SelectedIndex = 1; return; }
            Dialogos.Texto(this, c.arquivo + " — página " + c.pagina, c.texto, "Fechar");
        }

        void RemoverComprovantes()
        {
            var rem = new HashSet<Comprovante>(CompsSelecionados());
            if (rem.Count == 0) { Dialogos.Aviso(this, "Selecione os comprovantes na aba Comprovantes."); abas.SelectedIndex = 1; return; }
            comps.RemoveAll(rem.Contains);
            Revincular();
        }

        void LimparComprovantes()
        {
            if (comps.Count == 0) return;
            if (!Dialogos.SimNao(this, "Limpar comprovantes", "Tirar os " + comps.Count + " comprovante(s) anexado(s)?")) return;
            comps.Clear();
            origemComp.Clear();
            Revincular();
        }

        // ===================== aba Notas fiscais =====================

        TabPage CriarAbaNotas()
        {
            var p = new TabPage("Notas fiscais (XML)") { Padding = new Padding(3) };
            var bar = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true };
            bar.Controls.Add(Rot("CNPJ da empresa:"));
            bar.Controls.Add(txtCnpjEmp);
            dicas.SetToolTip(txtCnpjEmp, "CNPJ da empresa (cliente). Detectado pelos XMLs; define se a nota é compra (saída no banco) ou venda (entrada).");
            txtCnpjEmp.Leave += (s, e) => Revincular();
            txtCnpjEmp.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Revincular(); } };
            bar.Controls.Add(Rot("   Localizar no extrato: vencimento da duplicata ±"));
            bar.Controls.Add(spVenc);
            bar.Controls.Add(Rot("dias  |  nota sem duplicata: da emissão até"));
            bar.Controls.Add(spPrazo);
            bar.Controls.Add(Rot("dias depois (até 10 dias se o histórico não citar o fornecedor/cliente)"));
            spVenc.ValueChanged += (s, e) => Revincular();
            spPrazo.ValueChanged += (s, e) => Revincular();
            var g = PrepararGrade(gradeNF, COLS_NF, LARG_NF);
            g.ReadOnly = true;
            g.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            foreach (DataGridViewColumn c in g.Columns) c.SortMode = DataGridViewColumnSortMode.Programmatic;
            g.ColumnHeaderMouseClick += (s, e) =>
            {
                if (ordemNF == e.ColumnIndex) ordemNFDesc = !ordemNFDesc; else { ordemNF = e.ColumnIndex; ordemNFDesc = false; }
                foreach (DataGridViewColumn c in g.Columns) c.HeaderCell.SortGlyphDirection = SortOrder.None;
                g.Columns[ordemNF].HeaderCell.SortGlyphDirection = ordemNFDesc ? SortOrder.Descending : SortOrder.Ascending;
                FiltrarNotas();
            };
            g.CellValueNeeded += (s, e) => { var nf = Nota(e.RowIndex); if (nf != null) e.Value = ValorNF(nf, e.ColumnIndex); };
            g.CellToolTipTextNeeded += (s, e) => { var nf = Nota(e.RowIndex); if (nf != null && e.ColumnIndex >= 0) { string v = ValorNF(nf, e.ColumnIndex); e.ToolTipText = v.Length > 0 ? v : null; } };
            g.CellFormatting += (s, e) =>
            {
                var nf = Nota(e.RowIndex);
                if (nf == null) return;
                string sit = SituacaoNF(nf);
                bool exp = nf.titulos.Any(x => x.exportado.Length > 0);
                e.CellStyle.BackColor = nf.cancelada ? CINZA
                    : sit.StartsWith("Localizada", StringComparison.Ordinal) || exp && !sit.StartsWith("Parcial", StringComparison.Ordinal) ? VERDE
                    : sit.StartsWith("Parcial", StringComparison.Ordinal) ? AMARELO : ROSA;
                e.CellStyle.ForeColor = nf.cancelada ? Color.Gray : Color.Black;
            };
            var menu = new ContextMenuStrip();
            menu.Items.Add("Remover selecionadas", null, (s, e) => Executar(RemoverNotas));
            g.ContextMenuStrip = menu;
            g.CellMouseDown += (s, e) => { if (e.Button == MouseButtons.Right && e.RowIndex >= 0 && !g.Rows[e.RowIndex].Selected) { g.ClearSelection(); g.Rows[e.RowIndex].Selected = true; } };
            p.Controls.Add(g);
            p.Controls.Add(lblNF);
            p.Controls.Add(bar);
            return p;
        }

        NotaFiscal Nota(int r)
        {
            if (r < 0 || r >= visNF.Count) return null;
            int i = visNF[r];
            return i >= 0 && i < notas.Count ? notas[i] : null;
        }

        string CnpjEmpresa() { return Regex.Replace(txtCnpjEmp.Text, "[^0-9]", ""); }

        static string FormatarDoc(string d)
        {
            if (d == null) return "";
            if (d.Length == 14) return Regex.Replace(d, @"^([0-9]{2})([0-9]{3})([0-9]{3})([0-9]{4})([0-9]{2})$", "$1.$2.$3/$4-$5");
            if (d.Length == 11) return Regex.Replace(d, @"^([0-9]{3})([0-9]{3})([0-9]{3})([0-9]{2})$", "$1.$2.$3-$4");
            return d;
        }

        static string Dt(DateTime? d) { return d == null ? "" : d.Value.ToString("dd/MM/yyyy"); }

        string ValorNF(NotaFiscal nf, int c)
        {
            switch (c)
            {
                case 0: return nf.tipo;
                case 1: return nf.numero;
                case 2: return Dt(nf.emissao);
                case 3: return nf.emitNome;
                case 4: return FormatarDoc(nf.emitCnpj);
                case 5: return nf.destNome;
                case 6: return FormatarDoc(nf.destCnpj);
                case 7: { var cp = nf.Compra(CnpjEmpresa()); return cp == null ? "?" : cp.Value ? "Compra/Tomada" : "Venda/Prestação"; }
                case 8: return Moeda(nf.valor);
                case 9:
                    if (nf.titulos.Count == 1 && nf.titulos[0].vencimento == null) return "À vista / sem duplicata";
                    return nf.titulos.Count + "x: " + string.Join(" | ", nf.titulos.Select(t => Moeda(t.valor) + (t.vencimento == null ? "" : " (" + t.vencimento.Value.ToString("dd/MM") + ")")));
                case 10:
                    {
                        string sit = SituacaoNF(nf);
                        int e = nf.titulos.Count(t => t.exportado.Length > 0);
                        return e > 0 ? sit + " [" + e + " já exportada(s)]" : sit;
                    }
                case 11: return nf.arquivo;
            }
            return "";
        }

        static string SituacaoNF(NotaFiscal nf)
        {
            if (nf.cancelada) return "Cancelada";
            var loc = nf.titulos.Where(t => t.vinculo != null).ToList();
            if (loc.Count == 0) return "Não localizada no extrato";
            string sb = string.Join(" | ", loc.Select(t => t.vinculo.data + " — " + t.vinculo.historico));
            if (loc.Count < nf.titulos.Count) return "Parcial (" + loc.Count + "/" + nf.titulos.Count + " parcelas): " + sb;
            return "Localizada: " + sb;
        }

        bool PesquisaNota(NotaFiscal nf, List<string> termos)
        {
            if (termos.Count == 0) return true;
            string tx = nf.tipo + " " + nf.numero + " " + nf.emitNome + " " + nf.destNome + " " + nf.arquivo + " " + nf.chave;
            string dg = nf.emitCnpj + " " + nf.destCnpj + " " + nf.chave;
            var vals = new List<double> { nf.valor, nf.valorLiquido };
            vals.AddRange(nf.titulos.Select(t => t.valor));
            var va = vals.ToArray();
            if (termos.All(t => CasaTermo(t, tx, dg, Dt(nf.emissao), va))) return true;
            foreach (var ti in nf.titulos)
                if (ti.vencimento != null && termos.All(t => CasaTermo(t, tx, dg, Dt(ti.vencimento), va))) return true;
            return false;
        }

        void FiltrarNotas()
        {
            var termos = TermosPesquisa();
            var nova = new List<int>();
            for (int i = 0; i < notas.Count; i++) if (PesquisaNota(notas[i], termos)) nova.Add(i);
            if (ordemNF >= 0)
            {
                int c = ordemNF;
                Func<int, IComparable> chave = i =>
                {
                    var nf = notas[i];
                    if (c == 2) return nf.emissao ?? DateTime.MinValue;
                    if (c == 8) return nf.valor;
                    if (c == 1) { long n; return long.TryParse(nf.numero, out n) ? n.ToString("D20") : nf.numero; }
                    return ValorNF(nf, c);
                };
                nova = (ordemNFDesc ? nova.OrderByDescending(chave) : nova.OrderBy(chave)).ToList();
            }
            visNF = nova;
            gradeNF.RowCount = 0;
            gradeNF.RowCount = visNF.Count;
            gradeNF.Invalidate();
        }

        void AdicionarXml(bool pasta)
        {
            List<string> arqs;
            if (pasta)
            {
                using (var fb = new FolderBrowserDialog { Description = "Pasta com os XMLs (inclui subpastas e arquivos .zip)", SelectedPath = Prefs.Pasta("dirXml", Prefs.Pasta("dirExt", null)) })
                {
                    if (fb.ShowDialog(this) != DialogResult.OK) return;
                    Prefs.Put("dirXml", fb.SelectedPath);
                    arqs = ArquivosDaPasta(fb.SelectedPath);
                }
            }
            else
            {
                using (var fc = new OpenFileDialog { Multiselect = true, Filter = "XML / ZIP (*.xml, *.zip)|*.xml;*.zip|Todos os arquivos|*.*", Title = "XML das notas fiscais (NF-e, NFC-e, CT-e, NFS-e) ou .zip com XMLs", InitialDirectory = Prefs.Pasta("dirXml", Prefs.Pasta("dirExt", null)) })
                {
                    if (fc.ShowDialog(this) != DialogResult.OK) return;
                    Prefs.Put("dirXml", Path.GetDirectoryName(fc.FileNames[0]));
                    arqs = fc.FileNames.ToList();
                }
            }
            CarregarXml(arqs);
        }

        void CarregarXml(List<string> arquivos)
        {
            int antes = notas.Count, repetidas = 0, ignorados = 0, cancel = 0;
            var avisos = new StringBuilder();
            var chaves = new HashSet<string>(notas.Where(nf => nf.chave.Length > 0 && nf.tipo != "Evento").Select(nf => nf.tipo + nf.chave));
            var eventos = new List<NotaFiscal>();
            var lista = arquivos.Where(f => f.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".zip", StringComparison.OrdinalIgnoreCase)).ToList();
            var lidos = Dialogos.Pesado(this, "Lendo " + lista.Count + " arquivo(s) XML/ZIP...", () =>
            {
                var m = new List<KeyValuePair<string, object>>();
                foreach (var f in lista)
                {
                    try { m.Add(new KeyValuePair<string, object>(f, LeitorNotaXML.LerArquivo(f))); }
                    catch (Exception ex) { m.Add(new KeyValuePair<string, object>(f, ex)); }
                }
                return m;
            });
            foreach (var ent in lidos)
            {
                if (ent.Value is Exception) { avisos.Append("• ").Append(Path.GetFileName(ent.Key)).Append(": ").Append(((Exception)ent.Value).Message).Append("\r\n"); continue; }
                var ls = (List<NotaFiscal>)ent.Value;
                if (ls.Count == 0) ignorados++;
                foreach (var nf in ls)
                {
                    if (nf.tipo == "Evento") { eventos.Add(nf); continue; }
                    string k = nf.tipo + (nf.chave.Length == 0 ? nf.emitCnpj + "|" + nf.numero + "|" + nf.valor.ToString("R", CultureInfo.InvariantCulture) : nf.chave);
                    if (!chaves.Add(k)) { repetidas++; continue; }
                    notas.Add(nf);
                }
            }
            foreach (var ev in eventos)
                foreach (var nf in notas) if (nf.chave == ev.chave && !nf.cancelada) { nf.cancelada = true; cancel++; }
            if (txtCnpjEmp.Text.Trim().Length == 0 || antes == 0)
            {
                string c = VinculadorNotas.DetectarEmpresa(notas);
                if (c.Length > 0) txtCnpjEmp.Text = FormatarDoc(c);
            }
            abas.SelectedIndex = 2;
            Revincular();
            string msg = (notas.Count - antes) + " nota(s) adicionada(s)" + (repetidas > 0 ? " | " + repetidas + " já anexada(s) (ignoradas)" : "")
                         + (cancel > 0 ? " | " + cancel + " cancelada(s) pelo evento" : "") + (ignorados > 0 ? " | " + ignorados + " arquivo(s) sem nota reconhecida" : "");
            status.Text = status.Text + "   |   " + msg;
            if (repetidas > 0 || avisos.Length > 0) Dialogos.Escolher(this, "XML", msg + (avisos.Length > 0 ? "\r\n\r\n" + avisos : ""), 0, "OK");
        }

        void RemoverNotas()
        {
            var idx = new SortedSet<int>();
            foreach (DataGridViewRow r in gradeNF.SelectedRows) idx.Add(r.Index);
            var rem = new HashSet<NotaFiscal>(idx.Select(Nota).Where(x => x != null));
            if (rem.Count == 0) { Dialogos.Aviso(this, "Selecione as notas na aba Notas fiscais."); abas.SelectedIndex = 2; return; }
            notas.RemoveAll(rem.Contains);
            Revincular();
        }

        void LimparNotas()
        {
            if (notas.Count == 0) return;
            if (!Dialogos.SimNao(this, "Limpar notas", "Tirar as " + notas.Count + " nota(s) anexada(s)?")) return;
            notas.Clear();
            txtCnpjEmp.Text = "";
            Revincular();
        }

        string NotaTexto(Lancamento l)
        {
            if (l.notas.Count == 0) return notas.Count == 0 || l.descartada ? "" : "Sem nota";
            var x = l.notas.Select(t =>
            {
                string parc = t.nf.titulos.Count > 1 ? " parc. " + (t.parcela.Length == 0 ? (t.nf.titulos.IndexOf(t) + 1) + "/" + t.nf.titulos.Count : t.parcela) : "";
                return t.nf.Rotulo() + parc + " – " + t.nf.ContraparteNome(CnpjEmpresa());
            });
            return (l.notas.Count > 1 ? "[" + l.notas.Count + " títulos] " : "") + Juntar(x);
        }

        static string Juntar(IEnumerable<string> l)
        {
            var u = new List<string>();
            foreach (var s in l) if (!string.IsNullOrEmpty(s) && !u.Contains(s)) u.Add(s);
            return string.Join(" | ", u);
        }

        // ===================== planilha financeira =====================

        void AdicionarPlanilha()
        {
            using (var fc = new OpenFileDialog { Multiselect = true, Filter = "Planilha financeira (Excel, CSV ou PDF)|*.xlsx;*.xlsm;*.xls;*.ods;*.csv;*.txt;*.pdf|Todos os arquivos|*.*", Title = "Planilha financeira do cliente (Excel, CSV ou PDF) — pode selecionar várias", InitialDirectory = Prefs.Pasta("dirPlan", Prefs.Pasta("dirExt", null)) })
            {
                if (fc.ShowDialog(this) != DialogResult.OK) return;
                Prefs.Put("dirPlan", Path.GetDirectoryName(fc.FileNames[0]));
                CarregarPlanilha(fc.FileNames.ToList());
            }
        }

        void CarregarPlanilha(List<string> arquivos)
        {
            var avisos = new StringBuilder();
            int antes = planilha.Count;
            foreach (var f in arquivos)
            {
                string chave = Path.GetFullPath(f), nome = Path.GetFileName(f);
                if (arquivosPlanilha.Contains(chave)) { avisos.Append("• ").Append(nome).Append(": planilha já anexada.\r\n"); continue; }
                try
                {
                    var lidos = Dialogos.Pesado(this, "Lendo planilha " + nome + "...", () => LeitorPlanilhaFinanceira.Ler(f));
                    planilha.AddRange(lidos);
                    arquivosPlanilha.Add(chave);
                }
                catch (Exception ex)
                {
                    if (!(ex is IOException)) Log.Erro("planilha " + nome, ex);
                    avisos.Append("• ").Append(ex.Message).Append("\r\n");
                }
            }
            Revincular();
            status.Text = status.Text + "   |   " + (planilha.Count - antes) + " item(ns) da planilha lido(s)";
            if (avisos.Length > 0) Dialogos.Escolher(this, "Avisos", avisos.ToString().TrimEnd(), 0, "OK");
            if (planilha.Count > antes && extrato.Count > 0) ResumoConciliacao(true);
        }

        void RemoverPlanilha()
        {
            if (planilha.Count == 0) return;
            planilha.Clear();
            arquivosPlanilha.Clear();
            Revincular();
        }

        string Situacao(ItemPlanilha it)
        {
            if (it.exportado.Length > 0) return "Exportado — " + it.exportado;
            if (it.vinculo == null) return "Não localizado no extrato";
            var l = it.vinculo;
            string onde = l.data + " — " + l.historico + (l.planilha.Count > 1 ? " (agrupado, " + l.planilha.Count + " itens)" : "");
            if (Conciliado(l)) return "Conciliado: " + onde;
            if (l.conciliacao == null) return "Vinculado (taxa): " + onde;
            return "Divergente (" + string.Join("; ", l.conciliacao.Divergencias(ExigirNF())) + "): " + onde;
        }

        void VerPlanilha()
        {
            if (planilha.Count == 0) { Dialogos.Aviso(this, "Nenhuma planilha financeira anexada."); return; }
            int naoLoc = planilha.Count(it => it.vinculo == null && it.exportado.Length == 0);
            using (var f = new Form { Text = "Planilha financeira", Font = Dialogos.Fonte, Size = new Size(1240, 600), StartPosition = FormStartPosition.CenterParent, ShowInTaskbar = false, MinimizeBox = false })
            {
                var t = new DataGridView
                {
                    Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, RowHeadersVisible = false, BackgroundColor = SystemColors.Window,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect, ShowCellToolTips = true
                };
                string[] cols = { "Data", "Outras datas", "Favorecido / Cliente", "CPF/CNPJ", "Descrição", "NF", "Categoria", "Conta", "Valor", "E/S", "Situação", "Arquivo / linha" };
                int[] larg = { 78, 90, 220, 120, 240, 70, 140, 55, 90, 35, 320, 140 };
                for (int i = 0; i < cols.Length; i++) t.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = cols[i], Width = larg[i] });
                t.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                int nConc = 0, nDiv = 0;
                foreach (var it in planilha)
                {
                    string sit = Situacao(it);
                    int r = t.Rows.Add(it.Data(), string.Join(" ", it.datas.Skip(1)), it.favorecido, it.documento, it.descricao, string.Join(", ", it.nfs), it.categoria,
                        it.conta == null ? "" : it.conta.ToString(), Moeda(it.valor), it.saida == null ? "" : it.saida.Value ? "S" : "E", sit, it.arquivo + " : " + it.linha);
                    bool div = sit.StartsWith("Divergente", StringComparison.Ordinal);
                    if (div) nDiv++; else if (sit.StartsWith("Conciliado", StringComparison.Ordinal)) nConc++;
                    t.Rows[r].DefaultCellStyle.BackColor = sit.StartsWith("Não", StringComparison.Ordinal) ? ROSA : div ? AMARELO : VERDE;
                }
                var topo = new Label { Dock = DockStyle.Top, Height = 24, Padding = new Padding(6, 6, 6, 0),
                    Text = planilha.Count + " item(ns) | verde = conciliado/exportado (" + nConc + ") | amarelo = vinculado com divergência (" + nDiv + ") | rosa = não localizado no extrato (" + naoLoc + ")" };
                var bts = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(6) };
                var fe = new Button { Text = "Fechar", MinimumSize = new Size(90, 28), AutoSize = true };
                fe.Click += (s, e) => f.Close();
                bts.Controls.Add(fe);
                f.Controls.Add(t); f.Controls.Add(topo); f.Controls.Add(bts);
                f.CancelButton = fe;
                f.ShowDialog(this);
            }
        }

        // ===================== histórico / complemento =====================

        /// <summary>Favorecido/cliente da planilha (único) ou null.</summary>
        static string FavorecidoPlanilha(Lancamento l)
        {
            string f = null;
            foreach (var it in l.planilha)
            {
                if (it.favorecido.Length == 0) continue;
                if (f != null && Regras.Normalizar(f) != Regras.Normalizar(it.favorecido)) return null;
                f = it.favorecido;
            }
            return f;
        }

        /// <summary>Participante da(s) nota(s) localizada(s) (único) ou null.</summary>
        string ParticipanteNota(Lancamento l)
        {
            string f = null;
            foreach (var t in l.notas)
            {
                string n = t.nf.ContraparteNome(CnpjEmpresa());
                if (n.Length == 0) continue;
                if (f != null && Regras.Normalizar(f) != Regras.Normalizar(n)) return null;
                f = n;
            }
            return f;
        }

        static bool NomeCitado(string historico, string nome)
        {
            var tn = Classificador.Tokens(nome);
            var th = Classificador.TokSet(historico);
            int comum = tn.Count(th.Contains);
            return tn.Count > 0 && comum >= Math.Min(2, tn.Count);
        }

        static string NumerosNF(Lancamento l)
        {
            var nums = new List<string>();
            foreach (var t in l.notas) if (t.nf.numero.Length > 0 && !nums.Contains(t.nf.numero)) nums.Add(t.nf.numero);
            if (nums.Count == 0) return "";
            return (nums.Count == 1 ? "NF " : "NFs ") + string.Join(", ", nums);
        }
    }
}

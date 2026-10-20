using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExtratoComprovantes
{
    public partial class TelaPrincipal
    {
        // ===================== menus =====================

        ToolStripMenuItem Item(string texto, Action acao, Keys atalho = Keys.None, string dica = null)
        {
            var it = new ToolStripMenuItem(texto);
            if (atalho != Keys.None) it.ShortcutKeys = atalho;
            if (dica != null) it.ToolTipText = dica;
            it.Click += (s, e) => Executar(acao);
            return it;
        }

        /// <summary>Executa a ação do menu; erro inesperado vai para o log e para um aviso (o trabalho na tela não se perde).</summary>
        void Executar(Action a)
        {
            try
            {
                if (grade.IsCurrentCellInEditMode) grade.EndEdit();
                a();
            }
            catch (Exception ex)
            {
                Log.Erro("ação do menu", ex);
                Dialogos.Aviso(this, "Ocorreu um erro: " + ex.Message + "\r\n\r\nO trabalho na tela não foi perdido. Detalhes gravados em:\r\n" + Path.Combine(Log.PastaApp(), "log.txt"));
            }
        }

        MenuStrip CriarMenu()
        {
            var mb = new MenuStrip { Dock = DockStyle.Top };

            var arq = new ToolStripMenuItem("&Arquivo");
            arq.DropDownItems.Add(Item("Novo (mesmo cliente ou nova empresa)...", NovaSessao, Keys.Control | Keys.N));
            arq.DropDownItems.Add(Item("Abrir progresso...", () => AbrirProgresso(null), Keys.Control | Keys.O));
            arq.DropDownItems.Add(menuRecentes);
            arq.DropDownItems.Add(new ToolStripSeparator());
            arq.DropDownItems.Add(Item("Salvar", Salvar, Keys.Control | Keys.S, "Grava os lançamentos da tela e os já exportados (contas, edições, complemento), com o Plano e o Razão, para resgatar depois."));
            arq.DropDownItems.Add(Item("Salvar como...", () => SalvarProgresso(false), Keys.Control | Keys.Shift | Keys.S));
            arq.DropDownItems.Add(new ToolStripSeparator());
            arq.DropDownItems.Add(Item("Pasta de configurações e log", () => AbrirArquivo(Log.PastaApp())));
            arq.DropDownItems.Add(new ToolStripSeparator());
            arq.DropDownItems.Add(Item("Sair", Close, Keys.Alt | Keys.F4));
            mb.Items.Add(arq);

            var cli = new ToolStripMenuItem("&Cliente");
            cli.DropDownItems.Add(Item("Abrir regras do cliente...", AbrirRegras));
            cli.DropDownItems.Add(Item("Novas regras do cliente...", NovaRegras));
            cli.DropDownItems.Add(Item("Editar regras", EditarRegras));
            cli.DropDownItems.Add(new ToolStripSeparator());
            cli.DropDownItems.Add(Item("Carregar Razão...", CarregarRazao, Keys.Control | Keys.Shift | Keys.R));
            cli.DropDownItems.Add(Item("Carregar Plano de Contas...", CarregarPlano, Keys.Control | Keys.Shift | Keys.P));
            cli.DropDownItems.Add(Item("Folha e adiantamento (relação de líquidos)...", JanelaFolha, Keys.None,
                "Relação Geral dos Líquidos do Domínio (Folha Mensal e Adiantamento): separa salário e adiantamento de salário pelo nome, pelo líquido e pelo intervalo de pagamento."));
            cli.DropDownItems.Add(Item("Limpar Razão, Plano e funcionários", LimparBase));
            cli.DropDownItems.Add(new ToolStripSeparator());
            cli.DropDownItems.Add(Item("Aprendizado do cliente (o que o programa memorizou)...", VerAprendizado));
            mb.Items.Add(cli);

            var ext = new ToolStripMenuItem("&Extrato");
            ext.DropDownItems.Add(Item("Adicionar extrato (OFX, PDF, Excel, CSV...)...", AdicionarExtrato, Keys.Control | Keys.I));
            ext.DropDownItems.Add(Item("Inserir lançamento...", InserirLancamento, Keys.Control | Keys.I, "Inclui um lançamento que não veio no extrato (tecla Insert na tabela)."));
            ext.DropDownItems.Add(Item("Duplicar lançamento selecionado...", DuplicarLancamento, Keys.Control | Keys.D, "Abre a inclusão já preenchida com a linha selecionada (para desmembrar ou repetir um lançamento)."));
            ext.DropDownItems.Add(Item("Excluir lançamentos selecionados", ExcluirSelecionados, Keys.None, "Exclui as linhas selecionadas, com confirmação (tecla Delete na tabela)."));
            ext.DropDownItems.Add(Item("Limpar extrato", LimparExtrato));
            ext.DropDownItems.Add(chkCont);
            ext.DropDownItems.Add(new ToolStripSeparator());
            var lote = Item("Alterar contas em lote...", AlterarEmLote, Keys.Control | Keys.L, "Altera Débito e/ou Crédito de vários lançamentos de uma vez: os selecionados na tabela ou todos com o mesmo texto no histórico.");
            lote.Font = new Font(lote.Font, FontStyle.Bold);
            ext.DropDownItems.Add(lote);
            ext.DropDownItems.Add(Item("Criar regra da linha", CriarRegraDaLinha));
            ext.DropDownItems.Add(Item("Inverter D/C", InverterDC));
            ext.DropDownItems.Add(Item("Reaplicar regras", () => { foreach (var l in extrato) l.manual = false; Revincular(); }, Keys.None, "Desfaz as edições manuais de conta e classifica tudo de novo."));
            ext.DropDownItems.Add(Item("Classificar novamente", Revincular, Keys.F5));
            ext.DropDownItems.Add(new ToolStripSeparator());
            ext.DropDownItems.Add(Item("Limpar pesquisa", () => { txtPesquisa.Text = ""; cbFiltro.SelectedIndex = 0; cbSentido.SelectedIndex = 0; }));
            mb.Items.Add(ext);

            var cmp = new ToolStripMenuItem("C&omprovantes");
            cmp.DropDownItems.Add(Item("Adicionar comprovantes (PDF)...", AdicionarComprovantes));
            cmp.DropDownItems.Add(Item("Remover selecionados", RemoverComprovantes));
            cmp.DropDownItems.Add(Item("Limpar comprovantes", LimparComprovantes));
            cmp.DropDownItems.Add(new ToolStripSeparator());
            cmp.DropDownItems.Add(Item("Ver texto do comprovante", VerTexto));
            mb.Items.Add(cmp);

            var nf = new ToolStripMenuItem("&Notas fiscais");
            nf.DropDownItems.Add(Item("Adicionar XML / ZIP...", () => AdicionarXml(false)));
            nf.DropDownItems.Add(Item("Adicionar pasta...", () => AdicionarXml(true)));
            nf.DropDownItems.Add(Item("Remover selecionadas", RemoverNotas));
            nf.DropDownItems.Add(Item("Limpar notas", LimparNotas));
            mb.Items.Add(nf);

            var pl = new ToolStripMenuItem("&Planilha");
            pl.DropDownItems.Add(Item("Anexar planilha financeira...", AdicionarPlanilha, Keys.None, "Planilha financeira do cliente (contas pagas/recebidas, fluxo de caixa) em Excel, CSV ou PDF — ajuda a identificar a contrapartida."));
            pl.DropDownItems.Add(Item("Ver itens da planilha", VerPlanilha));
            pl.DropDownItems.Add(Item("Remover planilha", RemoverPlanilha));
            pl.DropDownItems.Add(new ToolStripSeparator());
            pl.DropDownItems.Add(Item("Conciliar novamente", () => { Revincular(); ResumoConciliacao(false); }));
            pl.DropDownItems.Add(chkExigirNF);
            mb.Items.Add(pl);

            var ex = new ToolStripMenuItem("E&xportar");
            var expTxt = Item("Exportar lançamentos (.txt)", ExportarLancamentos, Keys.Control | Keys.E, "Exportação principal: TXT no layout do lanctos.txt.");
            expTxt.Font = new Font(expTxt.Font, FontStyle.Bold);
            ex.DropDownItems.Add(expTxt);
            ex.DropDownItems.Add(Item("Exportar para Excel (análise)...", ExportarExcel, Keys.Control | Keys.Shift | Keys.E,
                "Planilha .xlsx com os lançamentos (contas, situação, comprovante, nota, planilha e conciliação), filtrados por período, e abas de apoio. Não marca nada como exportado."));
            ex.DropDownItems.Add(new ToolStripSeparator());
            ex.DropDownItems.Add(chkIncluirVAI);
            mb.Items.Add(ex);

            var aj = new ToolStripMenuItem("Aj&uda");
            aj.DropDownItems.Add(Item("Sobre", () => Dialogos.Info(this, "Sobre", "Extrato + Comprovantes + Planilha Financeira + XML → Lançamentos Contábeis\r\nVersão " + VERSAO + " (etapa 2: extrato, comprovantes, notas fiscais, planilha financeira e TXT)\r\n\r\n"
                + "Leitura de PDF: PdfPig (licença Apache 2.0).\r\nConfigurações e log: " + Log.PastaApp())));
            mb.Items.Add(aj);

            AtualizarRecentes();
            return mb;
        }

        // ===================== campos do topo =====================

        void CampoAlterado()
        {
            if (carregandoCampos) return;
            SincronizarCampos(false);
            ClassificarTodos();
            Atualizar();
        }

        bool SincronizarCampos(bool avisar)
        {
            regras.cliente = txtCliente.Text.Trim();
            string b = txtBanco.Text.Trim(), v = txtVAI.Text.Trim();
            int x;
            if (b.Length > 0 && !int.TryParse(b, out x) || v.Length > 0 && !int.TryParse(v, out x))
            {
                if (avisar) Dialogos.Aviso(this, "Código de conta inválido (use só o código numérico, ex.: 596).");
                return false;
            }
            regras.contaBanco = b.Length == 0 ? (int?)null : int.Parse(b);
            if (v.Length > 0) regras.contaVAI = int.Parse(v);
            return true;
        }

        void ColocarCampos()
        {
            carregandoCampos = true;
            try
            {
                txtCliente.Text = regras.cliente ?? "";
                txtBanco.Text = regras.contaBanco == null ? "" : regras.contaBanco.ToString();
                txtVAI.Text = regras.contaVAI.ToString();
            }
            finally { carregandoCampos = false; }
        }

        // ===================== classificação =====================

        /// <summary>Texto usado na classificação: histórico + comprovante + planilha financeira + participante da nota.</summary>
        string TextoBusca(Lancamento l)
        {
            var sb = new StringBuilder(l.historico);
            if (l.buscaSalva != null) sb.Append(' ').Append(l.buscaSalva);
            var cv = l.comprovante;
            if (cv != null) sb.Append(' ').Append(cv.favorecido).Append(' ').Append(cv.docFavorecido).Append(' ').Append(Regex.Replace(cv.docFavorecido, "[^0-9]", "")).Append(' ').Append(cv.descricao);
            foreach (var it in l.planilha)
                sb.Append(' ').Append(it.favorecido).Append(' ').Append(it.documento).Append(' ').Append(Regex.Replace(it.documento, "[^0-9]", ""))
                  .Append(' ').Append(it.descricao).Append(' ').Append(it.categoria);
            var vistas = new HashSet<NotaFiscal>();
            foreach (var t in l.notas)
                if (vistas.Add(t.nf))
                {
                    string c = t.nf.ContraparteCnpj(CnpjEmpresa());
                    sb.Append(' ').Append(t.nf.ContraparteNome(CnpjEmpresa())).Append(' ').Append(FormatarDoc(c)).Append(' ').Append(c);
                }
            return sb.ToString();
        }

        /// <summary>Complemento do histórico: histórico + favorecido (planilha, comprovante ou nota) + número da(s) NF(s).</summary>
        string Complemento(Lancamento l)
        {
            if (l.complementoSalvo != null && l.comprovante == null && l.planilha.Count == 0 && l.notas.Count == 0) return l.complementoSalvo;
            string h = l.historico;
            string f = FavorecidoPlanilha(l);
            if (f == null && l.comprovante != null && l.comprovante.favorecido.Length > 0) f = l.comprovante.favorecido;
            if (f == null) f = ParticipanteNota(l);
            if (f != null && !Regras.Normalizar(h).Contains(Regras.Normalizar(f)) && !NomeCitado(h, f)) h = h + " - " + f;
            string nf = NumerosNF(l);
            if (nf.Length > 0) h = h + " " + nf;
            return Regex.Replace(Regex.Replace(h, @"[\r\n\t]+", " "), " {2,}", " ").Trim();
        }

        void Revincular()
        {
            alterado = true;
            foreach (var l in extrato) if (!l.manual) { l.descartada = false; l.par = null; }
            DetectarTransferencias();
            Vinculador.Vincular(extrato, comps, (int)spDiasComp.Value);
            VinculadorPlanilha.Vincular(extrato, planilha, (int)spDias.Value);
            VinculadorNotas.Vincular(extrato, notas, CnpjEmpresa(), (int)spVenc.Value, (int)spPrazo.Value);
            AvaliarConciliacao();
            ClassificarTodos();
            DetectarTransferencias();
            Atualizar();
        }

        // ===================== conciliação com a planilha financeira =====================

        bool ExigirNF() { return chkExigirNF.Checked; }

        /// <summary>Linha conciliada: valor, data, descrição e NF da planilha batem com o extrato.</summary>
        bool Conciliado(Lancamento l) { return l.conciliacao != null && l.conciliacao.Completa(ExigirNF()); }

        /// <summary>Confere valor, data, descrição e NF de cada linha do extrato vinculada à planilha.</summary>
        void AvaliarConciliacao()
        {
            var taxas = new Dictionary<Lancamento, List<ItemPlanilha>>();
            foreach (var it in planilha)
                if (it.vinculo != null && !it.vinculo.planilha.Contains(it)) taxas.ObterOuCriar(it.vinculo).Add(it);
            string emp = CnpjEmpresa();
            foreach (var l in extrato) l.conciliacao = Conciliador.Avaliar(l, taxas.GetOr(l, null), emp);
        }

        /// <summary>Aviso com o resultado da conciliação automática (depois de anexar a planilha ou de conciliar novamente).</summary>
        void ResumoConciliacao(bool soSeHouver)
        {
            if (planilha.Count == 0) { if (!soSeHouver) Dialogos.Aviso(this, "Nenhuma planilha financeira anexada."); return; }
            int conc = 0, div = 0;
            var motivos = new Dictionary<string, int>();
            foreach (var l in extrato)
            {
                if (l.conciliacao == null) continue;
                if (Conciliado(l)) { conc++; continue; }
                div++;
                foreach (var m in l.conciliacao.Divergencias(ExigirNF()))
                    motivos.Somar(m.StartsWith("data", StringComparison.Ordinal) ? "data diferente" : m.StartsWith("NF ", StringComparison.Ordinal) ? "NF não confirmada (sem XML/comprovante/histórico com o número)" : m, 1);
            }
            int naoLoc = planilha.Count(it => it.vinculo == null && it.exportado.Length == 0);
            if (soSeHouver && conc + div == 0 && naoLoc == 0) return;
            var sb = new StringBuilder();
            sb.AppendFormat("Conciliados automaticamente: {0} lançamento(s) do extrato — valor, data, descrição e NF da planilha conferem.\r\n", conc);
            if (div > 0)
            {
                sb.AppendFormat("\r\nVinculados com divergência: {0} (a contrapartida da planilha é usada, mas confira):\r\n", div);
                foreach (var e in motivos.OrderByDescending(e => e.Value)) sb.Append("   • ").Append(e.Key).Append(": ").Append(e.Value).Append("\r\n");
            }
            if (naoLoc > 0) sb.AppendFormat("\r\nItens da planilha não localizados no extrato: {0} (menu Planilha > Ver itens da planilha).\r\n", naoLoc);
            sb.Append("\r\nFiltro \"Mostrar\": \"Conciliados com a planilha\" ou \"Planilha com divergência\"; detalhe na coluna Conciliação (planilha).");
            Dialogos.Info(this, "Conciliação com a planilha financeira", sb.ToString());
        }

        void ClassificarTodos()
        {
            classif.regras = regras;
            classif.plano = plano;
            classif.razao = razao;
            classif.funcionarios = funcionarios;
            classif.folha = folha;
            classif.contasBanco = new HashSet<int>(bancoPorArquivo.Where(e => e.Value != null).Select(e => e.Value.Value));
            classif.Preparar();
            classif.categoriaSessao = new Dictionary<string, int>();
            foreach (var l in extrato) classif.Classificar(l, TextoBusca(l));
            // categoria da planilha → conta, pelas linhas já classificadas (todas com a mesma conta); 2ª passada nas demais
            var cat = new Dictionary<string, Dictionary<int, int>>();
            foreach (var l in extrato)
            {
                if (l.descartada) continue;
                int? contra = l.saida ? l.debito : l.credito;
                string c = Classificador.CategoriaDaLinha(l);
                if (contra == null || contra == regras.contaVAI || string.IsNullOrEmpty(c)) continue;
                cat.ObterOuCriar((l.saida ? "S|" : "E|") + c).Somar(contra.Value, 1);
            }
            var sessao = new Dictionary<string, int>();
            foreach (var e in cat) if (e.Value.Count == 1) sessao[e.Key] = e.Value.Keys.First();
            if (sessao.Count > 0)
            {
                classif.categoriaSessao = sessao;
                foreach (var l in extrato) if (!l.manual && EmVAI(l)) classif.Classificar(l, TextoBusca(l));
            }
            AtualizarBase();
        }

        void AtualizarBase()
        {
            var sb = new StringBuilder();
            if (razao != null) sb.Append("Razão: ").Append(string.Join(", ", razao.arquivos)).Append(" (").Append(classif.QtdReferencias()).Append(" lançamentos com banco)");
            if (plano != null) sb.Append(sb.Length > 0 ? "   |   " : "").Append("Plano: ").Append(plano.arquivo).Append(" (").Append(plano.contas.Count).Append(" contas)");
            if (folha != null) sb.Append(sb.Length > 0 ? "   |   " : "").Append("Folha/adiantamento: ").Append(folha.itens.Select(x => x.ano * 100 + x.mes).Distinct().Count()).Append(" competência(s)");
            txtBase.Text = sb.Length == 0 ? "Nenhum carregado — carregue o Razão e o Plano de Contas do cliente (menu Cliente)" : sb.ToString();
            dicas.SetToolTip(txtBase, txtBase.Text);
            if (regras.contaBanco == null || plano == null) lblBanco.Text = "";
            else
            {
                var c = plano.Get(regras.contaBanco);
                lblBanco.Text = c == null ? "⚠ não existe no Plano" : (c.Analitica() ? c.nome : "⚠ " + c.nome + " é sintética");
                lblBanco.ForeColor = c != null && c.Analitica() ? Color.FromArgb(0, 110, 0) : Color.FromArgb(180, 0, 0);
            }
        }

        static bool HistoricoTransferencia(string h)
        {
            string n = " " + Regras.Normalizar(h) + " ";
            return n.Contains(" PIX") || n.Contains(" TED") || n.Contains("TRANSF") || n.Contains(" DOC ") || n.Contains(" TEF")
                   || n.Contains("REPASSE") || n.Contains("APLICAC") || n.Contains("RESGATE");
        }

        /// <summary>Transferência entre bancos: mesma data e valor, saída num extrato e entrada em outro. Mantém a saída e descarta a entrada.</summary>
        void DetectarTransferencias()
        {
            var bancos = new HashSet<int>();
            foreach (var l in extrato) if (BancoDe(l) != null) bancos.Add(BancoDe(l).Value);
            if (bancos.Count < 2) return;
            var usadas = new HashSet<Lancamento>();
            for (int passo = 0; passo < 2; passo++)
                foreach (var s in extrato)
                {
                    if (!s.saida || s.manual || usadas.Contains(s) || BancoDe(s) == null || !HistoricoTransferencia(s.historico)) continue;
                    foreach (var e in extrato)
                    {
                        if (e.saida || e.manual || usadas.Contains(e) || BancoDe(e) == null || BancoDe(e) == BancoDe(s)) continue;
                        double dif = Math.Round((s.valor - e.valor) * 100) / 100.0;
                        bool ok = passo == 0 ? Math.Abs(dif) <= 0.005 : dif > 0.005 && dif <= Math.Max(10, s.valor * 0.001);
                        if (e.data != s.data || !ok || !HistoricoTransferencia(e.historico)) continue;
                        usadas.Add(e); usadas.Add(s);
                        e.descartada = true;
                        e.par = s;
                        e.debito = null; e.credito = null;
                        e.obs = "Transferência descartada — lançada na saída do banco " + BancoDe(s) + " (" + s.origem + ")";
                        s.debito = BancoDe(e);
                        s.credito = BancoDe(s);
                        var pd = plano == null ? null : plano.Get(BancoDe(e));
                        s.obs = "Transferência entre bancos: D banco destino " + BancoDe(e) + (pd != null ? " " + pd.nome : "") + " | C banco origem";
                        if (passo == 1) s.obs += " | CONFERIR: chegou " + Moeda(e.valor) + " no banco " + BancoDe(e) + ", diferença de " + Moeda(dif)
                                                + " (tarifa) — lançar D tarifa / C " + BancoDe(e);
                        break;
                    }
                }
        }

        // ===================== extrato =====================

        static readonly string FILTRO_EXTRATO = "Extrato (OFX, PDF, Excel, CSV, QIF, MT940, CAMT, CNAB 240, ODS, HTML)|*.ofx;*.ofc;*.pdf;*.xlsx;*.xlsm;*.xls;*.csv;*.txt;*.qif;*.sta;*.mt940;*.940;*.942;*.xml;*.ret;*.cnab;*.rem;*.ods;*.htm;*.html|Todos os arquivos|*.*";

        void AdicionarExtrato()
        {
            using (var fc = new OpenFileDialog { Multiselect = true, Filter = FILTRO_EXTRATO, Title = "Adicionar extrato", InitialDirectory = Prefs.Pasta("dirExt", null) })
            {
                if (fc.ShowDialog(this) != DialogResult.OK) return;
                Prefs.Put("dirExt", Path.GetDirectoryName(fc.FileNames[0]));
                Carregar(fc.FileNames.ToList());
            }
        }

        void Arrastando(object s, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        }

        void Soltou(object s, DragEventArgs e)
        {
            var fs = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (fs == null) return;
            var todos = new List<string>();
            foreach (var f in fs)
            {
                if (Directory.Exists(f)) todos.AddRange(Directory.GetFiles(f, "*", SearchOption.AllDirectories).OrderBy(x => x, StringComparer.OrdinalIgnoreCase));
                else todos.Add(f);
            }
            var alvo = s;
            BeginInvoke((Action)(() => Executar(() => Distribuir(todos, alvo))));
        }

        /// <summary>
        /// Arquivos arrastados: na aba Comprovantes, PDFs viram comprovantes; .xml/.zip vão para as notas; OFX e formatos padrão
        /// de extrato vão para o extrato; os demais (PDF, Excel, CSV) perguntam se são extrato, planilha financeira ou comprovantes.
        /// </summary>
        void Distribuir(List<string> todos, object alvo)
        {
            if (alvo == gradeComp || abas.SelectedIndex == 1 && alvo != grade)
            {
                CarregarComprovantes(todos.Where(f => f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)).ToList());
                return;
            }
            var ext = new List<string>(); var xml = new List<string>(); var duvida = new List<string>();
            foreach (var f in todos)
            {
                string n = f.ToLowerInvariant();
                if (n.EndsWith(".ofx") || n.EndsWith(".ofc") || LeitorExtrato.FormatoPadrao(f) != null) ext.Add(f);
                else if (n.EndsWith(".xml") || n.EndsWith(".zip")) xml.Add(f);
                else if (Regex.IsMatch(n, @"\.(pdf|xlsx|xlsm|xls|csv|txt|ods|htm|html)$")) duvida.Add(f);
            }
            var plan = new List<string>();
            if (duvida.Count > 0)
            {
                var nomes = new StringBuilder();
                for (int i = 0; i < duvida.Count && i < 8; i++) nomes.Append("\r\n   • ").Append(Path.GetFileName(duvida[i]));
                if (duvida.Count > 8) nomes.Append("\r\n   ...");
                int r = Dialogos.Escolher(this, "Arquivo arrastado", "O que é este(s) arquivo(s)?" + nomes, 0, "Extrato bancário", "Planilha financeira", "Comprovantes", "Cancelar");
                if (r == 0) ext.AddRange(duvida);
                else if (r == 1) plan.AddRange(duvida);
                else if (r == 2) CarregarComprovantes(duvida.Where(f => f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)).ToList());
            }
            if (xml.Count > 0) CarregarXml(xml);
            if (ext.Count > 0) Carregar(ext);
            if (plan.Count > 0) CarregarPlanilha(plan);
        }

        /// <summary>Chave do lançamento para não importar o mesmo extrato duas vezes.</summary>
        string ChaveBase(Lancamento l)
        {
            return (BancoDe(l) == null ? "?" : BancoDe(l).ToString()) + "|" + l.data + "|" + l.valor.ToString("0.00", CultureInfo.InvariantCulture) + "|"
                   + (l.saida ? "S" : "E") + "|" + Regex.Replace(Regras.Normalizar(l.historico), @"\s+", " ").Trim();
        }

        List<Lancamento> FiltrarNovos(List<Lancamento> lote)
        {
            var ocorr = new Dictionary<string, int>();
            foreach (var l in lote)
            {
                if (l.chave != null) continue;
                string b = ChaveBase(l);
                ocorr.Somar(b, 1);
                l.chave = b + "#" + ocorr[b];
            }
            var ja = new HashSet<string>();
            foreach (var l in extrato.Concat(exportados)) if (l.chave != null) ja.Add(l.chave);
            var r = new List<Lancamento>();
            foreach (var l in lote) if (ja.Add(l.chave)) r.Add(l);
            return r;
        }

        int? BancoDoArquivo(string origem, out bool tem)
        {
            foreach (var e in bancoPorArquivo) if (e.Key == origem) { tem = true; return e.Value; }
            tem = false;
            return null;
        }

        void Carregar(List<string> arquivos)
        {
            var paraPlanilha = new List<string>();
            var avisos = new StringBuilder();
            var formatos = new List<string>();
            int provas = 0, antes = extrato.Count;
            foreach (var f in arquivos)
            {
                string nome = Path.GetFileName(f);
                try
                {
                    if (nome.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) && Dialogos.Pesado(this, "Verificando " + nome + "...", () => LeitorPDF.PareceRelatorioFinanceiro(f)))
                    {
                        int r = Dialogos.Escolher(this, "Relatório financeiro", nome + "\r\n\r\nEste PDF parece um relatório do sistema financeiro (colunas Terceiro, Documento,\r\n"
                            + "Histórico, Entrada, Saída e Saldo), não o extrato do banco.\r\n\r\nO recomendado é anexá-lo como planilha financeira e adicionar\r\n"
                            + "o extrato do banco (OFX, PDF ou Excel do banco/gateway).", 0, "Anexar como planilha financeira", "Ler como extrato mesmo assim", "Não carregar");
                        if (r == 0) { paraPlanilha.Add(f); continue; }
                        if (r != 1) { avisos.Append("• ").Append(nome).Append(": não carregado.\r\n"); continue; }
                    }
                    bool jc = chkCont.Checked;
                    var res = Dialogos.Pesado(this, "Lendo extrato " + nome + "...", () => LeitorExtrato.Ler(f, jc));
                    var l = res.lancamentos;
                    string prova = res.Prova();
                    formatos.Add(Regex.Replace(res.formato, @" \(.*", ""));
                    Log.Info("extrato " + nome + ": " + res.formato + " - " + l.Count + " lançamento(s)" + (prova.Length == 0 ? "" : " - " + prova));
                    if (prova.StartsWith("PROVA DO SALDO NÃO")) avisos.Append("• ").Append(nome).Append(": ").Append(prova).Append(".\r\n");
                    else if (prova.Length > 0) provas++;
                    if (res.aviso.Length > 0) avisos.Append("• ").Append(nome).Append(": ").Append(res.aviso).Append("\r\n");
                    int duv = l.Count(x => x.conferirSinal);
                    if (duv > 0) avisos.Append("• ").Append(nome).Append(": ").Append(duv).Append(" lançamento(s) com entrada/saída deduzida — confira a coluna Observação.\r\n");
                    if (l.Count == 0) continue;
                    string origem = l[0].origem;
                    bool tem;
                    int? cb = BancoDoArquivo(origem, out tem);
                    if (bancoPorArquivo.Count > 0 && !tem)
                    {
                        cb = PerguntarContaBanco(origem);
                        if (cb == null) { avisos.Append("• ").Append(nome).Append(": não carregado (conta banco não informada).\r\n"); continue; }
                        bancoPorArquivo.Add(new KeyValuePair<string, int?>(origem, cb));
                    }
                    else if (!tem) bancoPorArquivo.Add(new KeyValuePair<string, int?>(origem, null));
                    foreach (var x in l) { x.contaBanco = cb; x.seq = ++sequencia; }
                    var novos = FiltrarNovos(l);
                    if (novos.Count < l.Count)
                        avisos.Append("• ").Append(nome).Append(": ").Append(l.Count - novos.Count).Append(" lançamento(s) já importado(s) antes — ignorado(s) para não duplicar.\r\n");
                    extrato.AddRange(novos);
                }
                catch (Exception ex)
                {
                    if (!(ex is IOException)) Log.Erro("ler extrato " + nome, ex);
                    avisos.Append("• ").Append(ex.Message).Append("\r\n");
                }
            }
            Revincular();
            if (plano != null && regras.contaBanco == null && extrato.Count > 0) SugerirContaBanco();
            int n = extrato.Count - antes;
            status.Text = status.Text + "   |   " + n + " lançamento(s) lido(s)" + (formatos.Count > 0 ? " [" + string.Join("; ", formatos) + "]" : "")
                          + (provas > 0 ? " - prova do saldo confere" + (provas > 1 ? " (" + provas + " arquivos)" : "") : "");
            if (avisos.Length > 0) Dialogos.Escolher(this, "Avisos", avisos.ToString().TrimEnd(), 0, "OK");
            if (paraPlanilha.Count > 0) CarregarPlanilha(paraPlanilha);
        }

        int? PerguntarContaBanco(string arquivo)
        {
            string msg = "Extrato de outro banco: " + arquivo + "\r\nQual é a conta deste banco no Plano de Contas?";
            if (plano != null && plano.ContasDisponivel().Count > 0)
            {
                var disp = plano.ContasDisponivel();
                int i = Dialogos.Lista(this, "Conta banco", msg, disp.Select(c => c.codigo + " - " + c.nome + "   (" + c.classificacao + ")").ToList(), 0);
                return i < 0 ? (int?)null : disp[i].codigo;
            }
            string r = Dialogos.PedirTexto(this, "Conta banco", msg + "\r\n(código numérico, ex.: 597)", "");
            if (r == null) return null;
            int v;
            if (int.TryParse(r.Trim(), out v)) return v;
            Dialogos.Aviso(this, "Código inválido.");
            return null;
        }

        void LimparExtrato()
        {
            if (extrato.Count == 0) return;
            if (!Dialogos.SimNao(this, "Limpar", "Tirar os " + extrato.Count + " lançamento(s) da tela?\r\n(Os já exportados continuam guardados no progresso.)")) return;
            extrato.Clear();
            bancoPorArquivo.Clear();
            Revincular();
        }

        void InverterDC()
        {
            var sel = Selecionados();
            foreach (var l in sel)
            {
                l.saida = !l.saida;
                var d = l.debito; l.debito = l.credito; l.credito = d;
                l.conferirSinal = false;
                l.obs = Regex.Replace(l.obs, @"^Sinal[^|]*\| ?", "");
            }
            if (sel.Count > 0) Revincular();
        }

        string SugestaoTexto(Lancamento l)
        {
            string fav = FavorecidoPlanilha(l);
            if (fav == null && l.comprovante != null && l.comprovante.favorecido.Length > 0) fav = l.comprovante.favorecido;
            if (fav != null) return fav;
            string s = Regex.Replace(l.historico, @"^[^:]*:\s*", "");
            s = Regex.Replace(s, @"^cp\s*:\s*", "", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"[0-9][0-9./,\-]*", " ");
            s = Regex.Replace(s, @"^[\s\-:]+", "");
            return Regex.Replace(s, @"\s+", " ").Trim();
        }

        void CriarRegraDaLinha()
        {
            var l = Atual();
            if (l == null) { Dialogos.Aviso(this, "Selecione uma linha do extrato."); return; }
            if (!SincronizarCampos(true)) return;
            int? contraAtual = l.saida ? l.debito : l.credito;
            using (var f = new NovaRegra(l, SugestaoTexto(l), contraAtual == null || contraAtual == regras.contaVAI ? null : contraAtual, () => plano))
            {
                f.ShowDialog(this);
                if (!f.Ok) return;
                string texto = f.fTexto.Text.Trim();
                int conta;
                if (!int.TryParse(f.fConta.Text.Trim(), out conta)) { Dialogos.Aviso(this, "Conta inválida."); return; }
                if (texto.Length == 0) { Dialogos.Aviso(this, "Informe o texto."); return; }
                regras.lista.Add(new Regras.Regra(texto, conta, f.Tipo));
            }
            l.manual = false;
            int antes = extrato.Count(EmVAI);
            if (SalvarRegras())
            {
                Revincular();
                int depois = extrato.Count(EmVAI);
                status.Text = status.Text + "   |   Regra criada: " + (antes - depois) + " linha(s) reclassificada(s)";
            }
        }

        // ===================== regras do cliente =====================

        string PastaInicialRegras()
        {
            if (arquivoRegras != null && Path.GetDirectoryName(arquivoRegras) != null) return Path.GetDirectoryName(arquivoRegras);
            return Directory.Exists(@"Z:\CLIENTES") ? @"Z:\CLIENTES" : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        string PastaBase() { return Prefs.Pasta("dirBase", PastaInicialRegras()); }

        void AbrirRegras()
        {
            using (var fc = new OpenFileDialog { Filter = "Regras (*.csv)|*.csv|Todos os arquivos|*.*", Title = "Abrir regras do cliente", InitialDirectory = PastaInicialRegras() })
            {
                if (fc.ShowDialog(this) != DialogResult.OK) return;
                CarregarRegras(fc.FileName);
                Revincular();
            }
        }

        void NovaRegras()
        {
            using (var fc = new SaveFileDialog { Filter = "Regras (*.csv)|*.csv", Title = "Novas regras do cliente", InitialDirectory = PastaInicialRegras(), FileName = "regras_extrato.csv" })
            {
                if (fc.ShowDialog(this) != DialogResult.OK) return;
                string p = fc.FileName;
                if (!p.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)) p += ".csv";
                regras = new Regras();
                plano = null; razao = null; funcionarios.Clear(); folha = null;
                arquivoRegras = p;
                ColocarCampos();
                txtRegras.Text = p;
                Prefs.Put("regras", p);
                SalvarRegras();
                Revincular();
            }
        }

        void CarregarRegras(string p)
        {
            try
            {
                regras = Regras.Carregar(p);
                arquivoRegras = p;
                txtRegras.Text = p;
                dicas.SetToolTip(txtRegras, p);
                Prefs.Put("regras", p);
                ColocarCampos();
                CarregarBaseDasRegras();
            }
            catch (Exception ex)
            {
                carregandoCampos = false;
                Log.Erro("abrir regras", ex);
                Dialogos.Aviso(this, "Erro ao abrir regras: " + ex.Message);
            }
        }

        bool SalvarRegras()
        {
            SincronizarCampos(false);
            if (arquivoRegras == null)
            {
                using (var fc = new SaveFileDialog { Filter = "Regras (*.csv)|*.csv", Title = "Salvar regras do cliente", InitialDirectory = PastaInicialRegras(), FileName = "regras_extrato.csv" })
                {
                    if (fc.ShowDialog(this) != DialogResult.OK) return false;
                    string p = fc.FileName;
                    if (!p.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)) p += ".csv";
                    arquivoRegras = p;
                    txtRegras.Text = p;
                    Prefs.Put("regras", p);
                }
            }
            try { regras.Salvar(arquivoRegras); return true; }
            catch (Exception ex) { Dialogos.Aviso(this, "Erro ao salvar regras: " + ex.Message); return false; }
        }

        void EditarRegras()
        {
            if (!SincronizarCampos(true)) return;
            using (var f = new EditorRegras(regras, () => plano))
            {
                f.ShowDialog(this);
                if (f.Resultado == null) return;
                regras.lista.Clear();
                regras.lista.AddRange(f.Resultado);
            }
            if (SalvarRegras()) Revincular();
        }

        /// <summary>Recarrega Plano, Razão e funcionários gravados nas regras do cliente.</summary>
        void CarregarBaseDasRegras()
        {
            plano = null; razao = null;
            var err = new StringBuilder();
            if (!string.IsNullOrEmpty(regras.plano))
            {
                try
                {
                    string pp = regras.plano;
                    if (File.Exists(pp)) plano = Dialogos.Pesado(this, "Lendo Plano de Contas " + Path.GetFileName(pp) + "...", () => Plano.Carregar(pp));
                    else err.Append("• Plano não encontrado: ").Append(regras.plano).Append("\r\n");
                }
                catch (Exception ex) { err.Append("• ").Append(ex.Message).Append("\r\n"); }
            }
            foreach (var r in regras.razoes.ToList())
            {
                try
                {
                    if (!File.Exists(r)) { err.Append("• Razão não encontrado: ").Append(r).Append("\r\n"); continue; }
                    if (razao == null) razao = new Razao();
                    var rz = razao; var pl = plano;
                    Dialogos.Pesado(this, "Lendo Razão " + Path.GetFileName(r) + "...", () => { rz.Carregar(r, pl); return 0; });
                }
                catch (Exception ex) { err.Append("• ").Append(ex.Message).Append("\r\n"); }
            }
            if (razao != null && razao.movs.Count == 0) razao = null;
            ClienteDoRazao();
            funcionarios.Clear();   // quadro de funcionários substituído pela relação de líquidos (folha e adiantamento)
            CarregarFolhaDasRegras(err);
            if (err.Length > 0) BeginInvoke((Action)(() => Dialogos.Aviso(this, err.ToString().TrimEnd())));
        }

        void ClienteDoRazao()
        {
            if (razao == null || razao.empresa.Length == 0) return;
            bool antes = carregandoCampos;
            carregandoCampos = true;
            txtCliente.Text = razao.empresa;
            carregandoCampos = antes;
            regras.cliente = razao.empresa;
        }

        static readonly string FILTRO_BASE = "PDF, Excel ou CSV|*.pdf;*.xlsx;*.xlsm;*.xls;*.csv;*.txt;*.ods|Todos os arquivos|*.*";

        void CarregarPlano()
        {
            using (var fc = new OpenFileDialog { Filter = FILTRO_BASE, Title = "Plano de Contas do cliente (PDF, Excel ou CSV)", InitialDirectory = PastaBase() })
            {
                if (fc.ShowDialog(this) != DialogResult.OK) return;
                Prefs.Put("dirBase", Path.GetDirectoryName(fc.FileName));
                try
                {
                    string pp = fc.FileName;
                    plano = Dialogos.Pesado(this, "Lendo Plano de Contas " + Path.GetFileName(pp) + "...", () => Plano.Carregar(pp));
                    regras.plano = Path.GetFullPath(pp);
                    if (arquivoRegras != null) SalvarRegras();
                    int anal = plano.lista.Count(c => c.Analitica());
                    Revincular();
                    SugerirContaBanco();
                    status.Text = status.Text + "   |   Plano: " + plano.contas.Count + " contas (" + anal + " analíticas)";
                }
                catch (Exception ex) { Dialogos.Aviso(this, ex.Message); }
            }
        }

        void CarregarRazao()
        {
            using (var fc = new OpenFileDialog { Filter = FILTRO_BASE, Title = "Razão do cliente (PDF, Excel ou CSV) — pode selecionar vários", Multiselect = true, InitialDirectory = PastaBase() })
            {
                if (fc.ShowDialog(this) != DialogResult.OK) return;
                Prefs.Put("dirBase", Path.GetDirectoryName(fc.FileNames[0]));
                if (razao == null) razao = new Razao();
                var err = new StringBuilder();
                foreach (var f in fc.FileNames)
                {
                    try
                    {
                        var rz = razao; var pl = plano;
                        Dialogos.Pesado(this, "Lendo Razão " + Path.GetFileName(f) + "...", () => { rz.Carregar(f, pl); return 0; });
                        string abs = Path.GetFullPath(f);
                        if (!regras.razoes.Contains(abs)) regras.razoes.Add(abs);
                    }
                    catch (Exception ex) { err.Append("• ").Append(ex.Message).Append("\r\n"); }
                }
                if (razao.movs.Count == 0) razao = null;
                ClienteDoRazao();
                if (arquivoRegras != null) SalvarRegras();
                Revincular();
                if (err.Length > 0) Dialogos.Aviso(this, err.ToString().TrimEnd());
                if (plano == null) Dialogos.Aviso(this, "Carregue também o Plano de Contas: ele confirma que cada conta existe e é analítica,\r\ne identifica as contas de fornecedores/clientes.");
            }
        }

        void LimparBase()
        {
            plano = null; razao = null; funcionarios.Clear(); folha = null;
            regras.plano = ""; regras.razoes.Clear(); regras.funcionarios = ""; regras.folhaArquivos.Clear(); regras.folhaCfg = null;
            if (arquivoRegras != null) SalvarRegras();
            Revincular();
        }

        /// <summary>Conta banco vazia: lista as contas de disponível do Plano com a sugestão que combina com o extrato, e "Outra conta..." (F2).</summary>
        void SugerirContaBanco()
        {
            if (plano == null || regras.contaBanco != null) return;
            var disp = plano.ContasDisponivel().OrderBy(c => c.classificacao, StringComparer.Ordinal).ToList();
            string hist = extrato.Count == 0 ? "" : Regras.Normalizar(extrato[0].origem);
            Plano.Conta sug = null;
            foreach (var c in disp) foreach (var w in Classificador.Tokens(c.nome)) if (w.Length >= 4 && hist.Contains(w)) sug = c;
            var op = disp.Select(c => c.codigo + " - " + c.nome + "   (" + c.classificacao + ")").ToList();
            op.Add("Outra conta... (não está na lista: buscar em todo o Plano com F2)");
            int pre = sug == null ? 0 : disp.IndexOf(sug);
            int i = Dialogos.Lista(this, "Conta banco", "Qual é a conta deste banco no Plano de Contas?\r\n"
                + disp.Count + " conta(s) de banco/caixa/aplicação encontradas no Plano. Se o banco não estiver na lista,\r\n"
                + "escolha a última opção para procurar em todas as contas do Plano.", op, pre);
            if (i < 0) return;
            if (i == disp.Count)
            {
                var c = BuscaConta.Abrir(this, plano, "");
                if (c != null) txtBanco.Text = c.ToString();
                return;
            }
            txtBanco.Text = disp[i].codigo.ToString();
        }

        void VerAprendizado()
        {
            var linhas = new List<string[]>();
            foreach (var k0 in regras.ChavesAprendidas().ToList())
            {
                Dictionary<int, int> m;
                if (!regras.aprendido.TryGetValue(k0, out m)) continue;
                var k = k0.Split(new[] { '|' }, 3);
                if (k.Length < 3) continue;
                string tipo = k[0] == Regras.HIST ? "Histórico" : k[0] == Regras.DOC ? "CNPJ/CPF" : "Categoria da planilha";
                foreach (var c in m)
                {
                    var pc = plano == null ? null : plano.Get(c.Key);
                    linhas.Add(new[] { tipo, k[1] == "S" ? "Saída" : "Entrada", k[2], c.Key.ToString(), pc == null ? "" : pc.nome, c.Value.ToString(), k0 });
                }
            }
            linhas = linhas.OrderBy(x => x[0], StringComparer.Ordinal).ThenBy(x => x[2], StringComparer.Ordinal).ToList();
            List<string[]> sel;
            using (var f = new JanelaAprendizado(linhas))
            {
                f.ShowDialog(this);
                if (!f.Apagar) return;
                sel = f.Selecionadas();
            }
            if (sel.Count == 0) return;
            foreach (var l in sel)
            {
                Dictionary<int, int> mm;
                if (regras.aprendido.TryGetValue(l[6], out mm)) { mm.Remove(int.Parse(l[3])); if (mm.Count == 0) regras.Esquecer(l[6]); }
            }
            if (arquivoRegras != null) SalvarRegras();
            Revincular();
            status.Text = "Aprendizado: " + sel.Count + " item(ns) apagado(s).";
        }

        // ===================== exportação TXT =====================

        void AbrirArquivo(string p)
        {
            try { Process.Start(new ProcessStartInfo(p) { UseShellExecute = true }); }
            catch (Exception ex) { Dialogos.Aviso(this, "Não foi possível abrir " + p + ":\r\n" + ex.Message); }
        }

        void ExportarLancamentos()
        {
            if (extrato.Count == 0) { Dialogos.Aviso(this, "Adicione o extrato primeiro."); return; }
            if (!SincronizarCampos(true)) return;
            if (regras.contaBanco == null) { Dialogos.Aviso(this, "Informe o código da conta banco (Plano de Contas)."); txtBanco.Focus(); return; }
            if (plano != null && !plano.Valida(regras.contaBanco)) { Dialogos.Aviso(this, "A conta banco " + regras.contaBanco + " não existe no Plano de Contas ou é sintética."); return; }
            Revincular();
            if (arquivoRegras != null) SalvarRegras();

            int conf = extrato.Count(l => l.conferirSinal && !l.descartada);
            if (conf > 0 && !Dialogos.SimNao(this, "Conferir", conf + " lançamento(s) com entrada/saída deduzida (amarelos).\r\nJá conferiu? Exportar assim mesmo?")) return;
            bool incluirVAI = chkIncluirVAI.Checked;
            int identificados = extrato.Count(l => !l.descartada && l.debito != null && l.credito != null && (incluirVAI || !EmVAI(l)));
            if (identificados == 0)
            {
                int vai = extrato.Count(EmVAI);
                Dialogos.Aviso(this, "Nenhum lançamento para exportar." + (vai > 0 ? "\r\n" + vai + " lançamento(s) em Valores a Identificar continuam na tela."
                    + "\r\nCorrija as contas ou marque \"Exportar também os de Valores a Identificar (685)\" no menu Exportar." : ""));
                return;
            }
            int reexp = extrato.Count(l => !l.descartada && l.debito != null && l.credito != null && (incluirVAI || !EmVAI(l)) && l.exportadoEm.Length > 0);
            if (reexp > 0 && Dialogos.Escolher(this, "Reexportação", reexp + " lançamento(s) desta exportação JÁ FORAM EXPORTADOS antes (em azul).\r\n\r\n"
                + "Se o lote anterior já foi importado no sistema contábil, exclua/estorne esse lote lá antes de importar o novo,\r\n"
                + "para não duplicar. Continuar a exportação?", 1, "Continuar", "Cancelar") != 0) return;
            string dirLanc = @"Z:\CONTABIL\6. LANCTOS";
            string dirPadrao = Prefs.Pasta("dirExp", Directory.Exists(dirLanc) ? dirLanc : Prefs.Pasta("dirExt", null));
            string dest;
            using (var fc = new SaveFileDialog { Filter = "TXT de importação (*.txt)|*.txt", Title = "Salvar TXT de importação", InitialDirectory = dirPadrao, FileName = "lanctos.txt", OverwritePrompt = false })
            {
                if (fc.ShowDialog(this) != DialogResult.OK) return;
                dest = fc.FileName;
            }
            if (!dest.EndsWith(".txt", StringComparison.OrdinalIgnoreCase)) dest += ".txt";
            bool acrescentar = false;
            if (File.Exists(dest))
            {
                int r = Dialogos.Escolher(this, "Arquivo existente", Path.GetFileName(dest) + " já existe nesta pasta.\r\n\r\nSubstituir o arquivo ou acrescentar estes lançamentos ao final?",
                    1, "Substituir", "Acrescentar ao final", "Cancelar");
                if (r != 0 && r != 1) return;
                acrescentar = r == 1;
            }
            Prefs.Put("dirExp", Path.GetDirectoryName(dest));
            try
            {
                string rel = ExportarPara(dest, acrescentar);
                status.Text = "Exportado: " + dest + "   |   pendentes na tela: " + extrato.Count;
                if (Dialogos.Texto(this, "Relatório final", rel, "Abrir arquivo", "Fechar") == 0) AbrirArquivo(dest);
            }
            catch (Exception ex)
            {
                Log.Erro("exportar", ex);
                Dialogos.Aviso(this, "Erro ao gravar: " + ex.Message + "\r\n(O arquivo está aberto em outro programa?)");
            }
        }

        /// <summary>Grava o TXT, retira da tela os exportados (e as entradas de transferência descartadas), aprende e devolve o relatório.</summary>
        string ExportarPara(string dest, bool acrescentar)
        {
            bool incluirVAI = chkIncluirVAI.Checked;
            var saida = extrato.Where(l => !l.descartada && l.debito != null && l.credito != null && (incluirVAI || !EmVAI(l))).ToList();
            var exp = new HashSet<Lancamento>(saida);
            var descartadas = extrato.Where(l => l.descartada && l.par != null && exp.Contains(l.par)).ToList();
            int n = extrato.Count;
            int pendentes = n - saida.Count - descartadas.Count;
            var comp = new Dictionary<Lancamento, string>();
            foreach (var l in saida) comp[l] = Complemento(l);
            GravadorTxt.Gravar(dest, saida, comp, acrescentar);
            ultimaExportacao.Clear(); ultimaExportacao.AddRange(saida);
            string rel = Relatorio(dest, n, saida, descartadas, pendentes) + (acrescentar ? "\r\nLançamentos acrescentados ao final de " + Path.GetFileName(dest) : "");
            foreach (var l in saida) foreach (var it in l.planilha) it.exportado = l.data + " — " + Path.GetFileName(dest);
            foreach (var l in saida) if (l.comprovante != null) l.comprovante.exportado = l.data + " — " + Path.GetFileName(dest);
            foreach (var l in saida) foreach (var t in l.notas) t.exportado = l.data + " — " + l.historico + " (" + Path.GetFileName(dest) + ")";
            var sair = new HashSet<Lancamento>(saida.Concat(descartadas));
            string quando = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
            foreach (var l in saida.Concat(descartadas))
            {
                l.exportadoEm = Path.GetFileName(dest) + " (" + quando + ")";
                string c;
                l.complementoSalvo = comp.TryGetValue(l, out c) ? c : Complemento(l);
                string tb = TextoBusca(l);
                l.buscaSalva = tb.Substring(Math.Min(l.historico.Length, tb.Length)).Trim();
                if (l.buscaSalva.Length == 0) l.buscaSalva = null;
                if (!exportados.Contains(l)) exportados.Add(l);
            }
            int aprendidos = AprenderComExportados(saida);
            extrato.RemoveAll(sair.Contains);
            Revincular();
            string prog;
            try { string pp = SalvarProgresso(true); prog = pp == null ? "" : "\r\nProgresso salvo automaticamente: " + pp; }
            catch (Exception ex) { prog = "\r\n⚠ Progresso não salvo: " + ex.Message; }
            return rel + string.Format("\r\nRetirados da tela: {0} exportado(s) + {1} descartada(s).\r\nPermanecem na tela: {2} pendente(s) para corrigir e exportar depois.",
                       saida.Count, descartadas.Count, extrato.Count) + prog
                   + (aprendidos > 0 ? "\r\nAprendizado: " + aprendidos + " lançamento(s) memorizados para classificar os próximos extratos"
                                        + (arquivoRegras == null ? " (abra/crie as regras do cliente para guardar entre sessões)" : "") + "." : "");
        }

        /// <summary>Memoriza a conta de cada lançamento exportado (histórico, CNPJ/CPF, categoria). Correção manual pesa 3.</summary>
        int AprenderComExportados(List<Lancamento> saida)
        {
            int n = 0;
            foreach (var l in saida)
            {
                int? contra = l.saida ? l.debito : l.credito;
                if (contra == null || contra == regras.contaVAI || l.descartada || l.par != null) continue;
                if (bancoPorArquivo.Any(e => e.Value == contra) || contra == regras.contaBanco) continue;
                int peso = l.manual ? 3 : 1;
                string h = Classificador.ChaveHistorico(l.historico);
                if (h.Length > 0) regras.Aprender(Regras.HIST, l.saida, h, contra.Value, peso);
                var docs = Classificador.Documentos(TextoBusca(l));
                foreach (var it in l.planilha) foreach (var d0 in Classificador.Documentos(it.documento)) if (!docs.Contains(d0)) docs.Add(d0);
                foreach (var d in docs) regras.Aprender(Regras.DOC, l.saida, d, contra.Value, peso);
                string c = Classificador.CategoriaDaLinha(l);
                if (!string.IsNullOrEmpty(c)) regras.Aprender(Regras.CAT, l.saida, c, contra.Value, peso);
                if (h.Length > 0 || docs.Count > 0 || c != null) n++;
            }
            if (n > 0 && arquivoRegras != null) SalvarRegras();
            return n;
        }

        string Relatorio(string dest, int n, List<Lancamento> exp, List<Lancamento> descartadas, int pendentes)
        {
            var expSet = new HashSet<Lancamento>(exp);
            var descSet = new HashSet<Lancamento>(descartadas);
            var pend = extrato.Where(l => !expSet.Contains(l) && !descSet.Contains(l)).ToList();
            int vaiPend = pend.Count(EmVAI), vaiExp = exp.Count(EmVAI);
            int semBanco = pend.Count(l => !l.descartada && (l.debito == null || l.credito == null));
            var tiposVai = new Dictionary<string, int>();
            var criterios = new Dictionary<string, int>();
            foreach (var l in pend.Concat(exp))
            {
                if (!EmVAI(l)) continue;
                string t = Regex.Replace(Regex.Replace(l.historico, ":.*", ""), "[0-9].*", "").Trim();
                tiposVai.Somar(t.Length == 0 ? "(outros)" : t, 1);
            }
            foreach (var l in exp)
            {
                if (EmVAI(l)) continue;
                string o = Regex.Replace(l.obs, @"^Sinal[^|]*\| ?", "");
                string c = l.manual ? "Editado na tela"
                    : o.StartsWith("Razão") ? "Razão (lançamento anterior com banco)"
                    : o.StartsWith("Plano: conta do fornecedor") && o.Contains("despesa direta") ? "Plano: fornecedor no lugar da despesa direta do Razão"
                    : o.StartsWith("Plano: conta") ? "Plano: conta do terceiro (fornecedor/cliente)"
                    : o.StartsWith("Plano (natureza)") ? "Plano: natureza da operação"
                    : o.StartsWith("Funcionário") ? (o.Contains("salário") ? "Funcionário: salário" : "Funcionário: adiantamento")
                    : o.StartsWith("Regra") ? "Regra manual do cliente"
                    : o.StartsWith("Planilha financeira") ? "Planilha financeira: conta informada"
                    : o.StartsWith("Aprendido") ? "Aprendido em exportações anteriores"
                    : o.StartsWith("Transferência") ? "Transferência entre bancos" : Regex.Replace(o, "[:(\"\\[].*", "").Trim();
                criterios.Somar(c, 1);
            }
            int x = exp.Count, y = descartadas.Count;
            var sb = new StringBuilder();
            sb.Append("Arquivo: ").Append(Path.GetFileName(dest)).Append("\r\n\r\n");
            sb.AppendFormat("Controle: {0} na tela → {1} exportados + {2} descartadas + {3} pendentes = {4} {5}\r\n\r\n", n, x, y, pendentes, x + y + pendentes, x + y + pendentes == n ? "✓" : "✗ DIVERGENTE");
            sb.AppendFormat("1. Lançamentos processados: {0}\r\n", n);
            sb.AppendFormat("2. Exportados com Débito e Crédito identificados: {0}\r\n", x - vaiExp);
            sb.AppendFormat("3. Valores a Identificar ({0}): {1} pendente(s) na tela{2}\r\n", regras.contaVAI, vaiPend, vaiExp > 0 ? " + " + vaiExp + " exportado(s)" : " — não exportados");
            sb.AppendFormat("4. Sem banco/caixa identificado: {0} (pendentes na tela)\r\n", semBanco);
            sb.AppendFormat("5. Transferências descartadas: {0}\r\n", y);
            sb.Append("6. Principais tipos em Valores a Identificar:\r\n");
            foreach (var e in tiposVai.OrderByDescending(e => e.Value).Take(8)) sb.Append("     • ").Append(e.Key).Append(": ").Append(e.Value).Append("\r\n");
            if (tiposVai.Count == 0) sb.Append("     (nenhum)\r\n");
            sb.Append("7. Principais critérios nas classificações exportadas:\r\n");
            foreach (var e in criterios.OrderByDescending(e => e.Value).Take(8)) sb.Append("     • ").Append(e.Key).Append(": ").Append(e.Value).Append("\r\n");
            if (criterios.Count == 0) sb.Append("     (nenhum)\r\n");
            if (planilha.Count > 0)
            {
                int comPlan = exp.Count(Conciliado);
                int divPlan = exp.Count(l => l.conciliacao != null && !Conciliado(l));
                int naoLoc = planilha.Count(it => it.vinculo == null && it.exportado.Length == 0);
                sb.AppendFormat("\r\nPlanilha financeira: {0} exportado(s) conciliado(s) com a planilha | {1} vinculado(s) com divergência | {2} item(ns) da planilha não localizado(s) no extrato\r\n", comPlan, divPlan, naoLoc);
            }
            if (notas.Count > 0)
            {
                int comNF = exp.Count(l => l.notas.Count > 0);
                int nfNaoLoc = notas.Count(nf => !nf.cancelada && nf.titulos.All(t => t.vinculo == null && t.exportado.Length == 0));
                sb.AppendFormat("\r\nNotas fiscais (XML): {0} exportado(s) com nota localizada | {1} sem nota | {2} nota(s) não localizada(s) no extrato\r\n", comNF, exp.Count - comNF, nfNaoLoc);
            }
            return sb.ToString();
        }

        // ===================== progresso =====================

        void Salvar()
        {
            try { SalvarProgresso(arquivoProgresso != null); }
            catch (Exception ex) { Dialogos.Aviso(this, "Erro ao salvar o progresso: " + ex.Message); }
        }

        string PastaProgresso()
        {
            if (arquivoProgresso != null) return Path.GetDirectoryName(arquivoProgresso);
            if (arquivoRegras != null) return Path.GetDirectoryName(arquivoRegras);
            return Prefs.Pasta("dirProg", Prefs.Pasta("dirExp", PastaInicialRegras()));
        }

        string NomeProgressoPadrao()
        {
            var todos = exportados.Concat(extrato).ToList();
            string comp = todos.Count == 0 || todos[0].data == null || todos[0].data.Length < 10 ? "" : "_" + todos[0].data.Substring(6, 4) + "-" + todos[0].data.Substring(3, 2);
            string cli = Regex.Replace(Regex.Replace(string.IsNullOrEmpty(regras.cliente) ? "cliente" : regras.cliente, "[\\\\/:*?\"<>|]", ""), @"\s+", "_");
            if (cli.Length > 40) cli = cli.Substring(0, 40);
            return "progresso_" + cli + comp + DateTime.Now.ToString("_ddMMyyyy_HHmm") + ".ecprog";
        }

        /// <summary>Grava o progresso (tela + já exportados + Plano/Razão anexados). auto = grava sem perguntar no arquivo da sessão.</summary>
        string SalvarProgresso(bool auto)
        {
            if (extrato.Count == 0 && exportados.Count == 0) { if (!auto) Dialogos.Aviso(this, "Não há lançamentos para salvar."); return null; }
            SincronizarCampos(false);
            string dest = arquivoProgresso;
            if (!auto || dest == null)
            {
                string sug = dest != null ? Path.GetFileName(dest) : NomeProgressoPadrao();
                if (auto) dest = Path.Combine(PastaProgresso(), sug);
                else
                {
                    using (var fc = new SaveFileDialog { Filter = "Progresso do extrato (*.ecprog)|*.ecprog", Title = "Salvar progresso", InitialDirectory = PastaProgresso(), FileName = sug })
                    {
                        if (fc.ShowDialog(this) != DialogResult.OK) return null;
                        dest = fc.FileName;
                    }
                    if (!dest.EndsWith(".ecprog", StringComparison.OrdinalIgnoreCase)) dest += ".ecprog";
                }
            }
            var pr = new Progresso
            {
                cliente = regras.cliente ?? "",
                regras = arquivoRegras ?? "",
                contaBanco = regras.contaBanco,
                contaVAI = regras.contaVAI,
                salvoEm = DateTime.Now.ToString("dd/MM/yyyy HH:mm")
            };
            foreach (var e in bancoPorArquivo) pr.bancoPorArquivo[e.Key] = e.Value;
            pr.lancamentos.AddRange(exportados);
            pr.lancamentos.AddRange(extrato);
            var falhas = new StringBuilder();
            string pasta = Progresso.PastaAnexos(dest);
            if (plano != null && !string.IsNullOrEmpty(regras.plano)) pr.anexos.Add(ItemAnexo("PLANO", regras.plano, pasta, falhas).Campos());
            if (razao != null) foreach (var z in regras.razoes) pr.anexos.Add(ItemAnexo("RAZAO", z, pasta, falhas).Campos());
            foreach (var it in AnexosComprovantes(pasta, falhas)) pr.anexos.Add(it.Campos());
            var extra = new Dictionary<Lancamento, string[]>();
            foreach (var l in extrato)
            {
                string tb = TextoBusca(l), ex2 = tb.Length > l.historico.Length ? tb.Substring(l.historico.Length).Trim() : "";
                extra[l] = new[] { Complemento(l), ex2 };
            }
            try { Progresso.Gravar(dest, pr, extra); }
            catch (Exception ex)
            {
                if (auto) throw;
                Dialogos.Aviso(this, "Erro ao salvar o progresso: " + ex.Message);
                return null;
            }
            arquivoProgresso = dest;
            Prefs.Put("dirProg", Path.GetDirectoryName(dest));
            alterado = false;
            AdicionarRecente(dest);
            AtualizarTitulo();
            bool temPlano = pr.anexos.Any(a => a[0] == "PLANO");
            int nRaz = pr.anexos.Count(a => a[0] == "RAZAO"), nComp = pr.anexos.Count(a => a[0] == "COMP");
            var partes = new List<string>();
            if (temPlano) partes.Add("Plano");
            if (nRaz > 0) partes.Add(nRaz + " Razão");
            if (nComp > 0) partes.Add(nComp + " arquivo(s) de comprovantes");
            status.Text = status.Text + "   |   Progresso salvo: " + Path.GetFileName(dest) + (partes.Count == 0 ? "" : " (com " + string.Join(", ", partes) + ")");
            if (falhas.Length > 0 && !auto) Dialogos.Aviso(this, "Progresso salvo, mas não consegui guardar cópia de:\r\n" + falhas + "\r\nEsses arquivos serão procurados no local original ao reabrir.");
            return dest;
        }

        /// <summary>Arquivos de comprovantes da sessão (com as páginas que ficaram), copiados para a pasta de anexos.</summary>
        List<AnexosSessao.Item> AnexosComprovantes(string pasta, StringBuilder falhas)
        {
            var r = new List<AnexosSessao.Item>();
            var pags = new Dictionary<string, SortedSet<int>>();
            var ordem = new List<string>();
            foreach (var c in comps) { if (!pags.ContainsKey(c.arquivo)) { pags[c.arquivo] = new SortedSet<int>(); ordem.Add(c.arquivo); } pags[c.arquivo].Add(c.pagina); }
            foreach (var a in ordem)
            {
                string orig;
                if (!origemComp.TryGetValue(a, out orig)) { falhas.Append("• ").Append(a).Append(" (caminho do arquivo desconhecido)\r\n"); continue; }
                var it = ItemAnexo("COMP", orig, pasta, falhas);
                it.paginas = string.Join(",", pags[a]);
                r.Add(it);
            }
            return r;
        }

        static AnexosSessao.Item ItemAnexo(string tipo, string orig, string pasta, StringBuilder falhas)
        {
            var it = new AnexosSessao.Item(tipo, Path.GetFullPath(orig));
            it.copia = AnexosSessao.Copiar(orig, pasta);
            if (it.copia.Length == 0) falhas.Append("• ").Append(Path.GetFileName(orig)).Append("\r\n");
            return it;
        }

        /// <summary>Pergunta se salva antes de limpar/abrir/sair. false = cancelar.</summary>
        bool PodeDescartar(string titulo, string pergunta, string sim, string nao)
        {
            if (extrato.Count == 0 && exportados.Count == 0 || !alterado) return true;
            int r = Dialogos.Escolher(this, titulo, pergunta, 0, sim, nao, "Cancelar");
            if (r == 2 || r < 0) return false;
            if (r == 0)
            {
                try { return SalvarProgresso(false) != null; }
                catch (Exception ex) { Dialogos.Aviso(this, ex.Message); return false; }
            }
            return true;
        }

        void AbrirProgresso(string direto)
        {
            if (!PodeDescartar("Abrir progresso", "Há lançamentos na tela que não foram salvos no progresso.", "Salvar e abrir", "Abrir sem salvar")) return;
            string p = direto;
            if (p == null)
            {
                using (var fc = new OpenFileDialog { Filter = "Progresso do extrato (*.ecprog)|*.ecprog|Todos os arquivos|*.*", Title = "Abrir progresso salvo", InitialDirectory = PastaProgresso() })
                {
                    if (fc.ShowDialog(this) != DialogResult.OK) return;
                    p = fc.FileName;
                }
            }
            if (!File.Exists(p)) { Dialogos.Aviso(this, "Arquivo não encontrado:\r\n" + p); RemoverRecente(p); return; }
            try
            {
                var pr = Progresso.Ler(p);
                int exp = pr.lancamentos.Count(l => l.exportadoEm.Length > 0);
                int pend = pr.lancamentos.Count - exp;
                bool trazer = true;
                if (exp > 0)
                {
                    int r = Dialogos.Escolher(this, "Abrir progresso", "Progresso de " + (pr.cliente.Length == 0 ? "cliente" : pr.cliente) + " salvo em " + pr.salvoEm + ":\r\n\r\n"
                        + "   " + pend + " lançamento(s) pendente(s) (não exportados)\r\n"
                        + "   " + exp + " lançamento(s) já exportado(s)\r\n\r\n"
                        + "Trazer também os já exportados para a tela (para alterar e exportar de novo)?", 0, "Trazer todos para a tela", "Só os pendentes", "Cancelar");
                    if (r == 2 || r < 0) return;
                    trazer = r == 0;
                }
                CarregarProgresso(pr, p, trazer);
            }
            catch (Exception ex)
            {
                Log.Erro("abrir progresso", ex);
                Dialogos.Aviso(this, "Erro ao abrir o progresso: " + ex.Message);
            }
        }

        void CarregarProgresso(Progresso pr, string p, bool trazerExportados)
        {
            if (pr.regras.Length > 0 && File.Exists(pr.regras) && (arquivoRegras == null || !string.Equals(arquivoRegras, pr.regras, StringComparison.OrdinalIgnoreCase)))
                CarregarRegras(pr.regras);
            carregandoCampos = true;
            try
            {
                if (pr.cliente.Length > 0) txtCliente.Text = pr.cliente;
                if (pr.contaBanco != null) txtBanco.Text = pr.contaBanco.ToString();
                txtVAI.Text = pr.contaVAI.ToString();
            }
            finally { carregandoCampos = false; }
            SincronizarCampos(false);
            extrato.Clear(); exportados.Clear(); bancoPorArquivo.Clear();
            comps.Clear(); origemComp.Clear();
            foreach (var e in pr.bancoPorArquivo) bancoPorArquivo.Add(new KeyValuePair<string, int?>(e.Key, e.Value));
            foreach (var l in pr.lancamentos)
            {
                if (l.exportadoEm.Length == 0) extrato.Add(l);
                else if (trazerExportados) { l.manual = true; extrato.Add(l); }
                else exportados.Add(l);
            }
            extrato.OrdenarEstavel((a, b) => a.seq.CompareTo(b.seq));
            foreach (var l in pr.lancamentos) sequencia = Math.Max(sequencia, l.seq);
            FiltrarNovos(extrato.Concat(exportados).ToList());   // chaves para não importar de novo o mesmo extrato
            string faltando = RestaurarAnexos(pr.anexos.Select(AnexosSessao.Item.De).Where(x => x != null).ToList(), Progresso.PastaAnexos(p));
            arquivoProgresso = p;
            Revincular();
            alterado = false;
            AdicionarRecente(p);
            AtualizarTitulo();
            int az = extrato.Count(l => l.exportadoEm.Length > 0);
            status.Text = "Progresso aberto: " + Path.GetFileName(p) + " | " + extrato.Count + " lançamento(s) na tela"
                          + (az > 0 ? " (" + az + " já exportados, em azul)" : "") + (exportados.Count == 0 ? "" : " | " + exportados.Count + " exportado(s) guardados")
                          + (comps.Count == 0 ? "" : " | " + comps.Count + " comprovante(s)") + (razao != null ? " | Razão" : "") + (plano != null ? " | Plano" : "");
            if (faltando.Length > 0) Dialogos.Aviso(this, "Progresso aberto, mas estes anexos não foram encontrados (nem o original nem a cópia):\r\n" + faltando);
        }

        /// <summary>Recarrega Plano e Razão guardados no progresso (o original ou a cópia). Devolve os que não foram encontrados.</summary>
        string RestaurarAnexos(List<AnexosSessao.Item> anexos, string pasta)
        {
            var falta = new StringBuilder();
            var erros = new StringBuilder();
            bool mudouBase = false;
            foreach (var it in anexos.Where(a => a.tipo == "PLANO"))
            {
                string f = AnexosSessao.Resolver(it, pasta);
                if (f == null) { falta.Append("• Plano de Contas: ").Append(AnexosSessao.NomeOriginal(it)).Append("\r\n"); continue; }
                if (plano != null && (it.original == regras.plano || f == regras.plano)) continue;
                try { plano = Dialogos.Pesado(this, "Lendo Plano de Contas " + Path.GetFileName(f) + "...", () => Plano.Carregar(f)); regras.plano = Path.GetFullPath(f); mudouBase = true; }
                catch (Exception ex) { erros.Append("• ").Append(ex.Message).Append("\r\n"); }
            }
            foreach (var it in anexos.Where(a => a.tipo == "RAZAO"))
            {
                string f = AnexosSessao.Resolver(it, pasta);
                if (f == null) { falta.Append("• Razão: ").Append(AnexosSessao.NomeOriginal(it)).Append("\r\n"); continue; }
                if (razao != null && (regras.razoes.Contains(it.original) || regras.razoes.Contains(Path.GetFullPath(f)))) continue;
                try
                {
                    if (razao == null) razao = new Razao();
                    var rz = razao; var pl = plano;
                    Dialogos.Pesado(this, "Lendo Razão " + Path.GetFileName(f) + "...", () => { rz.Carregar(f, pl); return 0; });
                    regras.razoes.Add(Path.GetFullPath(f));
                    mudouBase = true;
                }
                catch (Exception ex) { erros.Append("• ").Append(ex.Message).Append("\r\n"); }
            }
            if (razao != null && razao.movs.Count == 0) razao = null;
            if (mudouBase)
            {
                if (string.IsNullOrEmpty(regras.cliente)) ClienteDoRazao();
                if (arquivoRegras != null) SalvarRegras();
            }
            var jaTem = new HashSet<string>(comps.Select(c => c.arquivo));
            foreach (var it in anexos.Where(a => a.tipo == "COMP"))
            {
                string nome = AnexosSessao.NomeOriginal(it);
                if (jaTem.Contains(nome)) continue;
                string f = AnexosSessao.Resolver(it, pasta);
                if (f == null) { falta.Append("• Comprovante: ").Append(nome).Append("\r\n"); continue; }
                var pg = new HashSet<int>();
                foreach (var x in it.paginas.Split(',')) { int v; if (int.TryParse(x.Trim(), out v)) pg.Add(v); }
                try
                {
                    foreach (var c in Dialogos.Pesado(this, "Lendo comprovantes " + nome + "...", () => LeitorComprovante.Ler(f)))
                    {
                        if (pg.Count > 0 && !pg.Contains(c.pagina)) continue;   // páginas excluídas pelo usuário não voltam
                        c.arquivo = nome;
                        comps.Add(c);
                    }
                    origemComp[nome] = Path.GetFullPath(f);
                    jaTem.Add(nome);
                }
                catch (Exception ex) { erros.Append("• ").Append(nome).Append(": ").Append(ex.Message).Append("\r\n"); }
            }
            if (erros.Length > 0) Dialogos.Aviso(this, "Erro ao reabrir anexos do progresso:\r\n" + erros);
            return falta.ToString();
        }

        void NovaSessao()
        {
            if (!PodeDescartar("Novo", "Salvar o progresso atual antes de limpar?", "Salvar", "Não salvar")) return;
            string cli = string.IsNullOrWhiteSpace(regras.cliente) ? "do cliente atual" : "de " + regras.cliente.Trim();
            int t = Dialogos.Escolher(this, "Novo", "Como deseja começar?\r\n\r\n"
                + "• Mesmo cliente: limpa extratos, comprovantes, notas e planilha, e mantém as regras, o Razão, o Plano de Contas e a conta banco " + cli + ".\r\n\r\n"
                + "• Nova empresa (do zero): limpa também regras, Razão, Plano de Contas, conta banco e cliente.\r\n"
                + "   Ao carregar o Plano de Contas da nova empresa, o programa pergunta o banco de novo.", 0, "Mesmo cliente", "Nova empresa (do zero)", "Cancelar");
            if (t == 2 || t < 0) return;
            extrato.Clear(); exportados.Clear(); bancoPorArquivo.Clear(); ultimaExportacao.Clear();
            comps.Clear(); origemComp.Clear(); notas.Clear(); planilha.Clear(); arquivosPlanilha.Clear();
            txtCnpjEmp.Text = "";
            arquivoProgresso = null;
            if (t == 1)
            {
                regras = new Regras();
                arquivoRegras = null;
                plano = null; razao = null; funcionarios.Clear(); folha = null;
                Prefs.Remove("regras");
                txtRegras.Text = "";
                ColocarCampos();
            }
            txtPesquisa.Text = "";
            Revincular();
            alterado = false;
            AtualizarTitulo();
            status.Text = t == 1 ? "Nova empresa: carregue o Razão e o Plano de Contas (o programa vai perguntar a conta do banco) e adicione o extrato."
                                 : "Novo: adicione o extrato do cliente.";
        }

        void Fechando(object s, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.WindowsShutDown) return;
            if (!PodeDescartar("Sair", "Salvar o progresso antes de sair?\r\n(Permite resgatar e alterar os lançamentos depois.)", "Salvar e sair", "Sair sem salvar")) e.Cancel = true;
        }

        // ---------------- recentes ----------------

        List<string> Recentes() { return Prefs.Get("recentes", "").Split('\n').Select(x => x.Trim()).Where(x => x.Length > 0).ToList(); }

        void AdicionarRecente(string p)
        {
            var r = Recentes();
            string a = Path.GetFullPath(p);
            r.RemoveAll(x => string.Equals(x, a, StringComparison.OrdinalIgnoreCase));
            r.Insert(0, a);
            while (r.Count > 8) r.RemoveAt(r.Count - 1);
            Prefs.Put("recentes", string.Join("\n", r));
            AtualizarRecentes();
        }

        void RemoverRecente(string p)
        {
            var r = Recentes();
            r.RemoveAll(x => string.Equals(x, Path.GetFullPath(p), StringComparison.OrdinalIgnoreCase));
            Prefs.Put("recentes", string.Join("\n", r));
            AtualizarRecentes();
        }

        void AtualizarRecentes()
        {
            menuRecentes.DropDownItems.Clear();
            var r = Recentes();
            if (r.Count == 0) { menuRecentes.DropDownItems.Add(new ToolStripMenuItem("(nenhum)") { Enabled = false }); return; }
            int i = 1;
            foreach (var s in r)
            {
                string p = s;
                var it = new ToolStripMenuItem(i++ + "  " + Path.GetFileName(p)) { ToolTipText = p };
                it.Click += (o, e) => Executar(() => AbrirProgresso(p));
                menuRecentes.DropDownItems.Add(it);
            }
        }
    }
}

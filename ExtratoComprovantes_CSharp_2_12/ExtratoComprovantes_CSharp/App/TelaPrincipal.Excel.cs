using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Exportar > Exportar para Excel: lançamentos do extrato (com contas, situação, comprovante, nota, planilha e conciliação)
    /// filtrados por escopo, sentido e período, e abas de apoio (resumo por conta, planilha financeira, notas e comprovantes).
    /// Não marca nada como exportado — é só para análise.
    /// </summary>
    public partial class TelaPrincipal
    {
        void ExportarExcel()
        {
            if (extrato.Count == 0 && planilha.Count == 0 && notas.Count == 0 && comps.Count == 0) { Dialogos.Aviso(this, "Nada para exportar: adicione o extrato primeiro."); return; }
            SincronizarCampos(false);
            var selecionados = Selecionados();
            var naTela = vis.Where(i => i < extrato.Count).Select(i => extrato[i]).ToList();

            var f = new Form
            {
                Text = "Exportar para Excel (análise)", Font = Dialogos.Fonte, Size = new Size(860, 420), StartPosition = FormStartPosition.CenterParent,
                ShowInTaskbar = false, MinimizeBox = false, MaximizeBox = false, MinimumSize = new Size(700, 400)
            };
            // ---------- 1. lançamentos ----------
            var g1 = new GroupBox { Text = "1. Quais lançamentos do extrato", Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 4, 8, 4) };
            var t1 = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1 };
            var rbTodos = new RadioButton { Name = "rbTodos", Text = "Todos os lançamentos do extrato (" + extrato.Count + ")", AutoSize = true, Checked = true };
            var rbTela = new RadioButton { Name = "rbTela", Text = "Os da tela, com a pesquisa e os filtros aplicados (" + naTela.Count + ")", AutoSize = true, Enabled = naTela.Count > 0 && naTela.Count < extrato.Count };
            var rbSel = new RadioButton { Name = "rbSel", Text = "Linhas selecionadas na tabela (" + selecionados.Count + ")", AutoSize = true, Enabled = selecionados.Count > 1 };
            var filt = new FlowLayoutPanel { AutoSize = true, WrapContents = true };
            var cbS = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 150 };
            cbS.Items.AddRange(new object[] { "Saídas e entradas", "Só saídas (D)", "Só entradas (C)" });
            cbS.SelectedIndex = 0;
            var chk685 = new CheckBox { Text = "Somente Valores a Identificar (" + regras.contaVAI + ")", AutoSize = true, Padding = new Padding(8, 3, 0, 0) };
            var chkDesc = new CheckBox { Text = "Incluir transferências descartadas", AutoSize = true, Checked = true, Padding = new Padding(8, 3, 0, 0) };
            filt.Controls.Add(Rot("Sentido:")); filt.Controls.Add(cbS); filt.Controls.Add(chk685); filt.Controls.Add(chkDesc);
            var periodo = new FiltroPeriodo(dicas);
            foreach (Control c in new Control[] { rbTodos, rbTela, rbSel, filt, periodo.Painel }) t1.Controls.Add(c);
            g1.Controls.Add(t1);

            // ---------- 2. abas ----------
            var g2 = new GroupBox { Text = "2. Abas da planilha", Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 4, 8, 4) };
            var t2 = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true };
            var cLanc = new CheckBox { Text = "Lançamentos", AutoSize = true, Checked = true, Enabled = false };
            var cResumo = new CheckBox { Text = "Resumo por contrapartida", AutoSize = true, Checked = true };
            var cPlan = new CheckBox { Text = "Planilha financeira (" + planilha.Count + ")", AutoSize = true, Checked = planilha.Count > 0, Enabled = planilha.Count > 0 };
            var cNF = new CheckBox { Text = "Notas fiscais (" + notas.Count + ")", AutoSize = true, Checked = notas.Count > 0, Enabled = notas.Count > 0 };
            var cComp = new CheckBox { Text = "Comprovantes (" + comps.Count + ")", AutoSize = true, Checked = comps.Count > 0, Enabled = comps.Count > 0 };
            foreach (var c in new[] { cLanc, cResumo, cPlan, cNF, cComp }) t2.Controls.Add(c);
            dicas.SetToolTip(cResumo, "Totais por conta de contrapartida (quantidade, entradas, saídas e líquido) dos lançamentos exportados.");
            dicas.SetToolTip(cPlan, "Todos os itens da planilha financeira, com a situação (conciliado, divergente, não localizado) — sem o filtro de período.");
            g2.Controls.Add(t2);

            var lblQtd = new Label { AutoSize = true, Padding = new Padding(4, 8, 4, 0), Font = new Font(Dialogos.Fonte, FontStyle.Bold) };
            var bts = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(6) };
            var bOk = new Button { Name = "bExportarExcel", Text = "Exportar...", MinimumSize = new Size(110, 30), AutoSize = true };
            var bCa = new Button { Text = "Cancelar", MinimumSize = new Size(100, 30), AutoSize = true };
            bts.Controls.Add(bCa); bts.Controls.Add(bOk); bts.Controls.Add(lblQtd);
            var lay = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(6) };
            lay.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            lay.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            lay.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            lay.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            lay.Controls.Add(g1, 0, 0); lay.Controls.Add(g2, 0, 1); lay.Controls.Add(new Panel(), 0, 2); lay.Controls.Add(bts, 0, 3);
            f.Controls.Add(lay);
            f.AcceptButton = bOk; f.CancelButton = bCa;

            Func<List<Lancamento>> escolhidos = () =>
            {
                var origem = rbSel.Checked ? selecionados : rbTela.Checked ? naTela : extrato;
                int sentido = cbS.SelectedIndex;
                return origem.Where(l => !(sentido == 1 && !l.saida || sentido == 2 && l.saida)
                                         && (!chk685.Checked || EmVAI(l))
                                         && (chkDesc.Checked || !l.descartada)
                                         && periodo.Contem(l.data)).ToList();
            };
            Action contar = () =>
            {
                var l = escolhidos();
                double ent = l.Where(x => !x.saida).Sum(x => x.valor), sai = l.Where(x => x.saida).Sum(x => x.valor);
                lblQtd.Text = l.Count + " lançamento(s) | Entradas R$ " + Moeda(ent) + " | Saídas R$ " + Moeda(sai) + "  (" + periodo.Descricao() + ")";
            };
            foreach (var rb in new[] { rbTodos, rbTela, rbSel }) rb.CheckedChanged += (s, e) => contar();
            cbS.SelectedIndexChanged += (s, e) => contar();
            chk685.CheckedChanged += (s, e) => contar();
            chkDesc.CheckedChanged += (s, e) => contar();
            periodo.Mudou += contar;
            bCa.Click += (s, e) => f.Close();
            bOk.Click += (s, e) =>
            {
                if (periodo.Erro) { Dialogos.Aviso(f, "Corrija o período (data incompleta/inválida ou inicial maior que a final)."); return; }
                var lista = escolhidos();
                if (lista.Count == 0 && !(cPlan.Checked || cNF.Checked || cComp.Checked)) { Dialogos.Aviso(f, "Nenhum lançamento com esses filtros."); return; }
                string dest = ArquivoExcel != null ? ArquivoExcel(lista) : PedirArquivoExcel(f);
                if (dest == null) return;
                // lê as opções antes: durante a gravação a janela fica desativada
                bool oRes = cResumo.Checked, oPlan = cPlan.Checked && planilha.Count > 0, oNF = cNF.Checked && notas.Count > 0, oComp = cComp.Checked && comps.Count > 0;
                string oPer = periodo.Descricao();
                try
                {
                    Dialogos.Pesado(f, "Gravando " + Path.GetFileName(dest) + "...", () =>
                    {
                        GravarExcel(dest, lista, oRes, oPlan, oNF, oComp, oPer);
                        return 0;
                    });
                }
                catch (IOException ex)
                {
                    Dialogos.Aviso(f, "Não foi possível gravar " + Path.GetFileName(dest) + ".\r\nSe o arquivo estiver aberto no Excel, feche e tente de novo.\r\n\r\n" + ex.Message);
                    return;
                }
                f.Tag = dest;
                f.Close();
            };
            f.Shown += (s, e) => contar();
            ExcelAberto = f;
            try { f.ShowDialog(this); }
            finally { ExcelAberto = null; f.Dispose(); }
            var arq = f.Tag as string;
            if (arq == null) return;
            status.Text = status.Text + "   |   Excel gravado: " + Path.GetFileName(arq);
            if (ArquivoExcel == null && Dialogos.SimNao(this, "Exportar para Excel", "Planilha gravada:\r\n" + arq + "\r\n\r\nAbrir agora?"))
            {
                try { Process.Start(new ProcessStartInfo(arq) { UseShellExecute = true }); }
                catch (Exception ex) { Dialogos.Aviso(this, "Não foi possível abrir: " + ex.Message); }
            }
        }

        /// <summary>Janela aberta e escolha do arquivo (usados pelo teste automático).</summary>
        Form ExcelAberto;
        Func<List<Lancamento>, string> ArquivoExcel;

        string PedirArquivoExcel(IWin32Window dono)
        {
            string cli = string.IsNullOrWhiteSpace(regras.cliente) ? "" : "_" + string.Join("_", regras.cliente.Split(Path.GetInvalidFileNameChars()));
            using (var fc = new SaveFileDialog
            {
                Filter = "Pasta de trabalho do Excel (*.xlsx)|*.xlsx", Title = "Exportar para Excel",
                InitialDirectory = Prefs.Pasta("dirExcel", Prefs.Pasta("dirExp", Prefs.Pasta("dirExt", null))),
                FileName = "analise_lancamentos" + cli + ".xlsx"
            })
            {
                if (fc.ShowDialog(dono) != DialogResult.OK) return null;
                Prefs.Put("dirExcel", Path.GetDirectoryName(fc.FileName));
                return fc.FileName;
            }
        }

        string NomeConta(int? c)
        {
            if (c == null || plano == null) return "";
            var pc = plano.Get(c);
            return pc == null ? "(não existe no Plano)" : pc.nome;
        }

        static object DataCel(string d) { var x = J.Data(d); return x == null ? (object)d : x.Value; }
        static object DataCel(DateTime? d) { return d == null ? null : (object)d.Value; }

        string SituacaoLanc(Lancamento l)
        {
            if (l.descartada) return "Transferência descartada";
            if ((l.saida ? l.credito : l.debito) == null) return "Sem banco/caixa";
            if (Em685(l)) return "Valores a Identificar";
            if (l.debito == null || l.credito == null) return "Sem contrapartida";
            return "Identificado";
        }

        void GravarExcel(string dest, List<Lancamento> lista, bool resumo, bool comPlan, bool comNF, bool comComp, string periodo)
        {
            var x = new GravadorXlsx();
            string emp = CnpjEmpresa();

            // ---------- lançamentos ----------
            var a = x.NovaAba("Lançamentos");
            a.Add("Data", "Histórico", "Complemento (TXT)", "D/C", "Valor", "Valor com sinal", "Débito", "Conta débito", "Crédito", "Conta crédito",
                  "Contrapartida", "Nome da contrapartida", "Banco/caixa", "Situação", "Observação", "Comprovante (favorecido)", "CPF/CNPJ favorecido",
                  "Nota fiscal (XML)", "Planilha financeira", "Categoria (planilha)", "Conta (planilha)", "Conciliação (planilha)", "Já exportado em", "Arquivo");
            foreach (var l in lista)
            {
                int? contra = l.descartada ? (int?)null : l.saida ? l.debito : l.credito;
                var k = l.comprovante;
                string contaPlan = Juntar(l.planilha.Where(it => it.conta != null).Select(it => it.conta.ToString()));
                a.Add(DataCel(l.data), l.historico, l.descartada ? "" : Complemento(l), l.saida ? "D" : "C", l.valor, l.saida ? -l.valor : l.valor,
                      (object)l.debito, NomeConta(l.debito), (object)l.credito, NomeConta(l.credito),
                      (object)contra, NomeConta(contra), (object)BancoDe(l), SituacaoLanc(l), l.manual ? "Editado manualmente" : l.obs,
                      k == null ? "" : k.favorecido, k == null ? "" : k.docFavorecido,
                      NotaTexto(l), Juntar(l.planilha.Select(it => it.Resumo())), Juntar(l.planilha.Select(it => it.categoria)),
                      contaPlan.Length > 0 && !contaPlan.Contains(" ") ? (object)int.Parse(contaPlan) : contaPlan,
                      l.conciliacao == null ? "" : l.conciliacao.Texto(ExigirNF()), l.exportadoEm, l.origem);
            }

            // ---------- resumo por contrapartida ----------
            if (resumo)
            {
                var r = x.NovaAba("Resumo por contrapartida");
                r.Add("Período: " + periodo + " | " + lista.Count + " lançamento(s)");
                r.Add();
                r.Add("Contrapartida", "Nome", "Qtde", "Entradas (C)", "Saídas (D)", "Líquido (entradas − saídas)");
                var grupos = lista.Where(l => !l.descartada).GroupBy(l => l.saida ? l.debito : l.credito)
                                  .OrderByDescending(g => g.Sum(l => l.valor));
                foreach (var g in grupos)
                {
                    double ent = g.Where(l => !l.saida).Sum(l => l.valor), sai = g.Where(l => l.saida).Sum(l => l.valor);
                    r.Add(g.Key == null ? (object)"(sem conta)" : g.Key.Value, g.Key == null ? "" : NomeConta(g.Key) + (g.Key == regras.contaVAI ? " (Valores a Identificar)" : ""),
                          g.Count(), Math.Round(ent, 2), Math.Round(sai, 2), Math.Round(ent - sai, 2));
                }
                var validos = lista.Where(l => !l.descartada).ToList();
                double te = validos.Where(l => !l.saida).Sum(l => l.valor), ts = validos.Where(l => l.saida).Sum(l => l.valor);
                r.Add("TOTAL", "", validos.Count, Math.Round(te, 2), Math.Round(ts, 2), Math.Round(te - ts, 2));
                int desc = lista.Count(l => l.descartada);
                if (desc > 0) { r.Add(); r.Add("Transferências descartadas (fora dos totais): " + desc); }
            }

            // ---------- planilha financeira ----------
            if (comPlan)
            {
                var p = x.NovaAba("Planilha financeira");
                p.Add("Data", "Outras datas", "Favorecido / Cliente", "CPF/CNPJ", "Descrição", "NF", "Categoria", "Conta", "Valor", "E/S",
                      "Situação", "Data no extrato", "Histórico no extrato", "Arquivo", "Linha");
                foreach (var it in planilha)
                    p.Add(DataCel(it.Data()), string.Join(" ", it.datas.Skip(1)), it.favorecido, it.documento, it.descricao, string.Join(", ", it.nfs),
                          it.categoria, (object)it.conta, it.valor, it.saida == null ? "" : it.saida.Value ? "S" : "E", Situacao(it),
                          it.vinculo == null ? null : DataCel(it.vinculo.data), it.vinculo == null ? "" : it.vinculo.historico, it.arquivo, it.linha);
            }

            // ---------- notas fiscais (uma linha por parcela) ----------
            if (comNF)
            {
                var n = x.NovaAba("Notas fiscais");
                n.Add("Tipo", "Número", "Série", "Emissão", "Emitente", "CNPJ emitente", "Destinatário / Tomador", "CNPJ destinatário", "Operação",
                      "Valor da nota", "Valor líquido (NFS-e)", "Parcela", "Valor da parcela", "Vencimento", "Cancelada", "Situação da parcela",
                      "Data no extrato", "Histórico no extrato", "Chave", "Arquivo");
                foreach (var nf in notas)
                {
                    var cp = nf.Compra(emp);
                    string op = cp == null ? "?" : cp.Value ? "Compra/Tomada" : "Venda/Prestação";
                    var tits = nf.titulos.Count > 0 ? nf.titulos : new List<NotaFiscal.Titulo> { null };
                    foreach (var t in tits)
                    {
                        string sit = nf.cancelada ? "Cancelada" : t == null ? "" : t.exportado.Length > 0 ? "Já exportada — " + t.exportado
                                   : t.vinculo != null ? "Localizada no extrato" : "Não localizada no extrato";
                        n.Add(nf.tipo, nf.numero, nf.serie, DataCel(nf.emissao), nf.emitNome, FormatarDoc(nf.emitCnpj), nf.destNome, FormatarDoc(nf.destCnpj), op,
                              nf.valor, nf.valorLiquido > 0 ? (object)nf.valorLiquido : null, t == null ? "" : t.parcela, t == null ? null : (object)t.valor,
                              t == null ? null : DataCel(t.vencimento), nf.cancelada ? "Sim" : "Não", sit,
                              t == null || t.vinculo == null ? null : DataCel(t.vinculo.data), t == null || t.vinculo == null ? "" : t.vinculo.historico, nf.chave, nf.arquivo);
                    }
                }
            }

            // ---------- comprovantes ----------
            if (comComp)
            {
                var c = x.NovaAba("Comprovantes");
                c.Add("Arquivo", "Página", "Tipo", "Data", "Valor", "Favorecido", "CPF/CNPJ", "Pagador", "Cód. barras / ID", "Descrição", "Situação");
                foreach (var k in comps)
                    c.Add(k.arquivo, k.pagina, k.tipo, DataCel(k.data), k.valor == 0 ? null : (object)k.valor, k.favorecido, k.docFavorecido, k.pagador,
                          k.codigoBarras.Length > 0 ? k.codigoBarras : k.idTransacao, k.descricao, SituacaoComp(k));
            }
            x.Gravar(dest);
        }
    }
}

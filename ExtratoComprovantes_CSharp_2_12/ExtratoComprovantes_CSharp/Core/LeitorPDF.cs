using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Leitura genérica de extrato em PDF (com texto, não escaneado): linhas "data + histórico + valor [+ saldo]";
    /// entrada/saída por sinal explícito, variação do saldo ou palavras-chave.
    /// </summary>
    public static class LeitorPDF
    {
        static readonly Regex DATA = new Regex(@"^([0-9]{2})[/.-]([0-9]{2})(?:[/.-]([0-9]{2,4}))?(?=\s|$)");
        static readonly Regex VALOR = new Regex(@"(?<![0-9.,/])(-\s?)?(?:R\$\s?)?(\()?([0-9]{1,3}(?:\.[0-9]{3})+|[0-9]+),([0-9]{2})(\))?(?![0-9/%]|\s%)(\s?-(?!\s?R\$|\s?[0-9]))?(?:\s?([DC])(?![A-Za-zÀ-ú]))?");
        static readonly Regex ANO = new Regex(@"\b[0-9]{2}/[0-9]{2}/([0-9]{4})\b");
        static readonly Regex DATA_EXTENSO = new Regex(@"^([0-9]{1,2})\s*(?:DE\s+)?(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)[A-Z]*\.?\s*(?:DE\s+)?([0-9]{4})\b");
        static readonly Regex TIRA_EXTENSO = new Regex(@"^\s*[0-9]{1,2}\s*(?i:de\s+)?\p{L}+\.?\s*(?i:de\s+)?[0-9]{4}\s*");
        static readonly string[] MESES = { "JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ" };

        static readonly string[] PALAVRAS_SAIDA = {
            "PAGTO", "PAGAMENTO", "PAG ", "PGTO", "TARIFA", "TAR ", "ENVIAD", "DEBITO", "DEB ", "SAQUE", "COMPRA",
            "BOLETO", "IOF", "JUROS", "DARF", "GPS", "DAS ", "SIMPLES NAC", "FGTS", "TAXA", "CESTA", "MENSALIDADE",
            "APLICACAO", "APLIC", "TRANSF ENV", "TED ENV", "DOC ENV", "PIX ENV", "PIX - ENV", "CHEQUE COMP", "CH COMP",
            "SEGURO", "EMPRESTIMO PARC", "PARCELA", "ENCARGO", "ANUIDADE"};
        static readonly string[] PALAVRAS_ENTRADA = {
            "RECEBID", "CREDITO", "CRED ", "DEPOSITO", "DEP ", "RESGATE", "REND", "ESTORNO", "TRANSF REC", "TED REC",
            "PIX REC", "PIX - REC", "LIQUIDACAO COBR", "COBRANCA", "VENDA", "CIELO", "REDE ", "STONE", "GETNET", "SALDO APLIC"};
        static readonly string[] CABECALHOS = {
            "SALDO", "PAGINA", "PAG.", "EXTRATO", "AGENCIA", "PERIODO", "HISTORICO", "LANCAMENTO", "OUVIDORIA", "SAC ",
            "EMITIDO", "DATA ", "TOTAL", "CONTA CORRENTE", "BANCO ", "ATENDIMENTO", "FALE COM", "SOLICITADO", "CPF/CNPJ",
            "INSTITUICAO"};

        public static List<Lancamento> Ler(string arquivo, bool juntarContinuacao)
        {
            string txt = Texto(arquivo);
            if (RelatorioFinanceiro(txt))
            {
                try { return LeitorRelatorioFinanceiro.Lancamentos(arquivo); }
                catch (IOException e) { Log.Erro("relatorio financeiro " + Path.GetFileName(arquivo), e); }
            }
            return LerTexto(txt, Path.GetFileName(arquivo), juntarContinuacao);
        }

        public static bool PareceRelatorioFinanceiro(string arquivo)
        {
            try { return RelatorioFinanceiro(Texto(arquivo)); } catch (IOException) { return false; }
        }

        static bool RelatorioFinanceiro(string texto)
        {
            foreach (var l in texto.SplitRe(@"\r?\n"))
            {
                string n = Regras.Normalizar(l);
                if (n.Contains("TERCEIRO") && n.Contains("HISTORICO") && n.Contains("ENTRADA") && n.Contains("SAIDA")) return true;
            }
            return false;
        }

        static string Texto(string arquivo)
        {
            string t = string.Join("\n", TextoPDF.Linhas(arquivo));
            if (t.Trim().Length == 0) throw new LeituraException("PDF sem texto (provavelmente escaneado/imagem). Peça o OFX ao banco.");
            return t;
        }

        static bool G(Match m, int i) { return m.Groups[i].Success; }

        /// <summary>Interpreta o texto de um extrato (uma linha por lançamento). Usado também para planilha sem cabeçalho.</summary>
        public static List<Lancamento> LerTexto(string texto, string origem, bool juntarContinuacao)
        {
            string anoPadrao = DateTime.Now.Year.ToString();
            var ma = ANO.Match(texto);
            if (ma.Success) anoPadrao = ma.Groups[1].Value;

            var linhasTexto = texto.SplitRe(@"\r?\n");
            int negativosNoValor = 0;
            foreach (var l in linhasTexto)
            {
                var mv0 = VALOR.Match(l.Replace(' ', ' '));
                if (mv0.Success && (G(mv0, 1) || G(mv0, 6)) && !Regras.Normalizar(l).Contains("SALDO")) negativosNoValor++;
            }
            bool sinalExplicitoNoDocumento = negativosNoValor >= 3;
            bool relatorio = RelatorioFinanceiro(texto);
            if (relatorio) { sinalExplicitoNoDocumento = false; juntarContinuacao = true; }

            var lista = new List<Lancamento>();
            string ultimaData = null;
            double? saldoAnterior = null;
            double somaDesdeSaldo = 0;
            Lancamento ultimo = null;
            int continuacoes = 0;

            foreach (var linhaBruta in linhasTexto)
            {
                string linha = Regex.Replace(linhaBruta.Replace(' ', ' '), @"\s{2,}", " ").Trim();
                if (linha.Length == 0) continue;
                string norm = Regras.Normalizar(linha);

                string data = null;
                string resto = linha;
                var mx = DATA_EXTENSO.Match(norm);
                var md = DATA.Match(linha);
                if (mx.Success)
                {
                    int mes = 0;
                    for (int i = 0; i < 12; i++) if (MESES[i] == mx.Groups[2].Value) mes = i + 1;
                    int dia = int.Parse(mx.Groups[1].Value);
                    if (mes > 0 && dia >= 1 && dia <= 31)
                    {
                        data = dia.ToString("00") + "/" + mes.ToString("00") + "/" + mx.Groups[3].Value;
                        ultimaData = data;
                        if (norm.Contains("SALDO") || !VALOR.IsMatch(linha)) { ultimo = null; continue; }
                        resto = TIRA_EXTENSO.Replace(linha, "", 1);
                    }
                }
                else if (md.Success && Valida(md.Groups[1].Value, md.Groups[2].Value))
                {
                    string ano = !G(md, 3) ? anoPadrao : (md.Groups[3].Value.Length == 2 ? "20" + md.Groups[3].Value : md.Groups[3].Value);
                    data = md.Groups[1].Value + "/" + md.Groups[2].Value + "/" + ano;
                    resto = linha.Substring(md.Index + md.Length).Trim();
                    var md2 = DATA.Match(resto);
                    if (md2.Success && Valida(md2.Groups[1].Value, md2.Groups[2].Value)) resto = resto.Substring(md2.Index + md2.Length).Trim();
                }

                var valores = new List<double[]>();
                foreach (Match mv in VALOR.Matches(resto))
                {
                    double v = J.ParseD(mv.Groups[3].Value.Replace(".", "") + "." + mv.Groups[4].Value);
                    bool neg = G(mv, 1) || (G(mv, 2) && G(mv, 5)) || G(mv, 6) || mv.Groups[7].Value == "D";
                    bool explicito = neg || mv.Groups[7].Value == "C";
                    valores.Add(new[] { neg ? -v : v, explicito ? 1 : 0, mv.Index });
                }

                if (norm.Contains("SALDO") && !norm.Contains("APLIC") && valores.Count > 0)
                {
                    saldoAnterior = valores[valores.Count - 1][0];
                    somaDesdeSaldo = 0;
                    ultimo = null;
                    if (data != null) ultimaData = data;
                    continue;
                }

                if (relatorio && (data == null || valores.Count < 2))
                {
                    if (ultimo != null && data == null && continuacoes < 8 && !EhCabecalho(norm))
                    { ultimo.historico = LeitorOFX.Limpar(ultimo.historico + " " + linha.Replace("\"", "")); continuacoes++; }
                    continue;
                }
                if (valores.Count == 0)
                {
                    if (juntarContinuacao && ultimo != null && data == null && continuacoes < 2 && !EhCabecalho(norm) && linha.Length <= 80)
                    {
                        ultimo.historico = LeitorOFX.Limpar(ultimo.historico + " " + linha.Replace("\"", ""));
                        continuacoes++;
                    }
                    else if (data == null) ultimo = null;
                    continue;
                }

                if (data == null) data = ultimaData;
                string hist = LeitorOFX.Limpar(resto.Substring(0, (int)valores[0][2]).Replace("\"", ""));
                if (data == null || hist.Length == 0 || EhCabecalho(Regras.Normalizar(hist))) { ultimo = null; continue; }
                ultimaData = data;

                double[] tx = valores[0];
                double? saldoLinha = valores.Count >= 2 ? (double?)valores[valores.Count - 1][0] : null;
                double valor = Math.Abs(tx[0]);
                if (valor < 0.005) { ultimo = null; continue; }
                bool? saida = null;
                string obs = "";

                if (tx[1] == 1 && !relatorio) saida = tx[0] < 0;
                else if (sinalExplicitoNoDocumento) saida = false;
                else if (saldoLinha != null && saldoAnterior != null)
                {
                    double diff = saldoLinha.Value - (saldoAnterior.Value + somaDesdeSaldo);
                    if (Math.Abs(diff + valor) < 0.01) saida = true;
                    else if (Math.Abs(diff - valor) < 0.01) saida = false;
                }
                if (saida == null)
                {
                    string nh = Regras.Normalizar(hist) + " ";
                    if (Contem(nh, PALAVRAS_SAIDA) && !Contem(nh, PALAVRAS_ENTRADA)) { saida = true; obs = "Sinal por palavra-chave"; }
                    else if (Contem(nh, PALAVRAS_ENTRADA)) { saida = false; obs = "Sinal por palavra-chave"; }
                    else { saida = false; obs = "Sinal presumido (entrada)"; }
                }

                var lc = new Lancamento(data, hist, valor, saida.Value) { origem = origem };
                if (obs.Length > 0) { lc.conferirSinal = true; lc.obs = obs; }
                lista.Add(lc);

                if (saldoLinha != null) { saldoAnterior = saldoLinha; somaDesdeSaldo = 0; }
                else somaDesdeSaldo += saida.Value ? -valor : valor;
                ultimo = lc;
                continuacoes = 0;
            }
            return lista;
        }

        static bool Valida(string d, string m)
        {
            int dd = int.Parse(d), mm = int.Parse(m);
            return dd >= 1 && dd <= 31 && mm >= 1 && mm <= 12;
        }

        static bool Contem(string texto, string[] palavras) { return palavras.Any(texto.Contains); }

        static bool EhCabecalho(string norm)
        {
            foreach (var c in CABECALHOS) if (norm.StartsWith(c) || norm.Contains(" " + c.Trim() + ":")) return true;
            return norm.Matches(@".*\b(PAGINA|FOLHA)\s*[0-9]+.*");
        }
    }
}

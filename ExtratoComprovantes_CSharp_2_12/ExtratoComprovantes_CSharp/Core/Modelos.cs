using System;
using System.Collections.Generic;
using System.Linq;

namespace ExtratoComprovantes
{
    /// <summary>Uma linha do extrato.</summary>
    public class Lancamento
    {
        public string data;            // DD/MM/AAAA
        public string historico;
        public double valor;           // sempre positivo
        public double valorBruto;      // gateway: valor antes das taxas (0 = igual ao valor)
        public bool saida;             // true = débito em conta (pagamento)
        public int? debito;
        public int? credito;
        public string obs = "";
        public bool conferirSinal;     // sinal deduzido por heurística
        public bool manual;            // contas editadas à mão: regras não sobrescrevem
        public string origem = "";
        public List<ItemPlanilha> planilha = new List<ItemPlanilha>();
        public List<NotaFiscal.Titulo> notas = new List<NotaFiscal.Titulo>();
        public Comprovante comprovante;
        public int? contaBanco;        // conta banco do extrato de origem (null = conta padrão)
        public bool descartada;        // entrada de transferência entre bancos
        public Lancamento par;
        public string exportadoEm = "";
        public string complementoSalvo;
        public long seq;
        public string buscaSalva;
        public string chave;           // identifica o lançamento para não importar duas vezes
        public Conciliacao conciliacao; // conciliação com a planilha financeira (null = sem planilha)

        public Lancamento(string data, string historico, double valor, bool saida)
        {
            this.data = data;
            this.historico = historico;
            this.valor = Math.Abs(valor);
            this.saida = saida;
        }
    }

    /// <summary>Uma linha da planilha financeira (próxima etapa).</summary>
    public class ItemPlanilha
    {
        public string arquivo = "";
        public int linha;
        public List<string> datas = new List<string>();
        public double valor;
        public List<double> alternativos = new List<double>();
        public bool? saida;
        public string favorecido = "";
        public string documento = "";
        public string descricao = "";
        public List<string> nfs = new List<string>();   // números de NF citados (coluna NF/Documento ou descrição)
        public string categoria = "";
        public int? conta;
        public Lancamento vinculo;
        public string exportado = "";

        public string Data() { return datas.Count == 0 ? "" : datas[0]; }

        public string Resumo()
        {
            var sb = new System.Text.StringBuilder(favorecido);
            if (descricao.Length > 0 && !Regras.Normalizar(favorecido).Contains(Regras.Normalizar(descricao)))
                sb.Append(sb.Length > 0 ? " – " : "").Append(descricao);
            return sb.ToString();
        }
    }

    /// <summary>Comprovante de pagamento (próxima etapa).</summary>
    public class Comprovante
    {
        public string arquivo = "";
        public int pagina = 1;
        public string tipo = "", data = "", favorecido = "", docFavorecido = "", pagador = "", docPagador = "", codigoBarras = "", idTransacao = "", descricao = "", texto = "";
        public double valor;
        public Lancamento vinculo;
        public bool conferir;
        public string exportado = "";
    }

    /// <summary>Nota fiscal (próxima etapa).</summary>
    /// <summary>Nota fiscal lida do XML (NF-e, NFC-e, CT-e, NFS-e) e seus títulos (parcelas).</summary>
    public class NotaFiscal
    {
        public string arquivo = "", tipo = "", chave = "", numero = "", serie = "";
        public DateTime? emissao;
        public string emitCnpj = "", emitNome = "", destCnpj = "", destNome = "";
        public double valor;          // vNF / vTPrest / vServ
        public double valorLiquido;   // NFS-e: líquido após retenções
        public bool cancelada;
        public readonly List<Titulo> titulos = new List<Titulo>();

        /// <summary>true = compra/tomada (saída no banco), false = venda/prestação (entrada), null = indefinida.</summary>
        public bool? Compra(string cnpjEmpresa)
        {
            if (string.IsNullOrEmpty(cnpjEmpresa)) return null;
            string raiz = cnpjEmpresa.Length == 14 ? cnpjEmpresa.Substring(0, 8) : cnpjEmpresa;
            bool emit = emitCnpj.StartsWith(raiz, StringComparison.Ordinal), dest = destCnpj.StartsWith(raiz, StringComparison.Ordinal);
            if (dest && !emit) return true;
            if (emit && !dest) return false;
            return null;
        }

        public string ContraparteNome(string cnpjEmpresa)
        {
            bool? c = Compra(cnpjEmpresa);
            if (c == null) return emitNome.Length == 0 ? destNome : emitNome;
            return c.Value ? emitNome : destNome;
        }

        public string ContraparteCnpj(string cnpjEmpresa)
        {
            bool? c = Compra(cnpjEmpresa);
            if (c == null) return emitCnpj;
            return c.Value ? emitCnpj : destCnpj;
        }

        public string Rotulo() { return tipo + " " + numero; }

        /// <summary>Uma parcela (duplicata) ou o valor total da nota quando não há duplicatas.</summary>
        public class Titulo
        {
            public NotaFiscal nf;
            public string parcela = "";
            public double valor;
            public List<double> alternativos = new List<double>();
            public DateTime? vencimento;
            public Lancamento vinculo;
            public string exportado = "";
            public DateTime? Referencia() { return vencimento ?? nf.emissao; }
        }
    }
}

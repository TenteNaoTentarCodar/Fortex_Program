using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace ExtratoComprovantes
{
    /// <summary>
    /// XML fiscal para conferência com o extrato: participantes, emissão, valor, duplicatas e cancelamento.
    /// NF-e/NFC-e (55/65), CT-e (57), CT-e OS (67), eventos e NFS-e (padrão nacional, ABRASF e Nota Paulistana).
    /// </summary>
    public static class LeitorNotaXML
    {
        class Doc
        {
            public string tipo = "", chave = "", numero = "", serie = "", emitCnpj = "", emitNome = "", destCnpj = "", destNome = "";
            public string tpEvento = "", chaveRef = "";
            public DateTime? emissao;
            public double valor, valorLiquido;
            public readonly List<Tuple<string, DateTime?, double>> dups = new List<Tuple<string, DateTime?, double>>();
            public readonly List<double> pags = new List<double>();
        }

        /// <summary>Lê .xml ou .zip (com XMLs dentro). Eventos de cancelamento vêm como notas com cancelada=true e chave preenchida.</summary>
        public static List<NotaFiscal> LerArquivo(string p)
        {
            var r = new List<NotaFiscal>();
            string n = Path.GetFileName(p);
            if (n.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
            {
                using (var z = ZipFile.OpenRead(p))
                    foreach (var e in z.Entries)
                    {
                        if (e.FullName.EndsWith("/") || !e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) continue;
                        byte[] b;
                        using (var s = e.Open())
                        using (var ms = new MemoryStream()) { s.CopyTo(ms); b = ms.ToArray(); }
                        var nf = Converter(Ler(b), n + " › " + Regex.Replace(e.FullName, @".*[/\\]", ""));
                        if (nf != null) r.Add(nf);
                    }
            }
            else
            {
                var nf = Converter(Ler(File.ReadAllBytes(p)), n);
                if (nf != null) r.Add(nf);
            }
            return r;
        }

        static NotaFiscal Converter(Doc d, string arquivo)
        {
            if (d == null) return null;
            var nf = new NotaFiscal { arquivo = arquivo };
            if (d.tipo == "Evento")
            {
                if (d.tpEvento != "110111" && !d.numero.ToUpperInvariant().Contains("CANCEL")) return null;
                nf.tipo = "Evento"; nf.cancelada = true; nf.chave = d.chaveRef; nf.emissao = d.emissao;
                return nf;
            }
            if (d.tipo == "MDF-e") return null;
            nf.tipo = d.tipo; nf.chave = d.chave; nf.numero = new Regex("^0+(?=[0-9])").Replace(d.numero, "", 1); nf.serie = d.serie;
            nf.emissao = d.emissao; nf.emitCnpj = d.emitCnpj; nf.emitNome = d.emitNome; nf.destCnpj = d.destCnpj; nf.destNome = d.destNome;
            nf.valor = d.valor; nf.valorLiquido = d.valorLiquido;
            foreach (var du in d.dups)
            {
                if (du.Item3 < 0.005) continue;
                nf.titulos.Add(new NotaFiscal.Titulo { nf = nf, parcela = du.Item1, vencimento = du.Item2, valor = du.Item3 });
            }
            if (nf.titulos.Count == 0 && nf.valor > 0.005)
            {
                var t = new NotaFiscal.Titulo { nf = nf, valor = nf.valor };
                if (nf.valorLiquido > 0.005 && Math.Abs(nf.valorLiquido - nf.valor) > 0.005) t.alternativos.Add(nf.valorLiquido);
                double somaPag = d.pags.Sum();
                foreach (var p in d.pags) if (Math.Abs(p - nf.valor) > 0.005 && !t.alternativos.Contains(p)) t.alternativos.Add(p);
                if (somaPag > 0.005 && Math.Abs(somaPag - nf.valor) > 0.005 && !t.alternativos.Contains(somaPag)) t.alternativos.Add(somaPag);
                nf.titulos.Add(t);
            }
            return nf;
        }

        static Doc Ler(byte[] bytes)
        {
            XmlElement raiz;
            try
            {
                var doc = new XmlDocument();
                var cfg = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
                using (var ms = new MemoryStream(bytes))
                using (var rd = XmlReader.Create(ms, cfg)) doc.Load(rd);
                raiz = doc.DocumentElement;
            }
            catch { return null; }
            if (raiz == null) return null;
            var d = new Doc();
            string nomeRaiz = raiz.LocalName.ToLowerInvariant();
            XmlElement inf;
            if (nomeRaiz.Contains("evento") || (Primeiro(raiz, "infEvento") != null && Primeiro(raiz, "infNFe") == null && Primeiro(raiz, "infCte") == null))
                return Evento(raiz, d);
            if ((inf = Primeiro(raiz, "infNFe")) != null) return Nfe(raiz, inf, d);
            if ((inf = Primeiro(raiz, "infCte")) != null) return Cte(raiz, inf, d);
            if ((inf = Primeiro(raiz, "infMDFe")) != null) { d.tipo = "MDF-e"; return d; }
            if (PareceNfse(raiz)) return Nfse(raiz, d);
            return null;
        }

        static Doc Nfe(XmlElement raiz, XmlElement inf, Doc d)
        {
            d.tipo = Texto(inf, "mod") == "65" ? "NFC-e" : "NF-e";
            d.chave = Chave(raiz, inf, "NFe", "chNFe");
            d.numero = Texto(inf, "nNF");
            d.serie = Texto(inf, "serie");
            d.emissao = Data(Nz(Texto(inf, "dhEmi"), Texto(inf, "dEmi")));
            XmlElement emit = Primeiro(inf, "emit"), dest = Primeiro(inf, "dest");
            d.emitCnpj = DocDe(emit); d.emitNome = Texto(emit, "xNome");
            d.destCnpj = DocDe(dest); d.destNome = Texto(dest, "xNome");
            d.valor = Num(Texto(Primeiro(inf, "ICMSTot"), "vNF"));
            foreach (XmlElement n in Todos(inf))
            {
                if (n.LocalName == "dup") d.dups.Add(Tuple.Create(TextoFilho(n, "nDup"), Data(TextoFilho(n, "dVenc")), Num(TextoFilho(n, "vDup"))));
                else if (n.LocalName == "detPag")
                {
                    double v = Num(TextoFilho(n, "vPag"));
                    if (v > 0.005 && TextoFilho(n, "tPag") != "90") d.pags.Add(v);
                }
            }
            return d;
        }

        static Doc Cte(XmlElement raiz, XmlElement inf, Doc d)
        {
            d.tipo = Texto(inf, "mod") == "67" ? "CT-e OS" : "CT-e";
            d.chave = Chave(raiz, inf, "CTe", "chCTe");
            d.numero = Texto(inf, "nCT");
            d.serie = Texto(inf, "serie");
            d.emissao = Data(Texto(inf, "dhEmi"));
            var emit = Primeiro(inf, "emit");
            d.emitCnpj = DocDe(emit); d.emitNome = Texto(emit, "xNome");
            d.valor = Num(Texto(Primeiro(inf, "vPrest"), "vTPrest"));
            XmlElement rem = Primeiro(inf, "rem"), dst = Primeiro(inf, "dest"), exp = Primeiro(inf, "exped"), rec = Primeiro(inf, "receb");
            XmlElement tomador = null;
            var ide = Primeiro(inf, "ide");
            var toma4 = Primeiro(ide, "toma4");
            if (toma4 != null) tomador = toma4;
            else
            {
                var toma3 = Primeiro(ide, "toma3") ?? Primeiro(ide, "toma03");
                string t = toma3 != null ? Texto(toma3, "toma") : "";
                switch (t)
                {
                    case "0": tomador = rem; break;
                    case "1": tomador = exp; break;
                    case "2": tomador = rec; break;
                    case "3": tomador = dst; break;
                }
                if (tomador == null && d.tipo == "CT-e OS") tomador = Primeiro(inf, "toma");
            }
            if (tomador != null) { d.destCnpj = DocDe(tomador); d.destNome = Texto(tomador, "xNome"); }
            else if (dst != null) { d.destCnpj = DocDe(dst); d.destNome = Texto(dst, "xNome"); }
            return d;
        }

        static Doc Evento(XmlElement raiz, Doc d)
        {
            d.tipo = "Evento";
            d.chaveRef = Nz(Texto(raiz, "chNFe"), Texto(raiz, "chCTe"), Texto(raiz, "chMDFe"));
            d.tpEvento = Texto(raiz, "tpEvento");
            string desc = Texto(raiz, "descEvento");
            d.numero = desc.Length == 0 ? d.tpEvento : desc;
            d.emissao = Data(Texto(raiz, "dhEvento"));
            if (d.chaveRef.Length == 44) d.emitCnpj = d.chaveRef.Substring(6, 14);
            var infEv = Primeiro(raiz, "infEvento");
            string autor = infEv != null ? Nz(TextoFilho(infEv, "CNPJ"), TextoFilho(infEv, "CPF")) : "";
            if (autor.Length > 0 && autor != d.emitCnpj) d.destCnpj = autor;
            return d.chaveRef.Length == 0 ? null : d;
        }

        static bool PareceNfse(XmlElement raiz)
        {
            string[] marcas = { "Nfse", "NFSe", "CompNfse", "InfNfse", "infNFSe", "infDPS", "ChaveNFe", "CPFCNPJPrestador", "PrestadorServico", "IdentificacaoPrestador" };
            foreach (var m in marcas) if (Primeiro(raiz, m) != null || string.Equals(raiz.LocalName, m, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        static Doc Nfse(XmlElement raiz, Doc d)
        {
            d.tipo = "NFS-e";
            var prest = PrimeiroDe(raiz, "emit", "PrestadorServico", "Prestador", "IdentificacaoPrestador", "prest", "CPFCNPJPrestador", "DadosPrestador");
            var toma = PrimeiroDe(raiz, "TomadorServico", "Tomador", "IdentificacaoTomador", "toma", "CPFCNPJTomador", "DadosTomador");
            d.emitCnpj = DocDe(prest);
            d.emitNome = Nz(Texto(prest, "RazaoSocial"), Texto(prest, "xNome"), Texto(raiz, "RazaoSocialPrestador"));
            d.destCnpj = DocDe(toma);
            d.destNome = Nz(Texto(toma, "RazaoSocial"), Texto(toma, "xNome"), Texto(raiz, "RazaoSocialTomador"));
            d.numero = Nz(Texto(raiz, "nNFSe"), TextoDe(raiz, "Numero"), Texto(raiz, "NumeroNFe"), Texto(raiz, "NumeroNfse"));
            var infN = Primeiro(raiz, "infNFSe");
            d.chave = Nz(infN == null ? "" : Regex.Replace(infN.GetAttribute("Id") ?? "", "^NFS", ""), Texto(raiz, "CodigoVerificacao"));
            d.emissao = Data(Nz(Texto(raiz, "dhEmi"), Texto(raiz, "DataEmissao"), Texto(raiz, "DataEmissaoNFe"), Texto(raiz, "dhProc"), Texto(raiz, "DataEmissaoNfse")));
            d.valor = Num(Nz(Texto(raiz, "vServ"), Texto(raiz, "ValorServicos"), Texto(raiz, "ValorServico"), Texto(raiz, "vLiq")));
            d.valorLiquido = Num(Nz(Texto(raiz, "vLiq"), Texto(raiz, "ValorLiquidoNfse"), Texto(raiz, "ValorLiquido")));
            return d;
        }

        // ---------------- utilitários ----------------

        static IEnumerable<XmlElement> Todos(XmlElement b)
        {
            if (b == null) yield break;
            foreach (XmlNode n in b.GetElementsByTagName("*")) { var e = n as XmlElement; if (e != null) yield return e; }
        }

        static XmlElement Primeiro(XmlElement b, string nome) { return Todos(b).FirstOrDefault(e => e.LocalName == nome); }

        static XmlElement PrimeiroDe(XmlElement b, params string[] nomes)
        {
            foreach (var n in nomes) { var e = Primeiro(b, n); if (e != null) return e; }
            return null;
        }

        static string Texto(XmlElement b, string nome) { var e = Primeiro(b, nome); return e == null ? "" : e.InnerText.Trim(); }

        static string TextoDe(XmlElement b, string nome)
        {
            foreach (var n in Todos(b))
            {
                if (n.LocalName != nome) continue;
                var pai = n.ParentNode;
                if (pai != null && pai.LocalName.ToLowerInvariant().Contains("rps")) continue;
                return n.InnerText.Trim();
            }
            return "";
        }

        static string TextoFilho(XmlElement b, string nome)
        {
            foreach (XmlNode n in b.ChildNodes) if (n is XmlElement && n.LocalName == nome) return n.InnerText.Trim();
            return "";
        }

        static string DocDe(XmlElement bloco)
        {
            if (bloco == null) return "";
            string c = Nz(Texto(bloco, "CNPJ"), Texto(bloco, "Cnpj"), Texto(bloco, "CPF"), Texto(bloco, "Cpf"));
            if (c.Length == 0) c = bloco.InnerText;
            c = Regex.Replace(c, "[^0-9]", "");
            return c.Length == 14 || c.Length == 11 ? c : "";
        }

        static string Chave(XmlElement raiz, XmlElement inf, string prefixo, string tagProt)
        {
            string id = inf.GetAttribute("Id") ?? "";
            if (id.StartsWith(prefixo, StringComparison.Ordinal)) return id.Substring(prefixo.Length);
            string p = Texto(raiz, tagProt);
            return p.Length == 0 ? id : p;
        }

        static DateTime? Data(string s)
        {
            if (s == null) return null;
            s = s.Trim();
            try
            {
                if (s.Matches(@"[0-9]{4}-[0-9]{2}-[0-9]{2}.*")) return new DateTime(int.Parse(s.Substring(0, 4)), int.Parse(s.Substring(5, 2)), int.Parse(s.Substring(8, 2)));
                if (s.Matches(@"[0-9]{2}/[0-9]{2}/[0-9]{4}.*")) return new DateTime(int.Parse(s.Substring(6, 4)), int.Parse(s.Substring(3, 2)), int.Parse(s.Substring(0, 2)));
                if (s.Matches(@"[0-9]{4}-[0-9]{2}")) return new DateTime(int.Parse(s.Substring(0, 4)), int.Parse(s.Substring(5, 2)), 1);
                if (s.Matches(@"[0-9]{6}")) return new DateTime(int.Parse(s.Substring(0, 4)), int.Parse(s.Substring(4)), 1);
                if (s.Matches(@"[0-9]{8}")) return new DateTime(int.Parse(s.Substring(0, 4)), int.Parse(s.Substring(4, 2)), int.Parse(s.Substring(6)));
            }
            catch { }
            return null;
        }

        static double Num(string s)
        {
            if (string.IsNullOrEmpty(s)) return 0;
            if (s.Contains(",")) s = s.Replace(".", "").Replace(',', '.');
            double v;
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v) ? v : 0;
        }

        static string Nz(params string[] v) { foreach (var x in v) if (!string.IsNullOrEmpty(x)) return x; return ""; }
    }
}

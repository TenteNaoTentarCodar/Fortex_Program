using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExtratoComprovantes
{
    /// <summary>Preferências do usuário (pastas, última regra, recentes, opções) em %LOCALAPPDATA%\ExtratoComprovantes\preferencias.txt.</summary>
    public static class Prefs
    {
        static readonly Dictionary<string, string> v = new Dictionary<string, string>();
        static bool lido;

        static string Arquivo() { return Path.Combine(Log.PastaApp(), "preferencias.txt"); }

        static void Ler()
        {
            if (lido) return;
            lido = true;
            try
            {
                if (!File.Exists(Arquivo())) return;
                foreach (var l in File.ReadAllLines(Arquivo(), Encoding.UTF8))
                {
                    int i = l.IndexOf('=');
                    if (i <= 0) continue;
                    v[l.Substring(0, i)] = Progresso.Des(l.Substring(i + 1));
                }
            }
            catch (Exception e) { Log.Erro("preferências", e); }
        }

        static void Gravar()
        {
            try
            {
                File.WriteAllLines(Arquivo(), v.Select(e => e.Key + "=" + Progresso.Esc(e.Value)), new UTF8Encoding(false));
            }
            catch (Exception e) { Log.Erro("preferências", e); }
        }

        public static string Get(string k, string padrao) { Ler(); string r; return v.TryGetValue(k, out r) ? r : padrao; }
        public static bool GetBool(string k, bool padrao) { var s = Get(k, null); return s == null ? padrao : s == "1"; }
        public static void Put(string k, string valor) { Ler(); if (valor == null) v.Remove(k); else v[k] = valor; Gravar(); }
        public static void PutBool(string k, bool b) { Put(k, b ? "1" : "0"); }
        public static void Remove(string k) { Put(k, null); }

        /// <summary>Pasta guardada, se ainda existir; senão a padrão.</summary>
        public static string Pasta(string k, string padrao)
        {
            var d = Get(k, null);
            if (d != null && Directory.Exists(d)) return d;
            return padrao != null && Directory.Exists(padrao) ? padrao : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }
    }
}

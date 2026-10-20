using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Arquivos anexados na sessão (Plano de Contas, Razão, comprovantes) guardados junto com o progresso.
    /// Cada arquivo é copiado para a pasta de anexos (nome = 10 primeiros dígitos do SHA-1 + nome original).
    /// Mesmo formato do programa em Java (linhas #ANEXO;tipo;original;copia;paginas).
    /// </summary>
    public static class AnexosSessao
    {
        public class Item
        {
            public string tipo = "", original = "", copia = "", paginas = "";
            public Item() { }
            public Item(string tipo, string original) { this.tipo = tipo; this.original = original; }
            public string[] Campos() { return new[] { tipo, original, copia, paginas }; }
            public static Item De(string[] c)
            {
                if (c == null || c.Length < 4) return null;
                return new Item { tipo = c[0], original = c[1], copia = c[2], paginas = c[3] };
            }
        }

        /// <summary>Copia o original para a pasta de anexos (se ainda não estiver lá) e devolve o nome da cópia; "" se não conseguiu.</summary>
        public static string Copiar(string orig, string pasta)
        {
            try
            {
                if (string.IsNullOrEmpty(orig) || !File.Exists(orig)) return "";
                string dirOrig = Path.GetDirectoryName(Path.GetFullPath(orig));
                if (string.Equals(Path.GetFullPath(pasta).TrimEnd(Path.DirectorySeparatorChar), dirOrig.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase))
                    return Path.GetFileName(orig);
                string nome = Hash(orig) + "_" + Path.GetFileName(orig);
                string dest = Path.Combine(pasta, nome);
                if (!File.Exists(dest))
                {
                    Directory.CreateDirectory(pasta);
                    string tmp = dest + ".tmp";
                    File.Copy(orig, tmp, true);
                    if (File.Exists(dest)) File.Delete(dest);
                    File.Move(tmp, dest);
                }
                return nome;
            }
            catch (Exception e)
            {
                Log.Erro("Copiar anexo " + orig, e);
                return "";
            }
        }

        /// <summary>Arquivo a usar ao reabrir: o original, se ainda existir; senão a cópia guardada. null se nenhum existe.</summary>
        public static string Resolver(Item it, string pasta)
        {
            try { if (it.original.Length > 0 && File.Exists(it.original)) return it.original; } catch { }
            if (it.copia.Length > 0 && pasta != null && File.Exists(Path.Combine(pasta, it.copia))) return Path.Combine(pasta, it.copia);
            return null;
        }

        public static string NomeOriginal(Item it)
        {
            if (it.original.Length > 0) return Path.GetFileName(it.original);
            return System.Text.RegularExpressions.Regex.Replace(it.copia, "^[0-9a-f]{10}_", "");
        }

        static string Hash(string p)
        {
            using (var sha = SHA1.Create())
            using (var f = File.OpenRead(p))
            {
                var h = sha.ComputeHash(f);
                var s = new StringBuilder();
                foreach (var b in h) s.Append(b.ToString("x2"));
                return s.ToString().Substring(0, 10);
            }
        }
    }
}

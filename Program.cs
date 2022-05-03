using System;
using System.Linq;
using ClosedXML.Excel;

namespace LerPlanilhaExcel {

    public static class ExtensaoString {

        public static string ParseHome(this string path) {
            string home = (Environment.OSVersion.Platform == PlatformID.Unix ||
                Environment.OSVersion.Platform == PlatformID.MacOSX)
                ? Environment.GetEnvironmentVariable("HOME")
                : Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%");

            return path.Replace("~", home);
        }

    }
    class Program {

        static void Main(string[] args) {

            string colunas = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var path = @"~\Documents\DTI\Ferramentas\planilha-exemplo.xlsx".ParseHome();

            var xls = new XLWorkbook(path);

            var planilha = xls.Worksheets.First(w => w.Name == "Planilha1");

            var totalLinhas = planilha.Rows().Count();
            var totalColunas = planilha.Columns().Count();

            for (int l = 1; l <= totalLinhas; l++) {

                for (int c = 0; c < totalColunas; c++) {

                    if (l == 1) {
                        var tituloColuna = planilha.Cell($"{colunas[c]}{l}").Value.ToString();
                        Console.Write($"{tituloColuna} - ");
                    } else {
                        var celula = planilha.Cell($"{colunas[c]}{l}").Value.ToString();
                        Console.Write($"{celula} - ");

                        //var codigo = int.Parse(planilha.Cell($"A{l}").Value.ToString());
                        //var descricao = planilha.Cell($"B{l}").Value.ToString();
                        //var preco = decimal.Parse(planilha.Cell($"C{l}").Value.ToString());
                        //Console.WriteLine($"{codigo} - {descricao} - {preco}");
                    }
                    
                }

                Console.WriteLine();

            }

        }
    }
}
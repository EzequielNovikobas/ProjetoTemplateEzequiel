using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomacaoGoogleAcademico.BLL
{
    class Arquivo
    {
        public string CriaPastaPesquisa( string sPesquisa)
        {

            try
            {
                if (Directory.Exists(@"C:\Pesquisas") == false) // Verifica se diretório existe
                {
                    Directory.CreateDirectory(@"C:\Pesquisas"); // Cria diretório
                }

                if (Directory.Exists(@"C:\Pesquisas\" + sPesquisa) == false) // Verifica se diretório existe
                {
                    Directory.CreateDirectory(@"C:\Pesquisas\" + sPesquisa); // Cria diretório
                }

                string Dt = DateTime.Now.ToString("MMddyyyyHHmmss");
                
                if (Directory.Exists(@"C:\Pesquisas\" + sPesquisa  + "\\" + sPesquisa + Dt) == false) // Verifica se diretório existe
                {
                    Directory.CreateDirectory(@"C:\Pesquisas\" + sPesquisa + "\\" + sPesquisa + Dt); // Cria diretório
                }

                return @"C:\Pesquisas\" + sPesquisa + "\\" + sPesquisa + Dt;

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Arquivo Método: CriaPastaPesquisa");
            }
        }
        public void MoveArq(string Caminho, string CaminhoPesq, string TipoArq)
        {
            

            try
            {
                var ARQ = new DirectoryInfo(Caminho).GetFiles("*."+ TipoArq);

                string PdfPesq = ARQ[0].Name;

                File.Move(Caminho + "//" + PdfPesq, CaminhoPesq + "//" + PdfPesq);
            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Arquivo Método: MoveArq");
            }
        }
    }
}

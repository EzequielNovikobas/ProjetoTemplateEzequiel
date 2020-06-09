using AutomacaoGoogleAcademico.BLL;
using AutomacaoGoogleAcademicoDB.DAL;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScieloEzequiel.BLL
{
    class Automacao
    {
        private string Caminho;
        private string sPesquisa;
        private string dDataPesq;
        private string Controle;
        private Boolean FimProcess = true;
        public bool ParaThread;
        private AccessDB DB = new AccessDB();

        public Automacao(string Caminho_, string sPesquisa_, string Controle_, AccessDB DB_)
        {
            DB = DB_;
            Caminho = Caminho_;
            sPesquisa = sPesquisa_;
            Controle = Controle_;
            dDataPesq = DateTime.Now.ToString("dd/mm/yyy hh:mm:ss");
        }
        public void PreparaCapturaTag()
        {
            IWebDriver driver = null;

            try
            {
                driver = new ChromeDriver();

                ReinciaTagsCapturados();

                int NumeroDeCptura = 0;

                while (FimProcess == true)
                {
                    string query = "select top 1 * from URLdoTemplate where DataEhora is null order by NumeroDaURL";

                    DataSet DS = DB.DS(query, "URLdoTemplate");

                    NumeroDeCptura++;

                    //IniciaURLdoTemplate(driver);

                    if (DS.Tables["URLdoTemplate"].Rows.Count > 0)
                    {
                        foreach (DataRow row in DS.Tables["URLdoTemplate"].Rows)
                        {

                            driver.Navigate().GoToUrl(row["URLdoLinks"].ToString().Replace("#", "")); //Navega na url especificada
                            Thread.Sleep(2500);

                            CapturaTagImput(driver);

                            CapturaTagA(driver, NumeroDeCptura, row["NumeroDaURL"].ToString());

                            PercorreTagsImput(driver, row["NumeroDaURL"].ToString());

                            AtualizaURLTemplatesExecução(row["NumeroDaURL"].ToString());

                        }
                    }
                    else
                    {
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + ", Classe: Automacao, Método: PreparaCapturaElemento");
            }
            finally
            {
                driver.Close(); // fecha pagina do ChromeDriver
                driver.Quit();  // fecha prompt do ChromeDriver
            }
        }

        private void PercorreTagsImput(IWebDriver driver, string NumeroDaURL)
        {
            string query = "";
            DataSet DS = new DataSet();
            DataSet DS1 = new DataSet();
            IWebElement element1 = null;
            int QtdPesquisas;
            int r = 0;



            try
            {


                query = " select * from TagsCapturados where TipoDeTag in (select Tags from TagsAtributos " +
                        " where DataValidade > #" + DateTime.Now.ToString("dd/MM/yyy hh:mm:ss") + "#) and " +
                        " DataEhora is null and NumeroDaURL_tela = " + NumeroDaURL + " order by SequenciaCapturada";

                DS = DB.DS(query, "TagsCapturados");

                query = "select * from TemplatesExecução where DataEhora is null";

                DS1 = DB.DS(query, "TemplatesExecução");



                if (DS1.Tables["TemplatesExecução"].Rows.Count > 0)
                {
                    foreach (DataRow RowTemplatesExecução in DS1.Tables["TemplatesExecução"].Rows)
                    {
                        if (DS.Tables["TagsCapturados"].Rows.Count > 0)
                        {
                            foreach (DataRow RowTagsCapturados in DS.Tables["TagsCapturados"].Rows)
                            {
                                if (RowTagsCapturados["NomeTag"].ToString() == RowTemplatesExecução["NomeTag"].ToString())
                                {
                                    var ElementosA = driver.FindElements(By.XPath(("//input[@name='" + RowTagsCapturados["NomeTag"] + "']")));

                                    QtdPesquisas = 0;
                                    QtdPesquisas = ElementosA.Count();


                                    element1 = ElementosA[r];

                                    element1.SendKeys(RowTemplatesExecução["DadoDoInpt"].ToString());

                                    r++;

                                    AtualizaTagsCapturados(RowTagsCapturados["SequenciaCapturada"].ToString());

                                    if (r == DS1.Tables["TemplatesExecução"].Rows.Count)
                                    {

                                        driver.FindElement(By.XPath("//input[@src='/iah/I/image/pesq.gif']")).Click();
                                        EsperaElemento("References", driver);

                                        ReinciaTagsCapturados();

                                        CapturaTagImput(driver);

                                        //CapturaTagA(driver);

                                        AtualizaTemplatesExecução(RowTemplatesExecução["CódigoDoAplicativo"].ToString());

                                        AtualizaTagsCapturados(RowTagsCapturados["SequenciaCapturada"].ToString());

                                        MessageBox.Show("Processo concluído");

                                        return;

                                    }

                                    break;
                                }
                                else
                                {
                                    return;
                                    //AtualizaTagsCapturados(RowTagsCapturados["SequenciaCapturada"].ToString());
                                }

                            }
                        }
                    }
                }
                else
                {
                    //MessageBox.Show("Não existe template(s) cadastrado(s) para iniciar a pesquisa.");
                    return;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: PercorreTagsImput");
            }
        }

        private string InsereURLdoTemplate(string Href)
        {
            string query = "";

            OleDbDataReader DR;

            try
            {
                query = " select top 1 NumeroDaCaptura, NumeroMaximoDaCaptura from TagsCapturados " +
                        " order by NumeroDaCaptura desc";

                DB.ExecutaQry(query);

                using (DR = DB.DR(query))
                {
                    DR.Read();

                    int NumeroDaCaptura = Int32.Parse(DR["NumeroDaCaptura"].ToString());
                    int NumeroMaximoDaCaptura = Int32.Parse(DR["NumeroMaximoDaCaptura"].ToString());

                    if (NumeroDaCaptura >= NumeroMaximoDaCaptura)
                    {
                        return "1";
                    }
                }

                query = "select * from URLdoTemplate where URLdoLinks = '" + Href + "'";

                DB.ExecutaQry(query);

                using (DR = DB.DR(query))
                {
                    if (DR.HasRows == false)
                    {
                        query = " insert into URLdoTemplate (URLdoLinks) values ('" + Href + "')";

                        DB.ExecutaQry(query);

                        //query = "select top 1 NumeroDaURL from URLdoTemplate order by NumeroDaURL desc";

                        //using (DR = DB.DR(query))
                        //{
                        //    DR.Read();

                        //    return DR["NumeroDaURL"].ToString();

                        //}

                        return "";
                    }
                    else
                    {
                        return "";
                    }

                }




            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: InsereURLdoTemplate");
            }
        }

        private void AtualizaTemplatesExecução(string CodTemplate)
        {

            string query = "";

            try
            {
                query = " update TemplatesExecução set DataEhora =" +
                        " '" + DateTime.Now.ToString("dd/MM/yyy hh:mm:ss") + "' ";


                DB.ExecutaQry(query);
            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: AtualizaTemplatesExecução");
            }
        }

        private void AtualizaURLTemplatesExecução(string NumeroDaURL)
        {

            string query = "";

            try
            {
                query = " update URLdoTemplate set DataEhora =" +
                        " '" + DateTime.Now.ToString("dd/MM/yyy hh:mm:ss") + "' where NumeroDaURL = " + NumeroDaURL + " ";


                DB.ExecutaQry(query);
            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: AtualizaURLTemplatesExecução");
            }
        }

        private void AtualizaTagsCapturados(string CodCapturados)
        {
            string query = "";

            try
            {
                if (CodCapturados == "")
                {
                    query = " update TagsCapturados set DataEhora =" +
                        " '" + DateTime.Now.ToString("dd/MM/yyy hh:mm:ss") + "' ";
                }
                else
                {
                    query = " update TagsCapturados set DataEhora =" +
                        " '" + DateTime.Now.ToString("dd/MM/yyy hh:mm:ss") + "' " +
                        " where SequenciaCapturada = " + CodCapturados + " ";
                }

                DB.ExecutaQry(query);
            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: AtualizaTagsCapturados");
            }
        }


        private void ReinciaTagsCapturados()
        {
            string query = "";

            try
            {
                query = " Delete from TagsCapturados";

                DB.ExecutaQry(query);


            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: ReiinciaTagsCapturados");
            }
        }

        private void CapturaTagA(IWebDriver driver, int NumeroDeCptura, string NumeroDaURL)
        {
            int QtdPesquisas = 1;
            string NomeTag = "";
            string LocalTag = "";
            string Href = "";
            string query = "";
            IWebElement element1 = null;
            //string NumeroDaURL = "";

            try
            {
                var ElementosA = driver.FindElements(By.XPath(("//a"))); //Pega os elemento referente aos links exibidos na pagina

                QtdPesquisas = 0;
                QtdPesquisas = ElementosA.Count();

                for (int r = 0; r < QtdPesquisas - 1; r++)
                {
                    element1 = ElementosA[r];

                    Href = element1.GetAttribute("href");

                    LocalTag = element1.Location.ToString();

                    if (Href != null && Href.Contains("www"))
                    {
                        string[] AUXLocalizacao = LocalTag.Split(',');

                        string LocalizacaoTagx = AUXLocalizacao[0].Replace("{X=", "");
                        string LocalizacaoTagy = AUXLocalizacao[1].Replace("Y=", "");

                        query = " insert into TagsCapturados (TipoDeTag,  LocalizaçãoTagX, LocalizaçãoTagY, NumeroDaURL_tela, NumeroDaCaptura, NumeroMaximoDaCaptura)" +
                                " values ('ref'," + LocalizacaoTagx + "," + LocalizacaoTagy.Replace("}", "") + "," + NumeroDaURL + "," + NumeroDeCptura + ", 5)";

                        DB.ExecutaQry(query);


                        InsereURLdoTemplate(Href.Replace("'", "₱"));      

                    }


                }
            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: CapturaTagA");
            }
        }

        private void CapturaTagImput(IWebDriver driver)
        {
            int QtdPesquisas = 1;
            string NomeTag = "";
            string LocalTag = "";
            string Href = "";
            string query = "";
            IWebElement element1 = null;
            int NumeroDaURL = 1;



            try
            {
                var ElementosImput = driver.FindElements(By.XPath(("//input"))); //Pega os elemento referente aos links exibidos na pagina

                QtdPesquisas = 0;
                QtdPesquisas = ElementosImput.Count();

                for (int i = 0; i < QtdPesquisas - 1; i++)
                {
                    element1 = ElementosImput[i];

                    NomeTag = element1.GetAttribute("name");

                    LocalTag = element1.Location.ToString();

                    string[] AUXLocalizacao = LocalTag.Split(',');

                    string LocalizacaoTagx = AUXLocalizacao[0].Replace("{X=", "");
                    string LocalizacaoTagy = AUXLocalizacao[1].Replace("Y=", "");

                    query = " insert into TagsCapturados (TipoDeTag, NomeTag, LocalizaçãoTagX, LocalizaçãoTagY, NumeroDaURL)" +
                            " values ('input','" + NomeTag.Replace("'", "₱") + "'," + LocalizacaoTagx + "," + LocalizacaoTagy.Replace("}", "") + "," + NumeroDaURL + ")";

                    DB.ExecutaQry(query);

                    NumeroDaURL++;

                }
            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: CapturaTagImput");
            }
        }

        private void IniciaURLdoTemplate(IWebDriver driver)
        {
            try
            {

                driver.Navigate().GoToUrl(ConsultaUrl()); //Navega na url especificada
                Thread.Sleep(2500);
            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: IniciaURLdoTemplate");
            }
        }

        private string ConsultaUrl()
        {
            string Qry;
            OleDbDataReader DR;

            try
            {
                Qry = "select top 1 URLdoLinks  from URLdoTemplate";

                using (DR = DB.DR(Qry))
                {
                    if (DR.HasRows == true)
                    {
                        DR.Read();

                        return DR["URLdoLinks"].ToString().Replace("#", "");
                    }
                    else
                    {
                        return "erro";
                    }

                }

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: ConsultaUrl");
            }
        }
        public string EsperaElemento(string Elemento, IWebDriver driver)
        {
            int qtd = 0;
            int TimeOut = 0;
            string Bib = "";

            try
            {

                while (qtd == 0)
                {
                    Thread.Sleep(500);

                    if (driver.PageSource.Contains(Elemento) == true)
                    {
                        return "ok";
                    }
                    else if (TimeOut == 120)
                    {
                        throw new Exception("Tempo de espera excedido, tela inesperada ou problemas com o navegador.");
                    }

                    TimeOut++;
                }

                return "ok";

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: EsperaElemento");
            }
        }

    }
}

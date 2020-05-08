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

                ReiinciaTagsCapturados();

                IniciaURLdoTemplate(driver);

                CapturaTagImput(driver);

                CapturaTagA(driver);

                PercorreTagsImput(driver);

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

        private void PercorreTagsImput(IWebDriver driver)
        {
            string query = "";
            DataSet DS = new DataSet();
            DataSet DS1 = new DataSet();
            IWebElement element1 = null;
            int QtdPesquisas;
            



            try
            {
                

                query = "select * from TagsCapturados where TipoDeTag = 'input' and DataEhora is null order by CodCapturados";

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

                                    for (int r = 0; r < QtdPesquisas - 1; r++)
                                    {
                                        element1 = ElementosA[r];
                                    }

                                    query = " update TemplatesExecução set DataEhora =" +
                                            " '" + DateTime.Now.ToString("dd/mm/yyy hh:mm:ss") + "' " +
                                            " where CodTemplate = '" + RowTemplatesExecução["CodTemplate"] + "' ";

                                    DB.ExecutaQry(query);

                                    query = " update TagsCapturados set DataEhora =" +
                                            " '" + DateTime.Now.ToString("dd/mm/yyy hh:mm:ss") + "' " +
                                            " where CodCapturados = '" + RowTagsCapturados["CodCapturados"] + "' ";

                                    DB.ExecutaQry(query);
                                }
                                else
                                {
                                    query = " update TagsCapturados set DataEhora =" +
                                            " '" + DateTime.Now.ToString("dd/mm/yyy hh:mm:ss") + "' " +
                                            " where CodCapturados = '" + RowTagsCapturados["CodCapturados"] + "' ";
                                }

                            }
                        }

                        query = " update TemplatesExecução set DataEhora =" +
                                           " '" + DateTime.Now.ToString("dd/mm/yyy hh:mm:ss") + "' " +
                                           " where CodTemplate = '" + RowTemplatesExecução["CodTemplate"] + "' ";

                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: Automacao Método: PercorreTagsImput");
            }
        }

        private void ReiinciaTagsCapturados()
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

        private void CapturaTagA(IWebDriver driver)
        {
            int QtdPesquisas = 1;
            string NomeTag = "";
            string LocalTag = "";
            string Href = "";
            string query = "";
            IWebElement element1 = null;

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

                    query = " insert into TagsCapturados (TipoDeTag, LinkDaTag, LocalizacaoTag)" +
                            " values ('ref','" + Href.Replace("'", "₱") + "','" + LocalTag + "')";

                    DB.ExecutaQry(query);
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

                    query = " insert into TagsCapturados (TipoDeTag, NomeTag, LocalizacaoTag)" +
                            " values ('input','" + NomeTag.Replace("'", "₱") + "','" + LocalTag + "')";

                    DB.ExecutaQry(query);

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
                Qry = "select top 1 URLdoAplicativo  from URLdoTemplate";

                using (DR = DB.DR(Qry))
                {
                    if (DR.HasRows == true)
                    {
                        DR.Read();

                        return DR["URLdoAplicativo"].ToString();
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


    }
}

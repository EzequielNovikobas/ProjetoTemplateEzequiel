using AutomacaoGoogleAcademico.BLL;
using AutomacaoGoogleAcademicoDB.DAL;
using ScieloEzequiel.BLL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScieloEzequiel
{
   
    public partial class Processamento : Form
    {
        Automacao automacao;
        public Processamento()
        {
            InitializeComponent();
        }

        private void cmdIniciar_Click(object sender, EventArgs e)
        {

            Arquivo arquivo = new Arquivo();
            AccessDB AccDB = new AccessDB();
            string Caminho;

            try
            {
                //Caminho = arquivo.CriaPastaPesquisa(TxtPesquisa.Text);

                string Controle = DateTime.Now.ToString("MMddyyyyHHmmss");

                AccDB.ConectaDB();

                automacao = new Automacao("", "", Controle, AccDB); //instância classe de testes

                automacao.ParaThread = true;

                automacao.PreparaCapturaTag();

                
            }
            catch (Exception ex)
            {
                MessageBox.Show("TESTE:"+ex.Message);


                cmdIniciar.Enabled = true;
            }
        }
    }
}

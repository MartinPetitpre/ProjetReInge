using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.OleDb;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Xml.Linq;
using System.IO;
using System.Reflection;

namespace TST_EXCL
{
    public partial class _Default : Page
    {
        string nom;
        string prenom;
        string nullStr = "";
        protected void Page_Load(object sender, EventArgs e)
        {
            string[] TabChaine = new string[20];

            int i = 0;

            String Fournisseur = "Provider=Microsoft.Jet.OLEDB.4.0";
            String Adresse_Donnees = "Data Source=C:\\Users\\marti\\Downloads\\TMP_EXCl_JAN_2021.xls";
            // TMP_EXCl_DEC2015.xls";

            String Outils_Concernes = " Extended Properties=Excel 8.0";
            String Specification_Connexion = Fournisseur + ";" + Adresse_Donnees + ";" + Outils_Concernes;


            OleDbConnection Obj_Interop = new OleDbConnection(Specification_Connexion);

            Obj_Interop.Open();
            DataTable dtExcelSchema;

            //Obtenir le Schema du fichier Excel
            dtExcelSchema = Obj_Interop.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string Nom_Tab_Excl = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            OleDbCommand Cmnd_Selection = new OleDbCommand("SELECT * FROM [" + Nom_Tab_Excl + "]", Obj_Interop);

            // Créér un adaptateur pour récupérer les valeurs des cellules Excel

            OleDbDataAdapter Adaptateur_Recp_Donnees = new OleDbDataAdapter();
            // transfert des données depuis le fichier Execl vers l'adaptateur
            Adaptateur_Recp_Donnees.SelectCommand = Cmnd_Selection;

            DataSet Ens_Donnees = new DataSet();
            // remplir le Data Set avec le contenu de l'adaptateur
            Adaptateur_Recp_Donnees.Fill(Ens_Donnees);
            Datagrid1.DataSource = Ens_Donnees.Tables[0].DefaultView;
            // la structure 'Web Form' type  Datagrid1  est chargé du contenu du Data Set (Ensemble flexible des Donnees)
            Datagrid1.DataBind();
            Response.Write("<center>");
            int k = -1;
            foreach (DataRow DdL in Ens_Donnees.Tables[0].Rows)
            {
                k = k + 1;
                nom = DdL[0].ToString();
                // accès à chaque colonne
                i = 0;
                foreach (Object v in DdL.ItemArray)
                {

                    TabChaine[i] = v.ToString();
                    if (TabChaine[i] != nullStr)
                    // if ((TabChaine[i].Length != 0))

                    { Response.Write("<br><hr width='30%'> TabChaine[" + k + "," + i + "] =  " + TabChaine[i]); }
                    i = i + 1;
                }

                // TabChaine[i] = nom;
                //  Response.Write("<br><hr>" + nom);  
                Label1.Text = nom;
                prenom = DdL[1].ToString();


            }
            Obj_Interop.Close();


        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            Label1.Text = nom;
        }
    }
}
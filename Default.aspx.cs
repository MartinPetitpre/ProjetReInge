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
using System.Xml;
using System.Data.SqlClient;
using System.Configuration;

namespace Projet
{
    public partial class _Default : Page
    {
        string nom;
        string prenom;
        string nullStr = "";
        protected void Page_Load(object sender, EventArgs e)
        {
            


        }

        protected void excel_Vers_XML(Object Sender, EventArgs e)
        {
            string[] TabChaine = new string[20];

            int i = 0;

            String Fournisseur = "Provider=Microsoft.Jet.OLEDB.4.0";
            String Adresse_Donnees = "Data Source=C:\\Users\\marti\\Desktop\\projet_reinge\\excel_projet.xls";
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

                //affichage
                foreach (Object v in DdL.ItemArray)
                {

                    TabChaine[i] = v.ToString();
                    if (TabChaine[i] != nullStr)
                        // if ((TabChaine[i].Length != 0))

                        //{ Response.Write("<br><hr width='30%'> TabChaine[" + k + "," + i + "] =  " + TabChaine[i]); }
                        i = i + 1;
                }

                // TabChaine[i] = nom;
                //  Response.Write("<br><hr>" + nom);  
                Label1.Text = nom;
                prenom = DdL[1].ToString();


            }

            using (XmlWriter writer = XmlWriter.Create("C:\\Users\\marti\\Desktop\\projet_reinge\\excel_xml.xml"))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("PROJET");

                foreach (DataRow DdL in Ens_Donnees.Tables[0].Rows)
                {
                    if (DdL[0].ToString() != nullStr)
                    {
                        writer.WriteStartElement("individu");
                        writer.WriteElementString("id", DdL[0].ToString());
                        writer.WriteElementString("nom", DdL[1].ToString());
                        writer.WriteElementString("prenom", DdL[2].ToString());
                        writer.WriteEndElement();
                    }
                }

                writer.WriteEndDocument();
            }

            Obj_Interop.Close();
        }

        protected void access_Vers_XML(Object Sender, EventArgs e)
        {
            OleDbDataAdapter adaptateur = new OleDbDataAdapter();
            DataSet Grand_Set;

            // Objet de Connexion avec la Base de Données
            OleDbConnection objConn = new OleDbConnection
                ("Provider=Microsoft.ACE.OLEDB.12.0;" +
                "Data Source=C:\\Users\\marti\\Desktop\\projet_reinge\\access_projet.accdb");

            objConn.Open();

            OleDbCommand objCmd = new OleDbCommand("select * from projet", objConn);
            adaptateur.SelectCommand = objCmd;

            OleDbCommandBuilder cb = new OleDbCommandBuilder(adaptateur);
            Grand_Set = new DataSet("projet");
            adaptateur.Fill(Grand_Set, "projet");
            DataTable dt = Grand_Set.Tables[0];
            // Création de fichier XML qui sera initié avec la variable writer de type   XmlWriter
            using (XmlWriter writer = XmlWriter.Create("C:\\Users\\marti\\Desktop\\projet_reinge\\access_xml.xml"))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("PROJET");

                foreach (DataRow dr in dt.Rows)
                {
                    // Affichage des valeurs des champs pour s'assurer 
                    // que la capture des données via le DataSet s'est bien passé
                    /*  Response.Write
                        (dr[0].ToString() + "   " + dr[1].ToString() + "   " +                dr[2].ToString() +
                                            "   " + dr[4].ToString() + "<br>"); */

                    writer.WriteStartElement("individu");
                    writer.WriteElementString("id", dr[0].ToString());
                    writer.WriteElementString("nom", dr[1].ToString());
                    writer.WriteElementString("prenom", dr[2].ToString());
                    writer.WriteEndElement();

                }

                writer.WriteEndDocument();
            }
            // Affichage des valeurs des données dans une table XML

            /*OleDbDataReader objReader;

            objReader = objCmd.ExecuteReader();

            Response.Write("<center><table cellpadding=5 cellspacing=5 bgcolor=lightblue border=4>");
            Response.Write("<td align=center> <font size=+3 color=green>" + objReader.GetName(0) + "</td>");
            Response.Write("<td align=center> <font size=+3 color=green>" + objReader.GetName(1) + "</td>");
            Response.Write("<td align=center> <font size=+3 color=green>" + objReader.GetName(2) + "</td>");
            // Response.Write("<td align=center> <font size=+3 color=green>" + objReader.GetName(3) + "</td>");
            // Response.Write("<td align=center> <font size=+3 color=green>" + objReader.GetName(4) + "</td>");

            while (objReader.Read())
            {
                Response.Write("<tr>");
                Response.Write("<td align=center> <font size=+3 color=green>" + objReader.GetInt16(0) + "</td>");
                Response.Write("<td align=center> <font size=+3 color=green>" + objReader.GetString(1) + "</td>");
                Response.Write("<td align=center> <font size=+3 color=green>" + objReader.GetString(2) + "</td>");
                // Response.Write("<td align=center> <font size=+3 color=green>" + "Date Quelconque" + "</td>");
                // Response.Write("<td align=center> <font size=+3 color=green>" + objReader.GetDateTime(3) + "</td>");
                //Response.Write("<td align=center> <font size=+3 color=green>" + objReader.GetString(4) + "</td>");
            }
            Response.Write("</table></center>");
            objConn.Close();*/

        }

        protected void OpenSqlConnection(object sender, EventArgs e)
        {
            // C:\Users\Henri Basson\source\repos\SQL1SRV\App_Data\Database1.mdf

            string PARAMS_INTEROP =
                  "Data Source = (LocalDB)\\MSSQLLocalDB;" +
                    "AttachDbFilename = C:\\Users\\marti\\source\\repos\\Projet\\App_Data\\bdProjet.mdf;" +
                     "Integrated Security = True";
            // "AttachDbFilename = C:\\Users\\Henri Basson\\source\\repos\\SQLDB2\\SQLDB2\\App_Data\\Database2.mdf;" +

            SqlConnection connection = new SqlConnection(PARAMS_INTEROP);


            connection.Open();
            Response.Write("<br> connection " + connection.ServerVersion);
            Response.Write("<br> connection " + connection.State);
            string TABLE_CONCERNEE = "TableProjet";
            string REQ_SQL = "SELECT[id], [nom], [prenom] FROM[" + TABLE_CONCERNEE + "]";


            SqlDataAdapter dataAdapt = new SqlDataAdapter(REQ_SQL, connection);

            DataSet ds = new DataSet();

            dataAdapt.Fill(ds, "individu");

            ds.WriteXml("C:\\Users\\marti\\Desktop\\projet_reinge\\sql_xml.xml");




            using (SqlDataAdapter adaptateur = new SqlDataAdapter(
                    REQ_SQL, connection))
            {
                //  adaptateur.MissingSchemaAction = MissingSchemaAction.AddWithKey;

                DataTable dt = new DataTable();
                adaptateur.Fill(dt);


                Response.Write("<center> <br>= = = = = = = = = = = = = = = = = = = = = =<br>  ");
                Response.Write("&emsp;&emsp; Affichage du  Contenu la table " + TABLE_CONCERNEE);
                Response.Write("<br> = = = = = = = = = = = = = = = = = = = = = =  ");
                Response.Write("<table cellpadding='3' cellspacing='3' border='3' bordercolor='green'>");
                int i, j, k = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    i = 0;
                    // Affichage des valeurs des champs pour s'assurer 
                    // de la capture des données via le DataSet 
                    Response.Write("<tr>");

                    Response.Write("<td align=center> <font size=+1 color=green>&emsp;" + k + "</td>");
                    Response.Write("<td align=center> <font size=+1 color=blue>&emsp;" + dr[i].ToString() + " </td>");
                    Response.Write("<td align=center> <font size=+1 color=red>&emsp;" + dr[i + 1].ToString() + "</td>");
                    Response.Write("<td align=center> <font size=+1 color=purple>&emsp;" + dr[i + 2].ToString() + "</td>");
                    k++;
                }
                Response.Write("<table>");
                Response.Write("<center>");
            }

        }

        protected void GridView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            Label1.Text = nom;
        }
    }
}
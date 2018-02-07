using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace id_duplicates
{
    class Program
    {
        static Excel.Workbook wb_tops, wb_childs;
        static Excel.Worksheet ws_tops, ws_childs;
        static string answer;
        static string path;
        static string pathToTopParent;
        static string pathToChild;
        static string pathToResult;
        static Excel.Application xlApp = new Excel.Application();
        static Dictionary<string, sheetHandler> sheets = new Dictionary<string, sheetHandler>();
        static Dictionary<string, Excel.Workbook> wbs = new Dictionary<string, Excel.Workbook>();

        static int openWorkbooks(Excel.Workbook wb_tops, string folderPath)
        {
            string ws_name = "Workbooks";
            Excel.Worksheet ws_wbs;
            Excel.Workbook wb_swap;
            try
            {
                ws_wbs = wb_tops.Sheets[ws_name];
            } catch (Exception e)
            {
                Console.Write("Impossible d'ouvrir l'onglet :" + ws_name + " du classeur : " + wb_tops.Name);
                Console.ReadLine();
                return 0;
            }
            for (int i = 2; ws_wbs.Cells[i, 1].Value != "<END>"; i++)
            {
                try
                {
                    wb_swap = xlApp.Workbooks.Open(folderPath + "/" + ws_wbs.Cells[i, 1].Value);
                    wbs.Add(ws_wbs.Cells[i, 1].Value, wb_swap);
                }
                catch (Exception e)
                {
                    Console.Write("Echec lors de l'ouverture de tous les classeurs de résultat. Vérifier que le noms de ces classeurs et de ces onglets correspond aux noms indiqués dans l'onglet Workbooks du classeur tops.xlsx \n Echec sur le classeur :" + ws_wbs.Cells[i, 1].Value);
                    return 0;
                }
            }
            return 1;
        }

        static int openWorksheets(Excel.Workbook wb_tops)
        {
            Excel.Worksheet ws;
            try
            {
               ws = wb_tops.Sheets["Workbooks"];
            } catch (Exception e)
            {
                Console.Write("Impossible d'ouvrir l'onglet Workbooks du classeur : " + wb_tops.Name);
                return 0;
            }
            sheetHandler c_sh;
            for (int i = 2; ws.Cells[i, 2].Value != "<END>"; i++)
            {
                try
                {
                    c_sh = new sheetHandler(wbs[ws.Cells[i, 2].Value], ws.Cells[i, 3].Value);
                } catch (Exception e)
                {
                    Console.Write("Impossible d'ouvrir la feuille : " + ws.Cells[i, 3].Value + " dans le classeur : " + wbs[ws.Cells[i, 2].Value]);
                    return 0;
                }
                sheets.Add(ws.Cells[i, 2].Value + ws.Cells[i, 3].Value, c_sh);
            }
            return 1;
        }
   
        public class sheetHandler
        {
            private Dictionary<string, int> col_top;
            private Dictionary<string, int> row_top;
            private int top_col_dist = 1;
            private Excel.Worksheet sheet;
            
            public sheetHandler(Excel.Workbook wb, string sheetName)
            {
                this.col_top = new Dictionary<string, int>();
                this.row_top = new Dictionary<string, int>();
                try
                {
                    this.sheet = wb.Sheets[sheetName];
                } catch (Exception e)
                {
                    throw new Exception(e.Message);
                }
            }

            public void writeChild(string topName, string childName, string childID)
            {
                if (this.col_top.ContainsKey(topName) == false || this.row_top.ContainsKey(topName) == false)
                {
                    this.col_top[topName] = this.top_col_dist;
                    this.sheet.Cells[1, this.top_col_dist].Value = topName;
                    this.row_top[topName] = 2;
                    this.top_col_dist += 2;
                }

                this.sheet.Cells[row_top[topName], col_top[topName]] = childName;
                this.sheet.Cells[row_top[topName], col_top[topName] + 1] = childID;
                this.row_top[topName]++;
            }
        }
        
        static void generate_excel_files()
        {
            Excel.Workbook wb_tops = xlApp.Workbooks.Add();
            Excel.Workbook wb_childs = xlApp.Workbooks.Add();
            Excel.Worksheet ws_tops;
            Excel.Worksheet ws_childs;
            string path = "";

            try
            {
                ws_tops = wb_tops.Sheets["Sheet1"];
            } catch (Exception ex)
            {
                ws_tops = wb_tops.Sheets.Add();
                ws_tops.Name = "Sheet1";
            }

            Excel.Worksheet ws_wbks = wb_tops.Sheets.Add();
            ws_wbks.Name = "Workbooks";

            try
            {
                ws_childs = wb_childs.Sheets["Sheet1"];
            } catch (Exception ex)
            {
                ws_childs = wb_childs.Sheets.Add();
                ws_childs.Name = "Sheet1";
            }

            ws_tops.Cells[1, 1] = "ID";
            ws_tops.Cells[1, 2] = "NAME";
            ws_tops.Cells[1, 3] = "TRIPLETS";
            ws_tops.Cells[1, 4] = "FILE NAME";
            ws_tops.Cells[1, 5] = "TAB NAME";
            ws_wbks.Cells[1, 1].Value = "FILES NAMES";
            ws_wbks.Cells[1, 2].Value = "FILES NAMES FOR TABS";
            ws_wbks.Cells[1, 3].Value = "TABS NAMES";
            ws_childs.Cells[1, 1] = "DATA ID";
            ws_childs.Cells[1, 2] = "DATA NAME";

            Console.Write("Please specify path for generating Excel files");
            path = Console.ReadLine();
            wb_tops.SaveAs(path + "/tops.xlsx");
            wb_childs.SaveAs(path + "/childs.xlsx");
            wb_tops.Close();
            wb_childs.Close();
            xlApp.Quit();
        }

        static int identify_duplicates(string pathToTopParent, string pathToChild, string pathToResult)
        {
            Excel.Workbook wb_top;
            Excel.Workbook wb_child;
            Excel.Worksheet s_top;
            Excel.Worksheet s_child;

            try
            {
                wb_top = xlApp.Workbooks.Open(pathToTopParent);
                wb_child = xlApp.Workbooks.Open(pathToChild);
            } catch (Exception e)
            {
                Console.Write("Fonction id_duplicates, echec d'ouverture des classeurs Tops & Childs : " + e.Message);
                Console.ReadLine();
                return 0;
            }
            
            try
            {
                s_top = wb_top.Sheets["Sheet1"];
                s_child = wb_child.Sheets["Sheet1"];
            } catch (Exception e)
            {
                Console.Write("Fonction id_duplicates, echec d'ouverture des feuilles excel des classeurs tops & childs");
                Console.ReadLine();
                return 0;
            }


            string path, top_swap;
            double result;
            int found_duplicates = 0;
            int child_n;
            Dictionary<int, string> childs_dicctionnary = new Dictionary<int, string>();
            
            Console.Write("Ouverture des classeurs en cours...\n");
            Console.ReadLine();
            if (openWorkbooks(wb_top, pathToResult) == 0) return 0;
            Console.Write("Classeurs ouverts avec succès...\n");
            Console.ReadLine();
            Console.Write("Ouverture des feuilles excel...\n");
            Console.ReadLine();
            if (openWorksheets(wb_top) == 0) return 0;
            Console.Write("Feuilles Excel ouvertes avec succès...\n");
            Console.ReadLine();
            Console.Write("Chargement des childs dans la mémoire tampon \n");
            Console.ReadLine();
            for (int j = 2; s_child.Cells[j, 2].Value != "<END>"; j++)
            {
                Console.Write("\r Chargement du child : {0}", j);
                childs_dicctionnary.Add(j, s_child.Cells[j, 3].Value);
            }
            for (int i = 2; s_top.Cells[i, 2].Value != "<END>"; i++)
            {
                top_swap = s_top.Cells[i, 3].Value;
                child_n = 0;
                foreach (KeyValuePair<int, string> child in childs_dicctionnary)
                {
                    child_n++;
                    result = trigram_compare(top_swap, child.Value);
                    //s_top.Cells[i, 2].Value + " " + s_child.Cells[j, 2].Value + " " + result + " ");
                    if (result > 0.60)
                    {
                        sheets[s_top.Cells[i, 4].Value + s_top.Cells[i, 5].Value].writeChild(s_top.Cells[i, 2].Value, s_child.Cells[child.Key, 2].Value, s_child.Cells[child.Key, 1].Value.ToString());
                        found_duplicates++;
                    }
                    Console.Write("\r Child : {0} - Top : {1}  |  Doublons : {2}    ", child_n, i - 1, found_duplicates);
                }
            }
            Console.Write("Traitement des données terminé avec succès.\n");
            Console.Write("Sauvegarde et fermeture des classeurs...\n");
            foreach (KeyValuePair<string, Excel.Workbook> entry in wbs)
            {
                entry.Value.Save();
                entry.Value.Close();
            }

            xlApp.Quit();
            Console.Write("Fin.");
            Console.ReadLine();
            return 1;
        }

        static double trigram_compare(string topParentTriplet, string childTriplets)
        {
            double result;
            double n_triplets_top = topParentTriplet.Split(';').Count();
            double n_triplets_child = childTriplets.Split(';').Count();
            double count_sim = 0.00;
            string[] top_triplets = topParentTriplet.Split(';');
            string[] child_triplets = childTriplets.Split(';');

            foreach (string top_triplet in top_triplets)
            {
                foreach (string child_triplet in child_triplets)
                {
                    if (child_triplet == top_triplet)
                    {
                        count_sim++;
                    }
                }
            }
            result = ((count_sim * 2) / (n_triplets_child + n_triplets_top));
            return result;
        }

        static int create_triplets(string pathToExcelFile)
        {
            Excel.Application xlApp = new Excel.Application();
            int row = 2;
            int col = 2;
            string full_string;
            string trigrams;
            int count_trim;

            try
            {
                wb_tops = xlApp.Workbooks.Open(pathToExcelFile);
                Console.Write("Fichier Excel ouvert avec succès \n");
                ws_tops = wb_tops.Sheets["Sheet1"];
                Console.Write("Programme prêt à traiter les données.. \n");
            }
            catch (Exception ex)
            {
                Console.Write("L'erreur suivante est survenue lors du démarrage de l'application : " + ex.Message + "\n");
                System.IO.File.WriteAllText("errors.log", ex.Message + "\n");
                Console.ReadLine();
                if (wb_tops != null) wb_tops.Close();
                xlApp.Quit();
                return -1;
            }
            Console.Write("Traitement des données en cours... \n");
            do
            {
                Console.Write("\r Traitement de la ligne : {0}", row);
                count_trim = 2;
                trigrams = "";
                full_string = " " + ws_tops.Cells[row, col].Value.ToUpper() + " ";

                while (count_trim < full_string.Length)
                {
                    if (count_trim + 1 == full_string.Length)
                        trigrams += full_string.Substring(count_trim - 2, 3);
                    else trigrams += full_string.Substring(count_trim - 2, 3) + ";";

                    count_trim++;
                }
                ws_tops.Cells[row, col + 1].Value = trigrams;
                row++;
            } while (ws_tops.Cells[row, col].Value + "" != "");
            Console.Write("\n");
            wb_tops.Save();
            wb_tops.Close();
            xlApp.Quit();
            Console.Write("Traitement fini avec succès. \n");
            Console.ReadLine();
            return 0;
        }

        static void Main(string[] args)
        {

            do
            {
                Console.Write("Please select option : \n");
                Console.Write("1 - Generate Excel Files for processing data\n");
                Console.Write("2 - Create Triplets\n");
                Console.Write("3 - Identify potential duplicates\n");
                Console.Write("4 - Quit Application\n");
                answer = Console.ReadLine();
            } while (answer != "4" && answer != "3" && answer != "2" && answer != "1");

            switch (answer)
            {
                case "1":
                    generate_excel_files();
                    break;
                case "2":
                    Console.Write("Please indicate path to the excel file \n");
                    path = Console.ReadLine();
                    create_triplets(path);
                    break;
                case "3":
                    Console.Write("Please indicate path to the file containing Top Parent information. Please note that triplets need to be created first. \n");
                    pathToTopParent = Console.ReadLine();
                    Console.Write("Please indicate path to the file containing child information. Please note that triplets need to be created first \n");
                    pathToChild = Console.ReadLine();
                    Console.Write("Please indicate path for result files.\n");
                    pathToResult = Console.ReadLine();
                    identify_duplicates(pathToTopParent, pathToChild, pathToResult);
                    break;
                case "4":
                    return;
                default:
                    return;
            }
        }
    }
}

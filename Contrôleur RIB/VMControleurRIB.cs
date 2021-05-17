using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using static Contrôleur_RIB.ViewModelBase;
using MessageBox = System.Windows.MessageBox;// Solves the ambiguous reference

namespace Contrôleur_RIB
{
    class VMControleurRIB : ViewModelBase
    {
        private ExcelApp excelApp;
        private String loadedFileText;
        private String processProgressText;
        public Command OpenExcelFile { get; set; }
        public Command AnalyseRIB { get; set; }
        public Command CloseExcelFile { get; set; }

        public VMControleurRIB()
        {
            OpenExcelFile = new Command(OpenExcelFile_Func);
            AnalyseRIB = new Command(AnalyseRIB_Func);
            CloseExcelFile = new Command(CloseExcelFile_Func);
            LoadedFileText = "Aucun fichier chargé";
            ProcessProgressText = "";
        }

        private void OpenExcelFile_Func()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();// Instanciates the class
            openFileDialog.Filter = "Tableau Excel |*.xlsx;*.xlsm";// Adds a filter for Excel files only
            if (openFileDialog.ShowDialog() == DialogResult.OK)// Enters the IF if the user selected a file and clicked OK
            {
                if(ExcelApp != null)
                {
                    ExcelApp.Terminate();// Automatically releases the previous file when the user loads a new one without releasing first.
                }
                ExcelApp = new ExcelApp(openFileDialog.FileName);// Initating my Excel object and passing the Excel file to open
                LoadedFileText = "Fichier chargé : "+openFileDialog.SafeFileName;//Displays the proper filename instead of the entire path for the UI
                OnPropertyChanged("LoadedFileText");// Updating the view
                // Couldn't find a way to select the button to re-enable it. Going to check if document is loaded directly inside the next function.
            }
        }

        private void CloseExcelFile_Func()// Release the file for use without closing the application
        {
            ExcelApp.Terminate();
            LoadedFileText = "Aucun fichier chargé";// Restoring default value
            OnPropertyChanged("LoadedFileText");
        }

        private void AnalyseRIB_Func()
        {
            if (ExcelApp != null)
            {
                const int columnToOverwrite = 4;// Column we want to write the results in. This is its only definition place, everything else works by passing the value through, for easier changes.
                if (ExcelApp.ColumnIsEmpty(columnToOverwrite))// If that column is empty in every line we say nothing
                {
                    List<String> listOfRIBs = new List<String>();
                    listOfRIBs = ExcelApp.GetAllRIBs();// Extracting all ribs from the Excel file
                    AnalyseAllRIBs(listOfRIBs, columnToOverwrite);
                }
                else// If we've found something in the column, display a message to ask user confirmation to overwrite it with the results of the alayse
                {
                    if (MessageBox.Show("La colonne N°"+columnToOverwrite+" n'est pas vide, poursuivre l'exécution écrasera le contenu. Continuer ?", "ControleurRIB", MessageBoxButton.YesNo) == MessageBoxResult.Yes)// Returns true is the user clicks "Yes", false if "No"
                    {
                        List<String> listOfRIBs = new List<String>();
                        listOfRIBs = ExcelApp.GetAllRIBs();// Extracting all ribs from the Excel file
                        AnalyseAllRIBs(listOfRIBs, columnToOverwrite);
                    }
                    else
                    {
                        return;// If the user clicks "No", the function stops
                    }
                }
            }
            else
            {
                MessageBox.Show("Veuillez charger un fichier à analyser", "ControleurRIB", MessageBoxButton.OK, MessageBoxImage.Information);// User clicks "Analyse" button without having a file loaded. No need for this if we can disable the button prior to loading a file.
            } 
        }
        private void AnalyseAllRIBs(List<String> listOfRIBs, int columnToOverwrite)
        {
            List<String> results = new List<String>();

            ProcessProgressText = "Nombre de RIB à traiter : "+listOfRIBs.Capacity.ToString();
            OnPropertyChanged("ProcessProgressText");

            foreach (var rib in listOfRIBs)
            {
                bool isValid = true;// Boolean for RIB validity
                String result = "";
                char[] separators = new char[] { '-', ' ' };
                String[] substrings;

                if (rib.Contains("-"))// Checking for - as a separating character
                {
                    substrings = rib.Split(separators[0]);
                }
                else // Checking for space as a separating character
                {
                    substrings = rib.Split(separators[1]);
                }
                // Would need a default case but I don't know which one.

                // Checking length and int for each substring
                if (substrings[0].Length != 5)
                {
                    result += "Problème de longueur sur le numéro de banque ";
                    isValid = false;
                }
                else
                {
                    int number;
                    if (!Int32.TryParse(substrings[0], out number))
                    {
                        result += "Le numéro de banque n'est pas un chiffre ";
                        isValid = false;
                    }
                }

                if (substrings[1].Length != 5)
                {
                    result += "Problème de longueur sur le numéro de guichet ";
                    isValid = false;
                }
                else
                {
                    int number;
                    if (!Int32.TryParse(substrings[1], out number))
                    {
                        result += "Le numéro de guichet n'est pas un chiffre ";
                        isValid = false;
                    }
                }

                if (substrings[2].Length != 11)
                {
                    result += "Problème de longueur sur le numéro de compte ";
                    isValid = false;
                }

                if (substrings[3].Length != 2)
                {
                    result += "Problème de longueur sur la clé RIB ";
                    isValid = false;
                }
                else
                {
                    int number;
                    if (!Int32.TryParse(substrings[3], out number))
                    {
                        result += "La clé RIB n'est pas un chiffre ";
                        isValid = false;
                    }
                }

                if(isValid)// No need to do the maths if we know we have incorrect values
                {
                    string b_s = substrings[0];
                    string g_s = substrings[1];
                    string c_s = substrings[2];
                    string k_s = substrings[3];

                    // Remplacement des lettres par des chiffres dans le numéro de compte
                    StringBuilder sb = new StringBuilder();
                    foreach (char ch in c_s.ToUpper())
                    {
                        if (char.IsDigit(ch))
                            sb.Append(ch);
                        else
                            sb.Append(RibLetterToDigit(ch));
                    }
                    c_s = sb.ToString();

                    // Séparation du numéro de compte pour tenir sur 32bits
                    string d_s = c_s.Substring(0, 6);
                    c_s = c_s.Substring(6, 5);

                    // Calcul de la clé RIB

                    int b = int.Parse(b_s);
                    int g = int.Parse(g_s);
                    int d = int.Parse(d_s);
                    int c = int.Parse(c_s);
                    int k = int.Parse(k_s);

                    int calculatedKey = 97 - ((89 * b + 15 * g + 76 * d + 3 * c) % 97);

                    if (k == calculatedKey)
                    {
                        result += "OK";
                    }
                    else
                    {
                        result += "Erreur dans le calcul de la clé de RIB";
                    }
                }
                results.Add(result);
            }
            ExcelApp.WriteResults(results, columnToOverwrite);
            ProcessProgressText += "\nTraitement terminé.";
            OnPropertyChanged("ProcessProgressText");
        }

        /// <summary>
        /// Convertit une lettre d'un RIB en un chiffre selon la table suivante :
        /// 1 2 3 4 5 6 7 8 9
        /// A B C D E F G H I
        /// J K L M N O P Q R
        /// _ S T U V W X Y Z
        /// </summary>
        /// <param name="letter">La lettre à convertir</param>
        /// <returns>Le chiffre de remplacement</returns>
        public char RibLetterToDigit(char letter)
        {
            if (letter >= 'A' && letter <= 'I')
            {
                return (char)(letter - 'A' + '1');
            }
            else if (letter >= 'J' && letter <= 'R')
            {
                return (char)(letter - 'J' + '1');
            }
            else if (letter >= 'S' && letter <= 'Z')
            {
                return (char)(letter - 'S' + '2');
            }
            else
                throw new ArgumentOutOfRangeException("Le caractère à convertir doit être une lettre majuscule dans la plage A-Z");
        }

        public ExcelApp ExcelApp {get; set; }
        public String LoadedFileText {get; set; }
        public String ProcessProgressText { get; set; }
    }
}

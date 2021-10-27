using IfcDemo;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Xbim.Ifc;

using Xbim.Ifc4.Interfaces;
//using Xbim.Ifc4.Interfaces;

namespace MantovaRocks
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private object debug;

		public MainWindow()
		{
			InitializeComponent();
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			MessageBox.Show("Ciao");
		}
	
		private static Dictionary<string, VoceDiListino> CaricaListinoExcel(string excelFile, string sheetName)
		{
			var TuttoIlListino = new Dictionary<string, VoceDiListino>();

			var sheet = AnExcelReader.GetExcelSheet(excelFile, sheetName);
			int rowCount = sheet.LastRowNum;
			for (int i = 1; i <= rowCount; i++)
			{
				IRow curRow = sheet.GetRow(i);
				var codice = curRow.GetCell(0).StringCellValue.Trim();
				var cost = curRow.GetCell(1).NumericCellValue;
				var logica = curRow.GetCell(2).StringCellValue.Trim();

				var t = VoceDiListino.Create(logica, cost);
				if (t != null && !TuttoIlListino.ContainsKey(codice))
				{
					TuttoIlListino.Add(codice, t);
				}
			}

			return TuttoIlListino;
		}

		private Dictionary<string, VoceDiListino> CaricaListinoAccess(string accessFileName)
		{
			var TuttoIlListino = new Dictionary<string, VoceDiListino>();
			OleDbConnection myConn = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={accessFileName};Persist Security Info=False;");
			myConn.Open();
			OleDbCommand myQuery = new OleDbCommand("select * from PrezziCorrenti", myConn);
			OleDbDataReader myReader = myQuery.ExecuteReader();
			while (myReader.Read())
			{
				var codice = myReader["codice"].ToString().Trim();
				var cost = Convert.ToDouble(myReader["Costo"]);
				var logica = myReader["Logica"].ToString().Trim();
				var t = VoceDiListino.Create(logica, cost);
				if (t != null && !TuttoIlListino.ContainsKey(codice))
				{
					TuttoIlListino.Add(codice, t);
				}
			}
			myConn.Close();
			return TuttoIlListino;
		}

		private void EseguIlComputo(object sender, RoutedEventArgs e)
		{
			// ogni contenuto dopo le doppie barre viene ignorato
			var IfcFileName = NomeFile.Text;
			if (!File.Exists(IfcFileName))
			{
				MessageBox.Show($"Il file di modello {IfcFileName} non e' stato trovato.");
				return;
			}

			var excelFileName = NomeListino.Text;
			if (!File.Exists(excelFileName))
			{
				MessageBox.Show($"Il file di listino {excelFileName} non e' stato trovato.");
				return;
			}

			// carico sia il modello che il listino
			var modello = IfcStore.Open(IfcFileName);
			var listinoCaricatoDaExcel = CaricaListinoExcel(excelFileName, "Ottobre"); // carico il listino di ottobre

			txtRisultato.Text = CalcolaPreventivo(modello, listinoCaricatoDaExcel);
		}

		private static string CalcolaPreventivo(IfcStore modello, Dictionary<string, VoceDiListino> listinoDaUsare)
		{
			// ora valuto tutti gli oggetti computabili (nello specifico gli IfcBuildingElements
			var totalePreventivo = 0.0; // se non ci sono elementi il preventivo e' 0
			var computableObjects = modello.Instances.OfType<IIfcBuildingElement>();

			StringBuilder sb = new StringBuilder();
			sb.AppendLine("Elenco degli oggetti computabili:");
			foreach (var oneObject in computableObjects)
			{
				// ora cerco il nome della famiglia e del tipo
				//
				string codice = PrendiIlNomeDellaFamiglia(oneObject);
				if (listinoDaUsare.TryGetValue(codice, out var voce))
				{
					var questoElemento = voce.Calcola(oneObject, out var errore);
					sb.AppendLine($"{oneObject.EntityLabel}\t{codice}\t{questoElemento:c2}\t{errore}");
					totalePreventivo += questoElemento;
				}
				else
				{
					sb.AppendLine($"{oneObject.EntityLabel}\t{codice}\t\tCodice di listino non trovato");
				}
			}
			sb.AppendLine("");
			sb.AppendLine($"Totale computato: {totalePreventivo:C2}");

			return sb.ToString();
		}

		private static string PrendiIlNomeDellaFamiglia(IIfcBuildingElement oneObject)
		{
			var partiDelNome = oneObject.Name.ToString().Split(new[] { ':' });
			partiDelNome = partiDelNome.Take(partiDelNome.Length - 1).ToArray();
			var codice = string.Join(":", partiDelNome);
			return codice;
		}

		private void EseguIlComputoAccess(object sender, RoutedEventArgs e)
		{
			// ogni contenuto dopo le doppie barre viene ignorato
			var IfcFileName = NomeFile.Text;
			if (!File.Exists(IfcFileName))
			{
				MessageBox.Show($"Il file di modello {IfcFileName} non e' stato trovato.");
				return;
			}

			var accessFileName = NomeListinoAccess.Text;
			if (!File.Exists(accessFileName))
			{
				MessageBox.Show($"Il file di listino {accessFileName} non e' stato trovato.");
				return;
			}

			// carico sia il modello che il listino
			var modello = IfcStore.Open(IfcFileName);
			var listinoCaricatoDaAccess = CaricaListinoAccess(accessFileName); // carico il listino corrente

			txtRisultato.Text = CalcolaPreventivo(modello, listinoCaricatoDaAccess);
		}

		
	}
}

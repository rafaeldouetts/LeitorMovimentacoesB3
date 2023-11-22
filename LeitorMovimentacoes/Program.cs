
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop;
using LinqToExcel;
using System.Reflection;
using LeitorMovimentacoes.Model;
using OfficeOpenXml;
using LeitorMovimentacoes.Enum;
using System.Text;
using LinqToExcel.Extensions;
using ClosedXML.Excel;

namespace NetCore_Monitor
{
	class Program
	{
		private static FileSystemWatcher _monitorar;
		public static void MonitorarArquivos(string path, string filtro)
		{
			Valid(path);

			_monitorar = new FileSystemWatcher(path, filtro)
			{
				IncludeSubdirectories = true
			};

			_monitorar.Created += OnFileChanged;
			//_monitorar.Changed += OnFileChanged;
			//_monitorar.Deleted += OnFileChanged;
			//_monitorar.Renamed += OnFileRenamed;

			_monitorar.EnableRaisingEvents = true;
			Console.WriteLine($"Monitorando arquivos e: {filtro}");
		}
		private static void OnFileChanged(object sender, FileSystemEventArgs e)
		{
			 Console.WriteLine($"O Arquivo {e.Name} {e.ChangeType}");
			Leitor(e.FullPath);
		}
		private static void OnFileRenamed(object sender, RenamedEventArgs e)
		{
			Console.WriteLine($"O Arquivo {e.OldName} {e.ChangeType} para {e.Name}");
		}
		static void Main(string[] args)
		{
			Console.WriteLine("Monitorando a pasta com o sistema : LeitorMovimentacoes");
			string path = @"c:\plataforma\movimentacoes";
			string filtro = "*.xlsx";
			MonitorarArquivos(path, filtro);
			Console.ReadLine();
		}

		private static void Valid (string path)
		{
			if (Directory.Exists(path)) return;

			Directory.CreateDirectory(path);
		}

		private static void Ler()
		{

		}

		public static List<Movimentacao> Leitor(string path)
		{
			var response = new List<Movimentacao>();
			var xls = new XLWorkbook(path);
			var planilha = xls.Worksheets.First(w => w.Name == "Movimentação");
			var totalLinhas = planilha.Rows().Count();
			// primeira linha é o cabecalho
			for (int l = 2; l <= totalLinhas; l++)
			{
				var movimentacao = new Movimentacao();

				movimentacao.Tipo =planilha.Cell($"A{l}").Value.ToString() == "Credito" ? TipoMovimentacao.Entrada: TipoMovimentacao.Saida ;
				movimentacao.Data = DateTime.Parse(planilha.Cell($"B{l}").Value.ToString());
				movimentacao.Descricacao = planilha.Cell($"C{l}").Value.ToString();
				movimentacao.Produto = planilha.Cell($"D{l}").Value.ToString();
				movimentacao.Instituto = planilha.Cell($"E{l}").Value.ToString();
				movimentacao.Quantidade = int.Parse(planilha.Cell($"F{l}").Value.ToString());
				movimentacao.Preco = Decimal.Parse(planilha.Cell($"G{l}").Value.ToString());
				movimentacao.TotalOperacao = Decimal.Parse(planilha.Cell($"h{l}").Value.ToString());
			
				response.Add(movimentacao);
			}

			return response;
		}
	}
}
using LeitorMovimentacoes.Enum;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;

namespace LeitorMovimentacoes.Model
{
	public class Movimentacao
	{
		public TipoMovimentacao Tipo { get; set; }
		public DateTime Data { get; set; }
		public string Descricacao { get; set; }
        public string Produto { get; set; }
        public string Instituto { get; set; }
        public int Quantidade { get; set; }
        public Decimal Preco { get; set; }
		public Decimal TotalOperacao { get; set; }

	}
}

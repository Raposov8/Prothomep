using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.Inpart.Models
{
    public class ItemCotacao
    {
         public int? idSeguradora { get; set; }
         public int? idCotacao { get; set; }
         public int? idSucursal { get; set; }
         public string? cdProdutoInpart { get; set; }
         public string? cdProdutoDistribuidor { get; set; }
         public string? cdProdutoFabricante { get; set; }
         public int? vlQuantidade { get; set; }
         public int? cdStatusItem { get; set; }
         public double? vlPrecoItem { get; set; }
         public int? idDistribuidor { get; set; }
         public int? idCotacaoItem { get; set; }
         public int? idFabricante { get; set; }
         public int? idCliente { get; set; }
         public bool flConsignadoPermanente { get; set; }
         public double? vlQuantidadeAutorizada { get; set; }
         public double? vlPrecoPrestador { get; set; }
         public int? iditemnegociacao { get; set; }
         public int? utilizaValorProcedimento { get; set; }
         public double? vlPrecoNegociadoInicial { get; set; }
         public Fabricante fabricante { get; set; }
         public Distribuidor distribuidor { get; set; }
    }
}

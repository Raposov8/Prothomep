using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class RecursoIntrumentador
    {
        [Key]
        public int Id { get; set; }

        public string? Nome { get; set; }
        public DateTime? Nascimento { get; set; }
        public string? Endereco { get; set; }
        public string? Bairro { get; set; }
        public string? Municipio { get; set; }
        public string? Estado { get; set; }
        public string? Pais { get; set; }
        public string? Email { get; set; }
        public string? Telefone { get; set; }
        public string? Telefone2 { get; set; }
        public string? CNPJ { get; set; }
        public string? NomeEmpresa { get; set; }
        public string? Tipo { get; set; }

        //DEMAIS DADOS
        public string? CBO { get; set; }
        public string? RG { get; set; }
        public string? CPF { get; set; }
        public string? PISPASEP { get; set; }
        public string? TipoChave { get; set; }
        public string? ChavePix { get; set; }
        public string? Banco { get; set; }
        public string? Ag { get; set; }
        public string? CC { get; set; }
        public string? Remuneracao { get; set; }

        //OBS
        public string? ServCon { get; set; }

        //Protheus
        public string? IdProtheus { get; set; }
        public string? EmpresaProtheus { get; set; }
    }
}

using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;

namespace SGID.Data.ViewModel
{
    public class Agendamentos
    {
        [Key]
        public int Id { get; set; }
        public int StatusCotacao { get; set; }
        public int StatusPedido { get; set; }

        //Protheus
        public string Empresa { get; set; }
        public string Cliente { get; set; }
        public string? CodCondPag { get; set; }
        public string? CondPag { get; set; }
        public string? CodTabela { get; set; }
        public string? Tabela { get; set; }
        public string Vendedor { get; set; }
        public string VendedorLogin { get; set; }
        public string Medico { get; set; }
        public string? Matricula { get; set; }
        public string Paciente { get; set; }
        public string CodHospital { get; set; }
        public string Hospital { get; set; }
        public string? CodConvenio { get; set; }
        public string? Convenio { get; set; }
        public string Instrumentador { get; set; }
        public DateTime? DataCirurgia { get; set; }
        public DateTime? DataAutorizacao { get; set; }
        public string Procedimento { get; set; }
        public string? NumAutorizacao { get; set; }
        public string? Senha { get; set; }
        public int Tipo { get; set; }
        public int Autorizado { get; set; }
        public string? Indicacao { get; set; }
        public string? Observacao { get; set; }
        [ForeignKey("AgendamentoId")]
        public ICollection<AnexosAgendamentos> Anexos { get; set; }

        public DateTime? DataCotacao { get; set; } 

        //ADM
        public string? UsuarioCriacao { get; set; }
        public DateTime DataCriacao { get; set; }
        public string? UsuarioAlterar { get; set; }
        public DateTime DataAlteracao { get; set; }

        //INSTRUMENTACAO
        public string? UsuarioInstrumentador { get; set; }
        public string? UsuarioGestorInstrumentador { get; set; }
        public DateTime? DataInstrumentador { get; set; }
        public int StatusInstrumentador { get; set; }
        public string? Latitude { get; set; }
        public string? Longitude { get; set; }

        //COMERCIAL
        public string? UsuarioComercial { get; set; }
        public DateTime? DataComercial { get; set; }

        //COMERCIAL VOLTA
        public string? UsuarioComercialAprova { get; set; }
        public DateTime? DataComercialAprova { get; set; }

        //ADM APROVA
        public string? UsuarioAprova { get; set; }
        public DateTime? DataAprova { get; set; }

        //Rejeição
        public string? UsuarioRejeicao { get; set; }
        public string? MotivoRejeicao { get; set; }
        public string? ObsRejeicao { get; set; }

        //Logistica
        public string? UsuarioLogistica { get; set; }
        public DateTime? DataLogistica { get; set; }

        //Logistica Loteamento
        public string? UsuarioLoteamento { get; set; }
        public DateTime? DataLoteamento { get; set; }
        public int StatusLogistica { get; set; }
        public string? MediaEntrega { get; set; }
        public DateTime? DataEntrega { get; set; }
        public DateTime? DataRetorno { get; set; }
    }
}

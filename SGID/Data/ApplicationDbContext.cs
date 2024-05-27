using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using SGID.Data.ViewModel;

namespace SGID.Data
{
    public class ApplicationDbContext : IdentityDbContext<UserInter>
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options)
        {
            this.Database.SetCommandTimeout(300);
        }

        //Venda
        public DbSet<Agendamentos> Agendamentos { get; set; }
        public DbSet<ObsAgendamento> ObsAgendamentos { get; set; }
        public DbSet<AnexosAgendamentos> AnexosAgendamentos { get; set; }
        public DbSet<PatrimonioAgendamento> PatrimoniosAgendamentos { get; set; }
        public DbSet<ProdutosAgendamentos> ProdutosAgendamentos { get; set; }
        public DbSet<RejeicaoMotivos> RejeicaoMotivos { get; set; }

        //Instrumentador
        public DbSet<AgendamentoInstrumentador> AgendamentoInstrumentadors { get; set; }
        public DbSet<RecursoIntrumentador> Instrumentadores { get; set; }
        public DbSet<DadosCirurgia> DadosCirurgias { get; set; }
        public DbSet<AnexosDadosCirurgia> AnexosDadosCirurgias { get; set; }
        public DbSet<DadosCirugiasProdutos> DadosCirugiasProdutos { get; set; }


        //Comercial Cotações
        public DbSet<AvulsosAgendamento> AvulsosAgendamento { get; set; }
        public DbSet<Procedimento> Procedimentos { get; set; }

        //Comercial Visitas

        public DbSet<Visitas> Visitas { get; set; }
        public DbSet<VisitaCliente> VisitaClientes { get; set; }

        //RH
        public DbSet<SolicitacaoAcesso> SolicitacaoAcessos { get; set; }
        public DbSet<AcessoTermo> AcessoTermos { get; set; }


        //Estoque
        public DbSet<FormularioAvulso> FormularioAvulsos { get; set; }
        public DbSet<FormularioAvulsoXProdutos> FormularioAvulsoXProdutos { get; set; }
        public DbSet<Ocorrencia> Ocorrencias { get; set; }
        public DbSet<AgendamentoCheck> AgendamentoChecks { get; set; }


        //Comissões
        public DbSet<Time> Times { get; set; }
        public DbSet<TimeProduto> TimeProdutos { get; set; }
        public DbSet<TimeADM> TimeADMs { get; set; }
        public DbSet<TimeDental> TimeDentals { get; set; }

        //Inventario
        public DbSet<Dispositivo> Dispositivos { get; set; }

        public DbSet<UsuarioDispositivo> UsuarioDispositivos { get; set; } 

        //Sistema
        public DbSet<Meta> Metas { get; set; }
        public DbSet<Log> Logs { get; set; }
    }
}
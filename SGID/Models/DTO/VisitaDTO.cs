namespace SGID.Models.DTO
{
    public class VisitaDTO
    {
        public int Id { get; set; }
        public DateTime DataCriacao { get; set; }
        public string DataHora { get; set; }
        public string DataUltima { get; set; }
        public string Local { get; set; }
        public string Assunto { get; set; }
        public string Observacao { get; set; }
        public string Endereco { get; set; }
        public string Medico { get; set; }
        public string Motvisita { get; set; }
        public string Bairro { get; set; }
        public string ResuVi { get; set; }
        public string Vendedor { get; set; }
        public int Status { get; set; }
        public string DataProxima { get; set; }
        public string UltimaResp1 { get; set; }
        public string DataResp1 { get; set; }
        public string UltimaResp2 { get; set; }
        public string DataResp2 { get; set; }

        public string Email { get; set; }
        public string Telefone { get; set; }
        public int Latitude { get; set; }
    }
}

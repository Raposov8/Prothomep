using System;
using System.Collections.Generic;

namespace SGID.Models.Cadastro
{
    public partial class Telefone
    {
        public int TelefoneId { get; set; }
        public string? Celular { get; set; }
        public string? Idcelular { get; set; }
        public int? UsuarioId { get; set; }
        public bool Ativo { get; set; }

        public virtual Usuario? Usuario { get; set; }
    }
}

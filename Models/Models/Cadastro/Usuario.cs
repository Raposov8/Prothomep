using System;
using System.Collections.Generic;

namespace SGID.Models.Cadastro
{
    public partial class Usuario
    {
        public Usuario()
        {
            Telefones = new HashSet<Telefone>();
        }

        public int UsuarioId { get; set; }
        public string? Nome { get; set; }
        public string? Senha { get; set; }
        public string? Email { get; set; }

        public virtual ICollection<Telefone> Telefones { get; set; }
    }
}

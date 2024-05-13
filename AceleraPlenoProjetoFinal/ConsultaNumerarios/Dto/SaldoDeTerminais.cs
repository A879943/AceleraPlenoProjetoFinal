using ConsultaNumerarios.Models;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.EntityFrameworkCore.Metadata.Conventions;

namespace ConsultaNumerarios.Dto
{
    public class SaldoDeTerminais
    {
        public decimal SaldoTotal { get; set; }
        public List<Terminal> Terminal { get; set; }

    }

    public class Terminal
    {
        public int Id { get; set; }
        public decimal Saldo { get; set; }
        public decimal LimiteMax { get; set; }
        public decimal LimiteMin { get; set; }
        public string Usuario { get; set; }
        public bool DentroDoLimite { get; set; }
    }
}

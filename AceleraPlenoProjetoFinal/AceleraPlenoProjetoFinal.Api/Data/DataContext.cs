using AceleraPlenoProjetoFinal.Api.Models;
using Microsoft.EntityFrameworkCore;

namespace AceleraPlenoProjetoFinal.Api.Data;

public class DataContext : DbContext
{
    public DataContext() { }

    public DataContext(DbContextOptions<DataContext> opt) : base(opt) { }

    public DbSet<TransportadoraValoresModel> TransportadoraValores { get; set; }

    public DbSet<UnidadeInstituicaoModel> UnidadeInstituicao { get; set; }

    public DbSet<UnidadeInstituicaoTransportadoraValoresModel> UnidadeInstituicaoTransportadoraValores { get; set; }

    public DbSet<TipoTerminalModel> TipoTerminal  { get; set; }

    public DbSet<TipoOperacaoModel> TipoOperacao { get; set; }

    public DbSet<OperacaoModel> Operacao { get; set; }
}
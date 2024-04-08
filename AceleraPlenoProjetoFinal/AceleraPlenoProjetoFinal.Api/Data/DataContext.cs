using AceleraPlenoProjetoFinal.Api.Models;
using Microsoft.EntityFrameworkCore;

namespace AceleraPlenoProjetoFinal.Api.Data;

public class DataContext : DbContext
{
    public DataContext() { }

    public DataContext(DbContextOptions<DataContext> opt) : base(opt) { }

    public DbSet<TransportadoraValoresModel> TransportadoraValores { get; set; }

    public DbSet<TipoTerminalModel> TipoTerminal  { get; set; }
}
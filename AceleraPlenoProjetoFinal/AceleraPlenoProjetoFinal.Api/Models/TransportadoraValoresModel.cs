using System.ComponentModel.DataAnnotations.Schema;

namespace AceleraPlenoProjetoFinal.Api.Models;

[Table("TBTRANSPORTADORAVALORES")]
public class TransportadoraValoresModel
{
    [Column("ID")]
    public int Id { get; set; }

    [Column("NUMCNPJ")]
    public string NumeroCnpj { get; set; }

    [Column("DESCTRANSPORTADORA")]
    public string DescricaoTransportadora { get; set; }

    [Column("PA")]
    public string PA { get; set; }

    [Column("DATAHORACARGA")]
    public DateTime DataHoraCarga { get; set; }
}
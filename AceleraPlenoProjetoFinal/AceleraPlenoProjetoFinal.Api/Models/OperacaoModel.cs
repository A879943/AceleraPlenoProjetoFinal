using System.ComponentModel.DataAnnotations.Schema;

namespace AceleraPlenoProjetoFinal.Api.Models;

[Table("TBOPERACAO")]
public class OperacaoModel
{
    [Column("ID")]
    public int Id { get; set; }

    [Column("CODTIPOTERMINAL")]
    public int CodigoTipoTerminal { get; set; }

    [Column("CODOPERACAO")]
    public string CodigoOperacao { get; set; }

    [Column("DESCOPERACAO")]
    public string DescricaoOperacao { get; set; }

    [Column("CODHISTORICO")]
    public int CodigoHistorico { get; set; }

    [Column("DESCHISTORICO")]
    public string DescricaoHistorico { get; set; }

    [Column("DATAOPERACAO")]
    public DateTime DataOperacao { get; set; }

    [Column("CODTERMINAL")]
    public string CodigoTerminal { get; set; }

    [Column("PA")]
    public int Pa { get; set; }

    [Column("CODAUT")]
    public string CodigoAut { get; set; }

    [Column("VALOR")]
    public decimal Valor { get; set; }

    [Column("DATAHORACARGA")]
    public DateTime DataHoraCarga { get; set; }
}
using System.ComponentModel.DataAnnotations.Schema;

namespace AceleraPlenoProjetoFinal.Api.Models
{
    [Table("TBTIPOTERMINAL")]
    public class TipoTerminalModel
    {
        [Column("ID")]
        public int Id { get; set; }

        [Column("CODTIPOTERMINAL")]
        public int CodigoTipoTerminal { get; set; }
        
        [Column("DESCTIPOTERMINAL")]
        public string DescricaoTipoTerminal { get; set; }
        
        [Column("BOLACESSOLIBERADO")]
        public int AcessoLiberado { get; set; }
        
        [Column("NUMCHEKALTERACAO")]
        public int NumCheckAlteracao { get; set; }
        
        [Column("PA")]
        public int PA { get; set; }
        
        [Column("LIMSUPERIOR")]
        public int LimiteSuperior { get; set; }
        
        [Column("LIMINFERIOR")]
        public int LimiteInferior { get; set; }

        [Column("CODCRIADOPOR")]
        public int CodigoCriadoPor { get; set; }

        [Column("DATAHORACRIACAO")]
        public DateTime DataHoraCriacao { get; set; }
    }
}
using ConsultaNumerarios.Models;

namespace ConsultaNumerarios.Interfaces
{
    public interface IOperacao
    {
        List<OperacaoModel> GetOperacoesPorPa(int idPa);

    }
}

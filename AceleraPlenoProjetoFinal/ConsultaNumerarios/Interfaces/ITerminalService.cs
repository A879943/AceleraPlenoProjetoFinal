using ConsultaNumerarios.Dto;
using ConsultaNumerarios.Models;

namespace ConsultaNumerarios.Interfaces
{
    public interface ITerminalService
    {
        SaldoDeTerminais CalculaSaldosDeTerminais(int operacaoModels);
    }
}

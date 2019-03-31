namespace ExportadorExcelExemplo.Models
{
    public class Pessoa
    {
        public Pessoa(int id, string nome, string cpf, string telefone)
        {
            Id = id;
            Nome = nome;
            Cpf = cpf;
            Telefone = telefone;
        }

        public int Id { get; set; }
        public string Nome { get; set; }
        public string Cpf { get; set; }
        public string Telefone { get; set; }
    }
}

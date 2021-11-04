using ClosedXML.Excel;
using Npgsql;
using System;

namespace puxarExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            //Chamando as ações
            InsertRecord();
            Console.ReadKey();

        }

        //Desenvolvendo a estrutura de inserção de dados 
        private static void InsertRecord()
        {
            //Realizando conexão
            using (NpgsqlConnection con = GetConnection())
            {
                //Abrindo arquivo xlsx 
                var bd = new XLWorkbook(@"C:\Users\ryanv\Desktop\puxarExcel\baseDeDados.xlsx");

                //manipulando celulas do Excel
                var planilha = bd.Worksheet(1);

                //iniciando variavel do contador
                var i = 0;

                //manipulando linha do Excel
                var linha = 2;

                //Iniciando Looping de importação e exportação de dados.
                while (true)
                {
                    //celula A equivale ao ID
                    var ID = planilha.Cell("A" + linha.ToString()).Value.ToString();

                    //quando o ID acabar o looping acaba
                    if (string.IsNullOrEmpty(ID)) break;

                    //celula B equivale a nome
                    var nome = planilha.Cell("B" + linha.ToString()).Value.ToString();

                    //celula C equivale a nome
                    var email = planilha.Cell("C" + linha.ToString()).Value.ToString();

                    //celula D equivale a nome
                    var cpf = planilha.Cell("D" + linha.ToString()).Value.ToString();

                    //celula E equivale a nome
                    var status = planilha.Cell("E" + linha.ToString()).Value.ToString();

                    //Contador para mostrar quantas linhas foram salvas
                    linha++;

                    //Verificando o status do usuario
                    if(status == "ATIVO")
                    {
                        //Realizando insert into a tabela no Postgres
                        string query = @"insert into public.tb_usuarioteste(Nome, Email, Cpf) values(@nome, @email, @cpf)";

                        NpgsqlCommand cmd = new NpgsqlCommand(query, con);

                        //Realizando conexão dos dados do Excel para a exportação ao Postgres
                        cmd.Parameters.AddWithValue("@nome", nome);
                        cmd.Parameters.AddWithValue("@email", email);
                        cmd.Parameters.AddWithValue("@cpf", cpf);

                        //Abrindo Banco de Dados
                        con.Open();

                        //Realizando execução
                        int n = cmd.ExecuteNonQuery();

                        //contador da linha
                        i++;

                        //se tudo ocorrer certo, avisa qual linha foi salva com sucesso
                        if (n == 1)
                        {
                            Console.WriteLine("A linha " + i + " foi salva");
                        }
                        else //se não, avisa qual linha não foi salva
                        {
                            Console.WriteLine("A linha " + i + " não foi salva");
                        }

                        //Fechando banco de dados
                        con.Close();
                    }
                    else
                    {
                        i++;
                        Console.WriteLine("A linha " + i + " não fol salva pois o usuario está inativo");
                    }
                    

                    
                }
                bd.Dispose();
            }
        }

        //Referencia e conexao com o Postgres
        private static NpgsqlConnection GetConnection()
        {
            return new NpgsqlConnection(@"Server=localhost;Port=5432;User Id=postgres;Password=root;Database=Desafio;");
        }
    }
}


using System.Data.OleDb;
using NUnit.Framework;
using System.Collections.Generic;
using System;
//using Npgsql;

namespace DDT
{
    public class TestDataReader
    {
        public TestDataReader() : base() { }

        public List<TestCaseData> LerDadosPlanilha(string caminho, string aba)
        {
            List<TestCaseData> listaDatas = new List<TestCaseData>();

            string sufixo = "";

            int e = caminho.LastIndexOf('.');
            if (e > 0)
            {
                sufixo = caminho.Substring(e + 1);
            }

            switch (sufixo)
            {
                case "xls":
                    listaDatas = LerPlanilhaExcel(caminho, aba);
                    break;
                case "accdb":
                    listaDatas = LerBancoAccess(caminho, aba);
                    break;
                default: throw new Exception("O arquivo informado não contém suporte para esta conexão. Tipo de arquivo: " + sufixo);
            }
            return listaDatas;
        }

        public List<TestCaseData> LerPlanilhaExcel(string caminho, string aba)
        {
            try
            {
                string connectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                "Data Source=" + caminho + ";Extended Properties=Excel 8.0;";

                List<TestCaseData> testDataList = new List<TestCaseData>();
                using (OleDbConnection connection = new OleDbConnection(connectionStr))
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand("SELECT * FROM [" + aba + "$]", connection);
                    OleDbDataReader reader = command.ExecuteReader();

                    int contagemDeColunas = reader.FieldCount;
                    contagemDeColunas = contagemDeColunas - 2;

                    while (reader.Read())
                    {
                        string[] args = new string[contagemDeColunas];

                        for (int i = 0; i < contagemDeColunas; i++)
                        {
                            if (i == contagemDeColunas)
                                args[i] = reader.GetValue(contagemDeColunas).ToString();
                            else
                            {
                                args[i] = reader.GetValue(i + 2).ToString();
                            }
                        }

                        TestCaseData testData = new TestCaseData(args);

                        if (reader.GetName(0) != "Caso de teste")
                            throw new FormatException("O campo Caso de teste é obrigatório como primeira coluna.\nAba: " + aba);

                        if (reader.GetName(1) != "Descricao")
                            throw new FormatException("O campo Descricao é obrigatório como segunda coluna.\nAba: " + aba);


                        testData.SetName(reader.GetString(0));
                        testData.SetDescription(reader.GetString(1));

                        testDataList.Add(testData);
                    }
                }
                return testDataList;
            }
            catch (FormatException e)
            {
                throw new Exception(e.Message);
            }
            catch(Exception e)
            {
                throw new Exception("Uma ou mais colunas obrigatórias podem estar ausentes na base de dados. Mensagem original: " + e.Message);
            }
        }

        public List<TestCaseData> LerBancoAccess(string caminho, string aba)
        {
            try
            {
                string connectionStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + caminho;

                List<TestCaseData> testDataList = new List<TestCaseData>();
                using (OleDbConnection connection = new OleDbConnection(connectionStr))
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand("SELECT * FROM [" + aba + "]", connection);
                    OleDbDataReader reader = command.ExecuteReader();

                    int contagemDeColunas = reader.FieldCount;
                    contagemDeColunas = contagemDeColunas - 3;

                    while (reader.Read())
                    {
                        string[] args = new string[contagemDeColunas];

                        for (int i = 0; i < contagemDeColunas; i++)
                        {
                            if (i == contagemDeColunas)
                                args[i] = reader.GetValue(contagemDeColunas).ToString();
                            else
                            {
                                args[i] = reader.GetValue(i + 3).ToString();
                            }
                        }

                        TestCaseData testData = new TestCaseData(args);

                        if (reader.GetName(1) != "Caso de teste")
                            throw new FormatException("O campo Caso de teste é obrigatório como segunda coluna.\nTabela: " + aba);

                        if (reader.GetName(2) != "Descricao")
                            throw new FormatException("O campo Descricao é obrigatório como terceira coluna.\nTabela: " + aba);

                        testData.SetName(reader.GetString(1));
                        testData.SetDescription(reader.GetString(2));

                        testDataList.Add(testData);
                    }
                }
                return testDataList;
            }
            catch (FormatException e)
            {
                throw new Exception(e.Message);
            }
            catch (Exception e)
            {
                throw new Exception("Uma ou mais colunas obrigatórias podem estar ausentes na base de dados. Mensagem original: " + e.Message);
            }
        }
    }
}

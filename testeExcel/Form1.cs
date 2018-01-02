using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.Common;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.Sql;
using System.Runtime.InteropServices;
using System.Configuration;
using Microsoft.Office.Interop.Excel;
using ExcelIt = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Drawing;

namespace testeExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static string path;
        public static string excelConnectionString;
        public string[] files;
        public string conexao;
        public string baseDeDados;
        public string tabela;
        public string caminho;
        public string directoryPath;
        private static Excel.Application MyApp = null;
        public List<string> filesAdionado = new List<string>();
        public List<string> colunas = new List<string>();
        public List<string> colunasCreate = new List<string>();
        public string tipoArquivo;

        private void button1_Click(object sender, EventArgs e)
        {

            MyApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;

            MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\ClientesOriginaisHwashin.xlsx");
            Workbook wb = MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\ClientesOriginaisHwashin.xlsx");
            Worksheet ws = wb.Sheets[1];
            MyApp.DisplayAlerts = false;
            ws.Range["A:A"].NumberFormat = "@";
            wb.SaveAs("c:\\a\\formatado\\clientesformatadoHwashin.xlsx");
            wb.Close();
            MyApp.Quit();
            SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");
            string sqlConnectionString = "Data Source=BRCAENRODRIGUES\\SQLEXPRESS;Initial Catalog=my_database;Integrated Security=True";

            SqlCommand cmdColuna = conn.CreateCommand();

            cmdColuna.CommandText =
              @"IF OBJECT_ID('dbo.clientes', 'U') IS NOT NULL 
                  DROP TABLE dbo.clientes; 
                    CREATE TABLE [dbo].[Clientes](
	                    [Cli_ID] [varchar](70) NULL,
	                    [Cli_Nome] [varchar](255) NULL,
	                    [Cli_Pss_ID] [int] NULL,
	                    [Cli_Vinc] [varchar](1) NULL,
	                    [Cli_Vinc_DT_Ini] [datetime] NULL,
	                    [Cli_Vinc_DT_Fim] [datetime] NULL,
	                    [Cli_CNPJ] [varchar](14) NULL,
	                    [Cli_Vinc_Justific] [varchar](2) NULL,
	                    [Cli_Paraiso_Fiscal] [varchar](1) NULL CONSTRAINT [DF_Clientes_Cli_Paraiso_Fiscal]  DEFAULT ('N'),
	                    [Arq_Origem_ID] [int] NULL,
	                    [Lin_Origem_ID] [int] NULL,
	                    [ID] [int] IDENTITY(1,1) NOT NULL
                    ) ON [PRIMARY]";

            SqlTransaction trA = null;

            conn.Open();
            trA = conn.BeginTransaction();
            cmdColuna.Transaction = trA;
            cmdColuna.ExecuteNonQuery();
            trA.Commit();
            conn.Close();

            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\a\\formatado\\clientesformatadoHwashin.xlsx; Extended Properties=Excel 12.0;";

            using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            {
                OleDbCommand cmd = new OleDbCommand("Select [Código do Cliente], Nome,  [Código do País], Vínculo, [Data Vínculo Inicial], [Data Vínculo Final], CNPJ from [Clientes$]", connection);

                connection.Open();
                OleDbDataReader dReader = cmd.ExecuteReader();

                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(sqlConnectionString))
                {
                    sqlBulk.DestinationTableName = "Clientes";
                    sqlBulk.WriteToServer(dReader);
                }

                SqlCommand cmdCopPedido = conn.CreateCommand();

                cmdCopPedido.CommandText =
                        @"INSERT INTO D_CLIENTES (CLI_ID, CLI_NOME, CLI_VINC, CLI_PSS_ID, [Cli_Vinc_DT_Ini], [Cli_Vinc_DT_Fim], [Cli_CNPJ], Lin_Origem_id)
                        SELECT CLI_ID, max(CLI_NOME), CLI_VINC, CLI_PSS_ID, [Cli_Vinc_DT_Ini], [Cli_Vinc_DT_Fim], max([Cli_CNPJ]), max(id)
                        FROM clientes
                        where cli_id is not null and cli_vinc is not null
                        GROUP BY CLI_ID, CLI_VINC, CLI_PSS_ID, [Cli_Vinc_DT_Ini], [Cli_Vinc_DT_Fim]";

                // select*
                //from Clientes a left join d_clientes b on a.id = b.lin_origem_id
                // where b.Cli_ID is null

                SqlTransaction tr = null;
                try
                {
                    conn.Open();
                    tr = conn.BeginTransaction();
                    cmdCopPedido.Transaction = tr;
                    cmdCopPedido.ExecuteNonQuery();
                    tr.Commit();

                    label1.Text = "Tabela clientes copiada ";
                }
                catch (Exception ex)
                {
                    tr.Rollback();
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    conn.Close();
                }

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            MyApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\a\\formatado\\fornecedoresFormatadoHwashin.xlsx; Extended Properties=Excel 12.0;";

            MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\FornecedoresOriginalHwashin.xlsx");
            Workbook wb = MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\FornecedoresOriginalHwashin.xlsx");
            Worksheet ws = wb.Sheets[1];
            MyApp.DisplayAlerts = false;
            ws.Range["A:A"].NumberFormat = "@";
            wb.SaveAs("c:\\a\\formatado\\fornecedoresFormatadoHwashin.xlsx");
            wb.Close();
            MyApp.Quit();
            SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");
            string sqlConnectionString = "Data Source=BRCAENRODRIGUES\\SQLEXPRESS;Initial Catalog=my_database;Integrated Security=True";

            SqlCommand cmdColuna = conn.CreateCommand();

            cmdColuna.CommandText =
              @"IF OBJECT_ID('dbo.fornecedores', 'U') IS NOT NULL 
                  DROP TABLE dbo.fornecedores; 
                    CREATE TABLE [dbo].[Fornecedores](
	                [For_ID] [varchar](70) NULL,
	                [For_Nome] [varchar](255) NULL,
	                [For_PSS_ID] [int] NULL,
	                [For_Vinc] [varchar](1) NULL,
	                [For_Vinc_DT_Ini] [datetime] NULL,
	                [For_Vinc_DT_Fim] [datetime] NULL,
	                [For_CNPJ] [varchar](14) NULL,
	                [For_Vinc_Just] [varchar](2) NULL CONSTRAINT [DF_D_Fornecedores_For_Vinc_Just]  DEFAULT ('0'),
	                [For_Paraiso_Fiscal] [varchar](1) NULL CONSTRAINT [DF_D_Fornecedores_For_Paraiso_Fiscal]  DEFAULT ('N'),
	                [Arq_Origem_ID] [int] NULL,
	                [Lin_Origem_ID] [int] NULL,
	                [ID] [int] IDENTITY(1,1) NOT NULL) ON [PRIMARY]";

            SqlTransaction trA = null;

            conn.Open();
            trA = conn.BeginTransaction();
            cmdColuna.Transaction = trA;
            cmdColuna.ExecuteNonQuery();
            trA.Commit();
            conn.Close();

            using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            {
                OleDbCommand cmd = new OleDbCommand("Select [Código do Fornecedor], Nome,  [Código do País], Vínculo, [Data Vínculo Inicial], [Data Vínculo Final], CNPJ from [Fornecedores$]", connection);

                connection.Open();
                OleDbDataReader dReader = cmd.ExecuteReader();

                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(sqlConnectionString))
                {
                    sqlBulk.DestinationTableName = "Fornecedores";
                    sqlBulk.WriteToServer(dReader);
                }

                SqlCommand cmdCopPedido = conn.CreateCommand();

                cmdCopPedido.CommandText =
                        @"INSERT INTO D_FORNECEDORES ([For_ID], [For_Nome], [For_PSS_ID], [For_Vinc], [For_Vinc_DT_Ini], [For_Vinc_DT_Fim], [For_CNPJ], Lin_Origem_id)
                        SELECT For_ID, max(For_Nome),For_PSS_ID, [For_Vinc], [For_Vinc_DT_Ini], [For_Vinc_DT_Fim], max([For_CNPJ]), max(id)
                        FROM fornecedores
                        WHERE FOR_ID IS NOT NULL AND FOR_VINC IS NOT NULL
						GROUP BY FOR_ID, FOR_VINC, FOR_PSS_ID, [FOR_Vinc_DT_Ini], [FOR_Vinc_DT_Fim]";

                // select*
                //from Clientes a left join d_clientes b on a.id = b.lin_origem_id
                // where b.Cli_ID is null

                SqlTransaction tr = null;
                try
                {
                    conn.Open();
                    tr = conn.BeginTransaction();
                    cmdCopPedido.Transaction = tr;
                    cmdCopPedido.ExecuteNonQuery();
                    tr.Commit();

                    label1.Text = "Tabela fornecedores copiada ";
                }
                catch (Exception ex)
                {
                    // se chegou aqui é porque deu erro
                    tr.Rollback();
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    conn.Close();
                }

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            MyApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;

            MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\produtosOriginalHwashin.xlsx");
            Workbook wb = MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\produtosOriginalHwashin.xlsx");
            Worksheet ws = wb.Sheets[1];
            MyApp.DisplayAlerts = false;
            ws.Range["A:A"].NumberFormat = "@";
            wb.SaveAs("c:\\a\\formatado\\produtosformatadoHwashin.xlsx");
            wb.Close();
            MyApp.Quit();
            SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");
            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\a\\formatado\\produtosformatadoHwashin.xlsx; Extended Properties=Excel 12.0;";
            SqlCommand cmdColuna = conn.CreateCommand();

            cmdColuna.CommandText =
              @"IF OBJECT_ID('dbo.produtos', 'U') IS NOT NULL 
                  DROP TABLE dbo.produtos; 
                   CREATE TABLE [dbo].[Produtos](
                 [Pro_ID] [varchar](70) NULL,
                 [Pro_Descricao] [varchar](255) NULL,
                 [Pro_Und_ID] [int] NULL,
                 [Pro_NCM] [varchar](max) NULL,
                 [Pro_Margem] [int] NULL,
                 [Lin_Origem_ID] [int] NULL,
                 [Arq_Origem_ID] [int] NULL,
                 [ID] [int] IDENTITY(1,1) NOT NULL) ON [PRIMARY]";

            SqlTransaction trA = null;

            conn.Open();
            trA = conn.BeginTransaction();
            cmdColuna.Transaction = trA;
            cmdColuna.ExecuteNonQuery();
            trA.Commit();
            conn.Close();

            using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            {
                OleDbCommand cmd = new OleDbCommand("Select [Código do Produto], [Descrição], [Unidade de Medida],  [ Classificação Fiscal (NCM)] from [Produtos$]", connection);

                connection.Open();
                OleDbDataReader dReader = cmd.ExecuteReader();
                conn.Open();
                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(conn))
                {
                    sqlBulk.DestinationTableName = "Produtos";
                    sqlBulk.WriteToServer(dReader);
                }
                connection.Close();
                conn.Close();
            }

            SqlCommand cmdDropProcedure = conn.CreateCommand();
            SqlCommand cmdCreateProcedure = conn.CreateCommand();
            SqlCommand cmdExecProcedure = conn.CreateCommand();


            cmdDropProcedure.CommandText =
                    @"IF EXISTS ( SELECT * 
                                FROM   sysobjects 
                                WHERE  id = object_id(N'[dbo].[SP_IncluiProdutos]') 
                                       and OBJECTPROPERTY(id, N'IsProcedure') = 1 )

                        DROP PROCEDURE [dbo].[SP_IncluiProdutos]";
            cmdCreateProcedure.CommandText = @"CREATE PROCEDURE SP_IncluiProdutos
                                AS
                                BEGIN
	                                SET NOCOUNT ON;
	                                INSERT INTO D_PRODUTOS ([Pro_ID], [Pro_Descricao], [Pro_Und_ID] , [Pro_NCM], [Pro_Margem], Lin_Origem_id)
                                                        SELECT [Pro_ID], MIN([Pro_Descricao]), MIN([Pro_Und_ID]) , MIN(SUBSTRING(Pro_NCM,0,8)), [Pro_Margem], MIN(id) 
                                                        FROM Produtos
                                                        WHERE PRO_ID IS NOT NULL  
						                                GROUP BY [Pro_ID],  [Pro_Margem]
                                END
                                ";
            cmdExecProcedure.CommandText = @"exec sp_incluiprodutos";
            // select*
            // from Clientes a left join d_clientes b on a.id = b.lin_origem_id
            // where b.Cli_ID is null
            cmdExecProcedure.CommandTimeout = 0;
            SqlTransaction tr = null;
            try
            {
                conn.Open();
                tr = conn.BeginTransaction();
                cmdDropProcedure.Transaction = tr;
                cmdDropProcedure.ExecuteNonQuery();
                tr.Commit();
                tr = conn.BeginTransaction();
                cmdCreateProcedure.Transaction = tr;
                cmdCreateProcedure.ExecuteNonQuery();
                tr.Commit();
                tr = conn.BeginTransaction();
                cmdExecProcedure.Transaction = tr;
                MessageBox.Show("3");
                cmdExecProcedure.ExecuteNonQuery();
                MessageBox.Show("4");
                tr.Commit();

                label1.Text = "Tabela produtos copiada ";
            }
            catch (Exception ex)
            {
                // se chegou aqui é porque deu erro
                tr.Rollback();
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }


        }

        private void button4_Click(object sender, EventArgs e)
        {

            MyApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;

            MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\SaldosInicialOriginalHwashin.xlsx");
            Workbook wb = MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\SaldosInicialOriginalHwashin.xlsx");
            Worksheet ws = wb.Sheets[1];
            MyApp.DisplayAlerts = false;
            ws.Range["A:A"].NumberFormat = "@";
            ws.Range["B:B"].Replace(".", "/");
            wb.SaveAs("c:\\a\\formatado\\SaldosInicialformatadoHwashin.xlsx");
            wb.Close();
            MyApp.Quit();
            SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");
            string sqlConnectionString = "Data Source=BRCAENRODRIGUES\\SQLEXPRESS;Initial Catalog=my_database;Integrated Security=True";

            SqlCommand cmdColuna = conn.CreateCommand();

            cmdColuna.CommandText =
              @"IF OBJECT_ID('dbo.Inventario_Carga_Inicial', 'U') IS NOT NULL 
                  DROP TABLE dbo.Inventario_Carga_Inicial; 
                    CREATE TABLE [dbo].[Inventario_Carga_Inicial](
	                [Inv_Pro_ID] [varchar](70) NULL,
	                [Inv_Data] [datetime] NULL,
	                [Inv_CNPJ] [varchar](20) NULL DEFAULT ('00000000000000'),
	                [Inv_Qtde] [numeric](24, 12) NULL ,
	                [Inv_Valor] [numeric](24, 12) NULL ,
	                [Inv_Tipo] [varchar](1) NULL ,
	                [Inv_Arq_Origem] [int] NULL,
	                [Inv_Registro_Origem] [varchar](1000) NULL,
	                [Inv_Und_Id] [int] NULL,
	                [Inv_Div_Id] [varchar](70) NULL,
	                [Inv_Local_Negocio] [varchar](70) NULL,
	                [Arq_Origem_ID] [int] NULL,
	                [Lin_Origem_ID] [int] NULL,
	                [ID] [int] IDENTITY(1,1) NOT NULL)";

            SqlTransaction trA = null;

            conn.Open();
            trA = conn.BeginTransaction();
            cmdColuna.Transaction = trA;
            cmdColuna.ExecuteNonQuery();
            trA.Commit();
            conn.Close();

            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\a\\formatado\\saldoInicialformatadoHwashin.xlsx; Extended Properties=Excel 12.0;";

            using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            {
                OleDbCommand cmd = new OleDbCommand("Select [Código do Produto], [Data Inventário], [CNPJ], [Quantidade em estoque], [Valor em Reais]  from [Saldos Iniciais$]", connection);

                connection.Open();
                OleDbDataReader dReader = cmd.ExecuteReader();

                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(sqlConnectionString))
                {
                    sqlBulk.DestinationTableName = "Inventario_Carga_Inicial";
                    sqlBulk.WriteToServer(dReader);
                }

                SqlCommand cmdCopPedido = conn.CreateCommand();

                cmdCopPedido.CommandText =
                        @"INSERT INTO D_INVENTARIO_CARGA 
						(Inv_Pro_ID, Inv_Data, Inv_CNPJ, Inv_Qtde, Inv_Valor, Inv_Tipo, Inv_Und_Id, Lin_Origem_id)
                          SELECT Inv_Pro_ID, max(Inv_Data), max(Inv_CNPJ), max(Inv_Qtde), max(Inv_Valor), Inv_Tipo, Inv_Und_Id, max(id)
                        FROM Inventario_Carga_Inicial
                        GROUP BY  Inv_Pro_ID, Inv_Und_Id, Inv_Tipo";

                // select*
                //from Clientes a left join d_clientes b on a.id = b.lin_origem_id
                // where b.Cli_ID is null

                SqlTransaction tr = null;
                try
                {
                    conn.Open();
                    tr = conn.BeginTransaction();
                    cmdCopPedido.Transaction = tr;
                    cmdCopPedido.ExecuteNonQuery();
                    tr.Commit();

                    label1.Text = "Tabela saldos iniciais copiada ";
                }
                catch (Exception ex)
                {
                    tr.Rollback();
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    conn.Close();
                }

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            MyApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;

            MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\SaldosFinalOriginalHwashin.xlsx");
            Workbook wb = MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\SaldosFinalOriginalHwashin.xlsx");
            Worksheet ws = wb.Sheets[1];
            MyApp.DisplayAlerts = false;
            ws.Range["A:A"].NumberFormat = "@";
            ws.Range["B:B"].Replace(".", "/");
            wb.SaveAs("c:\\a\\formatado\\SaldosFinalformatadoHwashin.xlsx");
            wb.Close();
            MyApp.Quit();
            SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");
            string sqlConnectionString = "Data Source=BRCAENRODRIGUES\\SQLEXPRESS;Initial Catalog=my_database;Integrated Security=True";

            SqlCommand cmdColuna = conn.CreateCommand();

            cmdColuna.CommandText =
              @"IF OBJECT_ID('dbo.Inventario_Carga_Final', 'U') IS NOT NULL 
                  DROP TABLE dbo.Inventario_Carga_Final; 
                    CREATE TABLE [dbo].[Inventario_Carga_Final](
	                [Inv_Pro_ID] [varchar](70) NULL,
	                [Inv_Data] [datetime] NULL,
	                [Inv_CNPJ] [varchar](20) NULL DEFAULT ('00000000000000'),
	                [Inv_Qtde] [numeric](24, 12) NULL ,
	                [Inv_Valor] [numeric](24, 12) NULL ,
	                [Inv_Tipo] [varchar](1) NULL ,
	                [Inv_Arq_Origem] [int] NULL,
	                [Inv_Registro_Origem] [varchar](1000) NULL,
	                [Inv_Und_Id] [int] NULL,
	                [Inv_Div_Id] [varchar](70) NULL,
	                [Inv_Local_Negocio] [varchar](70) NULL,
	                [Arq_Origem_ID] [int] NULL,
	                [Lin_Origem_ID] [int] NULL,
	                [ID] [int] IDENTITY(1,1) NOT NULL)";

            SqlTransaction trA = null;

            conn.Open();
            trA = conn.BeginTransaction();
            cmdColuna.Transaction = trA;
            cmdColuna.ExecuteNonQuery();
            trA.Commit();
            conn.Close();

            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\a\\formatado\\SaldosFinalformatadoHwashin.xlsx; Extended Properties=Excel 12.0;";

            using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            {
                OleDbCommand cmd = new OleDbCommand("Select [Código do Produto], [Data Inventário], [CNPJ], [Quantidade em estoque], [Valor em Reais]  from [Saldos Finais$]", connection);

                connection.Open();
                OleDbDataReader dReader = cmd.ExecuteReader();

                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(sqlConnectionString))
                {
                    sqlBulk.DestinationTableName = "Inventario_Carga_Final";
                    sqlBulk.WriteToServer(dReader);
                }

                SqlCommand cmdCopPedido = conn.CreateCommand();

                cmdCopPedido.CommandText =
                        @"INSERT INTO D_INVENTARIO_CARGA 
						(Inv_Pro_ID, Inv_Data, Inv_CNPJ, Inv_Qtde, Inv_Valor, Inv_Tipo, Inv_Und_Id, Lin_Origem_id)
                          SELECT Inv_Pro_ID, max(Inv_Data), max(Inv_CNPJ), max(Inv_Qtde), max(Inv_Valor), Inv_Tipo, Inv_Und_Id, max(id)
                        FROM Inventario_Carga_Final
                        GROUP BY  Inv_Pro_ID, Inv_Und_Id, Inv_Tipo";

                // select*
                //from Clientes a left join d_clientes b on a.id = b.lin_origem_id
                //where b.Cli_ID is null

                SqlTransaction tr = null;
                try
                {
                    conn.Open();
                    tr = conn.BeginTransaction();
                    cmdCopPedido.Transaction = tr;
                    cmdCopPedido.ExecuteNonQuery();
                    tr.Commit();

                    label1.Text = "Tabela saldos finais copiada ";
                }
                catch (Exception ex)
                {
                    tr.Rollback();
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    conn.Close();
                }

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

            MyApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;

            MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\ComprasOriginalHwashin.xlsx");
            Workbook wb = MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\ComprasOriginalHwashin.xlsx");
            Worksheet ws = wb.Sheets[1];
            MyApp.DisplayAlerts = false;
            ws.Range["A:A"].NumberFormat = "@";
            wb.SaveAs("c:\\a\\formatado\\ComprasformatadoHwashin.xlsx");
            wb.Close();
            MyApp.Quit();
            SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");
            string sqlConnectionString = "Data Source=BRCAENRODRIGUES\\SQLEXPRESS;Initial Catalog=my_database;Integrated Security=True";

            SqlCommand cmdColuna = conn.CreateCommand();

            cmdColuna.CommandText =
              @"IF OBJECT_ID('dbo.compras', 'U') IS NOT NULL 
                  DROP TABLE dbo.compras; 
                    CREATE TABLE [dbo].[compras](
	                   [Cmp_Pro_ID] [varchar](70) NULL,
                        [Cmp_Cod_Divisao] [varchar](50) NULL,
                        [Cmp_For_ID] [varchar](70) NULL,
                        [Cmp_Lanc_Cont] [varchar](70) NULL,
                        [Cmp_Fat_Coml] [varchar](70) NULL,
                        [Cmp_BL_DT] [datetime] NULL,
                        [Cmp_DI_ID] [varchar](70) NULL,
                        [Cmp_DI_DT_Emissao] [datetime] NULL,
                        [Cmp_NF_Entrada] [varchar](70) NULL,
                        [Cmp_NF_Serie] [varchar](70) NULL CONSTRAINT [DF_Compras_Cmp_NF_Serie]  DEFAULT (''),
                        [Cmp_NF_DT] [datetime] NULL,
                        [Cmp_CFOP] [int] NULL,
                        [Cmp_DI_DT_Vencimento] [datetime] NULL,
                        [Cmp_DI_Dias] [int] NULL,
                        [Cmp_Qtde] [numeric](24, 12) NULL CONSTRAINT [DF_Compras_Cmp_Qtde]  DEFAULT ((0)),
                        [Cmp_Valor_Fob] [numeric](24, 12) NULL CONSTRAINT [DF_Compras_Cmp_Valor_Fob]  DEFAULT ((0)),
                        [Cmp_Cod_Moeda] [int] NULL,
                        [Cmp_Vl_Frete_Moeda] [numeric](24, 12) NULL CONSTRAINT [DF_Compras_Cmp_VL_Frete_Reais]  DEFAULT ((0)),
                        [Cmp_VL_Seguro_Moeda] [numeric](24, 12) NULL CONSTRAINT [DF_Compras_Cmp_VL_Seguro_Moeda]  DEFAULT ((0)),
                        [Cmp_Cod_Moeda_Seguro] [int] NULL,
                        [Cmp_Cod_Moeda_Frete] [int] NULL,
                        [Cmp_Imposto_Import] [numeric](24, 12) NULL CONSTRAINT [DF_Compras_Cmp_Imposto_Import]  DEFAULT ((0)),
                        [Cmp_ICMS] [numeric](24, 12) NULL CONSTRAINT [DF_Compras_Cmp_ICMS]  DEFAULT ((0)),
                        [Cmp_PIS] [numeric](24, 12)  NULL CONSTRAINT [DF_Compras_Cmp_PIS]  DEFAULT ((0)),
                        [Cmp_COFINS] [numeric](24, 12) NULL CONSTRAINT [DF_Compras_Cmp_COFINS]  DEFAULT ((0)),
                        [Cmp_Und_Id] [int] NULL,
                        [Cmp_CNPJ] [varchar](20) NULL CONSTRAINT [DF_Compras_Cmp_CNPJ]  DEFAULT (''),
                        [Cmp_Incoterm] [varchar](70) NULL,
                        [Cmp_For_id_Seguro] [varchar](70) NULL,
                        [Cmp_For_id_Frete] [varchar](70) NULL,
                        [Arq_Origem_ID] [varchar](70) NULL,
                        [ID] [int] IDENTITY(1,1) NOT NULL)";

            SqlTransaction trA = null;

            conn.Open();
            trA = conn.BeginTransaction();
            cmdColuna.Transaction = trA;
            cmdColuna.ExecuteNonQuery();
            trA.Commit();
            conn.Close();

            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\a\\formatado\\ComprasformatadoHwashin.xlsx; Extended Properties=Excel 12.0;";

            using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            {
                OleDbCommand cmd = new OleDbCommand("Select [Código do Produto], [Código da Divisão] ,[Código do Fornecedor], [N˚ lançamento contábil]" +
                    ",[N˚ fatura comercial],[Data da BL],[Numero da DI]" +
                    ",[Data da Importação],[Número da NF de Entrada],[Série Nota fiscal de Entrada],[Data Entrada no Estoque]" +
                    ",[CFOP Nota Fiscal], [Data Vencimento Média], [Prazo de Vencimento Médio], [Quantidade]" +
                    ",[Valor \"FOB\" (moeda estrangeira)],[Código moeda estrangeira],[Frete],[Seguro],[Código moeda frete]" +
                    ",[código moeda seguro],[Imposto de Importação],[ICMS (reais)],[PIS (reais)],[COFINS (reais)]" +
                    ",[Unidade de Medida],[CNPJ],[Incoterm],[Código do fornecedor de Seguro],[Código do Fornecedor de Frete]" +
                    "  from [Compras$]", connection);

                connection.Open();
                OleDbDataReader dReader = cmd.ExecuteReader();

                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(sqlConnectionString))
                {
                    sqlBulk.DestinationTableName = "Compras";
                    sqlBulk.WriteToServer(dReader);
                }

                SqlCommand cmdCopPedido = conn.CreateCommand();

                cmdCopPedido.CommandText =
                        @"INSERT INTO D_Compras
						(Cmp_Pro_ID,
                            Cmp_Cod_Divisao,
                            Cmp_For_ID,
                            Cmp_Lanc_Cont,
                            Cmp_Fat_Coml,
                            Cmp_BL_DT,
                            Cmp_DI_ID,
                            Cmp_DI_DT_Emissao,
                            Cmp_NF_Entrada,
                            Cmp_NF_Serie,
                            Cmp_NF_DT,
                            Cmp_CFOP,
                            Cmp_DI_DT_Vencimento,
                            Cmp_DI_Dias,
                            Cmp_Qtde,
                            Cmp_Valor_Fob,
                            Cmp_Cod_Moeda,
                            Cmp_Vl_Frete_Moeda,
                            Cmp_VL_Seguro_Moeda,
                            Cmp_Cod_Moeda_Seguro,
                            Cmp_Cod_Moeda_Frete,
                            Cmp_Imposto_Import,
                            Cmp_ICMS,
                            Cmp_PIS,
                            Cmp_COFINS,
                            Cmp_Und_Id,
                            Cmp_CNPJ,
                            Cmp_Incoterm,
                            Cmp_For_id_Seguro,
                            Cmp_For_id_Frete,
                            [Arq_Origem_ID],
                            Lin_Origem_ID)
                          SELECT Cmp_Pro_ID,
                            Cmp_Cod_Divisao,
                            Cmp_For_ID,
                            Cmp_Lanc_Cont,
                            Cmp_Fat_Coml,
                            Cmp_BL_DT,
                            Cmp_DI_ID,
                            Cmp_DI_DT_Emissao,
                            Cmp_NF_Entrada,
                            Cmp_NF_Serie,
                            Cmp_NF_DT,
                            Cmp_CFOP,
                            Cmp_DI_DT_Vencimento,
                            Cmp_DI_Dias,
                            Cmp_Qtde,
                            Cmp_Valor_Fob,
                            Cmp_Cod_Moeda,
                            Cmp_Vl_Frete_Moeda,
                            Cmp_VL_Seguro_Moeda,
                            Cmp_Cod_Moeda_Seguro,
                            Cmp_Cod_Moeda_Frete,
                            Cmp_Imposto_Import,
                            Cmp_ICMS,
                            Cmp_PIS,
                            Cmp_COFINS,
                            Cmp_Und_Id,
                            Cmp_CNPJ,
                            Cmp_Incoterm,
                            Cmp_For_id_Seguro,
                            Cmp_For_id_Frete,
                            [Arq_Origem_ID],
                            [ID]
                        FROM Compras";

                // select*
                //from Clientes a left join d_clientes b on a.id = b.lin_origem_id
                //where b.Cli_ID is null

                SqlTransaction tr = null;
                try
                {
                    conn.Open();
                    tr = conn.BeginTransaction();
                    cmdCopPedido.Transaction = tr;
                    cmdCopPedido.ExecuteNonQuery();
                    tr.Commit();

                    label1.Text = "Tabela compras copiada ";
                }
                catch (Exception ex)
                {
                    tr.Rollback();
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    conn.Close();
                }

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {

            MyApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;

            MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\VendasOriginalHwashin.xlsx");
            Workbook wb = MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\VendasOriginalHwashin.xlsx");
            Worksheet ws = wb.Sheets[1];

            ws.Range["A:A"].NumberFormat = "@";
            ws.Range["J:J"].NumberFormat = "@";
            ws.Range["A1:AZ100000"].Copy();
            ws.Range["A1:AZ100000"].PasteSpecial(XlPasteType.xlPasteValues);
            MyApp.DisplayAlerts = false;
            wb.SaveAs("c:\\a\\formatado\\VendasFormatadoHwashin.xlsx");
            wb.Close();
            MyApp.Quit();
            SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");
            string sqlConnectionString = "Data Source=BRCAENRODRIGUES\\SQLEXPRESS;Initial Catalog=my_database;Integrated Security=True";

            SqlCommand cmdColuna = conn.CreateCommand();

            cmdColuna.CommandText =
              @"IF OBJECT_ID('dbo.Vendas_Itens', 'U') IS NOT NULL 
                  DROP TABLE dbo.Vendas_Itens; 
                    CREATE TABLE [dbo].[Vendas_Itens](
	                    [Vnd_Cli_ID] [varchar](70) NULL,
	                    [Vnd_NF_ID] [varchar](70) NULL,
	                    [Vnd_NF_Serie] [varchar](70) NULL CONSTRAINT [DF_Vendas_Itens_Vnd_NF_Serie]  DEFAULT (''),
	                    [Vnd_Cod_Divisao] [varchar](70) NULL,
	                    [Vnd_CFOP] [int] NULL,
	                    [Vnd_Dt_Emissao] [datetime] NULL,
	                    [Vnd_DT_Vencimento] [datetime] NULL,
	                    [Vnd_Dias] [int]  NULL CONSTRAINT [DF_Vendas_Itens_Vnd_Dias]  DEFAULT ((0)),
	                    [Vnd_Item] [varchar](70) NULL CONSTRAINT [DF_Vendas_Itens_Vnd_Item]  DEFAULT ((0)),
	                    [Vnd_Pro_id] [varchar](70) NULL,
	                    [Vnd_Qtde] [numeric](24, 12) NULL CONSTRAINT [DF_Vendas_Itens_Vnd_Qtde]  DEFAULT ((0)),
	                    [Vnd_Vl_Nota] [numeric](24, 12) NULL CONSTRAINT [DF_Vendas_Itens_Vnd_Vl_Nota]  DEFAULT ((0)),
	                    [Vnd_Desconto] [numeric](24, 12) NULL CONSTRAINT [DF_Vendas_Itens_Vnd_Desconto]  DEFAULT ((0)),
	                    [Vnd_ICMS] [numeric](24, 12)  NULL CONSTRAINT [DF_Vendas_Itens_Vnd_ICMS]  DEFAULT ((0)),
	                    [Vnd_PIS] [numeric](24, 12)  NULL CONSTRAINT [DF_Vendas_Itens_Vnd_PIS]  DEFAULT ((0)),
	                    [Vnd_COFINS] [numeric](24, 12)  NULL CONSTRAINT [DF_Vendas_Itens_Vnd_COFINS]  DEFAULT ((0)),
	                    [Vnd_ISS] [numeric](24, 12)  NULL CONSTRAINT [DF_Vendas_Itens_Vnd_ISS]  DEFAULT ((0)),
	                    [Vnd_Comissao] [numeric](24, 12)  NULL CONSTRAINT [DF_Vendas_Itens_Vnd_Comissao]  DEFAULT ((0)),
	                    [Vnd_Frete] [numeric](24, 12)  NULL CONSTRAINT [DF_Vendas_Itens_Vnd_Frete]  DEFAULT ((0)),
	                    [Vnd_Seguro] [numeric](24, 12)  NULL CONSTRAINT [DF_Vendas_Itens_Vnd_Seguro]  DEFAULT ((0)),
	                    [Vnd_Dt_Embarque] [datetime] NULL,
	                    [Vnd_Cod_Moeda] [int] NULL,
	                    [Vnd_Vl_Moeda] [numeric](24, 12)  NULL CONSTRAINT [DF_Vendas_Itens_Vnd_Vl_Moeda]  DEFAULT ((0)),
	                    [Vnd_Custo] [numeric](24, 12)  NULL CONSTRAINT [DF_Vendas_Itens_Vnd_Custo]  DEFAULT ((0)),
	                    [Vnd_RE] [varchar](50) NULL,
	                    [Vnd_CNPJ] [varchar](20) NULL CONSTRAINT [DF_Vendas_Itens_Vnd_CNPJ]  DEFAULT ('00000000000000'),
                        [ID] [int] IDENTITY(1,1) NOT NULL)";

            SqlTransaction trA = null;

            conn.Open();
            trA = conn.BeginTransaction();
            cmdColuna.Transaction = trA;
            cmdColuna.ExecuteNonQuery();
            trA.Commit();
            conn.Close();

            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\a\\formatado\\VendasFormatadoHwashin.xlsx; Extended Properties=Excel 12.0;";

            using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            {
                OleDbCommand cmd = new OleDbCommand("Select [Código do Cliente], [Número Nota Fiscal], [Série Nota Fiscal], [Código da Divisão]" +
                    ",[CFOP] ,[Data Emissão], [Data Vencimento]" +
                    ",[Prazo de Vencimento], [Item Nota Fiscal], [Código do Produto], [Quantidade]" +
                    ",[Valor Venda Sem o IPI (Reais)], [Descontos Incondicionais], [ICMS], [PIS], [COFINS], [ISS], [Comissão], [Frete], [Seguro]" +
                    ",[Data de Embarque], [Código Moeda estrangeira], [Valor em Moeda estrangeira]" +
                    ", [Custo da Venda Total] " +
                    " from [Vendas$]", connection);

                connection.Open();
                OleDbDataReader dReader = cmd.ExecuteReader();

                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(sqlConnectionString))
                {
                    sqlBulk.DestinationTableName = "Vendas_Itens";
                    sqlBulk.WriteToServer(dReader);
                }

                SqlCommand cmdCopPedido = conn.CreateCommand();

                cmdCopPedido.CommandText =
                        @"INSERT INTO D_Vendas_Itens
						([Vnd_Cli_ID],
	                        [Vnd_NF_ID],
	                        [Vnd_NF_Serie],
	                        [Vnd_Cod_Divisao],
	                        [Vnd_CFOP],
	                        [Vnd_Dt_Emissao],
	                        [Vnd_DT_Vencimento],
	                        [Vnd_Dias],
	                        [Vnd_Item],
	                        [Vnd_Pro_id],
	                        [Vnd_Qtde],
	                        [Vnd_Vl_Nota],
	                        [Vnd_Desconto],
	                        [Vnd_ICMS],
	                        [Vnd_PIS],
	                        [Vnd_COFINS],
	                        [Vnd_ISS],
	                        [Vnd_Comissao],
	                        [Vnd_Frete],
	                        [Vnd_Seguro],
	                        [Vnd_Dt_Embarque],
	                        [Vnd_Cod_Moeda],
	                        [Vnd_Vl_Moeda],
	                        [Vnd_Custo],
	                        [Vnd_RE],
	                        [Vnd_CNPJ],
	                        [Lin_Origem_ID])
                          SELECT [Vnd_Cli_ID],
	                        [Vnd_NF_ID],
	                        [Vnd_NF_Serie],
	                        [Vnd_Cod_Divisao],
	                        [Vnd_CFOP],
	                        [Vnd_Dt_Emissao],
	                        [Vnd_DT_Vencimento],
	                        [Vnd_Dias],
	                        [Vnd_Item],
	                        [Vnd_Pro_id],
	                        [Vnd_Qtde],
	                        [Vnd_Vl_Nota],
	                        [Vnd_Desconto],
	                        [Vnd_ICMS],
	                        [Vnd_PIS],
	                        [Vnd_COFINS],
	                        [Vnd_ISS],
	                        [Vnd_Comissao],
	                        [Vnd_Frete],
	                        [Vnd_Seguro],
	                        [Vnd_Dt_Embarque],
	                        [Vnd_Cod_Moeda],
	                        [Vnd_Vl_Moeda],
	                        [Vnd_Custo],
	                        [Vnd_RE],
	                        [Vnd_CNPJ],
	                        [ID]
                            FROM Vendas_Itens
                            WHERE VND_NF_ID IS NOT NULL AND VND_CFOP IS NOT NULL and VND_PRO_ID IS NOT NULL";

                // select*
                //from Clientes a left join d_clientes b on a.id = b.lin_origem_id
                //where b.Cli_ID is null

                SqlTransaction tr = null;
                try
                {
                    conn.Open();
                    tr = conn.BeginTransaction();
                    cmdCopPedido.Transaction = tr;
                    cmdCopPedido.ExecuteNonQuery();
                    tr.Commit();

                    label1.Text = "Tabela vendas copiada ";
                }
                catch (Exception ex)
                {
                    tr.Rollback();
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    conn.Close();
                }

            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            MyApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;

            MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\OrdensOriginalHwashin.xlsx");
            Workbook wb = MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\OrdensOriginalHwashin.xlsx");
            Worksheet ws = wb.Sheets[1];

            ws.Range["A:A"].NumberFormat = "@";
            ws.Range["D:D"].NumberFormat = "@";
            MyApp.DisplayAlerts = false;
            wb.SaveAs("c:\\a\\formatado\\OrdemProducaoformatadoHwashin.xlsx");
            wb.Close();
            MyApp.Quit();
            SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");
            string sqlConnectionString = "Data Source=BRCAENRODRIGUES\\SQLEXPRESS;Initial Catalog=my_database;Integrated Security=True";

            SqlCommand cmdColuna = conn.CreateCommand();

            cmdColuna.CommandText =
              @"IF OBJECT_ID('dbo.Insumo_Produto', 'U') IS NOT NULL 
                  DROP TABLE dbo.Insumo_Produto; 
                    CREATE TABLE [dbo].[Insumo_Produto](
	                [Ins_PA_Pro_Id] [varchar](70) NULL,
	                [Ins_Qtd_Produzida] [numeric](24, 12) NULL,
	                [Ins_PA_Und_Id] [int] NULL,
	                [Ins_MP_Pro_Id] [varchar](70) NULL,
	                [Ins_Qtd_Requisitada] [numeric](24, 12) NULL,
	                [Ins_MP_Und_Id] [int] NULL,
	                [Ins_Ordem_Prod] [varchar](70) NULL,
	                [Ins_DT_Ini] [datetime] NULL,
	                [Ins_DT_Fim] [datetime] NULL,
	                [Ins_CNPJ] [varchar](20) NULL CONSTRAINT [DF__D_Insumo___Ins_C__037C6257]  DEFAULT ('00000000000000'),
                    [ID] [int] IDENTITY(1,1) NOT NULL,
                    [Arq_Origem_ID] [int] NULL)";

            SqlTransaction trA = null;

            conn.Open();
            trA = conn.BeginTransaction();
            cmdColuna.Transaction = trA;
            cmdColuna.ExecuteNonQuery();
            trA.Commit();
            conn.Close();

            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\a\\formatado\\OrdemProducaoformatadoHwashin.xlsx; Extended Properties=Excel 12.0;";

            using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            {
                OleDbCommand cmd = new OleDbCommand("Select [Código do Produto Acabado/Projeto], [Quantidade Produzida], [Unidade de Medida do Produto Acabado], [Código Matéria-Prima], [Quantidade Requisitada], [Unidade de Medida Matéria-Prima]" +
                    ",[Número da Ordem de Produção], [Data de Início], [Data Fim], [CNPJ] from [Ordens de Produção$]", connection);

                connection.Open();
                OleDbDataReader dReader = cmd.ExecuteReader();

                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(sqlConnectionString))
                {
                    sqlBulk.DestinationTableName = "Insumo_Produto";
                    sqlBulk.WriteToServer(dReader);
                }

                SqlCommand cmdCopPedido = conn.CreateCommand();

                cmdCopPedido.CommandText =
                        @"INSERT INTO [D_Insumo_Produto]
						([Ins_PA_Pro_Id],
	                        [Ins_Qtd_Produzida],
	                        [Ins_PA_Und_Id],
	                        [Ins_MP_Pro_Id],
	                        [Ins_Qtd_Requisitada],
	                        [Ins_MP_Und_Id],
	                        [Ins_Ordem_Prod],
	                        [Ins_DT_Ini],
	                        [Ins_DT_Fim],
	                        [Ins_CNPJ],
	                        [Lin_Origem_ID],
	                        [Arq_Origem_ID])
                          SELECT [Ins_PA_Pro_Id],
	                        [Ins_Qtd_Produzida],
	                        [Ins_PA_Und_Id],
	                        [Ins_MP_Pro_Id],
	                        [Ins_Qtd_Requisitada],
	                        [Ins_MP_Und_Id],
	                        [Ins_Ordem_Prod],
	                        [Ins_DT_Ini],
	                        [Ins_DT_Fim],
	                        [Ins_CNPJ],
	                        [ID],
	                        [Arq_Origem_ID]
                            FROM [Insumo_Produto]
                            WHERE Ins_MP_Pro_Id IS NOT NULL";

                //WHERE VND_NF_ID IS NOT NULL AND VND_CFOP IS NOT NULL";

                // select*
                //from Clientes a left join d_clientes b on a.id = b.lin_origem_id
                //where b.Cli_ID is null

                SqlTransaction tr = null;
                try
                {
                    conn.Open();
                    tr = conn.BeginTransaction();
                    cmdCopPedido.Transaction = tr;
                    cmdCopPedido.ExecuteNonQuery();
                    tr.Commit();

                    label1.Text = "Tabela Insumo_Produto copiada ";
                }
                catch (Exception ex)
                {
                    tr.Rollback();
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    conn.Close();
                }

            }
        }

        private void button9_Click(object sender, EventArgs e)
        {

            MyApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;

            MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\RelacaoOriginalHwashin.xlsx");
            Workbook wb = MyApp.Workbooks.Add("C:\\a\\OriginalHwashin\\RelacaoOriginalHwashin.xlsx");
            Worksheet ws = wb.Sheets[1];

            ws.Range["A:A"].NumberFormat = "@";
           // ws.Range["D:D"].NumberFormat = "@";
            MyApp.DisplayAlerts = false;
            wb.SaveAs("c:\\a\\formatado\\relacaoProducaoformatadoHwashin.xlsx");
            wb.Close();
            MyApp.Quit();
            SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");
            string sqlConnectionString = "Data Source=BRCAENRODRIGUES\\SQLEXPRESS;Initial Catalog=my_database;Integrated Security=True";

            SqlCommand cmdColuna = conn.CreateCommand();

            cmdColuna.CommandText =
              @"IF OBJECT_ID('dbo.Relacao_Carga', 'U') IS NOT NULL 
                  DROP TABLE dbo.Relacao_Carga; 
                    CREATE TABLE [dbo].[Relacao_Carga](
	                [Rel_PA_Pro_ID] [varchar](70) NOT NULL,
	                [Rel_MP_Pro_ID] [varchar](70) NOT NULL,
	                [Rel_Produzida] [float] NULL CONSTRAINT [DF_Relacao_Rel_Produzida]  DEFAULT ('0'),
	                [Rel_Requisitada] [float] NULL CONSTRAINT [DF_Relacao_Rel_Requisitada]  DEFAULT ('0'),
	                [Rel_Relacao] [float] NULL,
	                [Rel_Tipo] [varchar](2) NULL,
	                [ID] [int] IDENTITY(1,1) NOT NULL,
	                [Arq_Origem_ID] [int] NULL)";

            SqlTransaction trA = null;

            conn.Open();
            trA = conn.BeginTransaction();
            cmdColuna.Transaction = trA;
            cmdColuna.ExecuteNonQuery();
            trA.Commit();
            conn.Close();

            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\\a\\formatado\\relacaoProducaoformatadoHwashin.xlsx; Extended Properties=Excel 12.0;";

            using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            {
                OleDbCommand cmd = new OleDbCommand("Select [Código Produto Acabado]," +
                    "[Código Matéria-prima]," +
                    "[Quantidade Produzida], [Quantidade Requisitada]," +
                    " [Relação de Produção], [Tipo de Relação]" +
                    "  from [Relação Produção$]", connection);

                connection.Open();
                OleDbDataReader dReader = cmd.ExecuteReader();

                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(sqlConnectionString))
                {
                    sqlBulk.DestinationTableName = "Relacao_Carga";
                    sqlBulk.WriteToServer(dReader);
                }

                SqlCommand cmdCopPedido = conn.CreateCommand();

                cmdCopPedido.CommandText =
                        @"INSERT INTO [D_Relacao_Carga]
						([Rel_PA_Pro_ID],
	                        [Rel_MP_Pro_ID],
	                        [Rel_Produzida],
	                        [Rel_Requisitada],
	                        [Rel_Relacao],
	                        [Rel_Tipo])
                          SELECT  max([Rel_PA_Pro_ID]),
	                                max([Rel_MP_Pro_ID]),
	                                max([Rel_Produzida]),
	                                max([Rel_Requisitada]),
	                                max([Rel_Relacao]),
	                                max([Rel_Tipo])
                            FROM [Relacao_Carga]
                            group by
                            [Rel_PA_Pro_ID],
	                        [Rel_MP_Pro_ID],
	                        [Rel_Tipo]";

                //WHERE VND_NF_ID IS NOT NULL AND VND_CFOP IS NOT NULL";

                // select*
                //from Clientes a left join d_clientes b on a.id = b.lin_origem_id
                //where b.Cli_ID is null

                SqlTransaction tr = null;
                try
                {
                    conn.Open();
                    tr = conn.BeginTransaction();
                    cmdCopPedido.Transaction = tr;
                    cmdCopPedido.ExecuteNonQuery();
                    tr.Commit();

                    label1.Text = "Tabela Relacao_Carga copiada ";
                }
                catch (Exception ex)
                {
                    tr.Rollback();
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    conn.Close();
                }

            }
        }

        private void buttonAbrir_Click(object sender, EventArgs e)
        {

        }
    }
}

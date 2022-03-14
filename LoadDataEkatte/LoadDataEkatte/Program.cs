using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npgsql;
using OfficeOpenXml;

namespace LoadDataEkatte
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string databaseInfo = "Host=localhost:5433;Username=******;Password=******;Database=******";

            using (NpgsqlConnection connection = new NpgsqlConnection(databaseInfo))
            {
                connection.Open();

                if (connection.State == ConnectionState.Open)
                {
                    LoadAreas(connection);
                    
                    LoadMunicipalities(connection);

                    LoadSettlements(connection);

                    LoadBelongsTo(connection);
                }

                connection.Close();
            }
        }

        private static void LoadAreas(NpgsqlConnection connection)
        {
            string filePath = "EKATTE\\Ek_obl.xlsx";
            byte[] fileBytes = File.ReadAllBytes(filePath);

            using (MemoryStream memoryStream = new MemoryStream(fileBytes))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(memoryStream))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.First();

                    string cell = string.Empty;
                    string[] record = new string[sheet.Dimension.End.Column];

                    string sqlSelect = "SELECT * FROM AREAS WHERE AREA = @AREA";
                    string sqlInsert = "INSERT INTO AREAS(AREA, EKATTE, NAME, REGION, DOCUMENT, ABC) VALUES(@AREA, @EKATTE, @NAME, @REGION, @DOCUMENT, @ABC)";

                    for (int row = sheet.Dimension.Start.Row + 1; row <= sheet.Dimension.End.Row; row++)
                    {
                        for (int column = sheet.Dimension.Start.Column; column <= sheet.Dimension.End.Column; column++)
                        {
                            cell = sheet.Cells[row, column].Value.ToString();
                            record[column - 1] = cell;
                        }

                        NpgsqlCommand commandSelect = new NpgsqlCommand(sqlSelect, connection);
                        NpgsqlCommand commandInsert = new NpgsqlCommand(sqlInsert, connection);

                        commandSelect.Parameters.AddWithValue("AREA", record[0]);
                        commandSelect.Prepare();

                        bool recordExists = false;
                        using (NpgsqlDataReader npgsqlDataReader = commandSelect.ExecuteReader())
                        {
                            if (npgsqlDataReader.Read())
                            {
                                recordExists = true;
                            }
                        };

                        if (recordExists == false) 
                        {
                            commandInsert.Parameters.AddWithValue("AREA", record[0]);
                            commandInsert.Parameters.AddWithValue("EKATTE", record[1]);
                            commandInsert.Parameters.AddWithValue("NAME", record[2]);
                            commandInsert.Parameters.AddWithValue("REGION", record[3]);
                            commandInsert.Parameters.AddWithValue("DOCUMENT", record[4]);
                            commandInsert.Parameters.AddWithValue("ABC", record[5]);

                            commandInsert.Prepare();
                            commandInsert.ExecuteNonQuery();
                        }
                    }
                }
            }
        }

        private static void LoadMunicipalities(NpgsqlConnection connection)
        {
            string filePath = "EKATTE\\Ek_obst.xlsx";
            byte[] fileBytes = File.ReadAllBytes(filePath);

            using (MemoryStream memoryStream = new MemoryStream(fileBytes))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(memoryStream))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.First();

                    string cell = string.Empty;
                    string[] record = new string[sheet.Dimension.End.Column];

                    string sqlSelect = "SELECT * FROM MUNICIPALITIES WHERE MUNICIPALITY = @MUNICIPALITY";
                    string sqlInsert = "INSERT INTO MUNICIPALITIES(MUNICIPALITY, EKATTE, NAME, CATEGORY, DOCUMENT, ABC, AREA) VALUES(@MUNICIPALITY, @EKATTE, @NAME, @CATEGORY, @DOCUMENT, @ABC, @AREA)";

                    for (int row = sheet.Dimension.Start.Row + 1; row <= sheet.Dimension.End.Row; row++)
                    {
                        for (int column = sheet.Dimension.Start.Column; column <= sheet.Dimension.End.Column; column++)
                        {
                            cell = sheet.Cells[row, column].Value.ToString();
                            record[column - 1] = cell;
                        }

                        NpgsqlCommand commandSelect = new NpgsqlCommand(sqlSelect, connection);
                        NpgsqlCommand commandInsert = new NpgsqlCommand(sqlInsert, connection);

                        commandSelect.Parameters.AddWithValue("MUNICIPALITY", record[0]);
                        commandSelect.Prepare();

                        bool recordExists = false;
                        using (NpgsqlDataReader npgsqlDataReader = commandSelect.ExecuteReader())
                        {
                            if (npgsqlDataReader.Read())
                            {
                                recordExists = true;
                            }
                        };
                        
                        if (recordExists == false)
                        {
                            commandInsert.Parameters.AddWithValue("MUNICIPALITY", record[0]);
                            commandInsert.Parameters.AddWithValue("EKATTE", record[1]);
                            commandInsert.Parameters.AddWithValue("NAME", record[2]);
                            commandInsert.Parameters.AddWithValue("CATEGORY", record[3]);
                            commandInsert.Parameters.AddWithValue("DOCUMENT", record[4]);
                            commandInsert.Parameters.AddWithValue("ABC", record[5]);
                            commandInsert.Parameters.AddWithValue("AREA", record[0].Substring(0, 3));

                            commandInsert.Prepare();
                            commandInsert.ExecuteNonQuery();
                        }
                    }
                }
            }
        }

        private static void LoadSettlements(NpgsqlConnection connection)
        {
            string filePath = "EKATTE\\Ek_sobr.xlsx";
            byte[] fileBytes = File.ReadAllBytes(filePath);

            using (MemoryStream memoryStream = new MemoryStream(fileBytes))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(memoryStream))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.First();

                    string cell = string.Empty;
                    string[] record = new string[sheet.Dimension.End.Column];

                    string sqlSelect = "SELECT * FROM SETTLEMENTS WHERE EKATTE = @EKATTE";
                    string sqlInsert = "INSERT INTO SETTLEMENTS(EKATTE, KIND, NAME, AREA1, AREA2, DOCUMENT, ABC) VALUES(@EKATTE, @KIND, @NAME, @AREA1, @AREA2, @DOCUMENT, @ABC)";

                    for (int row = sheet.Dimension.Start.Row + 1; row <= sheet.Dimension.End.Row; row++)
                    {
                        for (int column = sheet.Dimension.Start.Column; column <= sheet.Dimension.End.Column; column++)
                        {
                            cell = sheet.Cells[row, column].Value.ToString();
                            record[column - 1] = cell;
                        }

                        NpgsqlCommand commandSelect = new NpgsqlCommand(sqlSelect, connection);
                        NpgsqlCommand commandInsert = new NpgsqlCommand(sqlInsert, connection);

                        commandSelect.Parameters.AddWithValue("EKATTE", record[0]);
                        commandSelect.Prepare();

                        bool recordExists = false;
                        using (NpgsqlDataReader npgsqlDataReader = commandSelect.ExecuteReader())
                        {
                            if (npgsqlDataReader.Read())
                            {
                                recordExists = true;
                            }
                        };

                        if (recordExists == false)
                        {
                            commandInsert.Parameters.AddWithValue("EKATTE", record[0]);
                            commandInsert.Parameters.AddWithValue("KIND", record[1]);
                            commandInsert.Parameters.AddWithValue("NAME", record[2]);
                            commandInsert.Parameters.AddWithValue("AREA1", record[3]);
                            commandInsert.Parameters.AddWithValue("AREA2", record[4]);
                            commandInsert.Parameters.AddWithValue("DOCUMENT", record[5]);
                            commandInsert.Parameters.AddWithValue("ABC", record[6]);

                            commandInsert.Prepare();
                            commandInsert.ExecuteNonQuery();
                        }
                    }
                }
            }
        }

        private static void LoadBelongsTo(NpgsqlConnection connection)
        {
            string sqlSelect = "SELECT * FROM BELONGSTO WHERE SETTLEMENT_EKATTE = @SETTLEMENT_EKATTE AND MUNICIPALITY = @MUNICIPALITY";
            string sqlInsert = "INSERT INTO BELONGSTO(SETTLEMENT_EKATTE, MUNICIPALITY) VALUES(@SETTLEMENT_EKATTE, @MUNICIPALITY)";

            List<(string, string)> listBelongsTo = new List<(string, string)>();
            CreateListBelongsTo(connection, ref listBelongsTo);
            
            foreach ((string, string) pair in listBelongsTo)
            {
                NpgsqlCommand commandSelect = new NpgsqlCommand(sqlSelect, connection);
                NpgsqlCommand commandInsert = new NpgsqlCommand(sqlInsert, connection);

                commandSelect.Parameters.AddWithValue("SETTLEMENT_EKATTE", pair.Item1);
                commandSelect.Parameters.AddWithValue("MUNICIPALITY", pair.Item2);
                commandSelect.Prepare();

                bool recordExists = false;
                using (NpgsqlDataReader npgsqlDataReader = commandSelect.ExecuteReader())
                {
                    if (npgsqlDataReader.Read())
                    {
                        recordExists = true;
                    }
                };

                if (recordExists == false)
                {
                    commandInsert.Parameters.AddWithValue("SETTLEMENT_EKATTE", pair.Item1);
                    commandInsert.Parameters.AddWithValue("MUNICIPALITY", pair.Item2);

                    commandInsert.Prepare();
                    commandInsert.ExecuteNonQuery();
                }
            }
        }

        private static void CreateListBelongsTo(NpgsqlConnection connection, ref List<(string, string)> listBelongsTo)
        {
            string sqlSelectFromSettlements = "SELECT EKATTE, AREA1, AREA2 FROM SETTLEMENTS";
            NpgsqlCommand commandSelect = new NpgsqlCommand(sqlSelectFromSettlements, connection);

            Dictionary<string, (string, string)> settlementsAreas = new Dictionary<string, (string, string)>();

            using (NpgsqlDataReader npgsqlDataReader = commandSelect.ExecuteReader())
            {
                while (npgsqlDataReader.Read())
                {
                    settlementsAreas.Add(npgsqlDataReader.GetString(0), (npgsqlDataReader.GetString(1), npgsqlDataReader.GetString(2)));
                }
            };

            List<(string, string[])> ekatteAreaDataPairs = new List<(string, string[])>();

            foreach (var settlementAreas in settlementsAreas)
            {
                string ekatte = settlementAreas.Key;
                string area1 = settlementAreas.Value.Item1;
                string area2 = settlementAreas.Value.Item2;

                char[] separators = { '.', ',' };
                string[] splittedArea1 = area1.Split(separators, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string[] splittedArea2 = area2.Split(separators, StringSplitOptions.RemoveEmptyEntries).ToArray();

                string area1Field = splittedArea1[splittedArea1.Length - 1].Substring(1);
                string area1MunicipalityName = splittedArea1[splittedArea1.Length - 3].Substring(1);

                ekatteAreaDataPairs.Add((ekatte, new[] { area1Field, area1MunicipalityName }));

                if (splittedArea2.Length != 0)
                {
                    string area2Field = splittedArea2[splittedArea2.Length - 1].Substring(1);
                    string area2MunicipalityName = splittedArea2[splittedArea2.Length - 3].Substring(1);

                    if (area1MunicipalityName != area2MunicipalityName)
                    {
                        ekatteAreaDataPairs.Add((ekatte, new[] { area2Field, area2MunicipalityName }));
                    }
                }
            }

            foreach (var pair in ekatteAreaDataPairs)
            {
                listBelongsTo.Add((pair.Item1, DetermineMunicipality(connection, pair.Item2)));
            }
        }

        private static string DetermineMunicipality(NpgsqlConnection connection, string[] areaData)
        {
            string sqlSelectFromAreas = "SELECT AREA FROM AREAS WHERE NAME = @NAME";
            NpgsqlCommand commandSelectFromAreas = new NpgsqlCommand(sqlSelectFromAreas, connection);

            commandSelectFromAreas.Parameters.AddWithValue("NAME", areaData[0]);
            commandSelectFromAreas.Prepare();

            string areaCode = string.Empty;

            using (NpgsqlDataReader npgsqlDataReader = commandSelectFromAreas.ExecuteReader())
            {
                if (npgsqlDataReader.Read())
                {
                    areaCode = npgsqlDataReader.GetString(0);
                }
            };

            string sqlSelectFromMunicipalities = "SELECT MUNICIPALITY FROM MUNICIPALITIES WHERE NAME = @NAME";
            NpgsqlCommand commandSelectFromMunicipalities = new NpgsqlCommand(sqlSelectFromMunicipalities, connection);

            commandSelectFromMunicipalities.Parameters.AddWithValue("NAME", areaData[1]);
            commandSelectFromMunicipalities.Prepare();

            string municipalityCode = string.Empty;

            using (NpgsqlDataReader npgsqlDataReader = commandSelectFromMunicipalities.ExecuteReader())
            {
                while (npgsqlDataReader.Read())
                {
                    if (npgsqlDataReader.GetString(0).Substring(0, 3) == areaCode)
                    {
                        municipalityCode = npgsqlDataReader.GetString(0);
                    }
                }
            };

            return municipalityCode;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Collections.Specialized;
using System.Data;
using System.DirectoryServices;
using System.Text.RegularExpressions;

namespace responsibillityMonitoring
{
    class Program
    {
        static void Main(string[] args)
        {


            CheckPosition();
            Console.ReadLine();
        }


        public static string GetAdAttribute(string persId, string propertyName)
        {
            DirectoryEntry rootEntry = new DirectoryEntry("LDAP://OU=ACCOUNTS,DC=company,DC=ru");
            string searcherFilter = $"(sAMAccountName=SC{persId})";
            DirectorySearcher searcher = new DirectorySearcher(rootEntry, searcherFilter);
            SearchResult result = searcher.FindOne();
            if (result == null) { return null; }

            using (DirectoryEntry entry = result.GetDirectoryEntry())
            {
                return entry.Properties[propertyName].Value.ToString();
            }
        }

        public static void CheckPosition()
        {
            Program pr;

            var connectionString = ConfigurationManager.ConnectionStrings["SQLConnection"].ConnectionString;
            var connectionStringSQL3 = ConfigurationManager.ConnectionStrings["SQL3Connection"].ConnectionString;

            SqlConnection con = new SqlConnection(connectionString);
            SqlConnection con2 = new SqlConnection(connectionStringSQL3);
            con.Open();
            con2.Open();

            // получим максимальное значение id из таблицы ProductManagersResponsible на ESF_TQ

            string maxIdProfQuery = @"SELECT Max(id) FROM [ESF_TQ].[dbo].[ProductManagersResponsible]";

            SqlCommand cmdMax = new SqlCommand(maxIdProfQuery, con2);
            int maxIdProfId = (int)cmdMax.ExecuteScalar();


            List<int> ProfIdArr = new List<int>(); // Уникальные ProfessionID, которые соотвествуют дате

            for (int j = 1; j <= maxIdProfId; j++)
            {
                string arrQuery = $@"SELECT Distinct [ProductManagersResponsible].ProfessionID
                      FROM [ESF_TQ].[dbo].[ProductManagersResponsible]
                      WHERE id = {j} and [ProductManagersResponsible].[DateIn]<=getdate() and [ProductManagersResponsible].[DateOut]>=getdate()";
                SqlCommand cmdArr = new SqlCommand(arrQuery, con2);
                int ProfessionIdToArr = (int)cmdArr.ExecuteScalar();
                if (!ProfIdArr.Contains(ProfessionIdToArr))
                {
                    ProfIdArr.Add(ProfessionIdToArr);
                }
            }
            foreach (var profId in ProfIdArr)
            {
                // найдём personId
                string personIdQuery = $"SELECT PersonID FROM Applications.[dbo].[Staff_Movement] where[ProfessionID] = {profId} and DateIn <= GETDATE() and DateOut >= GETDATE()";

                SqlCommand cmdPerson = new SqlCommand(personIdQuery, con);

                try
                {
                    int personId = (int)cmdPerson.ExecuteScalar();
                    Console.WriteLine($"На должности {profId} работает человек с  id: {personId} ");
                }
                catch (Exception)
                {
                    // попадаем в этот блок если должность есть, но человека на должности нет
                    Console.WriteLine($"не найден человек");
                    // для начала определим DivisionId
   
                    string DivisionIdQeury = $@"SELECT [DivisionID]
                        FROM [Applications].[dbo].[Staff_Profession]
                        where ProfessionID = {profId}";
                    SqlCommand cmdDivis = new SqlCommand(DivisionIdQeury, con);

                    int division;
                    int divisionId;
                    try
                    {
                        division = (int)cmdDivis.ExecuteScalar();
                        // Проверяем есть ли такой отдел в данное время
                        string IsDivision = $@"SELECT[DivisionID]
                          FROM [Applications].[dbo].[Staff_Division]
                          where [DivisionID] = {division}
                          and DateIn <= GETDATE() and DateOut >= GETDATE();";
                        SqlCommand cmdIsDivis = new SqlCommand(IsDivision, con);
                        try
                        {
                            divisionId = (int)cmdIsDivis.ExecuteScalar();

                            // Имея код подразделения мы найдём ChiefProfessionId
                            string CheifProfessionIdQuery = $@"SELECT [ChiefProfessionID] FROM [Applications].[dbo].[Staff_Division] where [DivisionID] = {division} and DateIn <= GETDATE() and DateOut >= GETDATE()";
                            SqlCommand cmdCheif = new SqlCommand(CheifProfessionIdQuery, con);
                            int CheifId;

                            try
                            {
                                CheifId = (int)cmdCheif.ExecuteScalar();
                                // нашли руководителя, теперь найдём его personid
                                string personIdOfCheifQuery = $@"SELECT [PersonID]
                                                FROM [Applications].[dbo].[Staff_Movement]
                                                where [ProfessionID] = {CheifId}";
                                SqlCommand cmdPerChief = new SqlCommand(personIdOfCheifQuery, con);
                                int personIdOfCheif = (int)cmdPerChief.ExecuteScalar();
                                string personIdOfCheifStr = personIdOfCheif.ToString();
                                // получим email руководителя
                                string emailOfChief = GetAdAttribute(personIdOfCheifStr, "mail");
                                // получим ИО руководителя
                                string nameOfChief = GetAdAttribute(personIdOfCheifStr, "givenNameRus");
                                string middleNameOfChief = GetAdAttribute(personIdOfCheifStr, "middleNameRus");


                                // Найдём пол руководителя
                                string genderOfChiefQuery = $@"SELECT [Sex]
                                                    FROM[Applications].[dbo].[Staff_Personnel]
                                                    where PersonID = {personIdOfCheifStr}";
                                SqlCommand cmdGender = new SqlCommand(genderOfChiefQuery, con);
                                byte gender = (byte)cmdGender.ExecuteScalar();

                                // вместо ProfId выведем ObjectId


                                pr = new Program();
                                List<int> objId = new List<int>();
                                objId = pr.MakeListId(maxIdProfId, profId, con2);

              
                                Console.WriteLine($"Будет отправлено письмо на адрес: {emailOfChief}");
                                pr = new Program();
                                pr.SendEmailThereisntMan(emailOfChief, nameOfChief, middleNameOfChief, objId, gender);
                            }
                            catch (Exception)
                            {
                                string parentDivisionIdQuery = $@"SELECT[ParentDivisionID]
                                        FROM[Applications].[dbo].[Staff_Division]
                                        where[DivisionID] = {division}
                                        and DateIn <= GETDATE() and DateOut >= GETDATE()";
                                SqlCommand cmdparentDivis = new SqlCommand(parentDivisionIdQuery, con);
                                int parentDivisionId = (int)cmdparentDivis.ExecuteScalar();
                                // Имея код подразделения мы найдём ChiefProfessionId
                                string CheifProfessionIdQuery2 = $@"SELECT [ChiefProfessionID] FROM [Applications].[dbo].[Staff_Division] where [DivisionID] = {parentDivisionId} and DateIn <= GETDATE() and DateOut >= GETDATE()";
                                SqlCommand cmdCheif2 = new SqlCommand(CheifProfessionIdQuery2, con);
                                int CheifId2;
                                try
                                {
                                    CheifId2 = (int)cmdCheif2.ExecuteScalar();
                                    // нашли руководителя, теперь найдём его personid
                                    string personIdOfCheifQuery = $@"SELECT [PersonID]
                                                FROM [Applications].[dbo].[Staff_Movement]
                                                where [ProfessionID] = {CheifId2}";
                                    SqlCommand cmdPerChief2 = new SqlCommand(personIdOfCheifQuery, con);
                                    int personIdOfCheif2 = (int)cmdPerChief2.ExecuteScalar();
                                    string personIdOfCheif2Str = personIdOfCheif2.ToString();
                                    // получим email руководителя parentDivision
                                    string emailOfChief2 = GetAdAttribute(personIdOfCheif2Str, "mail");
                                    // получим ИО руководителя
                                    string nameOfChief2 = GetAdAttribute(personIdOfCheif2Str, "givenNameRus");
                                    string middleNameOfChief2 = GetAdAttribute(personIdOfCheif2Str, "middleNameRus");
                                    string genderOfChiefQuery2 = $@"SELECT [Sex]
                                                    FROM[Applications].[dbo].[Staff_Personnel]
                                                    where PersonID = {personIdOfCheif2Str}";
                                    SqlCommand cmdGender2 = new SqlCommand(genderOfChiefQuery2, con);
                                    byte gender2 = (byte)cmdGender2.ExecuteScalar();

                                    pr = new Program();
                                    List<int> objId = new List<int>();
                                    objId = pr.MakeListId(maxIdProfId, profId, con2);
                            
                                    Console.WriteLine($"Будет отправлено письмо на адрес: {emailOfChief2}");

                                    pr = new Program();

                                    List<int> objId2 = new List<int>();
                                    objId2 = pr.MakeListId(maxIdProfId, profId, con2);
                                    pr.SendEmailThereisntMan(emailOfChief2, nameOfChief2, middleNameOfChief2, objId2, gender2);
                                }
                                catch (Exception)
                                {
                                    // Если не нашли руководителя parentDivision отправляем письмо на: (olga.pavlyuk@company.ru)
                                    pr = new Program();
                                    List<int> objId5 = new List<int>();
                                    objId5 = pr.MakeListId(maxIdProfId, profId, con2);
                                    pr.SendEmailThereisntDivition(objId5);
                                }
                            }
                        }
                        catch (Exception)
                        {
                            // сюда попадаем если отдел закрыт
                            Console.WriteLine("Не найден отдел");
                            Console.WriteLine($"Будет отправлено письмо на адрес: olga.pavlyuk@company.ru");
                            // если есть должность, но нет человека и нет уже этого отдела - отправляем письмо


                            pr = new Program();

                            List<int> objId3 = new List<int>();
                            objId3 = pr.MakeListId(maxIdProfId, profId, con2);

                            pr.SendEmailThereisntDivition(objId3); // В письме пишем, что нет уже этого отдела
                        }
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("Не найден отдел");
                        Console.WriteLine($"Будет отправлено письмо на адрес: olga.pavlyuk@company.ru");
                        // если есть должность, но нет человека и нет уже этого отдела - отправляем письмо
                        pr = new Program();

                        List<int> objId4 = new List<int>();
                        objId4 = pr.MakeListId(maxIdProfId, profId, con2);

                        pr.SendEmailThereisntDivition(objId4); // В письме пишем, что нет уже этого отдела
                    }
                }
            }
            con.Close();
            con2.Close();
        }


        private List<int> MakeListId(int maxIdProfId, int profId, SqlConnection con2)
        {
            List<int> objIdList = new List<int>();
            for (int i = 1; i < maxIdProfId; i++)
            {
                string objIdQuery = $@"SELECT [ObjectID]
                                                  FROM [ESF_TQ].[dbo].[ProductManagersResponsible]
                                                  where ProfessionID = {profId}
                                                  and id = {i}
                                            and DateIn <= GETDATE() and DateOut >= GETDATE()";
                SqlCommand cmdObj = new SqlCommand(objIdQuery, con2);
                try
                {
                    int objectId = (int)cmdObj.ExecuteScalar();
                    if (!objIdList.Contains(objectId))
                    {
                        objIdList.Add(objectId);
                    }
                }
                catch (Exception) { }
            }
            return objIdList;
        }



        private void SendEmailThereisntMan(string emailOfChief, string firstName, string middleName, List<int> objList, byte gender)
        {

            string listId = String.Empty;
            foreach (var item in objList)
            {
                listId += "id = " + item + ", ";
            }
            listId = listId.Substring(0, listId.Length - 2);

            string recipient = String.Empty;
       
            string pattern = @"id";
            Regex rg = new Regex(pattern);
            MatchCollection matched = rg.Matches(listId);
            int amount = 0;
            for (int count = 0; count < matched.Count; count++)
            {
                amount++;
            }
            string brand = (amount > 1) ? "брендам" : "бренду";
            
            recipient = (gender == 1) ? "Уважаемый" : "Уважаемая";
    

            string writer = "noreplay@company.ru"; // Автори письма
            string title = "Мониторинг ответственности менеджеров-продукта";

            string text = $@"{recipient} {firstName} {middleName}, добрый день.
                            По {brand}: {listId} отсутствует ответственный менеджер-продукта.
                            Просьба назначить ответственного.

                            {emailOfChief}

                            С уважением, Робот.
                            Не отвечайте на это письмо.";

            string[] emals = new string[] { "aleksey.melentev@company.ru" }; // здесь нужно указать параметр emailOfChief

            MailingHelpers.SendMail(writer, title, text, emals);
        }

        private void SendEmailThereisntDivition(List<int> ListId)
        {
            string listId = String.Empty;
            foreach (var item in ListId)
            {
                listId += "id = " + item + ", ";
            }
            listId = listId.Substring(0, listId.Length - 2);

            string recipient = String.Empty;

            string pattern = @"id";
            Regex rg = new Regex(pattern);
            MatchCollection matched = rg.Matches(listId);
            int amount = 0;
            for (int count = 0; count < matched.Count; count++)
            {
                amount++;
            }
            string brand = (amount > 1) ? "бренды" : "бренд";
            string connected = (amount > 1) ? "привязаны" : "привязан";


            string writer = "noreplay@company.ru"; // Автори письма
            string title = "Мониторинг ответственности менеджеров-продукта";

            string text = $@"Добрый день. Cообщаем, что в данный момент {brand}: {listId} не {connected} к какому-либо отделу.
            
                        (olga.pavlyuk@company.ru)
                        С уважением, Робот.
                        Не отвечайте на это письмо.";

            string[] emals = new string[] { "aleksey.melentev@company.ru" };

            MailingHelpers.SendMail(writer, title, text, emals);
        }
    }
}
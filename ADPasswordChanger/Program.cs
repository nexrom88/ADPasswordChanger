using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADPasswordChanger
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string domainName;
            string ouString;
            string[] ouElements;
            string[] domainElements;

            //request for input
            Console.WriteLine("FQDN-Domain-Name:");
            domainName = Console.ReadLine();
            Console.WriteLine("OU-Path (e.g.: Schule\\Schueler):");
            ouString = Console.ReadLine();

            if (domainName == string.Empty || ouString == string.Empty) {
                Console.WriteLine("Wrong input");
                Console.ReadLine();
                return;
            }

            //split ouString
            if (ouString.Contains("\\"))
            {
                //multiple ous
                ouElements = ouString.Split("\\".ToCharArray());
                ouElements = Enumerable.Reverse(ouElements).ToArray();
            }
            else
            {
                //just one ou
                ouElements = new string[1];
                ouElements[0] = ouString;
            }

            //split domainString
            if (domainName.Contains("."))
            {
                //multiple domain parts
                domainElements = domainName.Split(".".ToCharArray());
            }
            else
            {
                //just one domain part
                domainElements = new string[1];
                domainElements[0] = domainName;
            }

            //get all users from ad within given ou
            List<ADUser> users;
            try
            {
                users = listUsers(domainElements, ouElements, domainName);
            }catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
                return;
            }

            //set passwords
            List<string> results = setPasswords(users);

            //write output file
            System.IO.File.WriteAllLines("output.txt", results);

            Console.WriteLine("done updating");
            Console.ReadLine();

        }

        //sets the passwords to a given list of users
        private static List<string> setPasswords(List<ADUser> users)
        {
            List<string> results = new List<string>();
            foreach (ADUser user in users)
            {
                Console.WriteLine("Setting new password for user: " + user.preName + " " + user.lastName);

                string newPassword;
                string[] specialChars = new string[] { "!", "?", "$", "#"};
                newPassword = user.preName.Substring(0, 1).ToUpper();
                newPassword += user.preName.Substring(1, 1).ToLower();

                newPassword += user.lastName.Substring(0, 1).ToUpper();
                newPassword += user.lastName.Substring(1, 1).ToLower();

                //generate random 4 digit number
                Random rand = new Random();
                newPassword += rand.Next(1000, 9999).ToString();

                //add random special char
                newPassword += specialChars[rand.Next(0, specialChars.Length -1)];

                user.principal.SetPassword(newPassword);

                user.principal.Save();

                results.Add(user.preName + " " + user.lastName + ": " + newPassword);
            }

            return results;
        }

        //get a list of all users
        private static List<ADUser> listUsers(string[] domainElements, string[] ouElements, string domainName)
        {
            List<ADUser> users = new List<ADUser>();

            //build container string
            string containerString = "";

            //build ou part
            foreach(string ou in ouElements)
            {
                containerString += "OU=" + ou + ",";
            }

            //build domain part
            foreach(string domain in domainElements)
            {
                containerString += "dc=" + domain + ",";
            }

            //remove last "," again
            containerString = containerString.Remove(containerString.Length - 1);

            //create domain context
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domainName,
                                                        containerString);
            //build a user principal
            UserPrincipal userPrincipal = new UserPrincipal(ctx);

            //create the searcher object   
            PrincipalSearcher srch = new PrincipalSearcher(userPrincipal);

            // find all matches
            foreach (UserPrincipal found in srch.FindAll())
            {
                // get distinguished name and save it to var            

                ADUser newUser = new ADUser();
                newUser.distinguishedName = found.DistinguishedName;
                newUser.lastName = found.Surname;
                newUser.preName = found.GivenName;
                newUser.principal = found;
                users.Add(newUser);
            }

            return users;
        }

        private struct ADUser
        {
            public UserPrincipal principal;
            public string distinguishedName;
            public string preName;
            public string lastName;
        }
    }
}

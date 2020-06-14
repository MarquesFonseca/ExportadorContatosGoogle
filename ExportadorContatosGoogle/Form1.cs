using MixERP.Net.VCards;
using MixERP.Net.VCards.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExportadorContatosGoogle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //var vcard = new VCard
            //{
            //    Version = MixERP.Net.VCards.Types.VCardVersion.V4,
            //    FormattedName = "John Doe",
            //    FirstName = "John",
            //    LastName = "Doe",
            //    Classification = MixERP.Net.VCards.Types.ClassificationType.Confidential,
            //    Categories = new[] { "Friend", "Fella", "Amsterdam" },
            //    Gender = MixERP.Net.VCards.Types.Gender.Female
            //};

            VCard vcard = new VCard();
            vcard.Version = MixERP.Net.VCards.Types.VCardVersion.V4;

            vcard.Prefix = "Meu Amigo";
            vcard.FirstName = "João";
            vcard.FormattedName = "João Silva Fonseca";
            vcard.LastName = "Fonseca";
            vcard.MiddleName = "Silva";
            vcard.NickName = "Marquim";
            //vcard.SortString = "Marques";
            //vcard.Suffix = 
            vcard.Addresses = new List<Address>()
            {
                new Address()
                {
                    Preference = 1,
                    Street = "Rua Venezuela",
                    PoBox = "17",
                    PostalCode = "77006-738",
                    ExtendedAddress = "Centro",
                    Region = "Tocantins",
                    Locality = "Palmas",
                    Country = "Brasil",
                    Type = MixERP.Net.VCards.Types.AddressType.Domestic
                }
                ,new Address()
                {
                    //Preference = 1,
                    Street = "Rua Paraguai",
                    PoBox = "45",
                    PostalCode = "77006-838",
                    ExtendedAddress = "Centro",
                    Region = "Tocantins",
                    Locality = "Palmas",
                    Country = "Brasil",
                    Type = MixERP.Net.VCards.Types.AddressType.Home
                }
            };
            //vcard.Anniversary = Convert.ToDateTime("02/01/1986");//casamento
            vcard.BirthDay = Convert.ToDateTime("02/01/1986");
            //vcard.CalendarAddresses = 
            //vcard.CalendarUserAddresses = 
            vcard.Categories = new string[] { "Grupo1", "Grupo2", "Grupo3" };
            vcard.Classification = MixERP.Net.VCards.Types.ClassificationType.Public;
            //vcard.CustomExtensions = 
            //vcard.DeliveryAddress = new MixERP.Net.VCards.Types.DeliveryAddress() { Address = "CAIXA POSTAL 3244, CEP 77022-971" };
            vcard.Emails = new List<Email>() {
                new Email() { EmailAddress = "marquessilvafonseca@bol.com.br", Preference = 1 }
                ,new Email() { EmailAddress = "marques.silva@correios.com.br" }
                //,new Email() { EmailAddress = "marques-fonseca@hotmail.com" }
            };
            vcard.Gender = MixERP.Net.VCards.Types.Gender.Male;
            //vcard.Impps = 
            //vcard.Key = 
            vcard.Kind = MixERP.Net.VCards.Types.Kind.Individual;
            //vcard.Languages = new List<Language>() { new Language() { Name = "pt-br", Preference = 1, Type = MixERP.Net.VCards.Types.LanguageType.Home } };
            //vcard.LastRevision = DateTime.Now;
            //vcard.Latitude = 
            //vcard.Logo = 
            //vcard.Longitude = 
            //vcard.Mailer = 
            vcard.Note = "Teste nota1";
            vcard.Organization = "Correios";
            vcard.OrganizationalUnit = "DR-TO";
            //vcard.Photo = new Photo(false, "jpg", @"https://mapacultural.secult.ce.gov.br/files/agent/6907/file/39820/1622235_697297746959776_1865972818_n-b4a796f893e8f3723e414d86128ebb8d.jpg");
            //vcard.Relations = 
            //vcard.Role = 
            //vcard.Sound = 
            //vcard.Source = 
            vcard.Telephones = new List<Telephone>() { new Telephone() { Number = "+55 63 99208-2269", Preference = 1, Type = MixERP.Net.VCards.Types.TelephoneType.Cell }, new Telephone() { Number = "+55 63 99290-6960", Type = MixERP.Net.VCards.Types.TelephoneType.Cell } };
            //vcard.TimeZone = 
            vcard.Title = "Funcionário Público Federal - ECT";
            //vcard.UniqueIdentifier = 
            //vcard.Url = 


            string serialized = MixERP.Net.VCards.Serializer.VCardSerializer.Serialize(vcard);
            string path = System.IO.Path.Combine(@"C:\Users\MARQUES\Desktop\eeeeeeeeee", "Marques.vcf");
            System.IO.File.WriteAllText(path, serialized);


            textBox1.Text = serialized;
            //MessageBox.Show("Finalizado");

        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ConsoleApplication2
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateDocument doc = new CreateDocument();

            Console.WriteLine("Enter the project number:");
            string projectNumber = Console.ReadLine();
            Console.WriteLine();

            Console.WriteLine("Enter the name of the requestor:");
            string requestor = Console.ReadLine();
            Console.WriteLine();

            Console.WriteLine("Enter the subject of the ticket:");
            string subject = Console.ReadLine();
            Console.WriteLine();

            Console.WriteLine("Enter the description of the ticket:");
            string description = Console.ReadLine();
            Console.WriteLine();

            doc.createWorkbook(projectNumber, requestor, subject, description);

            Console.WriteLine("Please choose from the menu:");
            Console.WriteLine("1. Claim Payment Error");

            int choice = 0;
            bool valid = int.TryParse(Console.ReadLine(), out choice);

            while (!valid)
            {
                Console.WriteLine();
                Console.WriteLine("Please choose from the menu:");
                Console.WriteLine("1. Claim Payment Error");
                Console.WriteLine("0. Exit");

                valid = int.TryParse(Console.ReadLine(), out choice);

                if (choice != 1)
                {
                    valid = false;
                }
            }

            switch (choice)
            {
                case 0:
                    doc.saveWorkbook(projectNumber);
                    break;
                case 1:
                    Console.WriteLine();
                    Console.WriteLine("How many payments are there to delete?");

                    int paymentsToDelete = int.Parse(Console.ReadLine());

                    List<string> draftNumbers = new List<string>();
                    List<string> claimNumbers = new List<string>();

                    for (int i = 1; i <= paymentsToDelete; i++)
                    {
                        Console.WriteLine();
                        Console.WriteLine("What is the draft number for payment " + i + "?");
                        draftNumbers.Add(Console.ReadLine());

                        Console.WriteLine();
                        Console.WriteLine("What is the claim number for payment " + i + "?");
                        claimNumbers.Add(Console.ReadLine());
                    }

                    doc.claimPaymentError(paymentsToDelete, draftNumbers, claimNumbers, projectNumber);
                    
                    break;
            }

            
        }

        
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadEmailsFromAD
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                var activeDirectory = new ActiveDirectory();
                Console.WriteLine("Read mail groups from AD...");
                var allMailGroups = activeDirectory.GetAllMailGroups();
                Console.WriteLine("Total mail group count is: " + allMailGroups.Count);
                if (allMailGroups.Count == 0)
                {
                    Console.WriteLine("ADMailGroup is empty. Nothing to do.");
                    Console.WriteLine("END");
                    return;
                }

                Console.WriteLine("Read Members from all groups...");

                var members = new List<Member>();

                foreach (var mailGroup in allMailGroups)
                {
                    members.AddRange(activeDirectory.BuildGroupMembers(mailGroup));
                }

                Console.WriteLine("Total mail group count is: " + allMailGroups.Count);
                Console.WriteLine("Total member result count is: " + members.Count);

                Console.WriteLine("Begin write to redis");
                WriteToRedis(members);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private static void WriteToRedis(List<Member> members)
        {
            throw new NotImplementedException();
        }
    }
}

using NPOIProject.Model;
using System.Collections.Generic;

namespace NPOIProject
{
    class Program
    {
        static void Main(string[] args)
        {
            var user = new List<UserClass>()
            {
               new UserClass{ id=1,Name="Time" },
               new UserClass{ id=1,Name="Time2" },
               new UserClass{ id=1,Name="Time3" }
            };
            ListToNpoiHelper.ListToNpoi(user, @"C:\Users\Administrator\source\repos\ListToNpoi\NPOIProject\bin\Debug\user.xls", true);
        }
    }
}

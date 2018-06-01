using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Combine2Code
{
    public class Record
    {
        public long ID { get; set; }
        //Add additional fields here
        

        public static Record MasterFromCsv(string csvLine)
        {
            try
            {
                string[] values = csvLine.Split(',');
                Record record = new Record();
                record.ID = Convert.ToInt64(values[0]);
                //bind additional fields here
                return record;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        public static Record BPTFromCsv(string csvLine)
        {
            try
            {
                string[] values = csvLine.Split(',');
                Record record = new Record();
                record.ID = Convert.ToInt64(values[0]);
                //bind additional fields here
                return record;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }
    }
}

using System;
namespace AveryAddressLabelTemplate
{
    public class Mailing
    {

        public Mailing(string ownerName, string street, string city, string state, string zipCode)
        {
            OwnerName = ownerName;
            Street = street;
            City = city;
            State = state;
            ZipCode = zipCode;
        }

        public string OwnerName { get; private set; }
        public string Street { get; private set; }
        public string City { get; private set; }
        public string State { get; private set; }
        public string ZipCode { get; private set; }

        public override string ToString()
        {
            return string.Format("Mailing : OwnerName={0}, Street={1}, City={2}, State={3}, ZipCode={4}", OwnerName, Street, City, State, ZipCode);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using Thought.vCards;

namespace Toxy
{
    public class ToxyBusinessCards
    {
        private List<ToxyBusinessCard> cards;
        public bool IncludeEmptyContact { get; set; }
        public ToxyBusinessCards()
        {
            this.cards = new List<ToxyBusinessCard>();
            this.IncludeEmptyContact = true;
        }
        public List<ToxyBusinessCard> Cards
        {
            get
            {
                if (IncludeEmptyContact)
                    return this.cards;

                return this.cards.FindAll(delegate(ToxyBusinessCard s) { return s.Contacts.Count > 0; });
            }
        }
        public void SaveAs(string path)
        {
            if (this.cards == null || this.cards.Count == 0)
            {
                throw new InvalidDataException("No data to be saved");
            }
            vCardStandardWriter writer = new vCardStandardWriter();
            using (TextWriter tw = File.CreateText(path))
            {
                foreach (var card in this.cards)
                {
                    writer.Write(card.ToVCard(), tw);
                }
            }

        }
        public DataTable ToDataTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("First Name");
            dt.Columns.Add("Last Name");
            dt.Columns.Add("Full Name");
            dt.Columns.Add("Nick Name");
            dt.Columns.Add("Title");
            dt.Columns.Add("Organization");
            dt.Columns.Add("Contact 1");
            dt.Columns.Add("Contact 2");
            dt.Columns.Add("Contact 3");
            dt.Columns.Add("Contact 4");
            dt.Columns.Add("Contact 5");
            dt.Columns.Add("Contact 6");
            foreach (var card in this.Cards)
            {
                var row = dt.NewRow();
                int i = 0;
                if (card.Name != null)
                {
                    row[i++] = card.Name.FirstName;
                    row[i++] = card.Name.LastName;
                    row[i++] = card.Name.FullName;
                }
                else
                {
                    i += 3;
                }
                if (card.NickName != null)
                {
                    row[i++] = card.NickName.FullName;
                }
                else
                {
                    i++;
                }
                row[i++] = card.Title;
                row[i++] = card.Orgnization;
                foreach (var contact in card.Contacts)
                {
                    row[i++] = contact.ToString();
                }
                dt.Rows.Add(row);
            }
            return dt;
        }
    }
}

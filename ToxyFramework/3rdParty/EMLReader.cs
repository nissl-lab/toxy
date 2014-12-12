using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Diagnostics;
using Toxy;

namespace HLIB.MailFormats
{
    class EMLReader
    {
        private string _x_Sender;
        public string X_Sender
        {
            get { return _x_Sender; }
            set { _x_Sender = value; }
        }

        private string[] _x_Receivers;
        public string[] X_Receivers
        {
            get { return _x_Receivers; }
            set { _x_Receivers = value; }
        }

        private string _Received;
        public string Received
        {
            get { return _Received; }
            set { _Received = value; }
        }

        private string _Mime_Version;
        public string Mime_Version
        {
            get { return _Mime_Version; }
            set { _Mime_Version = value; }
        }

        private string _From;
        public string From
        {
            get { return _From; }
            set { _From = value; }
        }

        private string _To;
        public string To
        {
            get { return _To; }
            set { _To = value; }
        }

        private string _CC;
        public string CC
        {
            get { return _CC; }
            set { _CC = value; }
        }

        private DateTime _Date = DateTime.MinValue;
        public DateTime Date
        {
            get { return _Date; }
            set { _Date = value; }
        }

        private string _Subject;
        public string Subject
        {
            get { return _Subject; }
            set { _Subject = value; }
        }
        
        private string _Content_Type;
        public string Content_Type
        {
            get { return _Content_Type; }
            set { _Content_Type = value; }
        }

        private string _Content_Transfer_Encoding;
        public string Content_Transfer_Encoding
        {
            get { return _Content_Transfer_Encoding; }
            set { _Content_Transfer_Encoding = value; }
        }

        private string _Return_Path;
        public string Return_Path
        {
            get { return _Return_Path; }
            set { _Return_Path = value; }
        }

        private string _Message_ID;
        public string Message_ID
        {
            get { return _Message_ID; }
            set { _Message_ID = value; }
        }

        private DateTime _x_OriginalArrivalTime = DateTime.MinValue;
        public DateTime X_OriginalArrivalTime
        {
            get { return _x_OriginalArrivalTime; }
            set { _x_OriginalArrivalTime = value; }
        }

        private string _Body;
        public string Body
        {
            get { return _Body; }
            set { _Body = value; }
        }

        private string _HTMLBody;
        public string HTMLBody
        {
            get { return _HTMLBody; }
            set { _HTMLBody = value; }
        }

        private Dictionary<string, string> _listUnsupported = null;
        public Dictionary<string, string> UnsupportedHeaders
        {
            get { return _listUnsupported; }
        }

        public EMLReader(Stream fsEML)
        {
            ParseEML(fsEML);
        }

        private void ParseEML(Stream fsEML)
        {
            StreamReader sr = new StreamReader(fsEML);
            string sLine;
            List<string> listAll = new List<string>();
            while ((sLine = sr.ReadLine()) != null)
            {
                listAll.Add(sLine);
            }

            List<string> list = new List<string>();
            int nStartBody = -1;
            string[] saAll = new string[listAll.Count];
            listAll.CopyTo(saAll);

            for (int i = 0; i < saAll.Length; i++)
            {
                if (saAll[i] == string.Empty)
                {
                    nStartBody = i;
                    break;
                }

                string sFullValue = saAll[i];
                GetFullValue(saAll, ref i, ref sFullValue);
                list.Add(sFullValue);


                //Debug.WriteLine(sFullValue);
            }

            SetFields(list.ToArray());

            if (nStartBody == -1)   // no body ?
                return;

            // Get the body info out of saAll and set the Body and/or HTMLBody properties
            if (Content_Type != null && Content_Type.ToLower().Contains("multipart/alternative"))   // set for HTMLBody messages
            {
                int ix = Content_Type.ToLower().IndexOf("boundary");        // boundary is used to separate the different body types
                if (ix == -1)
                    return;

                string sBoundaryMarker = Content_Type.Substring(ix + 8).Trim(new char[] {'=', '"', ' ', '\t' });

                // save this boundaries elements into a list of strings
                list = new List<string>();
                for (int n = nStartBody + 1; n < saAll.Length; n++)
                {
                    if (saAll[n].Contains(sBoundaryMarker))
                    {
                        if (list.Count > 0)
                        {
                            SetBody(list);
                            list = new List<string>();
                        }
                        continue;
                    }

                    list.Add(saAll[n]);
                }
            }
            else    // plain text body type only
            {
                Body = string.Empty;
                for (int n = nStartBody + 1; n < saAll.Length; n++)
                    Body += saAll[n] + "\r\n";
            }
        }

        private void SetBody(List<string> list)
        {
            bool bIsHTML = false;
            bool bIsBodyStart = false;
            List<string> listBody = new List<string>();

            foreach (string s in list)
            {
                // use to determine type of body
                if (s.ToLower().StartsWith("content-type"))
                {
                    if (s.ToLower().Contains("text/html"))
                        bIsHTML = true;
                    else if (!s.ToLower().Contains("text/plain"))
                        return;
                }
                else if (s == string.Empty && !bIsBodyStart)
                {
                    bIsBodyStart = true;
                }
                else if (bIsBodyStart)
                {
                    listBody.Add(s);
                }
            }

            string[] sa = new string[listBody.Count];
            listBody.CopyTo(sa);

            if (bIsHTML)
                HTMLBody = string.Join("\r\n", sa);
            else
                Body = string.Join("\r\n", sa);
        }

        private void GetFullValue(string[] sa, ref int i, ref string sValue)
        {
            if (i + 1 < sa.Length && sa[i + 1] != string.Empty && char.IsWhiteSpace(sa[i + 1], 0))   // spec says line's that begin with white space are continuation lines
            {
                i++;
                sValue += " " + sa[i].Trim();

                GetFullValue(sa, ref i, ref sValue);
            }
        }

        private void SetFields(string[] saLines)
        {
            //List<string> listUnsupported = new List<string>();
            _listUnsupported = new Dictionary<string, string>();
            List<string> listX_Receiver = new List<string>();
            foreach (string sHdr in saLines)
            {
                string[] saHdr = Split(sHdr);
                if (saHdr == null)  // not a valid header
                    continue;

                switch (saHdr[0].ToLower())
                {
                    case "x-sender":
                        X_Sender = saHdr[1];
                        break;
                    case "x-receiver":
                        listX_Receiver.Add(saHdr[1]);
                        break;
                    case "received":
                        Received = saHdr[1];
                        break;
                    case "mime-version":
                        Mime_Version = saHdr[1];
                        break;
                    case "from":
                        From = saHdr[1];
                        break;
                    case "to":
                        To = saHdr[1];
                        break;
                    case "cc":
                        CC = saHdr[1];
                        break;
                    case "date":
                        //DateTime dt=DateTime.MinValue;
                        //if(DateTime.TryParse(saHdr[1],out dt))
                        //{
                        //    Date = dt;
                        //}
                        Date = DateTimeParser.Parse(saHdr[1]);
                        break;
                    case "subject":
                        Subject = saHdr[1];
                        break;
                    case "content-type":
                        Content_Type = saHdr[1];
                        break;
                    case "content-transfer-encoding":
                        Content_Transfer_Encoding = saHdr[1];
                        break;
                    case "return-path":
                        Return_Path = saHdr[1];
                        break;
                    case "message-id":
                        Message_ID = saHdr[1];
                        break;
                    case "x-originalarrivaltime":
                        int ix = saHdr[1].IndexOf("FILETIME");
                        if (ix != -1)
                        {
                            string sOAT = saHdr[1].Substring(0, ix);
                            sOAT = sOAT.Replace("(UTC)", "-0000");
                            X_OriginalArrivalTime = DateTime.Parse(sOAT);
                        }
                        break;
                    //case "body":
                    //    Body = saHdr[1];
                    //    break;
                    default:
                        _listUnsupported.Add(saHdr[0], saHdr[1]);
                        break;
                }
            }

            X_Receivers = new string[listX_Receiver.Count];
            listX_Receiver.CopyTo(X_Receivers);
        }

        private string[] Split(string sHeader)  // because string.Split won't work here...
        {
            int ix;
            if((ix = sHeader.IndexOf(':')) == -1)
                return null;

            return new string[] { sHeader.Substring(0, ix).Trim(), sHeader.Substring(ix + 1).Trim() };
        }
    }
}


#if false
"x-sender: jamieson@rl.gov\r\nx-receiver: c00ab5e1-2ab9-4502-ae27-0b00bd189089@develop.rl.gov\r\nReceived: from Develop ([130.97.35.161]) by Develop.rl.gov with Microsoft SMTPSVC(6.0.2600.5512);\r\n\t Fri, 12 Sep 2008 10:07:45 -0700\r\nMIME-Version: 1.0\r\nFrom: jamieson@rl.gov\r\nTo: c00ab5e1-2ab9-4502-ae27-0b00bd189089@develop.rl.gov\r\nDate: 12 Sep 2008 10:07:45 -0700\r\nSubject: A test subjectcd5ba966-5d29-45b0-80f8-fca89e0379aa\r\nContent-Type: text/plain; charset=us-ascii\r\nContent-Transfer-Encoding: quoted-printable\r\nReturn-Path: jamieson@rl.gov\r\nMessage-ID: <Developkl894nUTDd2h00000662@Develop.rl.gov>\r\nX-OriginalArrivalTime: 12 Sep 2008 17:07:45.0445 (UTC) FILETIME=[13452D50:01C914FA]\r\n\r\nThis is a test\r\n\r\n"


-------
x-sender: jamieson@rl.gov
x-receiver: c00ab5e1-2ab9-4502-ae27-0b00bd189089@develop.rl.gov
Received: from Develop ([130.97.35.161]) by Develop.rl.gov with Microsoft SMTPSVC(6.0.2600.5512);
	 Fri, 12 Sep 2008 10:07:45 -0700
MIME-Version: 1.0
From: jamieson@rl.gov
To: c00ab5e1-2ab9-4502-ae27-0b00bd189089@develop.rl.gov
Date: 12 Sep 2008 10:07:45 -0700
Subject: A test subjectcd5ba966-5d29-45b0-80f8-fca89e0379aa
Content-Type: text/plain; charset=us-ascii
Content-Transfer-Encoding: quoted-printable
Return-Path: jamieson@rl.gov
Message-ID: <Developkl894nUTDd2h00000662@Develop.rl.gov>
X-OriginalArrivalTime: 12 Sep 2008 17:07:45.0445 (UTC) FILETIME=[13452D50:01C914FA]

This is a test

-------

#endif

﻿//    This file is part of OleViewDotNet.
//    Copyright (C) James Forshaw 2014. 2016
//
//    OleViewDotNet is free software: you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation, either version 3 of the License, or
//    (at your option) any later version.
//
//    OleViewDotNet is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with OleViewDotNet.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace OleViewDotNet
{
    public class COMCategory : IXmlSerializable
    {
        public Guid CategoryID { get; private set; }
        public string Name { get; private set; }
        public IEnumerable<Guid> Clsids { get; private set; }

        public COMCategory(Guid catid, IEnumerable<Guid> clsids)
        {
            CategoryID = catid;
            Clsids = clsids.ToArray();
            Name = COMUtilities.GetCategoryName(catid);
        }

        internal COMCategory()
        {
        }

        XmlSchema IXmlSerializable.GetSchema()
        {
            return null;
        }

        void IXmlSerializable.ReadXml(XmlReader reader)
        {
            Name = reader.ReadString("name");
            CategoryID = reader.ReadGuid("catid");
            Clsids = reader.ReadGuids("clsids").ToArray();
        }

        void IXmlSerializable.WriteXml(XmlWriter writer)
        {
            writer.WriteAttributeString("name", Name);
            writer.WriteGuid("catid", CategoryID);
            writer.WriteGuids("clsids", Clsids);
        }

        public override bool Equals(object obj)
        {
            if (base.Equals(obj))
            {
                return true;
            }

            COMCategory right = obj as COMCategory;
            if (right == null)
            {
                return false;
            }

            return Clsids.SequenceEqual(right.Clsids) && CategoryID == right.CategoryID && Name == right.Name;
        }

        public override int GetHashCode()
        {
            return CategoryID.GetHashCode() ^ Name.GetSafeHashCode() 
                ^ Clsids.GetEnumHashCode();
        }
    }
}
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace Excell
{
    
    class Field
    {
        public string UID = "";
        public string ObjectType = "";
        public string FieldKey = "";
        public string FieldValue = "";

        public Field(string UID, string objectType, string fieldKey, string fieldValue)
        {
            this.UID = UID;
            this.ObjectType = objectType;
            this.FieldKey = fieldKey;
            this.FieldValue = fieldValue;
        }
        public void AddField()
        {           
            DataService.AddField(this.UID, this.ObjectType, this.FieldKey, this.FieldValue);            
        }

    }
}

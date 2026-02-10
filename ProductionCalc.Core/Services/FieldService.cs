using System.Collections.Generic;
using System.Linq;
using ProductionCalc.Core.Models;

namespace ProductionCalc.Core.Services
{
    public class FieldService
    {
        private readonly List<Field> _fields;

        public FieldService()
        {
            _fields = new List<Field>
            {
                new Field { Id = 1, Name = "Field A", Description = "Primary oil production field" },
                new Field { Id = 2, Name = "Field B", Description = "Gas monitoring site" },
                new Field { Id = 3, Name = "Field C", Description = "Experimental test site" }
            };
        }

        public List<Field> GetAllFields() => _fields;

        public void DeleteField(int id)
        {
            var field = _fields.FirstOrDefault(f => f.Id == id);
            if (field != null)
                _fields.Remove(field);
        }
    }
}

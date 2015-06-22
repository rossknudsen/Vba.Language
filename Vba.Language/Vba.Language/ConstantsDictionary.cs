using System;
using System.Collections.Generic;

using Vba.Grammars;

namespace Vba.Language
{
    internal class ConstantsDictionary
    {
        private IDictionary<string, Data> catalog = new Dictionary<string, Data>();
 
        internal bool Contains(string key)
        {
            return catalog.ContainsKey(key.ToLower());
        }

        internal void Add(PreprocessorParser.ConstantDeclarationContext statement)
        {
            var key = statement.ID().GetText();
            var value = statement.expression();  
            // need an expression evaluator here to determine the correct value of the expression.
            throw new NotImplementedException();
            Add(key, value);
        }

        internal void Add(string key, object value)
        {
            if (Contains(key))
            {
                // TODO raise syntax error.
            }
            else
            {
                catalog[key.ToLower()] = new Data(key, value);
            }
        }

        internal object GetValue(string key)
        {
            return catalog[key.ToLower()].Value;
        }

        private class Data
        {
            internal Data(string definedKey, object value)
            {
                DefinedKey = definedKey;
                Value = value;
            }

            internal string DefinedKey { get; private set; }

            internal object Value { get; private set; }
        }
    }
}

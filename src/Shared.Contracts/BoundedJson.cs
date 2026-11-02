using System;
using System.Globalization;
using System.IO;
using System.Text;
using Newtonsoft.Json.Linq;

namespace PowerPointIntelligence.Contracts
{
    // Json.NET accepts JavaScript extensions (single quotes, comments, trailing
    // commas). This bounded parser deliberately accepts the JSON grammar only.
    internal sealed class BoundedJson
    {
        private readonly string input; private readonly string prefix;
        private int position;
        private BoundedJson(string input, string prefix) { this.input = input; this.prefix = prefix; }
        private void Fail() { throw new InvalidDataException(prefix + ".invalid-json"); }
        private char Peek { get { return position < input.Length ? input[position] : '\0'; } }
        private void Space() { while (Peek == ' ' || Peek == '\t' || Peek == '\r' || Peek == '\n') position++; }
        internal static JToken Parse(string json, int maximumCharacters, string prefix)
        {
            if (String.IsNullOrEmpty(json) || json.Length > maximumCharacters) throw new InvalidDataException(prefix + ".json-size");
            var parser = new BoundedJson(json, prefix);
            var value = parser.Value(1); parser.Space(); if (parser.position != json.Length) parser.Fail();
            return value;
        }
        private JToken Value(int depth)
        {
            Space(); if (depth > 32 || position >= input.Length) Fail();
            char c = Peek;
            if (c == '"') return new JValue(StringValue());
            if (c == '{')
            {
                position++; Space(); var result = new JObject();
                if (Peek == '}') { position++; return result; }
                while (position < input.Length)
                {
                    if (Peek != '"') Fail();
                    string key = StringValue(); Space();
                    if (Peek != ':' || result.Property(key, StringComparison.Ordinal) != null) Fail();
                    position++; result.Add(key, Value(depth + 1)); Space();
                    c = Peek; position++; if (c == '}') return result;
                    if (c != ',') Fail(); Space();
                }
                Fail();
            }
            if (c == '[')
            {
                position++; Space(); var result = new JArray();
                if (Peek == ']') { position++; return result; }
                while (position < input.Length)
                {
                    result.Add(Value(depth + 1)); Space(); c = Peek; position++;
                    if (c == ']') return result; if (c != ',') Fail();
                }
                Fail();
            }
            foreach (string literal in new[] { "true", "false", "null" })
                if (position + literal.Length <= input.Length && String.CompareOrdinal(input, position, literal, 0, literal.Length) == 0)
                { position += literal.Length; return literal == "null" ? JValue.CreateNull() : new JValue(literal == "true"); }
            return Number();
        }
        private static bool Digit(char c) { return c >= '0' && c <= '9'; }
        private JToken Number()
        {
            int start = position;
            if (Peek == '-') position++;
            if (Peek == '0') position++;
            else { if (Peek < '1' || Peek > '9') Fail(); while (Digit(Peek)) position++; }
            if (Peek == '.') { position++; if (!Digit(Peek)) Fail(); while (Digit(Peek)) position++; }
            if (Peek == 'e' || Peek == 'E')
            { position++; if (Peek == '+' || Peek == '-') position++; if (!Digit(Peek)) Fail(); while (Digit(Peek)) position++; }
            double number;
            if (!Double.TryParse(input.Substring(start, position - start), NumberStyles.Float, CultureInfo.InvariantCulture, out number) || Double.IsNaN(number) || Double.IsInfinity(number)) Fail();
            return new JValue(number);
        }
        private string StringValue()
        {
            position++; var text = new StringBuilder();
            while (position < input.Length)
            {
                char c = input[position++];
                if (c == '"')
                { string result = text.ToString(); if (!ContractText.Valid(result, Int32.MaxValue)) Fail(); return result; }
                if (c < 32) Fail();
                if (c != '\\') { text.Append(c); continue; }
                if (position >= input.Length) Fail();
                c = input[position++];
                switch (c)
                {
                    case '"': case '\\': case '/': text.Append(c); break;
                    case 'b': text.Append('\b'); break; case 'f': text.Append('\f'); break;
                    case 'n': text.Append('\n'); break; case 'r': text.Append('\r'); break; case 't': text.Append('\t'); break;
                    case 'u':
                        int code = 0;
                        for (int i = 0; i < 4; i++)
                        {
                            if (position >= input.Length) Fail(); c = input[position++];
                            int digit = c >= '0' && c <= '9' ? c - '0' : c >= 'a' && c <= 'f' ? c - 'a' + 10 : c >= 'A' && c <= 'F' ? c - 'A' + 10 : -1;
                            if (digit < 0) Fail(); code = code * 16 + digit;
                        }
                        text.Append((char)code); break;
                    default: Fail(); break;
                }
            }
            Fail(); return null;
        }
    }
    internal static class ContractText
    {
        internal static bool Valid(string value, int maxCodePoints)
        {
            if (value == null) return false; int count = 0;
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                if (Char.IsHighSurrogate(c)) { if (++i >= value.Length || !Char.IsLowSurrogate(value[i])) return false; }
                else if (Char.IsLowSurrogate(c)) return false;
                if (++count > maxCodePoints) return false;
            }
            return true;
        }
        internal static bool NonWhitespace(string value)
        {
            // Match ECMAScript \S used by the JSON schema and TypeScript validator.
            foreach (char c in value)
                if (!(c >= '\t' && c <= '\r') && c != ' ' && c != '\u00a0' && c != '\u1680' && !(c >= '\u2000' && c <= '\u200a') && c != '\u2028' && c != '\u2029' && c != '\u202f' && c != '\u205f' && c != '\u3000' && c != '\ufeff') return true;
            return false;
        }
    }
}

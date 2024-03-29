﻿// using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace UTDataValidator
{
    public class JsonDataValidator
    {
        public static IAssertion DefaultAssertion { get; set; }
        
        private readonly string _expected;
        private readonly IAssertion _assertion;
        private readonly List<string> _skipProperties;
        
        public JsonDataValidator(IAssertion assertion, FileInfo fileInfo)
        {
            _assertion = assertion;
            if (!File.Exists(fileInfo.FullName))
            {
                throw new Exception($"JSON file expected on path = '{fileInfo.FullName}' not found.");
            }

            using (FileStream fs = new FileStream(fileInfo.FullName, FileMode.Open, FileAccess.Read))
            {
                using (StreamReader sw = new StreamReader(fs))
                {
                    _expected = sw.ReadToEnd();
                }
            }
        }

        public JsonDataValidator(IAssertion assertion, string expectedJsonString)
        {
            _expected = expectedJsonString;
            _assertion = assertion;
        }
        
        public JsonDataValidator(FileInfo fileInfo)
        {
            _assertion = DefaultAssertion;
            if (!File.Exists(fileInfo.FullName))
            {
                throw new Exception($"JSON file expected on path = '{fileInfo.FullName}' not found.");
            }

            using (FileStream fs = new FileStream(fileInfo.FullName, FileMode.Open, FileAccess.Read))
            {
                using (StreamReader sw = new StreamReader(fs))
                {
                    _expected = sw.ReadToEnd();
                }
            }
        }
        
        public JsonDataValidator(FileInfo fileInfo, List<string> skipProperties)
        {
            _skipProperties = skipProperties;
            _assertion = DefaultAssertion;
            if (!File.Exists(fileInfo.FullName))
            {
                throw new Exception($"JSON file expected on path = '{fileInfo.FullName}' not found.");
            }

            using (FileStream fs = new FileStream(fileInfo.FullName, FileMode.Open, FileAccess.Read))
            {
                using (StreamReader sw = new StreamReader(fs))
                {
                    _expected = sw.ReadToEnd();
                }
            }
        }
        
        public JsonDataValidator(IAssertion assertion, FileInfo fileInfo, List<string> skipProperties)
        {
            _skipProperties = skipProperties;
            _assertion = assertion;
            if (!File.Exists(fileInfo.FullName))
            {
                throw new Exception($"JSON file expected on path = '{fileInfo.FullName}' not found.");
            }

            using (FileStream fs = new FileStream(fileInfo.FullName, FileMode.Open, FileAccess.Read))
            {
                using (StreamReader sw = new StreamReader(fs))
                {
                    _expected = sw.ReadToEnd();
                }
            }
        }

        public JsonDataValidator(string expectedJsonString)
        {
            _expected = expectedJsonString;
            _assertion = DefaultAssertion;
        }
        
        public JsonDataValidator(string expectedJsonString, List<string> skipProperties)
        {
            _expected = expectedJsonString;
            _assertion = DefaultAssertion;
            _skipProperties = skipProperties;
        }

        public JsonDataValidator(IAssertion assertion, string expectedJsonString, List<string> skipProperties)
        {
            _assertion = assertion;
            _expected = expectedJsonString;
            _skipProperties = skipProperties;
        }

        public void ValidateData(string actual)
        {
            ValidateJSON(_expected, actual);
        }

        private void ValidateJSON(string expected, string actual)
        {
            Dictionary<string, object> expectedJSON = JsonConvert.DeserializeObject<Dictionary<string, object>>(expected);
            Dictionary<string, object> actualJSON = JsonConvert.DeserializeObject<Dictionary<string, object>>(actual);
            ValidateJsonDictionary(expectedJSON, actualJSON, "");
        }

        private bool IsSkipValidate(string node)
        {
            if (_skipProperties == null || _skipProperties.Count == 0) return false;
            
            if (_skipProperties.Contains(node)) return true;
            
            var regexList = new Regex(@"(?:\[[0-9]*\])");
            var nodeCleanString = regexList.Replace(node, "[]");
            foreach (var skipProperty in _skipProperties)
            {
                var cleanString = regexList.Replace(skipProperty, "[]");
                if (nodeCleanString == cleanString)
                {
                    return true;
                }
            }
            
            return false;
        }

        private void ValidateJsonDictionary(Dictionary<string, object> expected, Dictionary<string, object> actual, string node)
        {
            foreach (KeyValuePair<string, object> keyValue in expected)
            {
                string nodeInfo = string.IsNullOrEmpty(node) ? keyValue.Key : $"{node}.{keyValue.Key}";
                if (IsSkipValidate(nodeInfo))
                {
                    continue;
                }
                
                _assertion.IsTrue(actual.ContainsKey(keyValue.Key), message: $"Node '{nodeInfo}' has different node between expected and actual, actual doesn't have '{nodeInfo}'.");
                _assertion.AreEqual(IsData(keyValue.Value), IsData(actual[keyValue.Key]), message: $"Node '{nodeInfo}' has different value between expected and actual.");
                _assertion.AreEqual(IsJsonDictionary(keyValue.Value), IsJsonDictionary(actual[keyValue.Key]), message: $"Node '{nodeInfo}' has different value between expected and actual.");
                _assertion.AreEqual(IsJsonList(keyValue.Value), IsJsonList(actual[keyValue.Key]), message: $"Node '{nodeInfo}' has different value between expected and actual.");
                _assertion.AreEqual(keyValue.Value == null, actual[keyValue.Key] == null, message: $"Node '{nodeInfo}' has different value between expected and actual.");
                _assertion.AreEqual(keyValue.Value?.ToString() == "", actual[keyValue.Key]?.ToString() == "", message: $"Node '{nodeInfo}' has different value between expected and actual.");

                if (keyValue.Value == null || keyValue.Value.ToString() == "")
                {
                    continue;
                }

                if (IsData(keyValue.Value))
                {
                    string expectedString = JsonObjectToString(keyValue.Value);
                    string actualString = JsonObjectToString(actual[keyValue.Key]);
                    _assertion.AreEqual(expectedString, actualString, message: $"Node '{nodeInfo}' has different value.");
                }

                if (IsJsonDictionary(keyValue.Value))
                {
                    Dictionary<string, object> expectedDict = JsonObjectToDictionary(keyValue.Value);
                    Dictionary<string, object> actualDict = JsonObjectToDictionary(actual[keyValue.Key]);

                    ValidateJsonDictionary(expectedDict, actualDict, nodeInfo);
                }

                if (IsJsonList(keyValue.Value))
                {
                    List<object> expectedList = JsonObjectToList(keyValue.Value);
                    List<object> actualList = JsonObjectToList(actual[keyValue.Key]);

                    ValidateJsonList(expectedList, actualList, nodeInfo);
                }
            }
        }

        private void ValidateJsonList(List<object> expected, List<object> actual, string node)
        {
            _assertion.AreEqual(expected.Count, actual.Count, $"Node {node} has different list count.");
            for (int i = 0; i < expected.Count; i++)
            {
                string nodeInfo = $"{node}[{i}]";

                object expectedObj = expected[i];
                object actualObj = actual[i];

                _assertion.AreEqual(IsData(expectedObj), IsData(actualObj), message: $"Node '{nodeInfo}' has different value between expected and actual.");
                _assertion.AreEqual(IsJsonDictionary(expectedObj), IsJsonDictionary(actualObj), message: $"Node '{nodeInfo}' has different value between expected and actual.");
                _assertion.AreEqual(IsJsonList(expectedObj), IsJsonList(actualObj), message: $"Node '{nodeInfo}' has different value between expected and actual.");
                _assertion.AreEqual(expectedObj == null, actualObj == null, message: $"Node '{nodeInfo}' has different value between expected and actual.");
                _assertion.AreEqual(expectedObj?.ToString() == "", actualObj?.ToString() == "", message: $"Node '{nodeInfo}' has different value between expected and actual.");

                if (expectedObj == null || expectedObj.ToString() == "")
                {
                    continue;
                }

                if (IsData(expectedObj))
                {
                    string expectedString = JsonObjectToString(expectedObj);
                    string actualString = JsonObjectToString(actualObj);
                    _assertion.AreEqual(expectedString, actualString, message: $"Node '{nodeInfo}' has different value between expected and actual.");
                }

                if (IsJsonDictionary(expectedObj))
                {
                    Dictionary<string, object> expectedDict = JsonObjectToDictionary(expectedObj);
                    Dictionary<string, object> actualDict = JsonObjectToDictionary(actualObj);

                    ValidateJsonDictionary(expectedDict, actualDict, nodeInfo);
                }

                if (IsJsonList(expectedObj))
                {
                    List<object> expectedList = JsonObjectToList(expectedObj);
                    List<object> actualList = JsonObjectToList(actualObj);

                    ValidateJsonList(expectedList, actualList, nodeInfo);
                }
            }
        }

        private bool IsJsonDictionary(object value)
        {
            string val = JsonConvert.SerializeObject(value);
            try
            {
                JsonConvert.DeserializeObject<Dictionary<string, object>>(val);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool IsJsonList(object value)
        {
            string val = JsonConvert.SerializeObject(value);
            try
            {
                JsonConvert.DeserializeObject<List<object>>(val);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool IsData(object value)
        {
            if (IsJsonDictionary(value))
            {
                return false;
            }

            if (IsJsonList(value))
            {
                return false;
            }

            return true;
        }

        private string JsonObjectToString(object value)
        {
            string data = JsonConvert.SerializeObject(value);
            return JsonConvert.DeserializeObject<string>(data);
        }

        private List<object> JsonObjectToList(object value)
        {
            string data = JsonConvert.SerializeObject(value);
            return JsonConvert.DeserializeObject<List<object>>(data);
        }
        
        private Dictionary<string, object> JsonObjectToDictionary(object value)
        {
            string data = JsonConvert.SerializeObject(value);
            return JsonConvert.DeserializeObject<Dictionary<string, object>>(data);
        }
    }
}

using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CaesarCipherTest
{
    [TestClass]
    public class CaesarCipherTest
    {
        void TestEncrypt(string text, int key, string encryptText)
        {
            CaesarCipher.Form1.SetKey(key);
            if (key < 0)
                CaesarCipher.Form1.setIsRight(false);
            var result = CaesarCipher.Form1.Encrypt(text);
            Assert.AreEqual(encryptText, result);
        }

        void TestDecrypt(string text, int key, string decryptText)
        {
            CaesarCipher.Form1.SetKey(key);
            if (key < 0)
                CaesarCipher.Form1.setIsRight(false);
            var result = CaesarCipher.Form1.Decrypt(text);
            Assert.AreEqual(decryptText, result);
        }

        [TestMethod]
        public void DecryptNumbersForRightShift()
        {
            TestDecrypt("0123456789", 1, "9012345678");
            TestDecrypt("0123456789", 2, "8901234567");
            TestDecrypt("0123456789", 3, "7890123456");
            TestDecrypt("0123456789", 4, "6789012345");
            TestDecrypt("0123456789", 5, "5678901234");
        }

        [TestMethod]
        public void DecryptNumbersForLeftShift()
        {
            TestDecrypt("0123456789", -1, "1234567890");
            TestDecrypt("0123456789", -2, "2345678901");
            TestDecrypt("0123456789", -3, "3456789012");
            TestDecrypt("0123456789", -4, "4567890123");
            TestDecrypt("0123456789", -5, "5678901234");
        }

        [TestMethod]
        public void EncryptNumbersForRightShift()
        {
            TestEncrypt("0123456789", 1, "1234567890");
            TestEncrypt("0123456789", 2, "2345678901");
            TestEncrypt("0123456789", 3, "3456789012");
            TestEncrypt("0123456789", 4, "4567890123");
            TestEncrypt("0123456789", 5, "5678901234");
        }

        [TestMethod]
        public void EncryptNumbersForLeftShift()
        {
            TestEncrypt("0123456789", -1, "9012345678");
            TestEncrypt("0123456789", -2, "8901234567");
            TestEncrypt("0123456789", -3, "7890123456");
            TestEncrypt("0123456789", -4, "6789012345");
            TestEncrypt("0123456789", -5, "5678901234");
        }

        [TestMethod]
        public void DecryptRussianLetters()
        {
            var text = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя";
            TestDecrypt(text, 1, "ЯАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮяабвгдеёжзийклмнопрстуфхцчшщъыьэю");
            TestDecrypt(text, 2, "ЮЯАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭюяабвгдеёжзийклмнопрстуфхцчшщъыьэ");
            TestDecrypt(text, 3, "ЭЮЯАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬэюяабвгдеёжзийклмнопрстуфхцчшщъыь");
            TestDecrypt(text, -1, "БВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯАбвгдеёжзийклмнопрстуфхцчшщъыьэюяа");
            TestDecrypt(text, -2, "ВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯАБвгдеёжзийклмнопрстуфхцчшщъыьэюяаб");
            TestDecrypt(text, -3, "ГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯАБВгдеёжзийклмнопрстуфхцчшщъыьэюяабв");
        }

        [TestMethod]
        public void EncryptRussianLetters()
        {
            var text = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя";
            TestEncrypt(text, 1, "БВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯАбвгдеёжзийклмнопрстуфхцчшщъыьэюяа");
            TestEncrypt(text, 2, "ВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯАБвгдеёжзийклмнопрстуфхцчшщъыьэюяаб");
            TestEncrypt(text, 3, "ГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯАБВгдеёжзийклмнопрстуфхцчшщъыьэюяабв");
            TestEncrypt(text, -1, "ЯАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮяабвгдеёжзийклмнопрстуфхцчшщъыьэю");
            TestEncrypt(text, -2, "ЮЯАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭюяабвгдеёжзийклмнопрстуфхцчшщъыьэ");
            TestEncrypt(text, -3, "ЭЮЯАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬэюяабвгдеёжзийклмнопрстуфхцчшщъыь");
        }

        [TestMethod]
        public void DecryptEnglishLetters()
        {
            var text = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
            TestDecrypt(text, 2, text);
            TestDecrypt(text, -2, text);
        }

        [TestMethod]
        public void EncryptEnglishLetters()
        {
            var text = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
            TestEncrypt(text, 2, text);
            TestEncrypt(text, -2, text);
        }

        [TestMethod]
        public void DecryptOtherSymbols()
        {
            var text = "~`@#$%^&*()_+-={}[]:;\"\'<,>.?/\\|!№";
            TestDecrypt(text, 2, text);
            TestDecrypt(text, -2, text);
        }

        [TestMethod]
        public void EncryptOtherSymbols()
        {
            var text = "~`@#$%^&*()_+-={}[]:;\"\'<,>.?/\\|!№";
            TestEncrypt(text, 2, text);
            TestEncrypt(text, -2, text);
        }

        [TestMethod]
        public void DecryptForZeroKey()
        {
            var text = "Hello, everyone!";
            TestDecrypt(text, 0, text);
        }

        [TestMethod]
        public void EncryptForZeroKey()
        {
            var text = "Hello, everyone!";
            TestEncrypt(text, 0, text);
        }
    }
}

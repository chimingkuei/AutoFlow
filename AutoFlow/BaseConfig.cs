using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows;

namespace AutoFlow
{
    public class BaseConfig<T>
    {
        
        private string publickey = "<RSAKeyValue><Modulus>5m9m14XH3oqLJ8bNGw9e4rGpXpcktv9MSkHSVFVMjHbfv+SJ5v0ubqQxa5YjLN4vc49z7SVju8s0X4gZ6AzZTn06jzWOgyPRV54Q4I0DCYadWW4Ze3e+BOtwgVU1Og3qHKn8vygoj40J6U85Z/PTJu3hN1m75Zr195ju7g9v4Hk=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";
        private string privatekey = "<RSAKeyValue><Modulus>5m9m14XH3oqLJ8bNGw9e4rGpXpcktv9MSkHSVFVMjHbfv+SJ5v0ubqQxa5YjLN4vc49z7SVju8s0X4gZ6AzZTn06jzWOgyPRV54Q4I0DCYadWW4Ze3e+BOtwgVU1Og3qHKn8vygoj40J6U85Z/PTJu3hN1m75Zr195ju7g9v4Hk=</Modulus><Exponent>AQAB</Exponent><P>/hf2dnK7rNfl3lbqghWcpFdu778hUpIEBixCDL5WiBtpkZdpSw90aERmHJYaW2RGvGRi6zSftLh00KHsPcNUMw==</P><Q>6Cn/jOLrPapDTEp1Fkq+uz++1Do0eeX7HYqi9rY29CqShzCeI7LEYOoSwYuAJ3xA/DuCdQENPSoJ9KFbO4Wsow==</Q><DP>ga1rHIJro8e/yhxjrKYo/nqc5ICQGhrpMNlPkD9n3CjZVPOISkWF7FzUHEzDANeJfkZhcZa21z24aG3rKo5Qnw==</DP><DQ>MNGsCB8rYlMsRZ2ek2pyQwO7h/sZT8y5ilO9wu08Dwnot/7UMiOEQfDWstY3w5XQQHnvC9WFyCfP4h4QBissyw==</DQ><InverseQ>EG02S7SADhH1EVT9DD0Z62Y0uY7gIYvxX/uq+IzKSCwB8M2G7Qv9xgZQaQlLpCaeKbux3Y59hHM+KpamGL19Kg==</InverseQ><D>vmaYHEbPAgOJvaEXQl+t8DQKFT1fudEysTy31LTyXjGu6XiltXXHUuZaa2IPyHgBz0Nd7znwsW/S44iql0Fen1kzKioEL3svANui63O3o5xdDeExVM6zOf1wUUh/oldovPweChyoAdMtUzgvCbJk1sYDJf++Nr0FeNW1RB1XG30=</D></RSAKeyValue>";
        public string Path { get; set; }
      
        public BaseConfig(): this(@"Config.json")
        { 
        }

        public BaseConfig(string _Path)
        {
            Path = _Path;
        }

        ~BaseConfig()
        { 
        }

        public void Save(List<T> record, bool encryption=false)
        {
            string jsonData = JsonConvert.SerializeObject(record);
            if (encryption)
            {
                string jsonData_RSAEncrypt = RSAEncrypt(jsonData);
                File.WriteAllText(Path, jsonData_RSAEncrypt);
            }
            else
            {
                File.WriteAllText(Path, jsonData);
            }
            
        }

        // 多組同參數儲存
        public void Save(List<T> record, int replace_index, bool encryption = false)
        {
            string jsonData = JsonConvert.SerializeObject(record);
            string[] contentarray = (File.ReadAllText(Path).Trim('[').Trim(']')).Split(new string[] { "}," }, StringSplitOptions.None);
            contentarray[replace_index] = jsonData.Trim('[').Trim(']').Trim('}');
            string content = null;
            if (replace_index == contentarray.Length - 1)
            {
                content = "[" + (string.Join("},", contentarray) + "}") + "]";
            }
            else
            {
                content = "[" + string.Join("},", contentarray) + "]";
            }
            if (encryption)
            {
                string jsonData_RSAEncrypt = RSAEncrypt(content);
                File.WriteAllText(Path, jsonData_RSAEncrypt);
            }
            else
            {
                File.WriteAllText(Path, content);
            }
        }

        public List<T> Load(bool encryption = false)
        {
            List<T> jsonData = null;
            if(File.Exists(Path))
            {
                string Record = File.ReadAllText(Path);
                if (encryption)
                {
                    string Record_RSADecrypt = RSADecrypt(Record);
                    jsonData = JsonConvert.DeserializeObject<List<T>>(Record_RSADecrypt);

                }
                else
                {
                    jsonData = JsonConvert.DeserializeObject<List<T>>(Record);
                }
            }
            else
            {
                string jsonEmpty = "[{}]";
                File.WriteAllText(Path, jsonEmpty);
            }
            return jsonData;
        }

        // 多組同參數Load一組
        public List<T> Load(int group_num, int index, bool encryption = false)
        {
            List<T> jsonData = null;
            if (File.Exists(Path))
            {
                string Record = null;
                string[] contentarray = (File.ReadAllText(Path).Trim('[').Trim(']')).Split(new string[] { "}," }, StringSplitOptions.None);
                if (index == contentarray.Length - 1)
                {
                    Record = "[" + contentarray[index] + "]";
                }
                else
                {
                    Record = "[" + contentarray[index] + "}]";
                }
                if (encryption)
                {
                    string Record_RSADecrypt = RSADecrypt(Record);
                    jsonData = JsonConvert.DeserializeObject<List<T>>(Record_RSADecrypt);

                }
                else
                {
                    jsonData = JsonConvert.DeserializeObject<List<T>>(Record);
                }
            }
            else
            {
                string jsonEmpty = "[" + string.Join(",", Enumerable.Repeat("{}", group_num)) + "]";
                File.WriteAllText(Path, jsonEmpty);
            }
            return jsonData;
        }

        // 加密
        private string RSAEncrypt(string content)
        {
            RSACryptoServiceProvider rsa = new RSACryptoServiceProvider();
            byte[] cipherbytes;
            rsa.FromXmlString(publickey);
            cipherbytes = rsa.Encrypt(Encoding.UTF8.GetBytes(content), false);
            return Convert.ToBase64String(cipherbytes);
        }

        // 解密
        private string RSADecrypt(string content)
        {
            RSACryptoServiceProvider rsa = new RSACryptoServiceProvider();
            byte[] cipherbytes;
            rsa.FromXmlString(privatekey);
            cipherbytes = rsa.Decrypt(Convert.FromBase64String(content), false);
            return Encoding.UTF8.GetString(cipherbytes);
        }

    }



}

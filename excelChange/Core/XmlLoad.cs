using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Deployment.Application;
using System.Windows;
using System.Collections.ObjectModel;
using excelChange.Entity;
using System.Xml.Serialization;
using System.IO;

namespace excelChange.Core
{
    public class XmlLoad : Window
    {
        private static XmlLoad _inst = null;
   
        public XmlLoad()
        {

        }
        //싱글톤패턴
        public static XmlLoad make()
        {
            if (_inst == null)
            {
                _inst = new XmlLoad();
                return _inst;
            }
            return _inst;
        }



        /************************************************************************************
        * 함  수  명      : GetBasicSetting
        * 내      용      : XML에서 리스트(컬렉션)형태로 값을 불러옵니다
        * 설      명      : 
        ************************************************************************************/
        public ObservableCollection<T> GetBasicSetting<T>(string fileName)
        {
            string file_path = this.GetXmlFilePath(fileName);
            if (!File.Exists(file_path))
            {
                return null;
            }
            XmlSerializer xs = new XmlSerializer(typeof(ObservableCollection<T>));

            ObservableCollection<T> BasicSetting = null;

            using (StreamReader rd = new StreamReader(file_path))
            {
                BasicSetting = xs.Deserialize(rd) as ObservableCollection<T>;
            }

            if (BasicSetting == null || BasicSetting.Count == 0)
            {
                return null;
            }
            return BasicSetting;

        }

        /************************************************************************************
        * 함  수  명      : SaveBasicSetting
        * 내      용      : 리스트형태로있을때 경우 사용하는 함수로 XML에 값을 저장합니다.
        * 설      명      : 
        ************************************************************************************/
        public bool SaveBasicSetting<T>(ObservableCollection<T> SaveSetting, string fileName)
        {
            try
            {
                string filePath = this.GetXmlFilePath(fileName);
                //Console.WriteLine(filePath);

                XmlSerializer xs = new XmlSerializer(typeof(ObservableCollection<T>));
                using (StreamWriter wr = new StreamWriter(filePath))
                {
                    xs.Serialize(wr, SaveSetting);
                    wr.Close();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        /************************************************************************************
        * 함  수  명      : GetXmlFilePath
        * 내      용      : XML의 파일이름을 받고 XML폴더에서 경로를 가져옵니다
        * 설      명      : 
        ************************************************************************************/
        public string GetXmlFilePath(string fileName)
        {
            string file_path = "";

            if (ApplicationDeployment.IsNetworkDeployed)
            {
                var deployment = ApplicationDeployment.CurrentDeployment;
                file_path = System.IO.Path.Combine(deployment.DataDirectory, "XML", fileName);
            }
            else
            {
                // 해당 문제점은 빌드 장소에 파일이 있어야한다는 점인데 이걸로 일시 해결
                //post-build
                //xcopy /s /y "$(ProjectDir)XML\*" "$(TargetDir)Config\"
                file_path = string.Format(@".\XML\" + fileName);

            }

            return file_path;
        }

        /************************************************************************************
        * 함  수  명      : SaveObjectToXml
        * 내      용      : 단일개체인경우 저장하는 함수
        * 설      명      : 파일이름을 받아 저장해줍니다.
        ************************************************************************************/
        public void SaveObjectToXml<T>(T obj, string fileName)
        {
            try
            {
                string filePath = this.GetXmlFilePath(fileName);
                XmlSerializer serializer = new XmlSerializer(typeof(T));
                using (StreamWriter writer = new StreamWriter(filePath))
                {
                    serializer.Serialize(writer, obj);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving XML: {ex.Message}");
            }
        }

        /************************************************************************************
        * 함  수  명      : LoadObjectFromXml
        * 내      용      : 파일이름을 받아서 리스트형식이아닌 단수타입인경우 값을 불러와줍니다.
        * 설      명      : XML에서 단일개체인경우에는 값을 반환해줍니다.
        ************************************************************************************/
        public T LoadObjectFromXml<T>(string fileName)
        {
            try
            {
                string filePath = this.GetXmlFilePath(fileName);

                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException("XML 파일이 존재하지 않습니다.", filePath);
                }

                XmlSerializer serializer = new XmlSerializer(typeof(T));
                using (StreamReader reader = new StreamReader(filePath))
                {
                    return (T)serializer.Deserialize(reader);
                }
            }
            catch (FileNotFoundException fnfEx)
            {
                MessageBox.Show($"파일 오류: {fnfEx.Message}");
                return default(T);
            }
            catch (InvalidOperationException invEx)
            {
                MessageBox.Show($"XML 형식 오류: {invEx.Message}");
                return default(T);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"오류 발생: {ex.Message}");
                return default(T);
            }
        }




    }
}

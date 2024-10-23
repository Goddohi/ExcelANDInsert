using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System;
using System.Data;
using System.IO;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;

namespace excelChange
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 전역변수
        /// </summary>

        // Quertlist는 만든 쿼리를 어디서든 부를수있게
        List<string> Querylist;

        /// <summary>
        /// Xnames에 있는 이름은 겹치는 무조건 해당값(X)에 걸러준다
        /// </summary>
        List<string> Cnames = new List<string> {  };
        List<string> Bnames = new List<string> {  };
        List<string> Anames = new List<string> {  };
        List<string> Zero_names = new List<string> {"open","OPEN"};


        /// <summary>
        /// Only_Xnames에 있는 이름은 아예동일한경우 무조건 해당값(X)에 걸러준다
        /// </summary>
        List<string> Only_Cnames = new List<string> { };
        List<string> Only_Bnames = new List<string> { };
        List<string> Only_Anames = new List<string> { };
        List<string> Only_Zero_names = new List<string> { };

        //저장할 테이블이름
        string tablename = "텟텟";


        /// <summary>
        /// 로직
        /// </summary>
     

        // 파일 열기 대화상자 생성하는 버튼 로직
        private void btnLoadExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenExcelFile();
        }

        //분리한 이유는 매번 새로운 인스턴를 생성하기 위해서 
        //(이프로그램은 특정한 규격의 엑셀만 돌아가게 설계)
        private void OpenExcelFile()
        {
            // 매번 새로운 OpenFileDialog 인스턴스 생성
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "엑셀파일 선택"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                LoadExcelFile(openFileDialog.FileName);
                btnQureyExcel.Visibility = Visibility.Visible;
            }
        }

        /*
         * 
         * 엑셀을 데이터 그리드로 변경하고 그리고나서 데이터를 String 2차배열으로 변환한다.
         * 
        */
        private void LoadExcelFile(string filePath)
        {
            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(stream);
                    ISheet sheet = workbook.GetSheetAt(0);

                    // 데이터 테이블 생성
                    var dataTable = new DataTable();

                    // 열 추가 ↓
                    IRow headerRow = sheet.GetRow(1);
                    for (int i = 0; i < headerRow.LastCellNum; i++)
                    {
                        dataTable.Columns.Add(headerRow.GetCell(i).ToString());
                    }

                    // 행 추가 i=1로 할경우 맨 처음 행이 사라짐 (컬럼으로 들어가있는것) ->
                    for (int i = 0; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row != null)
                        {
                            var newRow = dataTable.NewRow();
                            for (int j = 0; j < row.LastCellNum; j++)
                            {
                                newRow[j] = row.GetCell(j)?.ToString() ?? string.Empty; // null 체크
                            }
                            dataTable.Rows.Add(newRow);
                        }
                    }
                
          
                // 데이터 그리드에 데이터 바인딩
                dataGrid.ItemsSource = dataTable.DefaultView;

                // 2차원 배열로 변환
                string[,] dataArray = new string[dataTable.Rows.Count, dataTable.Columns.Count];

                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int column = 0; column < dataTable.Columns.Count; column++)
                    {
                        dataArray[row, column] = dataTable.Rows[row][column].ToString();
                    }
                }



                Querylist = excelToInsert(dataArray);
                /*
                foreach (string str in Querylist)
                {
                    Console.WriteLine(str);
                }
                */
                }
            }
            catch (Exception ex)
            {
                // 예외 처리
                MessageBox.Show($"오류 발생: {ex.Message}");
            }
           
       

        }

        /**
         * 2차배열을 받아서 List로 Insert해주는 것을 만들기 위한 로직
         * 즉, 커스텀을 하시면 될 듯합니다.
         * [↓,→]
         */
        public List<string> excelToInsert(string[,] exceldata) {
            //  커스텀해드림   2(m)  4*2*5  /,2(r1) r15
            bool count = false;
            List<string> result = new List<string>();
            for (int row=2;row < 2+40; row+=4)
            {
                int day = (row / 8) +1;

                if (count) { result.Add(Environment.NewLine + "======" + day.ToString() + "AM" + "======"); }
                else { result.Add(Environment.NewLine+"======" + day.ToString() + "PM" + "======"); }
            
                // 2 - 1(10없음) + 15 (마지막뽑아낼 rm)
                for (int column = 2; column < 2 - 1 + 15; column++)
                {
                    string data = "";
                    for (int i = 0; i < 4; i++)
                    {   //하루에 4개의 파트가 있어 4파트별로 사람이름 및 정보를 합칩니다. 구분자는 / 입니다.
                        data += exceldata[row + i, column].Trim() + "/";
                    }
                    //합친 사람내용을 가공하는 로직.
                    data = Filter(data);

                   
                    string dr_nm = "";
                    //해당 글귀가 나왓다면 검토를 해야한다.
                    string gubun = "정상적이지 않은 에러입니다 검토부탁드립니다.";

                    //가공한 사람내용이 없을시 구분0
                    if (string.IsNullOrEmpty(data))
                    {
                        gubun = "0";
                        //아래랑 같다
                        result.Add(string.Format("INSERT INTO {0}(dr_nm, day_time, week, op_no, gubun) VALUES('{1}', '{2}', '{3}','{4}','{5}');", tablename, dr_nm, count ? "AM" : "PM", day, "Rm " + (column-1<10 ? column-1:column).ToString(), gubun));
                        continue;
                    }
                    string[] splitData = data.Split('/');

                    for (int length = 0; length < splitData.Length; length++)
                    {

                        dr_nm = splitData[length];

                        //특수한 상황 주입
                        if (Cnames.Any(name => splitData[length].Contains(name)))
                        {
                            gubun = "C"; // 이름이 포함된 경우
                        }else if(Bnames.Any(name => splitData[length].Contains(name)))
                        {
                            gubun = "B"; // 이름이 포함된 경우
                        }
                        else if (Anames.Any(name => splitData[length].Contains(name)))
                        {

                            gubun = "A"; // 이름이 포함된 경우
                        }
                        else if (Zero_names.Any(name => splitData[length].Contains(name)))
                        {

                            gubun = "0"; // 이름이 포함된 경우
                        }
                        else if (ContainsEnglish(splitData[length]))
                        {
                            gubun = "0";
                         }
                        // 일반상황 (나눈이유는 특수한상황에&&을 같이해버리면 특수한것도 일반이랑 겹쳐버린다
                        else if (splitData.Length == 1)
                        {
                            gubun = "0";
                        }
                        else if(splitData.Length == 2)
                        {
                            gubun = "A";
                        }
                        else if (splitData.Length == 3)
                        {
                            gubun = length == 0 ? "C" : "B";
                        }

                        string query = string.Format("INSERT INTO {0}(dr_nm, day_time, week, op_no, gubun) VALUES('{1}', '{2}', '{3}','{4}','{5}');", tablename, dr_nm, count ? "AM" : "PM", day, "Rm " + (column - 1 < 10 ? column - 1 : column).ToString(), gubun);
                        result.Add(query);
                    }

                }
                count = !count; 
                

            }
            
            return result;
        }
             /**
              * 해당로직은 필터로직으로 이름의 구분을 하는 필터이다 사용자가 원하는 내용대로 커스텀을 해주었다.
              *
              */
         static string Filter(string input)
        {
            // 1단계: 숫자 다음이나 ')' 다음의 '주' 삭제
            string result = Regex.Replace(input, @"(\d|[)])주+", "$1");

            // 2단계 영어와 한글 사이의 '/' 삭제
            result = Regex.Replace(result, @"([a-zA-Z]+)/([가-힣])", "$1$2");

            // 3단계: 혼자 있는 '주' 삭제
            result = Regex.Replace(result, @"\b주\b", "");

            // 4단계: 문자 및 특수문자 '/'를 제외한 나머지 삭제
            result = Regex.Replace(result, @"[^가-힣a-zA-Z/]", ""); // 영어, 한글과 '/' 제외한 모든 문자 삭제
      
            // /삭제하는 로직이 뒤에서 이뤄야지 깔끔하기때문에 
            // 2-1단계: '//'를 '/'로 변경
            result = Regex.Replace(result, @"/{2,}", "/");

            // 2-2단계: 앞뒤의 '/' 제거
            result = result.Trim('/');

            return result;
        }


        //영어가 포함이 되어잇을경우 open, .. 분류해서 0처리해달라했다.
        static bool ContainsEnglish(string str)
        {
            // 정규 표현식을 사용하여 영어가 포함되어 있는지 확인
            return Regex.IsMatch(str, "[a-zA-Z]");
        }

        //쿼리를 띄워주는 화면을 생성 
        //전역리스트로 값을 불러온다.
        private void BtnQureyExcel_Click(object sender, RoutedEventArgs e)
        {
            QueryBox.Visibility = Visibility.Visible;
            LoadBox.Visibility = Visibility.Hidden;
            QueryBlock.Text = string.Join(Environment.NewLine, Querylist);
        }

        //가리기(껐으니까)
        private void BoxClose_Click(object sender, RoutedEventArgs e)
        {

            QueryBox.Visibility = Visibility.Hidden;
            LoadBox.Visibility = Visibility.Visible;
        }
    }
}

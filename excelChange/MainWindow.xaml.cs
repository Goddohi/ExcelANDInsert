﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Data;
using System.IO;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text.RegularExpressions;
using excelChange.View;
using excelChange.Core;

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
        List<string> Cnames = new List<string> {"d"};
        List<string> Bnames = new List<string> {  };
        List<string> Anames = new List<string> {  };
        List<string> Zero_names = new List<string> {"open","OPEN"};


        /// <summary>
        /// Only_Xnames에 있는 이름은 아예동일한경우 무조건 해당값(X)에 걸러준다
        /// 아직은 미사용
        /// </summary>
        List<string> Only_Cnames = new List<string> { };
        List<string> Only_Bnames = new List<string> { };
        List<string> Only_Anames = new List<string> { };
        List<string> Only_Zero_names = new List<string> { };

        //저장할 테이블이름
        string tablename = "d";

        //Setting 창닫는이벤트 인식용
        private readonly EventAggregator _eventAggregator = new EventAggregator();
        //추후 재 로직에 필요할 데이터
        string saveFileName = "";




        /* ********                ********   *
                  *                    *      *
           ********                  *   *    *
           *                        *     *   *
           ******** 
              *                       ******    
              *                           *
          ***********                   *            */


        /************************************************************************************
        * 함  수  명      : btnLoadExcel_Click
        * 내      용      : 엑셀불러오기 버튼을 눌렀을 경우
        * 작  성  자      : 최경태
        * 설      명      : 파일 열기 대화상자 생성하여 엑셀을 불러오고 데이터 가공을 시작한다.
        ************************************************************************************/
        private void btnLoadExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenExcelFile();
        }

     
        /************************************************************************************
        * 함  수  명      : OpenExcelFile
        * 내      용      : 엑셀을 선택하는 함수
        * 작  성  자      : 최경태
        * 설      명      : 엑셀 선택하여 데이터를 불러온다.
        *                   분리한 이유는 매번 새로운 인스턴를 생성하기 위해서 (이프로그램은 특정한 규격의 엑셀만 돌아가게 설계)   
        ************************************************************************************/
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
                saveFileName = openFileDialog.FileName;
                LoadExcelFile(openFileDialog.FileName,true);
                btnQureyExcel.Visibility = Visibility.Visible;

            }
        }


        /************************************************************************************
        * 함  수  명      : LoadExcelFile
        * 내      용      : 엑셀을 데이터를 그리드로 불러오는 함수 
        * 작  성  자      : 최경태
        * 설      명      : 엑셀을 (dataGrid_YN가 true일경우)데이터 그리드에 대입하고  
        *                      상관없이 그후  데이터를 String 2차배열으로 변환한다.
        ************************************************************************************/
        private void LoadExcelFile(string filePath,bool dataGrid_YN)
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
                    if (dataGrid_YN)
                    {
                        dataGrid.ItemsSource = dataTable.DefaultView;
                    }
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
                /* //테스트 코드 
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

        /************************************************************************************
        * 함  수  명      : excelToInsert
        * 내      용      : 엑셀 데이터를 string 2차배열[↓,→]로 받아  Insert로 변환하는 함수 (커스텀)
        * 작  성  자      : 최경태
        * 설      명      : 엑셀데이터를 받아 List로 변환하는 과정에서 사용자가 복사 붙여넣기 하기 편하게 Insert쿼리를 자동으로 제작해주는 로직으로
        *                   각종 사용자가 만든 예외상황으로 인해 예상치 못한 계산이 나올경우 쿼리에 주석으로 경고를 남겨놓도록 설계
        ************************************************************************************/
  
        public List<string> excelToInsert(string[,] exceldata) {
            //  커스텀해드림   2(m)  4*2*5  /,2(r1) r15
            bool count = false;
            List<string> result = new List<string>();
            for (int row=2;row < 2+40; row+=4)
            {
                int day = (row / 8) +1;

                if (count) { result.Add(Environment.NewLine + "/*======    " + day.ToString() + "AM" + "    ======"); }
                else { result.Add(Environment.NewLine+"/*======   " + day.ToString() + "PM" + "   ======*/"); }
            
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
                    int total = 0;

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
                        else if (OnlyEnglish(splitData[length]) || Zero_names.Any(name => splitData[length].Contains(name)))
                        {
                            //open인경우는 공란이다.
                            if (string.Equals("open", dr_nm.ToLower())) {
                                dr_nm = "";
                            }
                            gubun = "0"; // 이름이 포함된 경우
                        }
                   
                        // 일반상황 (나눈이유는 특수한상황에 ||을 같이해버리면 특수한것도 일반이랑 겹쳐버린다
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
                        switch (gubun)
                        {
                            case "0":
                                total += 100;
                                break;
                            case "A":
                                total += 50;
                                break;
                            case "B":
                                total += 30;
                                break;
                            case "C":
                                total += 40;
                                break;

                        }
                        string query = string.Format("INSERT INTO {0}(dr_nm, day_time, week, op_no, gubun) VALUES('{1}', '{2}', '{3}','{4}','{5}');", tablename, dr_nm, count ? "AM" : "PM", day, "Rm " + (column - 1 < 10 ? column - 1 : column).ToString(), gubun);
                        result.Add(query);
                    }
                    if (!(total == 100))
                    {
                        result.Add(string.Format("/* 상단의 RM {0}의 총합이 0.1이아닙니다 */", (column - 1 < 10 ? column - 1 : column).ToString()));
                    }
                }
                count = !count; 
                

            }
            
            return result;
        }

        /************************************************************************************
        * 함  수  명      : Filter
        * 내      용      : 커스텀 된 필터 로직
        * 작  성  자      : 최경태
        * 설      명      : 사용자가 원하는 내용을 담아 커스텀을 한 필터 로직이다. 해당 필터 로직으로 이름을 구분을 한다.
        ************************************************************************************/
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

        /************************************************************************************
        * 함  수  명      : OnlyEnglish
        * 내      용      : 영어만 있는 경우에 True반환
        * 작  성  자      : 최경태
        * 설      명      : 입력받은 string에서 영어만 있는 경우에 True반환한다.
        ************************************************************************************/
        static bool OnlyEnglish(string str)
        {
            // 정규 표현식을 사용하여 영어만.
            return Regex.IsMatch(str, @"^[a-zA-Z]+$");
        }

        /************************************************************************************
        * 함  수  명      : BtnQureyExcel_Click
        * 내      용      : 쿼리변환 버튼을 눌렀을때 나타나는 함수
        * 작  성  자      : 최경태
        * 설      명      : 쿼리을 보이게 하고 엑셀로드창은 숨기며 쿼리블럭에 현재 필터링 되어있는 값을 주입하여 사용자에게 보여준다.
        ************************************************************************************/
        private void BtnQureyExcel_Click(object sender, RoutedEventArgs e)
        {
            QueryBox.Visibility = Visibility.Visible;
            LoadBox.Visibility = Visibility.Hidden;
            QueryBlock.Text = string.Join(Environment.NewLine, Querylist);
        }

        /************************************************************************************
        * 함  수  명      : BoxClose_Click
        * 내      용      : 쿼리 블럭을 종료 했을경우
        * 작  성  자      : 최경태
        * 설      명      : 쿼리를 안보이게 하고 엑셀로드창을 다시 보이게 변환
        ************************************************************************************/
        private void BoxClose_Click(object sender, RoutedEventArgs e)
        {

            QueryBox.Visibility = Visibility.Hidden;
            LoadBox.Visibility = Visibility.Visible;
        }

        /************************************************************************************
        * 함  수  명      : Btn_Setting_Click
        * 내      용      : 설정창을 눌렀을 경우
        * 작  성  자      : 최경태
        * 설      명      : 설정창을 키고 설정창이 종료됨을 알 수 있도록 이벤트를 구독
        ************************************************************************************/
        private void Btn_Setting_Click(object sender, RoutedEventArgs e)
        {
            // 현재 창을 'this'로 참조하여 SettingView의 Owner 속성에 설정
            Setting setting = new Setting(_eventAggregator);
            // 이벤트 구독
            setting.SettingClosed += OnSettingClosed;
            setting.Owner = this;  // 현재 창을 소유자로 설정

            setting.ShowDialog();
        }



        /************************************************************************************
        * 함  수  명      : OnSettingClosed
        * 내      용      : 설정창이 종료 되었을 경우
        * 작  성  자      : 최경태
        * 설      명      : 자식 창이 닫힐 때 호출될 이벤트 핸들러를 이용하여 설정창이 종료 되었을 경우 적용 되도록 변경 
        ************************************************************************************/
        private void OnSettingClosed(object sender, EventArgs e)
        {
            MessageBox.Show("테스트");
            //데이터그리드에는 반영하지 않는다.
            if (!string.IsNullOrEmpty(saveFileName))
            {
                LoadExcelFile(saveFileName, false);
            }
        }
    }
}

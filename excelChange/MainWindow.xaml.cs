using System;
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
using System.Windows.Threading;
using System.Windows.Input;
using excelChange.Entity;

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
            LoadTableData();
            LoadXML_TypeName();
        }

        /// <summary>
        /// 전역변수
        /// </summary>

        // Quertlist는 만든 쿼리를 어디서든 부를수있게
        List<string> Querylist;

        /// <summary>
        /// Xnames에 있는 이름은 겹치는 무조건 해당값(X)에 걸러준다
        /// </summary>
        List<string> Cnames = new List<string> {};
        List<string> Bnames = new List<string> {  };
        List<string> Anames = new List<string> {  };
        List<string> Zero_names = new List<string> {};


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
        // 더블 클릭을 감지하기 위한 시간 (밀리초)
        private const int DoubleClickThreshold = 500;
        // 더블클릭 메소드 
        private int _clickCount = 0;
        //Setting 창닫는이벤트 인식용
        private readonly EventAggregator _eventAggregator = new EventAggregator();

        //XML로더
        XmlLoad xmlLoad = XmlLoad.make();
        //추후 재 로직에 필요할 데이터
        string saveFileName = "";





        /************************************************************************************
        * 함  수  명      : btnLoadExcel_Click
        * 내      용      : 엑셀불러오기 버튼을 눌렀을 경우
        * 설      명      : 파일 열기 대화상자 생성하여 엑셀을 불러오고 데이터 가공을 시작한다.
        ************************************************************************************/
        private void btnLoadExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenExcelFile();
        }

     
        /************************************************************************************
        * 함  수  명      : OpenExcelFile
        * 내      용      : 엑셀을 선택하는 함수
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
        * 함  수  명      : LoadTableData
        * 내      용      : 테이블 이름을 불러옵니다
        * 설      명      : 저장되어있는 테이블 이름을 불러옵니다.
        ************************************************************************************/
        private void LoadTableData()
        {

            try
            {
                TableEntity tableEntity = xmlLoad.LoadObjectFromXml<TableEntity>("Table.xml");
                tablename = tableEntity.Name;

            }
            catch
            {
                MessageBox.Show("테이블NAME을 불러오지 못했습니다.");
            }
        }


        /************************************************************************************
        * 함  수  명      : LoadXML_TypeName
        * 내      용      : 저장된 예외타입들을 불러옵니다.
        * 설      명      : 저장된 예외타입을 불러오는데 해당타입은 설정창을 닫을때와 시작할때 이루어집니다.
        ************************************************************************************/
        private void LoadXML_TypeName()
        {
            try
            {
                // 데이터 로드
                var GetTypeSetting = xmlLoad.GetBasicSetting<TypeSettingEntity>("TypeNames.xml");

                Only_Zero_names.Clear();

                Only_Cnames.Clear();
                Only_Bnames.Clear();
                Only_Anames.Clear();

                Zero_names.Clear();
                Cnames.Clear();
                Bnames.Clear();
                Anames.Clear();
                // 정렬 및 분류
                var TypeSettings = GetTypeSetting
                    .OrderBy(o => o.TYPE)
                    .ThenBy(o => o.CONTAIN)
                    .ThenBy(o => o.NAME);

                foreach (var item in TypeSettings)
                {
                    if (item.TYPE == "0" && item.CONTAIN == "N")
                    {
                        if (!Only_Zero_names.Contains(item.NAME))
                            Only_Zero_names.Add(item.NAME);
                    }
                    else if (item.TYPE == "0")
                    {
                        if (!Zero_names.Contains(item.NAME))
                            Zero_names.Add(item.NAME);
                    }
                    else if (item.TYPE == "C" && item.CONTAIN == "N")
                    {
                        if (!Only_Cnames.Contains(item.NAME))
                            Only_Cnames.Add(item.NAME);
                    }
                    else if (item.TYPE == "C")
                    {
                        if (!Cnames.Contains(item.NAME))
                             Cnames.Add(item.NAME);
                    }
                    else if (item.TYPE == "B" && item.CONTAIN == "N")
                    {
                        if (!Only_Bnames.Contains(item.NAME))
                            Only_Bnames.Add(item.NAME);
                    }
                    else if (item.TYPE == "B")
                    {
                        if (!Bnames.Contains(item.NAME))
                            Bnames.Add(item.NAME);
                    }
                    else if (item.TYPE == "A" && item.CONTAIN == "N")
                    {
                        if (!Only_Anames.Contains(item.NAME))
                            Only_Anames.Add(item.NAME);
                    }
                    else if (item.TYPE == "A")
                    {
                        if (!Anames.Contains(item.NAME))
                            Anames.Add(item.NAME);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"오류 발생: {ex.Message}");
            }
        }




        /************************************************************************************
        * 함  수  명      : LoadExcelFile
        * 내      용      : 엑셀을 데이터를 그리드로 불러오는 함수 
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
        * 설      명      : 엑셀데이터를 받아 List로 변환하는 과정에서 사용자가 복사 붙여넣기 하기 편하게 Insert쿼리를 자동으로 제작해주는 로직으로
        *                   각종 사용자가 만든 예외상황으로 인해 예상치 못한 계산이 나올경우 쿼리에 주석으로 경고를 남겨놓도록 설계
        ************************************************************************************/
  
        public List<string> excelToInsert(string[,] exceldata) {
            //  커스텀해드림   2(m)  4*2*5  /,2(r1) r15
            bool count = true;
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
                        if (Cnames.Any(name => splitData[length].Contains(name)) || Only_Cnames.Any(name => splitData[length].Equals(name)))
                        {
                            gubun = "C"; // 이름이 포함된 경우
                            
                        }else if(Bnames.Any(name => splitData[length].Contains(name)) || Only_Bnames.Any(name => splitData[length].Equals(name)))
                        {
                            gubun = "B"; // 이름이 포함된 경우
                        }
                        else if (Anames.Any(name => splitData[length].Contains(name)) || Only_Anames.Any(name => splitData[length].Equals(name)))
                        {

                            gubun = "A"; // 이름이 포함된 경우
                        }
                        else if (OnlyEnglish(splitData[length]) || Zero_names.Any(name => splitData[length].Contains(name)) || Only_Zero_names.Any(name => splitData[length].Equals(name)))
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
        * 설      명      : 쿼리을 새창으로 띄워준다. 필터링 되어있는 값을 주입하여 사용자에게 보여준다.
        * 수정  내용      : Codewindow를 제작하였기 때문이 이제는 로드창을 숨기지 않고 새창을 연다
        ************************************************************************************/
        private void BtnQureyExcel_Click(object sender, RoutedEventArgs e)
        {
          //  QueryBox.Visibility = Visibility.Visible;
          //  LoadBox.Visibility = Visibility.Hidden;
          //  QueryBlock.Text = string.Join(Environment.NewLine, Querylist);
            OpenCodeWIndow(string.Join(Environment.NewLine, Querylist));
        }



        /************************************************************************************
        * 함  수  명      : Btn_Setting_Click
        * 내      용      : 설정창을 눌렀을 경우
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
        * 설      명      : 자식 창이 닫힐 때 호출될 이벤트 핸들러를 이용하여 설정창이 종료 되었을 경우 적용 되도록 변경 
        ************************************************************************************/
        private void OnSettingClosed(object sender, EventArgs e)
        {
            LoadTableData();
            LoadXML_TypeName();
            if (!string.IsNullOrEmpty(saveFileName))
            {
                LoadExcelFile(saveFileName, false);
            }

        }


        /************************************************************************************
        * 함  수  명      : OpenCodeWIndow
        * 내      용      : 코드 새창열기
        * 설      명      : 코드를 받고 새창을 열어서 code창에 코드를 입력한 상태로 띄워준다.
        ************************************************************************************/
        protected void OpenCodeWIndow(string code)
        {
            using (CodeWindow window = new CodeWindow())
            {
                window.Owner = this.Parent as Window;
                window.Show();
                window.TxtCode.Text = code;
            }
        }



       


        /// <summary>
        /// 기본적인 환경 제작
        /// </summary>
        /************************************************************************************
        * 함  수  명      : MinimizeButton_Click
        * 내      용      : 최소화 버튼
        * 설      명      : 최소화 로직을 실행 시킨다.
        ************************************************************************************/
        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        /************************************************************************************
        * 함  수  명      : Maximize_window
        * 내      용      : 최대화
        * 설      명      : 최대화를 실행시킨다. 이미 최대일경우 일반상태로 변경시킨다
        ************************************************************************************/
        private void Maximize_window()
        {
            if (this.WindowState == WindowState.Maximized)
            {
                this.WindowState = WindowState.Normal;
            }
            else
            {
                this.WindowState = WindowState.Maximized;
            }
        }


        /************************************************************************************
        * 함  수  명      : DragBlock_MouseDown
        * 내      용      : 드래그 박스 더블인식
        * 설      명      : 드래그 박스에 더블인식을 추가하여 더블인식시 최대화 함수 실행
        ************************************************************************************/
        private void DragBlock_MouseDown(object sender, MouseButtonEventArgs e)
        {
            _clickCount++;
            if (_clickCount == 1)
            {
                // 첫 번째 클릭 후 타이머 시작
                DispatcherTimer timer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(DoubleClickThreshold)
                };
                timer.Tick += (s, args) =>
                {
                    timer.Stop();
                    _clickCount = 0; // 타이머가 만료되면 클릭 카운트 리셋
                };
                timer.Start();
            }
            else if (_clickCount == 2)
            {
                _clickCount = 0;
                Maximize_window();
            }
        }

        /************************************************************************************
        * 함  수  명      : CloseButton_Click
        * 내      용      : 프로그램종료
        * 설      명      : 잔여 프로세스가 남지 않도록 종료시킨다.
        ************************************************************************************/
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            // 어플리케이션을 종료
            Application.Current.Shutdown();
        }

        /************************************************************************************
        * 함  수  명      : MaximizeButton_Click
        * 내      용      : 최대화 버튼 클릭
        * 설      명      : 최대화 함수을 실행시킨다.
        ************************************************************************************/
        //최대화
        private void MaximizeButton_Click(object sender, RoutedEventArgs e)
        {
            Maximize_window();
        }
    }
}

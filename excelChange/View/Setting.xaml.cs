using excelChange.Core;
using excelChange.Entity;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace excelChange.View
{
    /// <summary>
    /// Setting.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Setting : Window
    {
        private ObservableCollection<TypeSettingEntity> ocTypeSetting;


        public ObservableCollection<TypeSettingEntity> OcTypeSetting { get => ocTypeSetting; set => ocTypeSetting = value; }
        XmlLoad xmlLoad = XmlLoad.make();
        private TableEntity tableEntity;
        private int _clickCount = 0;
        private const int DoubleClickThreshold = 500; // 더블 클릭을 감지하기 위한 시간 (밀리초)
        public event EventHandler SettingClosed;
        public List<string> TypeList { get; set; } = new List<string> { "0", "A", "B", "C" };
        public List<string> ContainList { get; set; } = new List<string> { "Y", "N"};


        /************************************************************************************
        * 함  수  명      : Setting
        * 내      용      : 기본생성자
        * 설      명      : 
        ************************************************************************************/
        public Setting()
        {
            InitializeComponent();
            ocTypeSetting = new ObservableCollection<TypeSettingEntity>();
            DataContext = this; // 현재 인스턴스를 DataContext로 설정
            this.Closed += Setting_Closed;
            LoadXML_TypeName();
            LoadTableData();
        }

        /************************************************************************************
        * 함  수  명      : Setting
        * 내      용      : 생성자 
        * 설      명      : 
        ************************************************************************************/
        public Setting(EventAggregator eventAggregator)
        {
            InitializeComponent();
            //닫혔을때 인식하는 이벤트 구독
            ocTypeSetting = new ObservableCollection<TypeSettingEntity>();
            DataContext = this; // 현재 인스턴스를 DataContext로 설정
            this.Closed += Setting_Closed;
            LoadXML_TypeName();
            LoadTableData();
            foreach (var column in DgdTypeName.Columns)
            {
                if (column is DataGridComboBoxColumn comboBoxColumn)
                {
                    if (comboBoxColumn.Header.ToString() == "Type")
                    {
                        comboBoxColumn.ElementStyle = new Style(typeof(ComboBox));
                        comboBoxColumn.EditingElementStyle = new Style(typeof(ComboBox));

                        // 드롭다운의 ItemsSource를 설정합니다.
                        comboBoxColumn.ElementStyle.Setters.Add(new Setter(ComboBox.ItemsSourceProperty, TypeList));
                        comboBoxColumn.EditingElementStyle.Setters.Add(new Setter(ComboBox.ItemsSourceProperty, TypeList));
                    }
                    if (comboBoxColumn.Header.ToString() == "Contain")
                    {
                        comboBoxColumn.ElementStyle = new Style(typeof(ComboBox));
                        comboBoxColumn.EditingElementStyle = new Style(typeof(ComboBox));

                        // 드롭다운의 ItemsSource를 설정합니다.
                        comboBoxColumn.ElementStyle.Setters.Add(new Setter(ComboBox.ItemsSourceProperty, ContainList));
                        comboBoxColumn.EditingElementStyle.Setters.Add(new Setter(ComboBox.ItemsSourceProperty, ContainList));
                    }
                }
            }
        }





        /************************************************************************************
        * 함  수  명      : LoadXML_TypeName
        * 내      용      : Type(DataGrid) 저장 버튼
        * 설      명      : 저장함수를 불러온다
        ************************************************************************************/
        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            this.SaveFavQueryInfo();

        }


        /************************************************************************************
         * 함  수  명      : LoadXML_TypeName
         * 내      용      : Type(DataGrid) 저장하는 함수
         * 설      명      : 중복이름이 있을경우 저장하지 않는다
         ************************************************************************************/
        private void SaveFavQueryInfo()
        {
            var is_dup = this.OcTypeSetting.GroupBy(x => new { x.NAME }).All(g => g.Count() > 1);

            if (is_dup)
            {
                MessageBox.Show("중복되는 이름이 있습니다.");
                return;
            }
            bool save_yn=xmlLoad.SaveBasicSetting(OcTypeSetting, "TypeNames.xml");

            if (save_yn)
            {
                MessageBox.Show("저장완료");
            }
            else
            {
                MessageBox.Show("저장실패");
            }
        }



        /************************************************************************************
        * 함  수  명      : Btn_Add_Click
        * 내      용      : 추가 버튼클릭
        * 설      명      : Type DataGrid row추가
        ************************************************************************************/
        private void Btn_Add_Click(object sender, RoutedEventArgs e)
        {
                var new_it = new TypeSettingEntity();
                new_it.NAME = "새로운사용자";
                new_it.TYPE = "0";
                new_it.CONTAIN = "Y";
                new_it.REMARK = "설명";

                this.OcTypeSetting.Add(new_it);
            
            DgdTypeName.ItemsSource = OcTypeSetting;
        }

        /************************************************************************************
        * 함  수  명      : Btn_Delete_Click
        * 내      용      : 삭제 버튼클릭
        * 설      명      : Type DataGrid row삭제
        ************************************************************************************/
        private void Btn_Delete_Click(object sender, RoutedEventArgs e)
        {
            var item = DgdTypeName.SelectedItem as TypeSettingEntity;
            if (item == null) return;

            this.OcTypeSetting.Remove(item);
            DgdTypeName.ItemsSource = OcTypeSetting;

        }


        /************************************************************************************
        * 함  수  명      : LoadXML_TypeName
        * 내      용      : TypeSetting(예외상황)을 XML에서 불러오는 함수
        * 설      명      : TypeNames에서 데이터를 불러와서 정렬을 하고 데이터 그리드에 바인딩 해준다.
        ************************************************************************************/
        private void LoadXML_TypeName()
        {
            try
            {
                // 데이터 로드
                
                var GetTypeSetting = xmlLoad.GetBasicSetting<TypeSettingEntity>("TypeNames.xml");

                OcTypeSetting.Clear();

                /*테스트코드
                    var settingsList = GetTypeSetting.ToList(); 
                     string result = string.Join(", ", settingsList.Select(s => s.NAME));
                     // MessageBox에 결과 표시
                     MessageBox.Show(result);
               */
                var TypeSettings = GetTypeSetting
                    .OrderBy(o => o.TYPE)
                    .ThenBy(o => o.CONTAIN)
                    .ThenBy(o => o.NAME);
                foreach (var TypeSetting in TypeSettings)
                {
                    OcTypeSetting.Add(TypeSetting);
                }
                DgdTypeName.ItemsSource = OcTypeSetting;
            }
            catch (Exception ex)
            {
                //MessageWindow.Instance.ShowMessage($"오류 발생: {ex.Message}");
                MessageBox.Show($"오류 발생: {ex.Message}");
            }
        }


        /************************************************************************************
        * 함  수  명      : Btn_Reset_Click
        * 내      용      : 새로고침함수
        * 설      명      : 데이터 그리드를 다시 새롭게 저장하기전 상태로 불러온다.
        ************************************************************************************/
        private void Btn_Reset_Click(object sender, RoutedEventArgs e)
        {
            LoadXML_TypeName();
        }

        /************************************************************************************
        * 함  수  명      : LoadTableData
        * 내      용      : 테이블명을 불러온다
        * 설      명      : 테이블명을 불러와서 tableEntity에 저장한다.
        ************************************************************************************/
        private void LoadTableData()
        {
            try
            {
                tableEntity = xmlLoad.LoadObjectFromXml<TableEntity>("Table.xml");
                txtName.Text = tableEntity.Name; // 텍스트박스에 이름 설정
            }
            catch
            {
                tableEntity = new TableEntity(); // 파일이 없으면 새로운 객체 생성
            }
        }

        /************************************************************************************
        * 함  수  명      : SaveButton_Click
        * 내      용      : 테이블명을 저장해주는 함수
        * 설      명      : 테이블 명을 저장해준다
        ************************************************************************************/
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            tableEntity.Name = txtName.Text; // 텍스트박스에서 수정된 이름 가져오기
            xmlLoad.SaveObjectToXml<TableEntity>(tableEntity, "Table.xml"); // XML 파일로 저장
            MessageBox.Show("테이블명이"+ tableEntity.Name+"로 저장되었습니다");
        }



        /************************************************************************************
        * 함  수  명      : Hyperlink_RequestNavigate
        * 내      용      : 하이퍼링크를 기본 브라우저에서 열게해주는 함수
        * 설      명      : 하이퍼링크를 누르게 될경우 기본 브라우저로 향하게 한다.
        ************************************************************************************/


        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            // 기본 브라우저에서 링크를 엽니다.
            Process.Start(new ProcessStartInfo
            {
                FileName = e.Uri.ToString(),
                UseShellExecute = true
            });

            // 이벤트가 처리되었음을 표시합니다.
            e.Handled = true;
        }




        /************************************************************************************
       * 함  수  명      : Btn_maker_Click
       * 내      용      : 제작자 버튼을 눌렀을 경우 함수
       * 설      명      : 제작자에 대한 설명 창을 띄워준다.
       ************************************************************************************/
        private void Btn_maker_Click(object sender, RoutedEventArgs e)
        {
            InformationGrid.Visibility = Visibility.Hidden;
            makerGrid.Visibility = Visibility.Visible;
            mainSettingGrid.Visibility = Visibility.Hidden;

        }


        /************************************************************************************
        * 함  수  명      : Btn_makerClose_Click
        * 내      용      : 제작자 종료을 눌렀을 경우 함수
        * 설      명      : 제작자에 대한  종료
        ************************************************************************************/
        private void Btn_makerClose_Click(object sender, RoutedEventArgs e)
        {
            InformationGrid.Visibility = Visibility.Hidden;
            makerGrid.Visibility = Visibility.Hidden;
            mainSettingGrid.Visibility = Visibility.Visible;
        }
        /************************************************************************************
        * 함  수  명      : Btn_Information_Click
        * 내      용      : 도움말을 눌렀을 경우 함수
        * 설      명      : 도움말창을 띄워준다
        ************************************************************************************/
        private void Btn_Information_Click(object sender, RoutedEventArgs e)
        {
            InformationGrid.Visibility = Visibility.Visible;
            makerGrid.Visibility = Visibility.Hidden;
            mainSettingGrid.Visibility = Visibility.Hidden;
            infoText.Text = Environment.NewLine + "Typenames에 대해서 설명드리겠습니다" +Environment.NewLine + 
                            "REMARK뺴고 NOTNULL입니다. " + Environment.NewLine +
                            "NAME: 예외처리할 이름" + Environment.NewLine +
                            "TYPE: 0(숫자 알파벳o아님) ,  A , B , C" + Environment.NewLine +
                            "CONTAIN  :  Y => 겹치기만 해도 OK   N => 무조건 이것만 있어야 가능" + Environment.NewLine +
                            "REMARK   :  왜 이런지 기록 남겨놓는 용도" + Environment.NewLine +
                            "e.g." + Environment.NewLine +
                            "NAME: 김땡땡 " + Environment.NewLine +
                            "TYPE : C" + Environment.NewLine +
                            "CONTAIN : Y" + Environment.NewLine +
                            "REMARK: 이분은 그냥 C로 해달라고 했다." + Environment.NewLine +
                            "=> 김땡땡은 겹치기만해도 C로 변환이 됩니다.";
        }

        /************************************************************************************
        * 함  수  명      : Btn_InformationClose_Click
        * 내      용      : 설명 종료을 눌렀을 경우 함수
        * 설      명      : 설명창 종료
        ************************************************************************************/
        private void Btn_InformationClose_Click(object sender, RoutedEventArgs e)
        {
            InformationGrid.Visibility = Visibility.Hidden;
            makerGrid.Visibility = Visibility.Hidden;
            mainSettingGrid.Visibility = Visibility.Visible;

        }






        /************************************************************************************
        * 함  수  명      : CloseButton_Click
        * 내      용      : 창닫기 버튼을 누른 창을 닫는다
        * 설      명      : Close를 통해 창닫기 버튼을 누른 창을 닫는다
        ************************************************************************************/
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }



        /************************************************************************************
        * 함  수  명      : Maximize_window
        * 내      용      : 최대화
        * 설      명      : 최대화를 실행시킨다. 이미 최대일경우 일반상태로 변경시킨다
        ************************************************************************************/
        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

       /************************************************************************************
        * 함  수  명      : MaximizeButton_Click
        * 내      용      : 최대화 버튼 클릭
        * 설      명      : 최대화 함수을 실행시킨다.
        ************************************************************************************/
        private void MaximizeButton_Click(object sender, RoutedEventArgs e)
        {
            Maximize_window();
        }


        /************************************************************************************
        * 함  수  명      : MaximizeButton_Click
        * 내      용      : 최대화 버튼 클릭
        * 설      명      : 최대화 함수을 실행시킨다.
        ************************************************************************************/
        private void Setting_Closed(object sender, EventArgs e)
        {
            // 창이 닫힐 때 SettingViewClosed 이벤트를 발생시킴
            SettingClosed?.Invoke(this, EventArgs.Empty);
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
        * 함  수  명      : MaximizeButton_Click
        * 내      용      : 최대화 버튼 클릭
        * 설      명      : 최대화 함수을 실행시킨다.
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


    }
}

using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;
using System.Windows.Threading;

namespace excelChange.View
{
    /// <summary>
    /// CodeWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class CodeWindow : Window, IDisposable
    {
        private int _clickCount = 0;
        private const int DoubleClickThreshold = 500; // 더블 클릭을 감지하기 위한 시간 (밀리초)

        public CodeWindow()
        {
            InitializeComponent();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

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
        //최대화
        private void MaximizeButton_Click(object sender, RoutedEventArgs e)
        {
            Maximize_window();
        }

        /************************************************************************************
        * 함  수  명      : CloseButton_Click
        * 내      용      : 창닫기 버튼을 누른 창을 닫는다
        * 설      명      : Close를 통해 창닫기 버튼을 누른 창을 닫는다
        ************************************************************************************/
        // 닫기 메소드
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
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




        /************************************************************************************
        * 함  수  명      : TxtCode_GotFocus
        * 내      용      : TextBox가 포커스를 받을 때 호출
        * 설      명      : TextBox가 포커스를 받을 때, SelectAll() 메서드를 사용하여 TextBox 안의 모든 텍스트를 선택하여
        *                   사용자가 복사 붙여 넣기 편하게 적용
        ************************************************************************************/
        private void TxtCode_GotFocus(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke
            (
                DispatcherPriority.ContextIdle,
                new Action
                (
                    delegate
                    {
                        (sender as TextBox).SelectAll();
                    }
                )
            );
        }

        public void Dispose()
        {

        }
    }
}
